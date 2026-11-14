#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""会员历史导出迁移脚本。

把 exports/ 下三代导出模板的文件归一为一张 members 表，写入 SQLite，
并在仓库根目录生成迁移报告 migration_report.md，以及带“拒”字的拒收留存
文件 members_rejected_拒收.csv。

归一基准为最新一代模板（v3）的字段与口径：
  - 会员等级统一为 mapping.yaml 中定义的等级名称（脚本内不写死任何对应关系，
    运营改完 mapping.yaml 重跑即可生效）；
  - 账户余额统一为元（v2 的分自动除以 100，四舍五入到分）；
  - 注册渠道仅最新模板提供，老模板归一后为 NULL；
  - 同一会员跨代出现时，以最新一代的记录为准，差异写入迁移报告。

映射表覆盖不到、或关键字段无法解析的行不写入 members，原始整行记录进入
members_rejected 表与 members_rejected_拒收.csv 备查，每行附拒收原因。

仅依赖 Python 标准库；PyYAML / openpyxl 不可用时也能运行
（mapping.yaml 用内置的简易解析器，xlsx 用 zipfile + ElementTree 读取）。
"""
from __future__ import annotations

import csv
import io
import json
import re
import sqlite3
import sys
import zipfile
from collections import OrderedDict
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Iterator
from xml.etree import ElementTree as ET

# ---------------------------------------------------------------------------
# 路径与常量
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
EXPORTS_DIR = BASE_DIR / "exports"
MAPPING_PATH = BASE_DIR / "mapping.yaml"
DB_PATH = BASE_DIR / "members.db"
REPORT_PATH = BASE_DIR / "migration_report.md"
REJECT_CSV_PATH = BASE_DIR / "members_rejected_拒收.csv"

MAPPING_SECTION = "等级映射"
CENT = Decimal("100")
Q2 = Decimal("0.01")

# 三代模板口径：unit 为该代余额列的标准单位（yuan=元 / cent=分）
GENERATIONS: dict[str, dict[str, str]] = {
    "v1": {"name": "v1（Excel 老模板）", "unit": "yuan"},
    "v2": {"name": "v2（CSV 中间代）", "unit": "cent"},
    "v3": {"name": "v3（CSV 最新模板）", "unit": "yuan"},
}

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


class MigrationError(Exception):
    """迁移无法继续时抛出，信息直接展示给使用者。"""


# ---------------------------------------------------------------------------
# mapping.yaml 读取（优先 PyYAML，缺失时用内置的简易解析器）
# ---------------------------------------------------------------------------
def _strip_comment(text: str) -> str:
    """去掉不在引号内的 # 注释。"""
    quote = ""
    for i, ch in enumerate(text):
        if quote:
            if ch == quote:
                quote = ""
        elif ch in ("'", '"'):
            quote = ch
        elif ch == "#":
            return text[:i]
    return text


def _split_key_value(text: str) -> tuple[str, str] | None:
    """按第一个不在引号内的冒号拆分 key: value。"""
    quote = ""
    for i, ch in enumerate(text):
        if quote:
            if ch == quote:
                quote = ""
        elif ch in ("'", '"'):
            quote = ch
        elif ch == ":":
            return text[:i], text[i + 1 :]
    return None


def _unquote(text: str) -> str:
    text = text.strip()
    if len(text) >= 2 and text[0] == text[-1] and text[0] in ("'", '"'):
        body = text[1:-1]
        if text[0] == '"':
            return body.replace('\\"', '"').replace("\\\\", "\\")
        return body.replace("''", "'")
    return text


def _parse_simple_yaml(path: Path) -> dict[str, dict[str, str]]:
    """解析 mapping.yaml 这种「一级段名 + 二级 key: value」的结构。"""
    root: dict[str, dict[str, str]] = {}
    section: str | None = None
    for lineno, raw in enumerate(path.read_text(encoding="utf-8-sig").splitlines(), 1):
        line = _strip_comment(raw).rstrip()
        if not line.strip():
            continue
        if "\t" in line[: len(line) - len(line.lstrip())]:
            raise MigrationError(f"mapping.yaml 第 {lineno} 行使用了制表符缩进，请改用空格")
        indent = len(line) - len(line.lstrip(" "))
        content = line.strip()
        pair = _split_key_value(content)
        if pair is None:
            raise MigrationError(f"mapping.yaml 第 {lineno} 行缺少冒号：{content}")
        key, value = pair
        if indent == 0:
            section = _unquote(key.strip())
            root.setdefault(section, {})
        else:
            if section is None:
                raise MigrationError(f"mapping.yaml 第 {lineno} 行缩进内容不属于任何段")
            root[section][_unquote(key.strip())] = _unquote(value.strip())
    return root


def load_level_mapping(path: Path) -> tuple[dict[str, str], list[tuple[str, str]]]:
    """读取等级映射，返回 (查找表, 有序条目列表)。

    查找表的键统一为去空白后的原始取值（如「金牌会员」「3」），
    值为归一后的等级名称。映射内容完全来自 mapping.yaml，代码不写死。
    """
    if not path.exists():
        raise MigrationError(f"找不到等级映射文件：{path}")
    try:
        import yaml  # type: ignore

        loaded = yaml.safe_load(path.read_text(encoding="utf-8-sig"))
        if not isinstance(loaded, dict):
            raise MigrationError("mapping.yaml 顶层必须是映射结构")
        parsed = loaded
    except ImportError:
        parsed = _parse_simple_yaml(path)

    if MAPPING_SECTION not in parsed:
        raise MigrationError(f"mapping.yaml 缺少顶层段「{MAPPING_SECTION}」")
    section = parsed[MAPPING_SECTION]
    if not isinstance(section, dict) or not section:
        raise MigrationError(f"mapping.yaml 的「{MAPPING_SECTION}」为空或格式不正确")

    mapping: dict[str, str] = {}
    entries: list[tuple[str, str]] = []
    for raw_key, raw_value in section.items():
        key = str(raw_key).strip()
        value = "" if raw_value is None else str(raw_value).strip()
        if key == "":
            raise MigrationError("mapping.yaml 中存在空的等级取值")
        if value == "":
            raise MigrationError(f"mapping.yaml 中等级取值 {key!r} 没有对应的等级名称")
        mapping[key] = value
        entries.append((key, value))
    return mapping, entries


# ---------------------------------------------------------------------------
# 导出文件读取（xlsx / csv，均只依赖标准库）
# ---------------------------------------------------------------------------
def _col_index(ref: str) -> int:
    """把单元格引用的列字母（A、B、AA…）转成从 0 开始的列序号。"""
    letters = re.match(r"[A-Za-z]+", ref).group().upper()
    index = 0
    for ch in letters:
        index = index * 26 + (ord(ch) - ord("A") + 1)
    return index - 1


def read_xlsx_rows(path: Path) -> Iterator[tuple[int, list[str]]]:
    """读取 xlsx 第一个工作表，返回 (行号, 单元格文本列表)，行号从 1 开始。"""
    try:
        zf = zipfile.ZipFile(path)
    except zipfile.BadZipFile as exc:
        raise MigrationError(f"{path.name} 不是有效的 xlsx 文件：{exc}") from exc
    with zf as z:
        names = z.namelist()

        # 共享字符串（老文件可能没有，本批样例全部用 inlineStr）
        shared: list[str] = []
        if "xl/sharedStrings.xml" in names:
            sroot = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in sroot.findall(f"{NS}si"):
                shared.append("".join(t.text or "" for t in si.iter(f"{NS}t")))

        # 找到第一个工作表的实际路径（兼容工作簿 rels 指向的文件名）
        sheet_path = "xl/worksheets/sheet1.xml"
        if "xl/workbook.xml" in names and "xl/_rels/workbook.xml.rels" in names:
            rels_root = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
            rel_map = {
                rel.get("Id"): rel.get("Target")
                for rel in rels_root.iter()
                if rel.tag.endswith("Relationship")
            }
            wb = ET.fromstring(z.read("xl/workbook.xml"))
            first_sheet = wb.find(f"{NS}sheets/{NS}sheet")
            if first_sheet is not None:
                rid = first_sheet.get(f"{REL_NS}id")
                target = rel_map.get(rid)
                if target:
                    candidate = target.lstrip("/")
                    if not candidate.startswith("xl/"):
                        candidate = "xl/" + candidate
                    if candidate in names:
                        sheet_path = candidate

        root = ET.fromstring(z.read(sheet_path))
        for row in root.iter(f"{NS}row"):
            rownum = int(row.get("r", "0"))
            cells: dict[int, str] = {}
            width = 0
            for cell in row.findall(f"{NS}c"):
                ref = cell.get("r", "A")
                col = _col_index(ref)
                width = max(width, col + 1)
                ctype = cell.get("t")
                if ctype == "inlineStr":
                    node = cell.find(f"{NS}is")
                    cells[col] = "".join(
                        t.text or "" for t in node.iter(f"{NS}t")
                    ) if node is not None else ""
                elif ctype == "s":
                    v = cell.find(f"{NS}v")
                    cells[col] = shared[int(v.text)] if v is not None and v.text else ""
                elif ctype == "b":
                    v = cell.find(f"{NS}v")
                    cells[col] = "TRUE" if v is not None and v.text == "1" else "FALSE"
                else:  # 数值 / 常规
                    v = cell.find(f"{NS}v")
                    cells[col] = v.text if v is not None and v.text is not None else ""
            yield rownum, [cells.get(i, "") for i in range(width)]


def read_csv_rows(path: Path) -> Iterator[tuple[int, list[str]]]:
    """读取 csv，自动识别 UTF-8 / GB18030，返回 (行号, 字段列表)。"""
    raw = path.read_bytes()
    try:
        text = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        try:
            text = raw.decode("gb18030")
        except UnicodeDecodeError as exc:
            raise MigrationError(
                f"{path.name} 编码无法识别（支持 UTF-8 / GB18030）"
            ) from exc
    reader = csv.reader(io.StringIO(text, newline=""))
    for lineno, row in enumerate(reader, 1):
        if not row or all(cell.strip() == "" for cell in row):
            continue
        yield lineno, row


def read_table(path: Path) -> Iterator[tuple[int, list[str]]]:
    suffix = path.suffix.lower()
    if suffix in (".xlsx", ".xlsm"):
        return read_xlsx_rows(path)
    if suffix == ".csv":
        return read_csv_rows(path)
    raise MigrationError(f"不支持的文件类型：{path.name}")


# ---------------------------------------------------------------------------
# 表头与值解析
# ---------------------------------------------------------------------------
def normalize_header(text: str) -> str:
    return (
        text.replace("﻿", "")
        .strip()
        .replace("（", "(")
        .replace("）", ")")
        .replace(" ", "")
    )


def find_column(
    headers: list[str], candidates: list[str], required: bool = True
) -> int | None:
    """按归一化后的表头模糊匹配列序号；candidates 为多个可接受的列名。"""
    normalized = [normalize_header(h) for h in headers]
    wanted = [normalize_header(c) for c in candidates]
    for i, header in enumerate(normalized):
        if any(w in header for w in wanted):
            return i
    if required:
        raise MigrationError(
            f"表头缺少列：{' / '.join(candidates)}（实际表头：{headers}）"
        )
    return None


def parse_amount(raw: str) -> Decimal:
    """把金额文本解析为 Decimal；无法解析时抛 ValueError。"""
    cleaned = raw.replace(",", "").replace("，", "").strip()
    if cleaned == "":
        raise ValueError("金额为空")
    try:
        return Decimal(cleaned)
    except InvalidOperation as exc:
        raise ValueError(f"金额无法解析：{raw!r}") from exc


def excel_serial_to_date(serial: int) -> str:
    # Excel 1900 日期系统：1 对应 1900-01-01，兼容其虚构的 1900-02-29
    base = date(1899, 12, 31)
    return (base + timedelta(days=serial)).strftime("%Y-%m-%d")


def parse_date(raw: str) -> str:
    """统一解析为 YYYY-MM-DD；支持 20240101、2024-01-01、2024/1/1 与 Excel 序列值。"""
    text = raw.strip()
    if text == "":
        raise ValueError("注册日期为空")
    if re.fullmatch(r"\d{8}", text):
        return datetime.strptime(text, "%Y%m%d").strftime("%Y-%m-%d")
    for fmt in ("%Y-%m-%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(text, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    if re.fullmatch(r"\d+(\.\d+)?", text):
        serial = int(Decimal(text))
        if 1 <= serial <= 2958465:
            return excel_serial_to_date(serial)
    raise ValueError(f"注册日期无法解析：{raw!r}")


# ---------------------------------------------------------------------------
# 数据结构
# ---------------------------------------------------------------------------
@dataclass
class Record:
    version: str
    version_rank: int
    source_file: str
    source_row: int
    member_id: str
    name: str
    level_raw: str
    level: str
    balance: Decimal          # 统一为元，已量化到分
    channel: str | None
    register_date: str
    amount_raw: str
    amount_unit: str          # yuan / cent，记录原始口径


@dataclass
class Rejection:
    source_file: str
    source_row: int
    version: str
    member_id: str | None
    name: str | None
    level_raw: str | None
    amount_raw: str | None
    category: str
    reason: str
    raw: dict[str, str]


@dataclass
class GenStats:
    name: str
    unit: str
    files: list[str] = field(default_factory=list)
    rows_read: int = 0
    accepted: int = 0
    rejected: int = 0
    rejected_by_category: dict[str, int] = field(default_factory=dict)
    superseded: int = 0       # 该代入库成功但被更新一代同会员覆盖的行数
    winners: int = 0          # 该代最终留在 members 里的行数
    accepted_amount: Decimal = Decimal("0")   # 归一成功行金额合计（元）
    winner_amount: Decimal = Decimal("0")     # 最终入库行金额合计（元）
    fractional_cents: list[tuple[str, str, str, str]] = field(default_factory=list)


@dataclass
class MergeGroup:
    member_id: str
    records: list[Record]
    winner: Record
    conflicts: dict[str, list[tuple[str, str]]]  # 字段标签 -> [(版本, 取值), ...]


# ---------------------------------------------------------------------------
# 单文件解析与归一
# ---------------------------------------------------------------------------
def parse_file(
    path: Path,
    mapping: dict[str, str],
    stats_by_gen: dict[str, GenStats],
) -> tuple[list[Record], list[Rejection]]:
    rows = list(read_table(path))
    if not rows:
        return [], []
    header_no, header = rows[0]
    data_rows = rows[1:]

    version_col = find_column(header, ["version"], required=False)
    if version_col is None:
        raise MigrationError(f"{path.name} 表头缺少 version 字段")
    id_col = find_column(header, ["会员编号"])
    name_col = find_column(header, ["姓名"])
    level_col = find_column(header, ["等级编码", "会员等级"])
    amount_col = find_column(header, ["账户余额"])
    date_col = find_column(header, ["注册日期"])
    channel_col = find_column(header, ["注册渠道"], required=False)

    amount_header = header[amount_col] if amount_col is not None else ""
    unit_hint: str | None = None
    if "分" in amount_header:
        unit_hint = "cent"
    elif "元" in amount_header:
        unit_hint = "yuan"

    records: list[Record] = []
    rejections: list[Rejection] = []

    def reject(
        row_no: int,
        cells: list[str],
        category: str,
        reason: str,
        version: str = "",
    ) -> None:
        raw_map = {
            header[i]: (cells[i] if i < len(cells) else "")
            for i in range(len(header))
        }
        stats = stats_by_gen.get(version)
        if stats is not None:
            stats.rejected += 1
            stats.rejected_by_category[category] = (
                stats.rejected_by_category.get(category, 0) + 1
            )
        rejections.append(
            Rejection(
                source_file=path.name,
                source_row=row_no,
                version=version,
                member_id=cells[id_col].strip() if id_col < len(cells) else None,
                name=cells[name_col].strip() if name_col < len(cells) else None,
                level_raw=cells[level_col].strip() if level_col < len(cells) else None,
                amount_raw=cells[amount_col].strip()
                if amount_col < len(cells)
                else None,
                category=category,
                reason=reason,
                raw=raw_map,
            )
        )

    for row_no, cells in data_rows:
        # 短行补齐，避免列下标越界
        if len(cells) < len(header):
            cells = cells + [""] * (len(header) - len(cells))

        version = cells[version_col].strip().lower() if version_col < len(cells) else ""
        if version not in GENERATIONS:
            raise MigrationError(
                f"{path.name} 第 {row_no} 行的 version={version!r} 不在已知代际 "
                f"{list(GENERATIONS)} 中，请先确认新模板口径并更新脚本"
            )
        gen = GENERATIONS[version]
        stats = stats_by_gen[version]
        stats.rows_read += 1

        member_id = cells[id_col].strip()
        name = cells[name_col].strip()
        level_raw = cells[level_col].strip()
        amount_raw = cells[amount_col].strip()
        date_raw = cells[date_col].strip()
        channel = (
            cells[channel_col].strip()
            if channel_col is not None and channel_col < len(cells)
            else ""
        ) or None

        # 1) 关键字段完整性
        if not member_id:
            reject(row_no, cells, "missing_id", "会员编号为空", version)
            continue
        if not name:
            reject(row_no, cells, "missing_name", "姓名为空", version)
            continue

        # 2) 等级映射（覆盖不到即拒收，映射完全由 mapping.yaml 决定）
        level = mapping.get(level_raw)
        if level is None:
            reject(
                row_no,
                cells,
                "unmapped_level",
                f"会员等级 {level_raw!r} 在 mapping.yaml 的「{MAPPING_SECTION}」"
                "中没有对应项",
                version,
            )
            continue

        # 3) 金额解析与单位归一
        try:
            amount = parse_amount(amount_raw)
        except ValueError as exc:
            reject(row_no, cells, "invalid_amount", str(exc), version)
            continue

        unit = {"元": "yuan", "分": "cent"}.get(unit_hint or "", gen["unit"])
        expected_unit = gen["unit"]
        if unit != expected_unit:
            raise MigrationError(
                f"{path.name} 的 {amount_header!r} 余额列写的是「{amount_header}」，"
                f"与该代已知口径（{'分' if expected_unit == 'cent' else '元'}）不符，"
                "请确认模板口径后再迁移"
            )

        if unit == "cent":
            # 记录带小数的“分”，报告里做数据质量提示；四舍五入到分后照常入库
            if amount != amount.to_integral_value():
                stats.fractional_cents.append(
                    (
                        member_id,
                        name,
                        amount_raw,
                        f"{(amount / CENT).quantize(Q2, ROUND_HALF_UP):.2f}",
                    )
                )
            amount = (amount / CENT).quantize(Q2, ROUND_HALF_UP)
        amount = amount.quantize(Q2, ROUND_HALF_UP)

        # 4) 注册日期
        try:
            register_date = parse_date(date_raw)
        except ValueError as exc:
            reject(row_no, cells, "invalid_date", str(exc), version)
            continue

        records.append(
            Record(
                version=version,
                version_rank=list(GENERATIONS).index(version),
                source_file=path.name,
                source_row=row_no,
                member_id=member_id,
                name=name,
                level_raw=level_raw,
                level=level,
                balance=amount,
                channel=channel,
                register_date=register_date,
                amount_raw=amount_raw,
                amount_unit=unit,
            )
        )
        stats.accepted += 1
        stats.accepted_amount += amount

    return records, rejections


# ---------------------------------------------------------------------------
# 跨代归并
# ---------------------------------------------------------------------------
def merge_records(
    records: list[Record],
) -> tuple[list[MergeGroup], int]:
    by_id: OrderedDict[str, list[Record]] = OrderedDict()
    for record in records:
        by_id.setdefault(record.member_id, []).append(record)

    same_gen_duplicates = 0
    groups: list[MergeGroup] = []
    compare_fields = (
        ("name", "姓名"),
        ("level", "等级"),
        ("balance", "余额"),
        ("register_date", "注册日期"),
        ("channel", "注册渠道"),
    )

    for member_id, rows in by_id.items():
        rows.sort(key=lambda r: (r.version_rank, r.source_file, r.source_row))

        # 同一代内重复：保留该代最后一条
        deduped: list[Record] = []
        seen_gen: set[str] = set()
        for record in rows:
            if record.version in seen_gen:
                same_gen_duplicates += 1
                deduped[-1] = record
            else:
                seen_gen.add(record.version)
                deduped.append(record)

        winner = deduped[-1]  # 最新一代为准
        conflicts: dict[str, list[tuple[str, str]]] = {}
        for attr, label in compare_fields:
            values: list[tuple[str, str]] = []
            seen: set[str] = set()
            for record in deduped:
                value = getattr(record, attr)
                if value is None:
                    continue  # 老模板没有注册渠道，NULL 不参与冲突比较
                text = f"{value:.2f}" if isinstance(value, Decimal) else str(value)
                if text not in seen:
                    seen.add(text)
                    values.append((record.version, text))
            if len(values) > 1:
                conflicts[label] = values

        groups.append(MergeGroup(member_id, deduped, winner, conflicts))

    return groups, same_gen_duplicates


# ---------------------------------------------------------------------------
# SQLite 与拒收留存文件
# ---------------------------------------------------------------------------
def write_database(
    groups: list[MergeGroup],
    rejections: list[Rejection],
) -> None:
    conn = sqlite3.connect(DB_PATH)
    try:
        conn.execute("DROP TABLE IF EXISTS members")
        conn.execute("DROP TABLE IF EXISTS members_rejected")
        conn.execute(
            """
            CREATE TABLE members (
                member_id        TEXT PRIMARY KEY,
                name             TEXT NOT NULL,
                level            TEXT NOT NULL,
                balance_cents    INTEGER NOT NULL,
                balance_yuan     REAL NOT NULL,
                register_channel TEXT,
                register_date    TEXT NOT NULL,
                source_versions  TEXT NOT NULL,
                source_refs      TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE members_rejected (
                source_file  TEXT NOT NULL,
                source_row   INTEGER NOT NULL,
                version      TEXT NOT NULL,
                member_id    TEXT,
                name         TEXT,
                level_raw    TEXT,
                balance_raw  TEXT,
                reason_code  TEXT NOT NULL,
                reason       TEXT NOT NULL,
                raw_json     TEXT NOT NULL,
                PRIMARY KEY (source_file, source_row)
            )
            """
        )

        for group in sorted(groups, key=lambda g: g.member_id):
            winner = group.winner
            versions = ",".join(r.version for r in group.records)
            refs = "; ".join(
                f"{r.source_file}:{r.source_row}" for r in group.records
            )
            conn.execute(
                """
                INSERT INTO members
                    (member_id, name, level, balance_cents, balance_yuan,
                     register_channel, register_date, source_versions, source_refs)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    winner.member_id,
                    winner.name,
                    winner.level,
                    int((winner.balance * 100).to_integral_value(ROUND_HALF_UP)),
                    float(winner.balance),
                    winner.channel,
                    winner.register_date,
                    versions,
                    refs,
                ),
            )

        for rej in rejections:
            conn.execute(
                """
                INSERT INTO members_rejected
                    (source_file, source_row, version, member_id, name,
                     level_raw, balance_raw, reason_code, reason, raw_json)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    rej.source_file,
                    rej.source_row,
                    rej.version,
                    rej.member_id,
                    rej.name,
                    rej.level_raw,
                    rej.amount_raw,
                    rej.category,
                    rej.reason,
                    json.dumps(rej.raw, ensure_ascii=False),
                ),
            )
        conn.commit()
    finally:
        conn.close()


def write_rejected_csv(rejections: list[Rejection]) -> None:
    """带“拒”字的留存件：原始整行 + 拒收原因，每行一条。"""
    raw_headers: list[str] = []
    for rej in rejections:
        for key in rej.raw:
            if key not in raw_headers:
                raw_headers.append(key)
    # 原始列里本身含 version，元数据列改名为“代际”避免重名
    columns = (
        ["来源文件", "原始行号", "代际"]
        + raw_headers
        + ["拒收原因代码", "拒收原因"]
    )
    with open(REJECT_CSV_PATH, "w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow(columns)
        for rej in rejections:
            writer.writerow(
                [rej.source_file, rej.source_row, rej.version]
                + [rej.raw.get(h, "") for h in raw_headers]
                + [rej.category, rej.reason]
            )


# ---------------------------------------------------------------------------
# 迁移报告
# ---------------------------------------------------------------------------
def money(value: Decimal) -> str:
    return f"{value:,.2f}"


def build_report(
    stats_by_gen: dict[str, GenStats],
    rejections: list[Rejection],
    groups: list[MergeGroup],
    same_gen_duplicates: int,
    mapping_entries: list[tuple[str, str]],
    total_balance: Decimal,
) -> str:
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    total_read = sum(s.rows_read for s in stats_by_gen.values())
    total_accepted = sum(s.accepted for s in stats_by_gen.values())
    total_rejected = sum(s.rejected for s in stats_by_gen.values())
    total_superseded = sum(s.superseded for s in stats_by_gen.values())
    merged_groups = [g for g in groups if len({r.version for r in g.records}) > 1]
    conflicting = [g for g in merged_groups if g.conflicts]
    channel_null = sum(1 for g in groups if g.winner.channel is None)

    lines: list[str] = []
    lines.append("# 会员数据迁移报告")
    lines.append("")
    lines.append(f"- 生成时间：{now}（由 `migrate.py` 自动生成，可重复运行）")
    lines.append(
        f"- 输出库：`{DB_PATH.name}`，表：`members`（归一结果）、"
        "`members_rejected`（未归一原始记录）"
    )
    lines.append(
        f"- 拒收留存件：`{REJECT_CSV_PATH.name}`（每行附拒收原因，原始整行备查）"
    )
    lines.append(
        "- 归一基准：最新一代模板 v3 的字段与口径（余额单位：元，含注册渠道）"
    )
    lines.append(
        f"- 等级映射：`{MAPPING_PATH.name}`，本次载入 {len(mapping_entries)} 条对应关系"
    )
    lines.append("")

    # 1. 各代读取与入库
    lines.append("## 1. 各代读取与入库")
    lines.append("")
    lines.append(
        "| 代际 | 文件 | 读取行数 | 入库成功 | 覆盖不到 | 被更新一代覆盖 | 最终入库 |"
    )
    lines.append(
        "| --- | --- | ---: | ---: | ---: | ---: | ---: |"
    )
    for ver in GENERATIONS:
        stats = stats_by_gen[ver]
        files = "<br>".join(stats.files) if stats.files else "—"
        lines.append(
            f"| {stats.name} | {files} | {stats.rows_read} | {stats.accepted} | "
            f"{stats.rejected} | {stats.superseded} | {stats.winners} |"
        )
    lines.append(
        f"| **合计** | — | **{total_read}** | **{total_accepted}** | "
        f"**{total_rejected}** | **{total_superseded}** | **{len(groups)}** |"
    )
    lines.append("")
    lines.append("说明：")
    lines.append("- “读取行数”不含表头；“覆盖不到”指无法按 mapping.yaml 归一的等级值等异常行。")
    lines.append("- “入库成功”是各代自身可归一的记录数；同一会员跨代重复时，只保留最新一代的记录。")
    lines.append(
        f"- 各代入库成功合计 {total_accepted} 行，其中 {total_superseded} 行被更新一代的"
        f"同会员记录覆盖，`members` 最终 {len(groups)} 行。"
    )
    if same_gen_duplicates:
        lines.append(
            f"- 另有 {same_gen_duplicates} 行在同一代内重复出现同一会员，已保留后一条。"
        )
    lines.append("")

    # 2. 覆盖不到的行
    lines.append("## 2. 覆盖不到的行")
    lines.append("")
    if rejections:
        categories = sorted({r.category for r in rejections})
        cat_text = "、".join(categories)
        lines.append(
            f"本次共 {len(rejections)} 行未能归一（原因代码：{cat_text}），明细如下："
        )
        lines.append("")
        lines.append(
            "| 来源文件 | 行号 | 代际 | 会员编号 | 姓名 | 原始等级 | 原始余额 | 原因 |"
        )
        lines.append("| --- | ---: | --- | --- | --- | --- | --- | --- |")
        for rej in rejections:
            lines.append(
                f"| {rej.source_file} | {rej.source_row} | {rej.version or '—'} | "
                f"{rej.member_id or '—'} | {rej.name or '—'} | "
                f"{rej.level_raw or '—'} | {rej.amount_raw or '—'} | {rej.reason} |"
            )
        lines.append("")
        lines.append(
            "处理方式：这些行**不写入 `members`**，整行原始记录（含全部原始列）保存在 "
            "`members_rejected` 表与 "
            f"`{REJECT_CSV_PATH.name}` 中备查。等级值属于 mapping.yaml 的维护范围："
            "运营补上对应关系后直接重跑 `python3 migrate.py`，这些行即会按新映射归一入库。"
        )
    else:
        lines.append("本次没有覆盖不到的行。")
    lines.append("")

    # 3. 金额口径
    lines.append("## 3. 金额口径")
    lines.append("")
    lines.append(
        "| 代际 | 原始口径 | 归一方式 | 入库成功行金额合计（元） | 最终入库行金额合计（元） |"
    )
    lines.append("| --- | --- | --- | ---: | ---: |")
    unit_text = {"cent": "分", "yuan": "元"}
    for ver in GENERATIONS:
        stats = stats_by_gen[ver]
        method = "÷100 后四舍五入到分" if stats.unit == "cent" else "原样保留，四舍五入到分"
        lines.append(
            f"| {stats.name} | {unit_text[stats.unit]} | {method} | "
            f"{money(stats.accepted_amount)} | {money(stats.winner_amount)} |"
        )
    lines.append(
        f"| **合计** | — | — | — | **{money(sum(s.winner_amount for s in stats_by_gen.values()))}** |"
    )
    lines.append("")
    lines.append("金额合计均为各代逐行归一到分后的 Decimal 精确求和（v2 的分值先 ÷100 再四舍五入到分）。")
    lines.append("")
    lines.append(
        f"**入库金额合计：¥{money(total_balance)}**（`members` 表 {len(groups)} 行余额之和，"
        "按 Decimal 精确求和；SQLite 中 `balance_cents` 为整数分，报表求和请用此列避免浮点误差）。"
    )
    lines.append("")

    # 3.1 v2 小数分数据质量提示
    v2_stats = stats_by_gen["v2"]
    if v2_stats.fractional_cents:
        fc = v2_stats.fractional_cents
        lines.append(
            f"数据质量提示：v2 的“分”列中有 {len(fc)} 行带小数，按“1 元 = 100 分”换算后"
            "不足一分，已四舍五入到分。明细："
        )
        lines.append("")
        lines.append("| 会员编号 | 姓名 | 原始值（分） | 归一后（元） |")
        lines.append("| --- | --- | ---: | ---: |")
        for member_id, name, raw, normalized in fc:
            lines.append(f"| {member_id} | {name} | {raw} | {normalized} |")
        lines.append("")
        lines.append(
            f"这些值在 v3 中没有对应记录，无法交叉验证。若实际是导出时误按“元”填写，"
            "需运营确认后修正源数据；按当前口径（分）换算，这 "
            f"{len(fc)} 行对合计的影响已包含在上表中。"
        )
        lines.append("")

    # 4. 同会员跨代归并
    lines.append("## 4. 同会员跨代归并")
    lines.append("")
    if merged_groups:
        lines.append(
            f"同一会员在多个代际文件中出现的共 {len(merged_groups)} 人，"
            "按“最新一代记录为准”归并。其中 "
            f"{len(conflicting)} 人存在字段不一致，`members` 保留最新一代的值，差异如下："
        )
        lines.append("")
        lines.append("| 会员编号 | 姓名（最新） | 出现代际 | 冲突字段（旧 → 新） |")
        lines.append("| --- | --- | --- | --- |")
        for group in merged_groups:
            versions = "、".join(r.version for r in group.records)
            if group.conflicts:
                detail = "；".join(
                    f"{label}："
                    + " → ".join(f"{ver}={val}" for ver, val in vals)
                    for label, vals in group.conflicts.items()
                )
            else:
                detail = "无（各字段一致）"
            lines.append(
                f"| {group.member_id} | {group.winner.name} | {versions} | {detail} |"
            )
        lines.append("")
        consistent = len(merged_groups) - len(conflicting)
        lines.append(
            f"归并结果：{len(merged_groups)} 人中 {consistent} 人各字段完全一致，"
            f"{len(conflicting)} 人存在冲突并已按最新代取值；冲突只记录在本报告中，不阻塞入库。"
        )
    else:
        lines.append("本次没有同一会员跨代出现的情况。")
    lines.append("")

    # 5. 输出字段口径
    lines.append("## 5. 输出字段口径")
    lines.append("")
    lines.append("| members 列 | 含义 | 来源与口径 |")
    lines.append("| --- | --- | --- |")
    lines.append("| `member_id` | 会员编号 | 各代的“会员编号”，作为跨代归并键 |")
    lines.append("| `name` | 姓名 | 取最新一代记录 |")
    lines.append("| `level` | 会员等级 | 文字等级或数字编码经 mapping.yaml 映射后的等级名称 |")
    lines.append("| `balance_cents` | 账户余额（分） | 精确到分的整数，报表求和请用此列避免浮点误差 |")
    lines.append("| `balance_yuan` | 账户余额（元） | 各代按表头单位换算、四舍五入到分；v2 为分 ÷100 |")
    lines.append("| `register_channel` | 注册渠道 | 仅 v3 提供；v1/v2 无此列，合并后仍无来源的为 NULL |")
    lines.append("| `register_date` | 注册日期 | 统一为 YYYY-MM-DD |")
    lines.append("| `source_versions` | 该会员出现过的代际 | 归并元数据，按代际升序 |")
    lines.append("| `source_refs` | 来源文件:行号 | 归并元数据，便于回查原始记录 |")
    lines.append("")

    # 6. 其他说明
    lines.append("## 6. 其他说明")
    lines.append("")
    lines.append(
        f"- v1/v2 模板没有注册渠道字段，`members` 中有 {channel_null} 人的 "
        "`register_channel` 为 NULL。"
    )
    lines.append(
        "- 每次运行都会重建 `members`、`members_rejected` 与拒收留存 CSV，"
        f"结果只取决于 `{EXPORTS_DIR.name}/` 与 `{MAPPING_PATH.name}` 的当前内容。"
    )
    lines.append(
        f"- `{MAPPING_PATH.name}` 是等级口径的唯一来源，脚本内不写死任何等级对应关系；"
        "修改后重跑即可生效。"
    )
    lines.append("")

    # 附录：本次使用的等级映射
    lines.append("## 附录：本次使用的等级映射")
    lines.append("")
    lines.append("| 原始等级取值 | 归一后等级 |")
    lines.append("| --- | --- |")
    for key, value in mapping_entries:
        lines.append(f"| {key} | {value} |")
    lines.append("")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------
def main() -> int:
    if not EXPORTS_DIR.is_dir():
        raise MigrationError(f"找不到导出目录：{EXPORTS_DIR}")

    files = sorted(p for p in EXPORTS_DIR.iterdir() if p.is_file())
    files = [
        p
        for p in files
        if p.suffix.lower() in (".csv", ".xlsx", ".xlsm", ".xls")
    ]
    for p in files:
        if p.suffix.lower() == ".xls":
            raise MigrationError(
                f"发现旧版 Excel 文件 {p.name}，本脚本不支持 .xls，"
                "请先另存为 .xlsx 后重试"
            )
    if not files:
        raise MigrationError(f"{EXPORTS_DIR} 中没有可读取的 csv/xlsx 文件")

    mapping, mapping_entries = load_level_mapping(MAPPING_PATH)

    stats_by_gen: dict[str, GenStats] = {
        ver: GenStats(name=cfg["name"], unit=cfg["unit"])
        for ver, cfg in GENERATIONS.items()
    }

    all_records: list[Record] = []
    rejections: list[Rejection] = []
    for path in files:
        file_records, file_rejections = parse_file(path, mapping, stats_by_gen)
        versions = sorted({r.version for r in file_records} | {r.version for r in file_rejections if r.version})
        for ver in versions:
            if path.name not in stats_by_gen[ver].files:
                stats_by_gen[ver].files.append(path.name)
        all_records.extend(file_records)
        rejections.extend(file_rejections)

    # 跨代归并
    groups, same_gen_duplicates = merge_records(all_records)

    # 统计每代覆盖/最终入库情况
    winner_keys: set[tuple[str, str]] = set()
    for group in groups:
        winner = group.winner
        winner_keys.add((winner.version, winner.member_id))
        stats_by_gen[winner.version].winners += 1
        stats_by_gen[winner.version].winner_amount += winner.balance
        for record in group.records:
            if record is not winner:
                stats_by_gen[record.version].superseded += 1

    write_database(groups, rejections)
    write_rejected_csv(rejections)

    total_balance = sum((g.winner.balance for g in groups), Decimal("0"))
    report = build_report(
        stats_by_gen,
        rejections,
        groups,
        same_gen_duplicates,
        mapping_entries,
        total_balance,
    )
    REPORT_PATH.write_text(report, encoding="utf-8")

    print(f"迁移完成：{DB_PATH.name} 的 members 表 {len(groups)} 行，"
          f"members_rejected 表 {len(rejections)} 行")
    print(f"拒收留存件：{REJECT_CSV_PATH.name}（{len(rejections)} 行）")
    print(f"迁移报告：{REPORT_PATH.name}")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except MigrationError as exc:
        print(f"迁移失败：{exc}", file=sys.stderr)
        sys.exit(1)
    except Exception as exc:  # noqa: BLE001
        print(f"迁移失败（未预期错误）：{exc}", file=sys.stderr)
        sys.exit(2)

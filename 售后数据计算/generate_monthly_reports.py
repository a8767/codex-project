#!/usr/bin/env python3
"""Generate monthly after-sales reports from the files in ``原始数据``.

The script is deliberately data-driven: all column names written to an Excel
template come from that template's first row.  It does not use positions from
the old ``数据计算`` files as a source of truth.

Examples
--------
    python generate_monthly_reports.py --month 2026-09
    python generate_monthly_reports.py --month 2026-09 --created-date 20260910

Outputs are written to ``数据计算`` as ``表格名称_创建日期.xlsx``.
"""

from __future__ import annotations

import argparse
import re
from collections import Counter, defaultdict
from copy import copy
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Iterable

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils.datetime import from_excel, WINDOWS_EPOCH


ROOT = Path(__file__).resolve().parent
RAW_DIR_NAME = "原始数据"
TEMPLATE_DIR_NAME = "数据表模板"
OUTPUT_DIR_NAME = "数据计算"

TEMPLATE_FILES = {
    "总数据计算": "总数据计算.xlsx",
    "小组数据计算": "小组统计表.xlsx",
    "大组数据计算": "大组统计表.xlsx",
    "工厂数据计算": "工厂统计表.xlsx",
    "车型数据计算": "车型统计表.xlsx",
    "配件数据计算": "配件统计表.xlsx",
}

# 输出文件名跟随“数据表模板”中的实际模板名称；左侧仍保留内部报表逻辑名，
# 以免改变来源依赖、备注和下游字段映射。
OUTPUT_FILE_STEMS = {name: Path(filename).stem for name, filename in TEMPLATE_FILES.items()}
OUTPUT_FILE_STEMS["店铺数据计算"] = "店铺统计表"

# 每次运行收集原始数据的可用性。缺失数据只影响相关指标，不中断整批报表。
MISSING_INPUTS: dict[str, set[str]] = {}
DATA_ENTRY_ERRORS: list[dict[str, Any]] = []
RAW_RECORD_CACHE: dict[Path, list[dict[str, Any]]] = {}
RAW_TUPLE_CACHE: dict[Path, list[tuple[Any, ...]]] = {}
INVALID_SOURCE_ROWS: dict[Path, list[tuple[int, list[str]]]] = {}
REQUIRED_COLUMNS = {
    "店铺规范.xlsx": {"店铺名称", "售后组", "当日真实营业额", "车子销量"},
    "工厂数据.xlsx": {"工厂"},
    "退货数据.xlsx": {"登记时间", "店铺", "售后组", "工厂", "退货退款原因"},
    # 补发售后组为空时会按店铺规范回填，因此不是缺失告警条件。
    "补发数据.xlsx": {"登记时间", "店铺"},
    "万里牛补发.xlsx": {"售后组", "数量", "日期"},
    "补偿数据.xlsx": {"登记时间", "店铺", "售后组", "补偿金额", "是否不参与计算判断"},
    "永久日销.xlsx": {"编码表车型", "销量", "销售额"},
    "菲利普日销.xlsx": {"型号", "销量", "销售额"},
    "大众日销.xlsx": {"车型", "销量", "销售额"},
    "玛莎拉蒂日销.xlsx": {"车型", "汇总", "销售额"},
    "配件日销.xlsx": {"标题", "单选", "销量", "销售额"},
}
# 用户确认：单条登记是否完整，只以店铺名称是否填写为准。
SHOP_NAME_FIELD = {
    "店铺规范.xlsx": "店铺名称",
    "退货数据.xlsx": "店铺",
    "补发数据.xlsx": "店铺",
    "补偿数据.xlsx": "店铺",
}

# 原始数据文件名可能带日期后缀或使用业务表名。优先按关键词识别，旧文件名继续兼容。
SOURCE_KEYWORDS = {
    "店铺规范.xlsx": ("店铺规范",), "工厂数据.xlsx": ("工厂数据匹配", "工厂数据"),
    "退货数据.xlsx": ("退货登记表", "退货数据"), "补发数据.xlsx": ("补发登记表", "补发数据"),
    "补偿数据.xlsx": ("补偿登记表", "补偿数据"), "配件日销.xlsx": ("配件日销表", "配件日销"),
    "售后分组.xlsx": ("售后分组",), "万里牛补发.xlsx": ("万里牛补发",),
    "永久车型英文编码.xlsx": ("永久车型编码",), "菲利普车型英文编码.xlsx": ("菲利普车型编码",),
    "大众车型英文编码.xlsx": ("大众车型编码",), "玛莎拉蒂英文编码.xlsx": ("玛莎拉蒂车型编码",),
    "配件英文编码.xlsx": ("配件编码",),
    "永久编码.xlsx": ("永久车型编码",), "菲利普编码.xlsx": ("菲利普车型编码",),
    "大众编码.xlsx": ("大众车型编码",), "玛莎拉蒂编码.xlsx": ("玛莎拉蒂车型编码",),
}

def resolve_source_path(path: Path) -> Path:
    if path.is_file():
        return path
    keywords = SOURCE_KEYWORDS.get(path.name, (path.stem,))
    # WPS/Excel 打开工作簿时会生成 ``~$...xlsx`` 锁定文件；它不是数据源，
    # 不能参与关键字匹配，否则可能被误选为当前输入表。
    candidates = [p for p in path.parent.glob("*.xlsx") if p.is_file() and not p.name.startswith("~$")]
    ranked = sorted(candidates, key=lambda p: max((len(k) for k in keywords if k in p.stem), default=0), reverse=True)
    return ranked[0] if ranked and max((len(k) for k in keywords if k in ranked[0].stem), default=0) > 0 else path


def reset_missing_inputs() -> None:
    MISSING_INPUTS.clear()
    DATA_ENTRY_ERRORS.clear()
    RAW_RECORD_CACHE.clear()
    RAW_TUPLE_CACHE.clear()
    INVALID_SOURCE_ROWS.clear()


def record_missing(filename: str, reason: str) -> None:
    MISSING_INPUTS.setdefault(filename, set()).add(reason)


def missing_note(*filenames: str) -> str:
    """Return a concise, row-safe explanation of unavailable source inputs."""
    notes = []
    for filename in filenames:
        reasons = MISSING_INPUTS.get(filename, set())
        if reasons:
            notes.append(f"{filename}（{'、'.join(sorted(reasons))}）")
    return "缺少：" + "；".join(notes) if notes else ""


def row_missing_fields(filename: str, record: dict[str, Any]) -> list[str]:
    """Only a missing shop name excludes an individual source record."""
    field = SHOP_NAME_FIELD.get(filename)
    return [field] if field and clean(record.get(field)) == "" else []


def is_red_cell(cell: Any) -> bool:
    """Identify a visibly red fill without treating light error highlighting as a tag."""
    color = getattr(getattr(cell, "fill", None), "fgColor", None)
    rgb = clean(getattr(color, "rgb", ""))[-6:].upper()
    if len(rgb) != 6:
        return False
    try:
        red, green, blue = int(rgb[:2], 16), int(rgb[2:4], 16), int(rgb[4:], 16)
    except ValueError:
        return False
    return red >= 180 and green <= 120 and blue <= 150


def clear_previous_error_marks(raw_dir: Path) -> None:
    """Remove only the prior auto-generated registration-error flags."""
    for path in raw_dir.glob("*.xlsx"):
        if path.name.startswith("~$"):
            continue
        try:
            workbook = load_workbook(path)
            sheet = workbook.active
            headers = [clean(cell.value) for cell in sheet[1]]
            if "备注" not in headers:
                continue
            column = headers.index("备注") + 1
            changed = False
            for row in range(2, sheet.max_row + 1):
                cell = sheet.cell(row, column)
                if clean(cell.value).startswith("登记错误："):
                    cell.value = None
                    for item in sheet[row]:
                        item.fill = PatternFill()
                        item.font = Font(color="000000")
                    changed = True
            if changed:
                workbook.save(path)
        except Exception:
            record_missing(path.name, "旧登记错误标记无法清理")


def mark_invalid_rows(path: Path, invalid_rows: list[tuple[int, list[str]]]) -> bool:
    """Flag excluded source rows in red and append a reusable error remark."""
    if not invalid_rows:
        return True
    try:
        workbook = load_workbook(path)
        if not workbook.worksheets:
            return
        sheet = workbook.active
        headers = [clean(cell.value) for cell in sheet[1]]
        if "备注" in headers:
            remark_column = headers.index("备注") + 1
        else:
            remark_column = sheet.max_column + 1
            header = sheet.cell(1, remark_column, "备注")
            header.font = Font(color="FFFFFF", bold=True)
            header.fill = PatternFill("solid", fgColor="C00000")
            header.alignment = Alignment(horizontal="center", vertical="center")
            sheet.column_dimensions[header.column_letter].width = 36
        red_fill = PatternFill("solid", fgColor="FFC7CE")
        red_font = Font(color="9C0006")
        for row_number, fields in invalid_rows:
            for column in range(1, remark_column + 1):
                cell = sheet.cell(row_number, column)
                cell.fill = red_fill
                cell.font = copy(red_font)
            sheet.cell(row_number, remark_column).value = "登记错误：缺少" + "、".join(fields) + "；仅不参与需要店铺名称的统计"
        workbook.save(path)
        return True
    except Exception:
        return False


def apply_error_marks(raw_dir: Path) -> list[str]:
    """Write source-cell flags after read-only calculation and report export."""
    failures: list[str] = []
    clear_previous_error_marks(raw_dir)
    for path, rows in INVALID_SOURCE_ROWS.items():
        if not mark_invalid_rows(path, rows):
            failures.append(path.name)
    return failures


def write_data_entry_errors(output_path: Path) -> None:
    """Export excluded source records with their source and missing fields."""
    fixed_headers = ["来源表格", "原始行号", "缺失字段", "备注"]
    source_headers = sorted({key for row in DATA_ENTRY_ERRORS for key in row if key not in fixed_headers})
    headers = fixed_headers + source_headers
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "数据登记错误"
    sheet.freeze_panes = "A2"
    header_fill = PatternFill("solid", fgColor="C00000")
    for column, header in enumerate(headers, start=1):
        cell = sheet.cell(1, column, header)
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        sheet.column_dimensions[cell.column_letter].width = max(14, min(36, len(header) * 2 + 8))
    red_fill = PatternFill("solid", fgColor="FFC7CE")
    for row_index, record in enumerate(DATA_ENTRY_ERRORS, start=2):
        for column, header in enumerate(headers, start=1):
            cell = sheet.cell(row_index, column, record.get(header, ""))
            cell.fill = red_fill
            cell.alignment = Alignment(vertical="center", wrap_text=True)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)


REPORT_SOURCES = {
    "总数据计算": ("店铺规范.xlsx", "退货数据.xlsx", "补发数据.xlsx", "万里牛补发.xlsx", "补偿数据.xlsx"),
    "小组数据计算": ("店铺规范.xlsx", "退货数据.xlsx", "补发数据.xlsx", "万里牛补发.xlsx", "补偿数据.xlsx"),
    "大组数据计算": ("店铺规范.xlsx", "退货数据.xlsx", "补发数据.xlsx", "万里牛补发.xlsx", "补偿数据.xlsx"),
    "工厂数据计算": ("工厂数据.xlsx", "退货数据.xlsx", "补发数据.xlsx", "补偿数据.xlsx", "万里牛补发.xlsx"),
    "车型数据计算": ("永久编码.xlsx", "菲利普编码.xlsx", "大众编码.xlsx", "玛莎拉蒂编码.xlsx", "永久日销.xlsx", "菲利普日销.xlsx", "大众日销.xlsx", "玛莎拉蒂日销.xlsx", "退货数据.xlsx", "补发数据.xlsx", "补偿数据.xlsx"),
    "配件数据计算": ("配件编码.xlsx", "配件日销.xlsx", "补发数据.xlsx", "补偿数据.xlsx", "退货数据.xlsx"),
    "店铺数据计算": ("店铺规范.xlsx", "退货数据.xlsx", "补发数据.xlsx", "补偿数据.xlsx"),
}


def clean(value: Any) -> str:
    """Return a trimmed display/key string without turning None into 'None'."""
    if value is None:
        return ""
    return str(value).strip()


# 2026-09-22 用户确认的同店别名。不得泛化删除其他店名的关店标记。
SHOP_NAME_ALIASES = {
    "快手PHILLIPS户外装备旗舰店": "快手PHILLIPS户外装备旗舰店",
    "快手PHILLIPS户外装备旗舰店（已关店）": "快手PHILLIPS户外装备旗舰店",
}


def canonical_shop_name(value: Any) -> str:
    """Only the explicitly approved pair is merged; other names stay exact."""
    name = clean(value)
    return SHOP_NAME_ALIASES.get(name, name)


def read_shop_records(path: Path) -> list[dict[str, Any]]:
    """Normalize copies for aggregation; cached/source records stay unchanged."""
    rows = []
    for original in read_records(path):
        row = dict(original)
        for field in ("店铺", "店铺名称", "店铺名"):
            if field in row:
                original_name = clean(row[field])
                row[field] = canonical_shop_name(original_name)
                if row[field] != original_name:
                    row["__原始店铺名称__"] = original_name
        rows.append(row)
    return rows


def number(value: Any) -> Decimal:
    """Convert spreadsheet numbers stored as strings, blanks, or numbers."""
    if value is None or value == "":
        return Decimal("0")
    if isinstance(value, bool):
        return Decimal(int(value))
    if isinstance(value, Decimal):
        return value
    try:
        return Decimal(str(value).replace(",", "").replace("¥", "").strip())
    except (InvalidOperation, ValueError):
        return Decimal("0")


def excel_number(value: Decimal | int | float) -> int | float:
    """Keep quantities integral and write amounts with two decimal places."""
    amount = Decimal(value).quantize(Decimal("0.01"))
    return int(amount) if amount == amount.to_integral_value() else float(amount)


def ratio(numerator: Decimal, denominator: Decimal) -> float:
    return float(numerator / denominator) if denominator else 0.0


def month_key(year: int, month: int) -> str:
    return f"{year:04d}-{month:02d}"


def normalize_excel_date(value: Any, epoch: datetime = WINDOWS_EPOCH) -> Any:
    """Convert numeric Excel dates, preserving text and existing date objects."""
    if isinstance(value, (int, float, Decimal)) and not isinstance(value, bool):
        try:
            converted = from_excel(float(value), epoch=epoch)
            return converted if isinstance(converted, (date, datetime)) else None
        except (ValueError, OverflowError, TypeError):
            return None
    return value


def parse_month(value: Any, fallback_year: int) -> str | None:
    """Parse common Excel/Chinese date forms to YYYY-MM.

    补发数据 currently records dates such as ``09月09日`` without a year.  Its
    year is therefore supplied by the selected report month.
    """
    value = normalize_excel_date(value)
    if isinstance(value, datetime):
        return month_key(value.year, value.month)
    if isinstance(value, date):
        return month_key(value.year, value.month)
    text = clean(value)
    if not text:
        return None
    match = re.search(r"(20\d{2})\D*(\d{1,2})", text)
    if match:
        return month_key(int(match.group(1)), int(match.group(2)))
    match = re.search(r"(\d{1,2})\s*月", text)
    if match:
        return month_key(fallback_year, int(match.group(1)))
    return None

def parse_concrete_date(value: Any, fallback_year: int) -> date | None:
    value = normalize_excel_date(value)
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = clean(value)
    m = re.search(r"(20\d{2})\D*(\d{1,2})\D*(\d{1,2})", text)
    if m:
        try: return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError: return None
    m = re.search(r"(\d{1,2})\s*月\s*(\d{1,2})\s*日", text)
    if m:
        try: return date(fallback_year, int(m.group(1)), int(m.group(2)))
        except ValueError: return None
    return None


def read_records(path: Path) -> list[dict[str, Any]]:
    """Read the active sheet as header-keyed records, using cached values."""
    canonical_name = path.name
    path = resolve_source_path(path).resolve()
    if path in RAW_RECORD_CACHE:
        return RAW_RECORD_CACHE[path]
    if not path.is_file():
        if canonical_name != "万里牛补发.xlsx":
            record_missing(canonical_name, "缺失表格")
        RAW_RECORD_CACHE[path] = []
        return RAW_RECORD_CACHE[path]
    try:
        workbook = load_workbook(path, data_only=True, read_only=True)
        if not workbook.worksheets:
            record_missing(canonical_name, "无可用工作表")
            RAW_RECORD_CACHE[path] = []
            return RAW_RECORD_CACHE[path]
        sheet = workbook.active
        headers = [clean(cell) for cell in next(sheet.iter_rows(min_row=1, max_row=1, values_only=True))]
    except Exception:
        record_missing(canonical_name, "无法读取")
        RAW_RECORD_CACHE[path] = []
        return RAW_RECORD_CACHE[path]
    if not any(headers):
        record_missing(canonical_name, "无表头")
        RAW_RECORD_CACHE[path] = []
        return RAW_RECORD_CACHE[path]
    missing_columns = REQUIRED_COLUMNS.get(canonical_name, set()) - set(headers)
    if missing_columns:
        record_missing(canonical_name, "缺少字段：" + "、".join(sorted(missing_columns)))
    records: list[dict[str, Any]] = []
    invalid_rows: list[tuple[int, list[str]]] = []
    tag_index = headers.index("匹配标签") if canonical_name == "补发数据.xlsx" and "匹配标签" in headers else None
    for row_number, cells in enumerate(sheet.iter_rows(min_row=2), start=2):
        row = tuple(cell.value for cell in cells)
        if not any(value is not None and clean(value) != "" for value in row):
            continue
        record = {headers[index]: value for index, value in enumerate(row) if index < len(headers) and headers[index]}
        # General-formatted date exports retain numeric Excel serials.
        # Use this workbook's epoch so 1900 and 1904 date systems both work.
        date_field = "日期" if canonical_name == "万里牛补发.xlsx" else "登记时间"
        if canonical_name in {"补发数据.xlsx", "补偿数据.xlsx", "退货数据.xlsx", "万里牛补发.xlsx"}:
            raw_date = record.get(date_field)
            record[date_field] = normalize_excel_date(raw_date, workbook.epoch)
            if parse_month(record[date_field], datetime.now().year) is None:
                record_missing(canonical_name, f"第{row_number}行{date_field}无法识别：{raw_date!r}，未计入")
        if tag_index is not None and tag_index < len(cells):
            record["__匹配标签红色__"] = is_red_cell(cells[tag_index])
        fields = row_missing_fields(canonical_name, record)
        if fields:
            invalid_rows.append((row_number, fields))
            DATA_ENTRY_ERRORS.append({
                "来源表格": path.name,
                "原始行号": row_number,
                "缺失字段": "、".join(fields),
                "备注": "登记错误：缺少" + "、".join(fields) + "；仅不参与需要店铺名称的统计",
                **record,
            })
        record["__source_row__"] = row_number
        records.append(record)
    workbook.close()
    if invalid_rows:
        record_missing(canonical_name, f"登记错误 {len(invalid_rows)} 行（已标记，按指标条件处理）")
        INVALID_SOURCE_ROWS[path] = invalid_rows
    RAW_RECORD_CACHE[path] = records
    return RAW_RECORD_CACHE[path]


def read_tuple_rows(path: Path) -> list[tuple[Any, ...]]:
    path = resolve_source_path(path).resolve()
    if path in RAW_TUPLE_CACHE:
        return RAW_TUPLE_CACHE[path]
    if not path.is_file():
        record_missing(path.name, "缺失表格")
        RAW_TUPLE_CACHE[path] = []
        return RAW_TUPLE_CACHE[path]
    try:
        workbook = load_workbook(path, data_only=True, read_only=True)
        if not workbook.worksheets:
            record_missing(path.name, "无可用工作表")
            RAW_TUPLE_CACHE[path] = []
            return RAW_TUPLE_CACHE[path]
        RAW_TUPLE_CACHE[path] = list(workbook.active.iter_rows(min_row=2, values_only=True))
        return RAW_TUPLE_CACHE[path]
    except Exception:
        record_missing(path.name, "无法读取")
        RAW_TUPLE_CACHE[path] = []
        return RAW_TUPLE_CACHE[path]


def metric() -> dict[str, Decimal]:
    return {
        "revenue": Decimal("0"),
        "orders": Decimal("0"),
        "part_orders": Decimal("0"),
        "returns": Decimal("0"),
        "signed_returns": Decimal("0"),
        "resends": Decimal("0"),
        "part_resends": Decimal("0"),
        "compensation": Decimal("0"),
    }


def add_metrics(target: dict[str, Decimal], **values: Decimal) -> None:
    for key, value in values.items():
        target[key] += value


def large_group_name(small_group: str) -> str:
    """Confirmed grouping: A1/A2 -> A, B1/B2 -> B, C1/C2 -> C."""
    match = re.fullmatch(r"售后([ABC])\d+组", clean(small_group))
    return f"售后{match.group(1)}组" if match else clean(small_group)


def write_template_rows(template_path: Path, output_path: Path, rows: list[dict[str, Any]]) -> None:
    """Copy one template and fill its first sheet by the template's headers."""
    workbook = load_workbook(template_path)
    if not workbook.worksheets:
        raise ValueError(f"模板没有工作表：{template_path.name}")
    sheet = workbook.active
    headers = [clean(cell.value) for cell in sheet[1]]
    if not any(headers):
        raise ValueError(f"模板没有表头：{template_path.name}")
    # 已通过分表计算的新增记录正常输出；模板名单用于备注提示，不在导出阶段删行。

    if "备注" not in headers:
        remark_column = len(headers) + 1
        source = sheet.cell(1, len(headers))
        remark = sheet.cell(1, remark_column, "备注")
        if source.has_style:
            remark._style = copy(source._style)
        remark.alignment = copy(source.alignment) if source.alignment else Alignment(horizontal="center", vertical="center")
        sheet.column_dimensions[remark.column_letter].width = 42
        headers.append("备注")
    original_max_row = max(sheet.max_row, 2)
    required_last_row = max(1 + len(rows), 2)
    if required_last_row > original_max_row:
        source_row = original_max_row
        for new_row in range(original_max_row + 1, required_last_row + 1):
            for column in range(1, len(headers) + 1):
                source = sheet.cell(source_row, column)
                destination = sheet.cell(new_row, column)
                if source.has_style:
                    destination._style = copy(source._style)
                if source.number_format:
                    destination.number_format = source.number_format
                if source.alignment:
                    destination.alignment = copy(source.alignment)

    for row_index in range(2, max(original_max_row, required_last_row) + 1):
        for column_index in range(1, len(headers) + 1):
            sheet.cell(row_index, column_index).value = None

    for row_index, record in enumerate(rows, start=2):
        for column_index, header in enumerate(headers, start=1):
            value = record.get(header, 0)
            sheet.cell(row_index, column_index).value = value
            if "率" in header:
                sheet.cell(row_index, column_index).number_format = "0.00%"

    if template_path.name == "售后个人数据统计表.xlsx":
        sheet.title = "售后个人数据统计"
        sheet.freeze_panes = "D2"
        sheet.auto_filter.ref = sheet.dimensions
        for index, record in enumerate(rows, start=2):
            group_lines = max(1, clean(record.get("售后组")).count("、") + 1)
            sheet.row_dimensions[index].height = max(90, group_lines * 18)
            for cell in sheet[index]:
                alignment = copy(cell.alignment)
                alignment.wrap_text = True
                cell.alignment = alignment

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)


def write_store_report(output_path: Path, rows: list[dict[str, Any]]) -> None:
    """Create the approved new 店铺数据计算 structure from raw shop data."""
    headers = [
        "店铺名称", "售后组", "大组", "运营组", "配件组", "平台", "只销配件店铺",
        "营业额", "订单量", "配件店铺订单量", "退货单量", "签收退货单量",
        "补发次数", "配件店铺补发次数", "补偿金额", "退货率", "签收退货率",
        "补偿率", "补发率", "配件店铺补发率", "日期", "判断月份", "售后率", "备注",
    ]
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "月度店铺数据汇总"
    sheet.freeze_panes = "A2"
    header_fill = PatternFill("solid", fgColor="1F4E78")
    for column_index, header in enumerate(headers, start=1):
        cell = sheet.cell(1, column_index, header)
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for row_index, record in enumerate(rows, start=2):
        for column_index, header in enumerate(headers, start=1):
            cell = sheet.cell(row_index, column_index, record.get(header, ""))
            if "率" in header:
                cell.number_format = "0.00%"
    for column_index, header in enumerate(headers, start=1):
        width = max(12, min(28, len(header) * 2 + 4))
        sheet.column_dimensions[chr(64 + column_index) if column_index <= 26 else "A"].width = width
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)


def code_maps(raw_dir: Path) -> tuple[dict[str, dict[str, str]], dict[str, set[str]]]:
    """Return Chinese/English-to-model maps and all known vehicle model names."""
    file_by_brand = {
        "永久": "永久编码.xlsx",
        "菲利普": "菲利普编码.xlsx",
        "大众": "大众编码.xlsx",
        "玛莎拉蒂": "玛莎拉蒂编码.xlsx",
    }
    maps: dict[str, dict[str, str]] = {}
    all_models: dict[str, set[str]] = {"永久": set(), "菲利普": set(), "大众": set(), "玛莎拉蒂": set()}
    for brand, filename in file_by_brand.items():
        chinese: dict[str, str] = {}
        english: dict[str, str] = {}
        for row in read_tuple_rows(raw_dir / filename):
            if len(row) < 2:
                continue
            model = clean(row[0]).split()[0] if brand == "玛莎拉蒂" else (clean(row[2]) if len(row) >= 3 else "")
            if not model:
                continue
            all_models[brand].add(model)
            if clean(row[0]):
                chinese[clean(row[0])] = model
            if clean(row[1]):
                english[clean(row[1])] = model
        maps[f"{brand}_中文"] = chinese
        maps[f"{brand}_英文"] = english
    return maps, all_models


def part_standard_names(raw_dir: Path, template_dir: Path | None = None) -> list[str]:
    """Read the unique approved accessory names from 配件日销表的“单选”列.

    “单选” is the business-maintained canonical name.  The output must not
    introduce names from the old template or from raw titles, because those
    names can bypass the daily-sales classification maintained by the user.
    """
    names: list[str] = []
    for row in read_records(raw_dir / "配件日销.xlsx"):
        name = clean(row.get("单选"))
        if name and name not in names:
            names.append(name)
    return names


def part_name_key(value: Any) -> str:
    """Build a conservative comparison key while retaining model/spec identifiers."""
    text = clean(value).replace("蓝球", "球").replace("海欧", "海鸥")
    text = re.sub(r"^(菲利普|永久|PHILLIPS|PHILIPS)[\\-－_\\s]*", "", text, flags=re.I)
    text = re.sub(r"(黑色|白色|红色|黄色|蓝色|绿色|银色|金色|橙色|橘色|灰色|浅蓝|浅绿|浅灰|粉色|粉蓝|全黑|黑红|黑蓝|白粉|紫色|咖色)", "", text)
    text = re.sub(r"(\\||\\+|/|\\\\|\\(|\\)|（|）|_|，|,|。|、)", "", text)
    text = re.sub(r"(小头|大头|长款|短款|一套|单支|套装|款|带表|气压表版|电子表|全铝|气压)", "", text)
    text = re.sub(r"\\d+(?:CM|cm|ML|ml|毫安|瓶|只|支|个|对)", "", text)
    return text


def canonical_part_name(value: Any, standard_names: list[str]) -> str:
    """Map a raw part value to one and only one 配件日销表“单选” name.

    Returning an empty string is intentional: an unresolved value must not
    become a new output row outside the single-select standard list.
    """
    source = clean(value)
    if not source or not standard_names:
        return ""
    # Exact and prefix matches retain color/size-specific reference names.
    exact = [name for name in standard_names if source == name]
    if exact:
        return exact[0]
    source_no_brand = re.sub(r"^(菲利普|永久|PHILLIPS|PHILIPS)[\\-－_\\s]*", "", source, flags=re.I)
    exact = [name for name in standard_names if source_no_brand == name]
    if exact:
        return exact[0]
    pref = [name for name in standard_names if source_no_brand.startswith(name + "|") or source_no_brand.startswith(name + "(")]
    if pref:
        return max(pref, key=len)
    source_key = part_name_key(source)
    candidates = [name for name in standard_names if part_name_key(name) and part_name_key(name) in source_key]
    if candidates:
        ranked = sorted(candidates, key=lambda name: (len(part_name_key(name)), len(name)), reverse=True)
        if len(ranked) == 1 or (
            len(part_name_key(ranked[0])) > len(part_name_key(ranked[1]))
        ):
            return ranked[0]
    # Known word-order variants in the reference table.
    aliases = {
        "挡泥板系列飞龙26A铁卡": "飞龙26A铁卡挡泥板",
        "挡泥板系列勇士210": "挡泥板勇士系列",
        "挡泥板系列勇士300": "挡泥板勇士系列",
        "挡泥板系列勇士900": "挡泥板勇士系列",
        "挡泥板系列勇士310": "挡泥板勇士系列",
        "彩色挡泥板": "彩挡泥板",
    }
    alias = aliases.get(source, "")
    return alias if alias in standard_names else ""


def selected_records(records: Iterable[dict[str, Any]], date_column: str, report_month: str, year: int) -> list[dict[str, Any]]:
    return [record for record in records if parse_month(record.get(date_column), year) == report_month]


def is_signed_return(record: dict[str, Any]) -> bool:
    """Return whether a return is counted as signed, excluding non-receipts."""
    reason = clean(record.get("退货退款原因"))
    return not any(
        keyword in reason for keyword in ("截回", "拦截", "拒收")
    )


def collect_wanli(
    records: Iterable[dict[str, Any]],
    report_month: str,
    year: int,
    registration_dates: set[date],
):
    """仅归集补发登记表中出现过的同一登记日期，再按组和工厂汇总数量。"""
    groups: defaultdict[str, Decimal] = defaultdict(Decimal)
    factories: defaultdict[str, Decimal] = defaultdict(Decimal)
    for record in records:
        raw_date = record.get("日期")
        day = parse_concrete_date(raw_date, year)
        if day is None:
            record_missing("万里牛补发.xlsx", "存在无法识别日期的记录，未计入")
            continue
        if day.strftime("%Y-%m") != report_month or day not in registration_dates:
            continue
        raw_quantity = clean(record.get("数量"))
        try:
            quantity = Decimal(raw_quantity.replace(",", ""))
            if not quantity.is_finite() or quantity < 0 or quantity != quantity.to_integral_value():
                raise ValueError("数量须为非负整数")
        except (InvalidOperation, ValueError):
            record_missing("万里牛补发.xlsx", "存在无效数量，未计入")
            continue
        group = clean(record.get("售后组"))
        factory = (clean(record.get("工厂匹配")) or clean(record.get("工厂"))
                   or clean(record.get("仓库")) or clean(record.get("补发仓库")))
        if group:
            groups[group] += quantity
        elif quantity:
            record_missing("万里牛补发.xlsx", "存在未归属售后组的数量，未计入总表及小组大组")
        if factory:
            factories[factory] += quantity
        elif quantity:
            record_missing("万里牛补发.xlsx", "存在未归属工厂的数量，未计入工厂表")
    return groups, factories


def annotate_new_records(reports, template_dir: Path | None) -> None:
    """模板只作为新增提示基准，不据此删除已按分表算法生成的行。"""
    if template_dir is None:
        return
    schemas = [
        ("总数据计算.xlsx", "已确认登记表无误"),
        ("小组统计表.xlsx", "售后组"), ("大组统计表.xlsx", "售后组"),
        ("工厂统计表.xlsx", "工厂"), ("车型统计表.xlsx", "车型"),
        ("配件统计表.xlsx", "配件型号名称"), ("店铺统计表.xlsx", "店铺名称"),
    ]
    for rows, (filename, key_field) in zip(reports, schemas):
        path = template_dir / filename
        if not path.is_file() or filename == "总数据计算.xlsx":
            continue
        workbook = load_workbook(path, read_only=True, data_only=True)
        try:
            existing = {clean(r[0]) for r in workbook.active.iter_rows(min_row=2, values_only=True)
                        if r and clean(r[0])}
        finally:
            workbook.close()
        if filename == "店铺统计表.xlsx":
            existing = {canonical_shop_name(name) for name in existing}
        if not existing:
            continue  # 空模板没有可比较的历史名单，不能把所有行称为新增。
        for row in rows:
            key = clean(row.get(key_field))
            if not key or key == "数据缺失提示" or key in existing:
                continue
            locations = []
            # 只引用已经读取的原始记录，不读取答案值填充计算结果。
            for source, records in RAW_RECORD_CACHE.items():
                for record in records:
                    if any(clean(v) == key for k, v in record.items() if not k.startswith("__")):
                        location = f"{source.name}第{record.get('__source_row__', '?')}行"
                        if location not in locations:
                            locations.append(location)
            note = "新增记录（相对模板名单）：按本表算法及原始数据计算，请核实"
            if locations:
                note += "；来源：" + "、".join(locations[:5])
                if len(locations) > 5:
                    note += f"等{len(locations)}行"
            else:
                note += "；由原始编码/组织映射归集，请核实对应关系"
            if filename == "店铺统计表.xlsx":
                note += "；模板外门店的补发仍按现行模板准入规则排除，其他指标独立计算"
            row["备注"] = "；".join(x for x in [clean(row.get("备注")), note] if x)


def match_return_part(record: dict[str, Any], code_titles, title_parts, standard_names) -> str:
    """Match return identifiers exactly; conflicting mappings stay unresolved."""
    candidates = set()
    def add_title(value):
        text = clean(value)
        if text in standard_names:
            candidates.add(text)
        candidates.update(title_parts.get(text, set()))
    for field in ("车型英文编码", "英文编码整合", "英文编码"):
        for title in code_titles.get(clean(record.get(field)), set()):
            add_title(title)
    for field in ("车型", "中文编码"):
        add_title(record.get(field))
    return next(iter(candidates)) if len(candidates) == 1 else ""


PERSONAL_HEADERS = [
    "姓名", "售后组", "统计月份", "数据登记日期",
    "补发登记次数", "补发涉及订单数", "重复补发订单数", "配件店铺补发次数",
    "退货登记次数", "退货涉及订单数", "签收退货单量", "拒收／拦截单量",
    "质量问题退货次数", "仓库／物流原因退货次数", "买家原因退货次数", "其他／未明原因退货次数",
    "补偿登记次数", "补偿涉及订单数", "补偿总金额", "每单平均补偿金额", "备注",
]
PERSONAL_DETAIL_HEADERS = [
    "记录ID", "登记人", "原登记人", "店铺", "订单号", "登记日期", "售后类型",
    "原始原因", "具体原因", "补偿金额", "计入口径", "异常提示",
    "责任分类", "责任人", "复核结果", "复核备注", "来源表格", "原始行号",
    "复核状态", "责任订单计数",
]


PERSONAL_HIGH_COMPENSATION_ALERTS = False  # 2026-09-30：暂时关闭高额补偿提示。


def personal_settings(template_dir):
    """模板设置为唯一输入；不擅自推断昵称，也不默认设定扣分阈值。"""
    aliases, threshold, notes = {}, None, []
    path = template_dir / "售后个人数据统计表.xlsx" if template_dir else None
    if not path or not path.is_file():
        return aliases, threshold, notes
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if "考核设置" not in wb.sheetnames:
            return aliases, threshold, notes
        sheet = wb["考核设置"]
        raw = sheet["B2"].value
        if PERSONAL_HIGH_COMPENSATION_ALERTS and clean(raw):
            try:
                threshold = Decimal(str(raw))
                if not threshold.is_finite() or threshold <= 0:
                    raise ValueError("高额补偿阈值必须为正数")
            except (InvalidOperation, ValueError):
                threshold = None
                notes.append("高额补偿阈值无效，未启用识别")
        for row in sheet.iter_rows(min_row=7, max_col=3, values_only=True):
            alias, name, confirmed = map(clean, row)
            if confirmed != "是":
                continue
            if not alias or not name:
                raise ValueError("人员映射已确认但缺少登记名或正式姓名")
            if alias in aliases and aliases[alias] != name:
                raise ValueError(f"人员映射冲突：{alias}")
            aliases[alias] = name
        # 仅支持直接映射，避免多级链式归属产生不一致。
        for alias, name in aliases.items():
            if name in aliases and aliases[name] != name:
                raise ValueError(f"人员映射存在链式或循环关系：{alias}，请直接填写正式姓名")
    finally:
        wb.close()
    return aliases, threshold, notes


def personal_report_rows(raw_dir, template_dir, returns, resends, compensations,
                         part_only_shops, report_month, output_date_text,
                         compensation_source=None):
    """处理量与异常分开；订单量仅对店铺、订单号均完整的记录去重。"""
    from collections import Counter
    import hashlib
    import json

    aliases, threshold, settings_notes = personal_settings(template_dir)
    roster_path = template_dir / "售后分组.xlsx" if template_dir else raw_dir / "售后分组.xlsx"
    if not roster_path.is_file():
        roster_path = raw_dir / "售后分组.xlsx"
    groups = {}
    for record in read_records(roster_path):
        group = clean(record.get("售后组"))
        for member in re.split(r"[,，、;；\n]+", clean(record.get("组员详情"))) + [clean(record.get("组长"))]:
            member = aliases.get(clean(member), clean(member))
            if member and group and group not in groups.setdefault(member, []):
                groups[member].append(group)

    events = []
    occurrences = Counter()
    eligible_comp = {id(r) for r in compensations}
    sources = [("退货", "退货数据.xlsx", returns), ("补发", "补发数据.xlsx", resends),
               ("补偿", "补偿数据.xlsx", compensation_source if compensation_source is not None else compensations)]
    dates = []
    for kind, source, records in sources:
        for raw in records:
            original = clean(raw.get("登记人"))
            attribution_name = original
            attribution_note = ""
            primary_field = {"退货": "实际同意退货售后", "补偿": "实际同意补偿售后"}.get(kind)
            if primary_field:
                attribution_name = clean(raw.get(primary_field))
                if not attribution_name:
                    attribution_name = clean(raw.get("操作售后"))
                    primary_issue = f"{primary_field}为空" if primary_field in raw else f"缺少{primary_field}字段"
                    fallback_issue = "操作售后为空" if "操作售后" in raw else "缺少操作售后字段"
                    attribution_note = (
                        f"{kind}：{primary_issue}，按操作售后归属"
                        if attribution_name else
                        f"{kind}：{primary_issue}，{fallback_issue}，未填写归属人员")
            elif kind == "补发":
                attribution_name = clean(raw.get("操作售后"))
                if not attribution_name:
                    operator_issue = "操作售后为空" if "操作售后" in raw else "缺少操作售后字段"
                    attribution_note = f"补发：{operator_issue}，未填写归属人员"
            person = aliases.get(attribution_name, attribution_name) or "未填写登记人"
            shop = canonical_shop_name(raw.get("店铺"))
            order = clean(raw.get("订单号"))
            day = parse_concrete_date(raw.get("登记时间"), int(report_month[:4]))
            if day:
                dates.append(day)
            key = (shop, order) if shop and order else None
            reason = clean(raw.get("退货退款原因" if kind == "退货" else "原因"))
            fields = ("质量问题原因", "仓库错发原因", "个人具体原因", "售后备注") if kind == "退货" else (
                ("仓库具体问题", "质量具体原因", "工厂具体问题", "补发内容", "补发具体原因", "配件补发原因") if kind == "补发" else ("原因概括",))
            specifics = "；".join(f"{field}：{clean(raw.get(field))}" for field in fields if clean(raw.get(field)))
            included = kind != "补偿" or id(raw) in eligible_comp
            amount = number(raw.get("补偿金额")) if kind == "补偿" else Decimal(0)
            identity = [kind, original, shop, order, str(raw.get("登记时间")), reason, specifics, str(amount)]
            digest = hashlib.sha256(json.dumps(identity, ensure_ascii=False).encode("utf-8")).hexdigest()[:24]
            occurrences[digest] += 1
            events.append({"person": person, "key": key, "kind": kind, "included": included,
                           "归属原名": attribution_name, "归属说明": attribution_note,
                           "amount": amount, "raw": raw, "source": source,
                           "记录ID": f"{digest}-{occurrences[digest]}", "登记人": person, "原登记人": original,
                           "店铺": shop, "订单号": order, "登记日期": day.isoformat() if day else "",
                           "售后类型": kind, "原始原因": reason, "具体原因": specifics,
                           "补偿金额": excel_number(amount) if kind == "补偿" else None,
                           "计入口径": "计入" if included else "未计入",
                           "来源表格": resolve_source_path(raw_dir / source).name,
                           "原始行号": raw.get("__source_row__")})

    order_counts = Counter((e["kind"], e["key"]) for e in events if e["key"] and e["included"])
    order_people = defaultdict(set)
    order_money = defaultdict(Decimal)
    for event in events:
        if event["key"] and event["included"]:
            order_people[event["key"]].add(event["person"])
            if event["kind"] == "补偿":
                order_money[event["key"]] += event["amount"]
    people = {name: [] for name in groups}
    for event in events:
        people.setdefault(event["person"], []).append(event)

    rows = []
    # 旧报表可能仍在原始目录查找分组，不能把它的缺失告警误报给已读取模板名单的个人表。
    source_note = missing_note("退货数据.xlsx", "补发数据.xlsx", "补偿数据.xlsx")
    if not groups:
        source_note = "；".join(n for n in (source_note, missing_note("售后分组.xlsx"), "未读取到有效人员分组") if n)
    for person, own in people.items():
        by_kind = {kind: [e for e in own if e["kind"] == kind and e["included"]] for kind in ("退货", "补发", "补偿")}
        keys = {kind: {e["key"] for e in items if e["key"]} for kind, items in by_kind.items()}
        compensation = sum((e["amount"] for e in by_kind["补偿"]), Decimal(0))
        valid_money = sum((e["amount"] for e in by_kind["补偿"] if e["key"]), Decimal(0))
        excluded = [e for e in own if e["kind"] == "补偿" and not e["included"]]
        signed = sum(is_signed_return(e["raw"]) for e in by_kind["退货"])
        part_resends = sum(e["店铺"] in part_only_shops for e in by_kind["补发"])
        categories = Counter()
        for event in by_kind["退货"]:
            reason = event["原始原因"]
            category = ("质量问题退货次数" if reason == "质量问题" else
                        "买家原因退货次数" if reason in {"个人原因", "7天无理由"} else
                        "仓库／物流原因退货次数" if any(k in reason for k in ("仓库", "物流", "快递")) else
                        "其他／未明原因退货次数")
            categories[category] += 1
        details = []
        for event in own:
            alerts = []
            if event["归属说明"]:
                alerts.append(event["归属说明"])
            if person not in groups:
                alerts.append("归属人员未匹配名单" if event["归属原名"] else "未填写归属人员")
            if not event["key"]:
                alerts.append("缺少店铺或订单号，未计入去重订单数")
            if not event["原始原因"]:
                alerts.append("原因未填写")
            if event["key"] and event["included"]:
                if event["kind"] in {"补发", "补偿"} and order_counts[(event["kind"], event["key"])] > 1:
                    alerts.append(f"同订单多次{event['kind']}，待核实")
                if len(order_people[event["key"]]) > 1:
                    alerts.append("同订单多人登记，待核实")
                if event["kind"] == "补偿" and threshold is not None and order_money[event["key"]] >= threshold:
                    alerts.append("同订单累计补偿达到高额阈值")
            if not event["included"]:
                flag = clean(event["raw"].get("是否不参与计算判断"))
                alerts.append("不参与计算标记为是" if flag == "是" else "未计入补偿口径，请核实参与标记、工厂及店铺")
            if alerts:
                detail = {field: event.get(field) for field in PERSONAL_DETAIL_HEADERS}
                detail.update({"异常提示": "；".join(alerts), "责任分类": "待核实", "责任人": "",
                               "复核结果": "待复核", "复核备注": ""})
                details.append(detail)
        notes = [source_note] + settings_notes
        attribution_notes = Counter(e["归属说明"] for e in own if e["归属说明"])
        notes.extend(f"{note}（{count}条）" for note, count in attribution_notes.items())
        if person not in groups:
            notes.append("登记名未匹配正式人员，请确认映射")
        missing = sum(not e["key"] for e in own if e["included"])
        if missing:
            notes.append(f"{missing}条记录缺少店铺或订单号：计入登记量及金额，不计去重订单数；平均补偿仅基于有效订单")
        if not own:
            notes.append("本期可用记录中无对应登记")
        row = {field: 0 for field in PERSONAL_HEADERS}
        row.update({"姓名": person, "售后组": "、".join(groups.get(person, [])) or "未匹配分组",
                    "统计月份": report_month, "数据起始日期": min(dates).isoformat() if dates else None,
                    "数据截至日期": max(dates).isoformat() if dates else None,
                    "数据登记日期": output_date_text,
                    "补发登记次数": len(by_kind["补发"]), "补发涉及订单数": len(keys["补发"]),
                    "重复补发订单数": sum(order_counts[("补发", k)] > 1 for k in keys["补发"]),
                    "配件店铺补发次数": part_resends, "退货登记次数": len(by_kind["退货"]),
                    "退货涉及订单数": len(keys["退货"]), "签收退货单量": signed,
                    "拒收／拦截单量": len(by_kind["退货"]) - signed,
                    "补偿登记次数": len(by_kind["补偿"]), "补偿涉及订单数": len(keys["补偿"]),
                    "补偿总金额": excel_number(compensation),
                    "每单平均补偿金额": excel_number(valid_money / len(keys["补偿"])) if keys["补偿"] else None,
                    "多次补偿订单数": sum(order_counts[("补偿", k)] > 1 for k in keys["补偿"]),
                    "高额补偿订单数": sum(order_money[k] >= threshold for k in keys["补偿"]) if threshold is not None else None,
                    "未计入补偿次数": len(excluded), "未计入补偿金额": excel_number(sum((e["amount"] for e in excluded), Decimal(0))),
                    "待核实记录数": len(details), "已确认客服责任单数": 0, "考核备注": "",
                    "备注": "；".join(n for n in notes if n), "_考核明细": details})
        row.update(categories)
        rows.append(row)
    return rows


def write_personal_report(template_path, output_path, rows):
    """填写绩效工作簿，保留同名文件中的人工复核和个人备注。"""
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.formatting.rule import FormulaRule

    previous, comments, archive = {}, {}, []
    if output_path.is_file():
        old = load_workbook(output_path, read_only=True, data_only=False)
        try:
            for title in ("考核异常明细", "复核留存"):
                if title not in old.sheetnames:
                    continue
                data = old[title].iter_rows(values_only=True)
                headers = next(data)
                for values in data:
                    record = dict(zip(headers, values))
                    if record.get("记录ID"):
                        previous[record["记录ID"]] = record
            if "个人绩效汇总" in old.sheetnames:
                data = old["个人绩效汇总"].iter_rows(values_only=True)
                headers = next(data)
                for values in data:
                    record = dict(zip(headers, values))
                    if record.get("姓名"):
                        comments[(record.get("统计月份"), record["姓名"])] = (
                            record.get("考核备注") or record.get("备注") or "")
        finally:
            old.close()

    wb = load_workbook(template_path)
    for title in ("个人绩效汇总", "考核异常明细", "考核设置"):
        if title not in wb.sheetnames:
            raise ValueError(f"个人模板缺少 {title} 工作表，请使用新版绩效模板")
    summary, detail_sheet = wb["个人绩效汇总"], wb["考核异常明细"]
    summary_headers = [clean(summary.cell(1, col).value) for col in range(1, summary.max_column + 1)]
    detail_headers = [clean(detail_sheet.cell(1, col).value) for col in range(1, detail_sheet.max_column + 1)]
    if not summary_headers or any(not header for header in summary_headers):
        raise ValueError("个人绩效汇总模板首行存在空表头")
    if not detail_headers or any(not header for header in detail_headers):
        raise ValueError("考核异常明细模板首行存在空表头")
    missing_summary = set(PERSONAL_HEADERS) - set(summary_headers)
    missing_detail = set(PERSONAL_DETAIL_HEADERS) - set(detail_headers)
    if missing_summary or missing_detail:
        missing = []
        if missing_summary:
            missing.append("个人绩效汇总：" + "、".join(sorted(missing_summary)))
        if missing_detail:
            missing.append("考核异常明细：" + "、".join(sorted(missing_detail)))
        raise ValueError("个人统计模板缺少必要字段：" + "；".join(missing))
    details = [dict(d) for row in rows for d in row.get("_考核明细", [])]
    current = {d["记录ID"] for d in details}
    manual = ("责任分类", "责任人", "复核结果", "复核备注")
    for detail in details:
        if detail["记录ID"] in previous:
            for field in manual:
                detail[field] = previous[detail["记录ID"]].get(field)
    for key, record in previous.items():
        if key not in current and any(clean(record.get(k)) not in {"", "待核实", "待复核"} for k in manual):
            record["异常提示"] = "当前统计未匹配到此记录，仅保留历史复核，不计入本次统计；" + clean(record.get("异常提示"))
            archive.append(record)

    def fill(sheet, headers, records):
        column_count = len(headers)
        for col, header in enumerate(headers, 1):
            if sheet.cell(1, col).value is None:
                sheet.cell(1, col, header)
        body_styles = [copy(sheet.cell(2, col)._style) for col in range(1, column_count + 1)]
        base_height = sheet.row_dimensions[2].height or 15
        # 清空旧模板内容时保留其行列样式、验证和条件格式。
        for row_index in range(2, sheet.max_row + 1):
            for col in range(1, column_count + 1):
                sheet.cell(row_index, col).value = None
        for index, record in enumerate(records, 2):
            for col, header in enumerate(headers, 1):
                cell = sheet.cell(index, col)
                if index > 2:
                    cell._style = copy(body_styles[col - 1])
                value = record.get(header)
                if header in {"数据起始日期", "数据截至日期", "数据登记日期", "登记日期"} and clean(value) and not isinstance(value, (date, datetime)):
                    try:
                        value = date.fromisoformat(clean(value))
                    except ValueError:
                        pass
                cell.value = value
                if isinstance(value, str) and value.startswith(("=", "+", "-", "@")):
                    cell.data_type = "s"
                if cell.number_format in {"General", ""}:
                    if "金额" in header:
                        cell.number_format = "0.00"
                    elif header in {"数据起始日期", "数据截至日期", "数据登记日期", "登记日期"}:
                        cell.number_format = "yyyy-mm-dd"
                    elif header in {"记录ID", "订单号", "原登记人"}:
                        cell.number_format = "@"
            longest = max((len(clean(record.get(header))) / max(
                8, (sheet.column_dimensions[get_column_letter(col)].width
                    or sheet.sheet_format.defaultColWidth or 13) / 2)
                           for col, header in enumerate(headers, 1)), default=1)
            sheet.row_dimensions[index].height = max(base_height, min(409, int(longest + 1) * 15))
        sheet.auto_filter.ref = f"A1:{get_column_letter(column_count)}{max(1, len(records) + 1)}"
        if sheet.freeze_panes is None:
            sheet.freeze_panes = "C2"

    for row in rows:
        note = clean(row.get("备注"))
        comment = clean(comments.get((row.get("统计月份"), row["姓名"])))
        if not PERSONAL_HIGH_COMPENSATION_ALERTS:
            retired_notes = {"高额阈值未启用", "高额补偿阈值无效，未启用识别",
                             "同订单累计补偿达到高额阈值"}
            comment = "；".join(note for note in comment.split("；") if note not in retired_notes)
        if "考核备注" in summary_headers:
            row["考核备注"] = comment
        elif comment and "备注" in summary_headers:
            row["备注"] = "；".join(dict.fromkeys(value for value in (note + "；" + comment).split("；") if value))
    fill(summary, summary_headers, rows)
    fill(detail_sheet, detail_headers, details)
    end = max(2, len(details) + 1)
    detail_col = {header: get_column_letter(index) for index, header in enumerate(detail_headers, 1)}
    required_formula_headers = {"责任分类", "责任人", "复核结果", "店铺", "订单号", "复核状态", "责任订单计数"}
    if required_formula_headers.issubset(detail_col):
        category = detail_col["责任分类"]
        assignee = detail_col["责任人"]
        result = detail_col["复核结果"]
        shop = detail_col["店铺"]
        order = detail_col["订单号"]
        status = detail_col["复核状态"]
        count = detail_col["责任订单计数"]
        for index in range(2, len(details) + 2):
            detail_sheet[f"{status}{index}"] = (
                f'=IF({result}{index}="不成立","不成立",IF(AND({result}{index}="已确认",'
                f'{category}{index}<>"",{category}{index}<>"待核实",'
                f'OR({category}{index}<>"客服责任",AND({assignee}{index}<>"",'
                f'{shop}{index}<>"",{order}{index}<>""))),"已确认","待复核"))')
            detail_sheet[f"{count}{index}"] = (
                f'=IF(AND({status}{index}="已确认",{category}{index}="客服责任"),'
                f'IF(COUNTIFS(${shop}$2:{shop}{index},{shop}{index},'
                f'${order}$2:{order}{index},{order}{index},${assignee}$2:{assignee}{index},{assignee}{index},'
                f'${category}$2:{category}{index},"客服责任",${status}$2:{status}{index},"已确认")=1,1,0),0)')
        for index in range(2, len(rows) + 2):
            if {"待核实记录数", "姓名"}.issubset(summary_headers) and "登记人" in detail_col:
                name_col = get_column_letter(summary_headers.index("姓名") + 1)
                summary.cell(index, summary_headers.index("待核实记录数") + 1,
                             f'=COUNTIFS(\'考核异常明细\'!${detail_col["登记人"]}$2:${detail_col["登记人"]}${end},'
                             f'{name_col}{index},\'考核异常明细\'!${status}$2:${status}${end},"待复核")')
            if {"已确认客服责任单数", "姓名"}.issubset(summary_headers):
                name_col = get_column_letter(summary_headers.index("姓名") + 1)
                summary.cell(index, summary_headers.index("已确认客服责任单数") + 1,
                             f'=SUMIFS(\'考核异常明细\'!${count}$2:${count}${end},'
                             f'\'考核异常明细\'!${assignee}$2:${assignee}${end},{name_col}{index})')
    for field, options in (("责任分类", "待核实,客服责任,商品质量,仓库问题,物流问题,买家原因,其他"),
                           ("复核结果", "待复核,已确认,不成立")):
        if field not in detail_col:
            continue
        column = detail_col[field]
        validation = DataValidation(type="list", formula1='"' + options + '"', allow_blank=True)
        validation.errorTitle = "请选择有效选项"
        validation.error = "请使用下拉列表中的选项"
        validation.showErrorMessage = True
        detail_sheet.add_data_validation(validation)
        validation.add(f"{column}2:{column}{end}")
    if rows and "责任人" in detail_col:
        # 责任人限定为当前正式名单或独立保留的未匹配登记人。
        validation = DataValidation(type="list", formula1=f'INDIRECT("\'个人绩效汇总\'!$A$2:$A${len(rows)+1}")', allow_blank=True)
        validation.showErrorMessage = True
        detail_sheet.add_data_validation(validation)
        validation.add(f"{detail_col['责任人']}2:{detail_col['责任人']}{end}")
    if {"复核结果", "复核状态"}.issubset(detail_col):
        result_col = detail_col["复核结果"]
        status_col = detail_col["复核状态"]
        detail_sheet.conditional_formatting.add(
            f"{result_col}2:{result_col}{end}",
            FormulaRule(formula=[f'${status_col}2="待复核"'], fill=PatternFill("solid", fgColor="FFF2CC")))
    if archive:
        archived = wb.create_sheet("复核留存") if "复核留存" not in wb.sheetnames else wb["复核留存"]
        fill(archived, PERSONAL_DETAIL_HEADERS[:18], archive)
        for col in range(1, 19):
            archived.cell(1, col)._style = copy(detail_sheet.cell(1, col)._style)
            archived.column_dimensions[get_column_letter(col)].width = detail_sheet.column_dimensions[get_column_letter(col)].width
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temporary = output_path.with_name(output_path.stem + ".writing.xlsx")
    try:
        wb.save(temporary)
        temporary.replace(output_path)
    finally:
        wb.close()
        if temporary.exists():
            temporary.unlink()
def aggregate_reports(raw_dir: Path, report_month: str, template_dir: Path | None = None) -> tuple[
    list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]],
    list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]],
]:
    """Build total, group, factory, model, part, and shop report records."""
    reset_missing_inputs()
    report_year = int(report_month[:4])
    compact_month = report_month.replace("-", "")
    shops = read_shop_records(raw_dir / "店铺规范.xlsx")
    factories = read_records(raw_dir / "工厂数据.xlsx")
    return_excluded_shops = {
        "阿里巴巴国际站", "抖音菲利普全致专卖店", "快手PHILLIPS菲利普自行车旗舰店",
        "视频号PHILLIPS菲利普自行车", "京东自营入仓长库龄退回",
    }
    compensation_excluded_shops = {
        "抖音菲利普全致专卖店", "快手PHILLIPS菲利普自行车旗舰店",
        "视频号PHILLIPS菲利普自行车",
    }
    resend_excluded_shops = compensation_excluded_shops
    returns = [row for row in selected_records(read_shop_records(raw_dir / "退货数据.xlsx"), "登记时间", report_month, report_year)
               if clean(row.get("店铺")) not in return_excluded_shops
               and clean(row.get("工厂"))
               and clean(row.get("退货退款原因")) != "京东自营入仓破损退回"]
    # 店铺表退货按门店完全匹配正常归属，不套用总表/小组的门店排除名单。
    # 工厂退货独立筛选，不能复用按门店排除的总表/车型集合。
    returns_factory = [row for row in selected_records(read_shop_records(raw_dir / "退货数据.xlsx"), "登记时间", report_month, report_year)
                       if clean(row.get("工厂"))
                       and clean(row.get("退货退款原因")) != "京东自营入仓破损退回"]
    returns_store = [row for row in selected_records(read_shop_records(raw_dir / "退货数据.xlsx"), "登记时间", report_month, report_year)
                     if clean(row.get("工厂")) and clean(row.get("退货退款原因")) != "京东自营入仓破损退回"]
    # ``补发数据`` and ``补偿数据`` are now separate source tables.  The former
    # provides every resend event; the latter provides only included monetary
    # compensation records.  Do not use the compensation table as a proxy for
    # resend count.
    resend_candidates = selected_records(read_shop_records(raw_dir / "补发数据.xlsx"), "登记时间", report_month, report_year)
    resend_dates = [parse_concrete_date(row.get("登记时间"), report_year) for row in resend_candidates]
    output_date = max((d for d in resend_dates if d), default=None)
    output_date_text = output_date.isoformat() if output_date else report_month
    # 总表、店铺表沿用门店排除规则；小组/大组按有效补发记录正常统计，不排除门店。
    resends = [row for row in resend_candidates
               if clean(row.get("店铺")) not in resend_excluded_shops
               and clean(row.get("工厂"))
               and clean(row.get("原因")) != "自购"
               and not clean(row.get("匹配标签")) == "红色"
               and not row.get("__匹配标签红色__", False)]
    # 小组/大组与店铺表使用完全一致的订单补发数据源及筛选口径
    resends_group = list(resends)
    compensation_source_rows = selected_records(read_shop_records(raw_dir / "补偿数据.xlsx"), "登记时间", report_month, report_year)
    # 工厂表的补偿按工厂独立归属：补偿表中因店铺名称为空而被一般统计排除的
    # 记录，只要工厂和参与计算标记完整，仍可用于对应工厂的补偿金额。
    invalid_compensation_rows = [
        {key: value for key, value in row.items() if key not in {"来源表格", "原始行号", "缺失字段", "备注"}}
        for row in DATA_ENTRY_ERRORS if row.get("来源表格") == "补偿数据.xlsx"
    ]
    factory_compensations = [
        row for row in compensation_source_rows
        if clean(row.get("是否不参与计算判断")) == "否" and clean(row.get("工厂"))
    ]
    compensations = [
        row for row in compensation_source_rows
        if clean(row.get("是否不参与计算判断")) == "否" and clean(row.get("店铺")) not in compensation_excluded_shops and clean(row.get("工厂"))
    ]

    shop_by_name = {clean(row.get("店铺名称")): row for row in shops if clean(row.get("店铺名称"))}
    store_template_shops = set()
    if template_dir:
        store_template_path = template_dir / "店铺统计表.xlsx"
        if store_template_path.is_file():
            template_sheet = load_workbook(store_template_path, read_only=True, data_only=True).active
            store_template_shops = {canonical_shop_name(row[0]) for row in template_sheet.iter_rows(min_row=2, values_only=True) if row and clean(row[0])}
    # 新补发表偶有“自行车/车”字样缺失；只在唯一匹配时修正，避免误归属。
    def shop_key(value: Any) -> str:
        return re.sub(r"自行车|车", "", clean(value))

    shop_name_by_key = {
        key: names[0] for key, names in (
            (key, [name for name in shop_by_name if shop_key(name) == key])
            for key in {shop_key(name) for name in shop_by_name}
        ) if len(names) == 1
    }

    def resolved_shop_name(name: str) -> str:
        # 门店名称严格完全匹配，不进行“自行车/车”等模糊修正
        name = canonical_shop_name(name)
        return name if name in shop_by_name else ""
    # 补发统计仅保留店铺统计模板中登记的门店
    if store_template_shops:
        resends = [row for row in resends if resolved_shop_name(clean(row.get("店铺"))) in store_template_shops]
        resends_group = list(resends)
    # 工厂表独立口径：工厂有值、原因非自购、匹配标签非红色，不受店铺模板名单影响
    resends_factory = [row for row in resend_candidates
                       if clean(row.get("工厂"))
                       and clean(row.get("原因")) != "自购"
                       and clean(row.get("匹配标签")) != "红色"
                       and not row.get("__匹配标签红色__", False)]
    # 仅这两家店铺不计入总表和店铺表；其他店铺沿用既有统计口径。
    excluded_total_store_shops = {"阿里巴巴国际站", "阿里巴巴天津锦瑞运动器材有限公司"}
    part_only_shops = {
        name for name, row in shop_by_name.items() if clean(row.get("只销配件店铺"))
    }

    small: defaultdict[str, dict[str, Decimal]] = defaultdict(metric)
    store_metrics: defaultdict[str, dict[str, Decimal]] = defaultdict(metric)
    total = metric()
    member_group: dict[str, str] = {}
    for record in read_records(raw_dir / "售后分组.xlsx"):
        group_name = clean(record.get("售后组"))
        for member in clean(record.get("组员详情")).replace("，", ",").split(","):
            if member.strip() and group_name:
                member_group[member.strip()] = group_name
    wanli_registration_dates = {
        day for day in (
            parse_concrete_date(record.get("登记时间"), report_year)
            for record in resend_candidates
        ) if day is not None
    }
    wanli_group, wanli_factory = collect_wanli(
        read_records(raw_dir / "万里牛补发.xlsx"), report_month, report_year,
        wanli_registration_dates)
    wanli_total = sum(wanli_group.values(), Decimal("0"))
    # 只有万里牛、没有销售或订单补发的小组也必须输出。
    for group_name in wanli_group:
        small[group_name]
    for row in shops:
        name = clean(row.get("店铺名称"))
        responsible = clean(row.get("实际负责"))
        group = member_group.get(responsible) or clean(row.get("售后组")) or responsible
        revenue = number(row.get("当日真实营业额"))
        orders = number(row.get("车子销量"))
        part_orders = number(row.get("配件销量")) if name in part_only_shops else Decimal("0")
        if name and name not in excluded_total_store_shops:
            add_metrics(total, revenue=revenue, orders=orders, part_orders=part_orders)
        if name:
            add_metrics(store_metrics[name], revenue=revenue, orders=orders, part_orders=part_orders)
        if group and name and name not in excluded_total_store_shops:
            add_metrics(small[group], revenue=revenue, orders=orders, part_orders=part_orders)

    for row in returns:
        name = resolved_shop_name(clean(row.get("店铺")))
        group = clean(row.get("售后组"))
        signed = Decimal("1") if is_signed_return(row) else Decimal("0")
        if name and name not in excluded_total_store_shops:
            add_metrics(total, returns=Decimal("1"), signed_returns=signed)
        if group and name and name not in excluded_total_store_shops:
            add_metrics(small[group], returns=Decimal("1"), signed_returns=signed)

    for row in returns_store:
        name = resolved_shop_name(clean(row.get("店铺")))
        if name:
            signed = Decimal("1") if is_signed_return(row) else Decimal("0")
            add_metrics(store_metrics[name], returns=Decimal("1"), signed_returns=signed)

    for row in resends:
        name = resolved_shop_name(clean(row.get("店铺")))
        shop_info = shop_by_name.get(name, {})
        group = member_group.get(clean(row.get("登记人"))) or clean(row.get("售后组")) or clean(shop_info.get("实际负责")) or clean(shop_info.get("售后组"))
        part_resend = Decimal("1") if resolved_shop_name(clean(row.get("店铺"))) in part_only_shops else Decimal("0")
        if name and name not in excluded_total_store_shops and group:
            add_metrics(total, resends=Decimal("1"), part_resends=part_resend)
        if name:
            add_metrics(store_metrics[name], resends=Decimal("1"), part_resends=part_resend)
    for row in resends_group:
        name = resolved_shop_name(clean(row.get("店铺")))
        group = member_group.get(clean(row.get("登记人"))) or clean(row.get("售后组")) or clean(shop_by_name.get(name, {}).get("实际负责")) or clean(shop_by_name.get(name, {}).get("售后组"))
        part_resend = Decimal("1") if resolved_shop_name(clean(row.get("店铺"))) in part_only_shops else Decimal("0")
        if group:
            add_metrics(small[group], resends=Decimal("1"), part_resends=part_resend)

    for row in compensations:
        name = resolved_shop_name(clean(row.get("店铺")))
        group = clean(row.get("售后组")) or clean(shop_by_name.get(name, {}).get("实际负责")) or clean(shop_by_name.get(name, {}).get("售后组"))
        compensation = number(row.get("补偿金额"))
        if name and name not in excluded_total_store_shops:
            add_metrics(total, compensation=compensation)
        if name:
            add_metrics(store_metrics[name], compensation=compensation)
        if group and name and name not in excluded_total_store_shops:
            add_metrics(small[group], compensation=compensation)

    large: defaultdict[str, dict[str, Decimal]] = defaultdict(metric)
    for group, values in small.items():
        add_metrics(large[large_group_name(group)], **values)

    def group_row(group: str, values: dict[str, Decimal], include_large: bool) -> dict[str, Any]:
        wanli_count = wanli_group.get(group, Decimal("0")) if include_large else sum(
            (quantity for small_group, quantity in wanli_group.items()
             if large_group_name(small_group) == group), Decimal("0")
        )
        order_resends = values["resends"] - values["part_resends"]
        combined_resends = order_resends + wanli_count
        base = {
            "售后组": group,
            "营业额": excel_number(values["revenue"]),
            "订单量": excel_number(values["orders"]),
            "配件店铺订单量": excel_number(values["part_orders"]),
            "退货单量": excel_number(values["returns"]),
            "店铺退货": excel_number(values["returns"]),
            "签收退货单量": excel_number(values["signed_returns"]),
            "补发次数": excel_number(combined_resends),
            "配件店铺补发次数": excel_number(values["part_resends"]),
            "配件补发次数": excel_number(values["part_resends"]),
            "补偿金额": excel_number(values["compensation"]),
            "退货率": ratio(values["returns"], values["orders"]),
            "签收退货率": ratio(values["signed_returns"], values["orders"]),
            "补偿率": ratio(values["compensation"], values["revenue"]),
            "补发率": ratio(combined_resends, values["orders"]),
            "配件店铺补发率": ratio(values["part_resends"], values["part_orders"]),
            # 新版小组/大组模板使用“配件补发率”表头，保留旧别名兼容旧模板。
            "配件补发率": ratio(values["part_resends"], values["part_orders"]),
            "订单补发": excel_number(order_resends),
            "万里牛补发": excel_number(wanli_count),
            "售后率": ratio(values["returns"] + combined_resends, values["orders"]),
            "日期": output_date_text,
            "数据月份": report_month,
            "月份": compact_month,
            "月份判断": compact_month,
            "判断月份": compact_month,
        }
        if include_large:
            base["大组"] = large_group_name(group)
        return base

    small_rows = [group_row(group, values, True) for group, values in small.items()]
    large_rows = [group_row(group, values, False) for group, values in large.items()]
    small_rows.sort(key=lambda row: (-number(row["营业额"]), clean(row["售后组"])))
    large_rows.sort(key=lambda row: (-number(row["营业额"]), clean(row["售后组"])))

    # 工厂表严格以工厂数据中的名单为准，保留名单内零值工厂。
    factory_metrics: defaultdict[str, dict[str, Decimal]] = defaultdict(metric)
    factory_names: list[str] = []
    factory_order_excluded = {"河北仓库", "永康云仓"}
    for row in factories:
        name = clean(row.get("工厂"))
        if not name:
            continue
        if name not in factory_metrics:
            factory_names.append(name)
        add_metrics(
            factory_metrics[name],
            # 优先使用原始字段，避免引用源表已处理的汇总列。
            revenue=number(row.get("营业额", row.get("营业额汇总"))),
            # 河北仓库、永康云仓只保留营业额，不把订单量计入工厂表。
            orders=Decimal("0") if name in factory_order_excluded else number(row.get("订单量", row.get("订单量汇总"))),
        )
    for row in returns_factory:
        name = clean(row.get("工厂"))
        if name in factory_metrics:
            add_metrics(factory_metrics[name], returns=Decimal("1"))
    for row in resends_factory:
        name = clean(row.get("工厂匹配")) or clean(row.get("工厂"))
        if name in factory_metrics:
            add_metrics(factory_metrics[name], resends=Decimal("1"))
    for row in factory_compensations:
        name = clean(row.get("工厂"))
        if name in factory_metrics:
            add_metrics(factory_metrics[name], compensation=number(row.get("补偿金额")))
    for factory, quantity in wanli_factory.items():
        if factory not in factory_metrics and quantity:
            record_missing("万里牛补发.xlsx", f"工厂名单外：{factory}，数量{quantity}未计入工厂表")
    factory_rows = []
    for factory in factory_names:
        values = factory_metrics[factory]
        factory_rows.append({
            "工厂名称": factory,
            "工厂": factory,
            "营业额": excel_number(values["revenue"]),
            "订单量": excel_number(values["orders"]),
            "退货单量": excel_number(values["returns"]),
            "补发次数": excel_number(values["resends"]),
            "总补发次数": excel_number(values["resends"] + wanli_factory.get(factory, Decimal("0"))),
            "补发次数 （除万里牛）": excel_number(values["resends"]),
            "补偿金额": excel_number(values["compensation"]),
            "退货率": ratio(values["returns"], values["orders"]),
            "补偿率": ratio(values["compensation"], values["revenue"]),
            # 规范待确认：保留现有订单补发分子，不擅自改为含万里牛。
            "补发率": ratio(values["resends"], values["orders"]),
            "万里牛补发": excel_number(wanli_factory.get(factory, Decimal("0"))),
            "日期": output_date_text,
            "月份": compact_month,
        })

    maps, known_by_brand = code_maps(raw_dir)
    known_vehicle_models = set().union(*known_by_brand.values())
    vehicle_metrics: defaultdict[str, dict[str, Decimal]] = defaultdict(metric)
    sales_specs = [
        # 永久日销的“销售额（汇总）”已逐行取整。必须汇总原始“销售额”后
        # 再统一四舍五入，避免车型层累计产生取整偏差。
        ("永久日销.xlsx", "编码表车型", "销量", "销售额", "销量", "销售额"),
        ("菲利普日销.xlsx", "型号", "销量", "销售额", "销量", "销售额"),
        ("大众日销.xlsx", "车型", "销量", "销售额", "销量", "销售额"),
        ("玛莎拉蒂日销.xlsx", "车型", "汇总", "销售额", "汇总", "销售额"),
    ]
    for filename, model_column, quantity_column, revenue_column, fallback_quantity, fallback_revenue in sales_specs:
        for row in read_records(raw_dir / filename):
            model = clean(row.get(model_column))
            if not model:
                continue
            quantity = number(row.get(quantity_column, row.get(fallback_quantity)))
            revenue = number(row.get(revenue_column, row.get(fallback_revenue)))
            if quantity or revenue:
                add_metrics(vehicle_metrics[model], orders=quantity, revenue=revenue)
                known_vehicle_models.add(model)

    unmatched_vehicle_returns = []
    for row in returns:
        brand = clean(row.get("品牌"))
        # 退货登记表使用统一字段“车型英文编码/中文编码”。兼容旧导出中
        # 按品牌拆分的字段，并优先使用已登记且已存在于车型编码表的“车型”。
        english_code = clean(row.get("车型英文编码")) or clean(row.get("英文编码整合"))
        chinese_code = clean(row.get("中文编码"))
        named_model = clean(row.get("车型"))
        model = named_model if named_model in known_vehicle_models else ""
        if brand == "永久":
            model = model or maps["永久_英文"].get(english_code, "") or maps["永久_英文"].get(clean(row.get("永久车型英文编码")), "") or maps["永久_中文"].get(chinese_code, "") or maps["永久_中文"].get(clean(row.get("永久中文编码")), "")
        elif brand == "菲利普":
            model = model or maps["菲利普_英文"].get(english_code, "") or maps["菲利普_英文"].get(clean(row.get("菲利普车型英文编码")), "") or maps["菲利普_中文"].get(chinese_code, "") or maps["菲利普_中文"].get(clean(row.get("菲利普中文编码")), "")
        elif brand == "大众":
            model = model or maps["大众_中文"].get(chinese_code, "") or maps["大众_中文"].get(clean(row.get("大众中文编码")), "")
        if model:
            add_metrics(vehicle_metrics[model], returns=Decimal("1"))
        else:
            unmatched_vehicle_returns.append(row)

    part_code_map = {
        clean(row[1]): clean(row[0])
        for row in read_tuple_rows(raw_dir / "配件编码.xlsx")
        if len(row) >= 2 and clean(row[0]) and clean(row[1])
    }
    part_metrics: defaultdict[str, dict[str, Decimal]] = defaultdict(metric)
    standard_part_names = part_standard_names(raw_dir, template_dir)
    # “单选”是唯一标准名称。标题仅作为 after-sales 明细反查日销归类的键。
    part_name_by_title: dict[str, str] = {}
    for row in read_records(raw_dir / "配件日销.xlsx"):
        title = clean(row.get("标题"))
        part = canonical_part_name(row.get("单选"), standard_part_names)
        quantity = number(row.get("销量"))
        revenue = number(row.get("销售额"))
        if title and part:
            part_name_by_title[title] = part
        if part and (quantity or revenue):
            add_metrics(part_metrics[part], orders=quantity, revenue=revenue)
        elif title and not part:
            record_missing("配件日销.xlsx", "存在“单选”为空或不在标准名称集合的记录，未计入")

    # Return mapping uses exact, unique code/title -> 单选 relationships.
    # Keep all candidates instead of silently overwriting duplicate mappings.
    return_code_titles = defaultdict(set)
    for code_row in read_tuple_rows(raw_dir / "配件编码.xlsx"):
        if len(code_row) >= 2 and clean(code_row[0]) and clean(code_row[1]):
            return_code_titles[clean(code_row[1])].add(clean(code_row[0]))
    return_title_parts = defaultdict(set)
    for sales_row in read_records(raw_dir / "配件日销.xlsx"):
        title, part = clean(sales_row.get("标题")), clean(sales_row.get("单选"))
        if title and part in standard_part_names:
            return_title_parts[title].add(part)
    for row in unmatched_vehicle_returns:
        part = match_return_part(row, return_code_titles, return_title_parts, standard_part_names)
        if part:
            add_metrics(part_metrics[part], returns=Decimal("1"))
        else:
            record_missing("退货数据.xlsx", f"第{row.get('__source_row__', '?')}行未匹配车型或唯一配件标准名，未计入车型/配件退货")

    model_by_english = {
        code: model for key, mapping in maps.items() if key.endswith("_英文") for code, model in mapping.items()
    }

    def add_after_sales_detail(row: dict[str, Any], *, resend: bool = False, compensation: bool = False) -> None:
        """Allocate a resend/compensation record to a vehicle or accessory."""
        named_model = clean(row.get("车型"))
        english_code = clean(row.get("英文编码整合")) or clean(row.get("英文编码"))
        vehicle_model = named_model if named_model in known_vehicle_models else model_by_english.get(english_code, "")
        amount = number(row.get("补偿金额")) if compensation else Decimal("0")
        if vehicle_model:
            add_metrics(vehicle_metrics[vehicle_model], **({"resends": Decimal("1")} if resend else {}), **({"compensation": amount} if compensation else {}))
            return
        part_source = part_code_map.get(english_code, "") or named_model or clean(row.get("中文编码"))
        part = canonical_part_name(part_name_by_title.get(part_source, part_source), standard_part_names)
        if not part:
            record_missing("配件日销.xlsx", "存在补发或补偿配件无法按“单选”对应，未计入")
            return
        add_metrics(part_metrics[part], **({"resends": Decimal("1")} if resend else {}), **({"compensation": amount} if compensation else {}))

    # 车型/配件表补发采用最新有效口径，与工厂表一致，不受店铺名单过滤
    for row in resends_factory:
        add_after_sales_detail(row, resend=True)
    for row in compensations:
        add_after_sales_detail(row, compensation=True)

    def detail_row(name: str, values: dict[str, Decimal], label: str) -> dict[str, Any]:
        return {
            "车型": name,
            "型号": name,
            "配件型号名称": name,
            "销售额": excel_number(values["revenue"]),
            "营业额": excel_number(values["revenue"]),
            "订单量": excel_number(values["orders"]),
            "配件订单量": excel_number(values["orders"]),
            "退货单量": excel_number(values["returns"]),
            "退货量": excel_number(values["returns"]),
            "补发次数": excel_number(values["resends"]),
            "补发次数量": excel_number(values["resends"]),
            "补偿金额": excel_number(values["compensation"]),
            "补偿金额量": excel_number(values["compensation"]),
            "退货率": ratio(values["returns"], values["orders"]),
            "补偿率": ratio(values["compensation"], values["revenue"]),
            "补发率": ratio(values["resends"], values["orders"]),
            "日期": output_date_text,
            "月份": report_month,
        }

    vehicle_rows = [detail_row(name, values, "车型") for name, values in vehicle_metrics.items() if any(values.values())]
    part_rows = [detail_row(name, values, "配件") for name, values in part_metrics.items() if any(values.values())]
    vehicle_rows.sort(key=lambda row: (-number(row["营业额"]), clean(row["车型"])))
    part_rows.sort(key=lambda row: (-number(row["销售额"]), clean(row["配件型号名称"])))

    store_rows = []
    for name, shop in shop_by_name.items():
        # 用户确认：仅店铺表不统计该店，其他表继续使用各自口径。
        if name == "阿里巴巴天津锦瑞运动器材有限公司":
            continue
        values = store_metrics[name]
        group = clean(shop.get("实际负责")) or clean(shop.get("售后组"))
        store_rows.append({
            "店铺名": name,
            "店铺名称": name,
            "售后组": group,
            "大组": large_group_name(group),
            "运营组": clean(shop.get("运营组")),
            "配件组": clean(shop.get("配件组")),
            "平台": clean(shop.get("平台")),
            "只销配件店铺": clean(shop.get("只销配件店铺")),
            "营业额": excel_number(values["revenue"]),
            "订单量": excel_number(values["orders"]),
            "配件店铺订单量": excel_number(values["part_orders"]),
            "退货单量": excel_number(values["returns"]),
            "签收退货单量": excel_number(values["signed_returns"]),
            # 店铺表的“补发次数”不含只销配件店补发；配件补发单列统计。
            "补发次数": excel_number(values["resends"] - values["part_resends"]),
            "配件店铺补发次数": excel_number(values["part_resends"]),
            "补偿金额": excel_number(values["compensation"]),
            "退货率": ratio(values["returns"], values["orders"]),
            "签收退货率": ratio(values["signed_returns"], values["orders"]),
            "补偿率": ratio(values["compensation"], values["revenue"]),
            "补发率": ratio(values["resends"] - values["part_resends"], values["orders"]),
            "配件店铺补发率": ratio(values["part_resends"], values["part_orders"]),
            "日期": output_date_text,
            "判断月份": compact_month,
            "售后率": ratio(values["returns"] + values["resends"] - values["part_resends"], values["orders"]),
        })
    store_rows.sort(key=lambda row: (-number(row["营业额"]), clean(row["店铺名称"])))

    total_row = [{
        "已确认登记表无误": "已确认",
        "营业额": excel_number(total["revenue"]),
        "订单量": excel_number(total["orders"]),
        "退货单量": excel_number(total["returns"]),
        "补发次数": excel_number(total["resends"] + wanli_total),
        "补偿金额": excel_number(total["compensation"]),
        "退货率": ratio(total["returns"], total["orders"]),
        "补偿率": ratio(total["compensation"], total["revenue"]),
        "补发率": ratio(total["resends"] + wanli_total, total["orders"]),
        "万里牛补发": excel_number(wanli_total),
        "订单补发": excel_number(total["resends"]),
        "日期": output_date_text,
        "数据日期": output_date_text,
    }]

    reports = (total_row, small_rows, large_rows, factory_rows, vehicle_rows, part_rows, store_rows)
    for report_name, rows in zip((*TEMPLATE_FILES.keys(), "店铺数据计算"), reports, strict=True):
        note = missing_note(*REPORT_SOURCES[report_name])
        for row in rows:
            row["备注"] = note
            if row.get("店铺名称") == "快手PHILLIPS户外装备旗舰店":
                row["备注"] = "；".join(x for x in (note, "同店合并：快手PHILLIPS户外装备旗舰店（已关店）；原始名称保留在源表") if x)
        # 当依赖源全部缺失而没有可聚合的明细时，保留一行可见提示，避免输出空表。
        if not rows and note:
            key = {"小组数据计算": "售后组", "大组数据计算": "售后组", "工厂数据计算": "工厂名称", "车型数据计算": "车型", "配件数据计算": "配件型号名称", "店铺数据计算": "店铺名称"}[report_name]
            rows.append({key: "数据缺失提示", "备注": note})
    annotate_new_records(reports, template_dir)
    personal_rows = personal_report_rows(
        raw_dir, template_dir, returns, resends_group, compensations,
        part_only_shops, report_month, output_date_text, compensation_source_rows)
    return (*reports, personal_rows)


def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="生成售后月度数据计算表")
    parser.add_argument("--month", default=datetime.now().strftime("%Y-%m"), help="统计月份，格式 YYYY-MM")
    parser.add_argument("--created-date", default=datetime.now().strftime("%Y%m%d"), help="文件创建日期，格式 YYYYMMDD")
    parser.add_argument("--root", type=Path, default=ROOT, help="包含原始数据、模板和输出目录的根目录")
    parser.add_argument("--raw-dir", type=Path, default=None, help="可选：单独指定原始数据目录，支持带日期后缀的目录")
    args = parser.parse_args()
    if not re.fullmatch(r"20\d{2}-(0[1-9]|1[0-2])", args.month):
        parser.error("--month 必须是 YYYY-MM，例如 2026-09")
    if not re.fullmatch(r"20\d{2}(0[1-9]|1[0-2])([0-2]\d|3[01])", args.created_date):
        parser.error("--created-date 必须是 YYYYMMDD，例如 20260910")
    return args


def main() -> None:
    args = parse_arguments()
    root = args.root.resolve()
    raw_dir = args.raw_dir.resolve() if args.raw_dir else root / RAW_DIR_NAME
    if not raw_dir.is_dir():
        # 原始数据目录按关键字识别，不要求目录名完全等于“原始数据”。
        candidates = [p for p in root.iterdir() if p.is_dir() and ("原始数据" in p.name or "原始" in p.name)]
        dated_raw_dirs = sorted(candidates, key=lambda p: p.name, reverse=True)
        if dated_raw_dirs:
            raw_dir = dated_raw_dirs[0]
    template_dir = root / TEMPLATE_DIR_NAME
    output_dir = root / OUTPUT_DIR_NAME
    if not raw_dir.is_dir() or not template_dir.is_dir():
        raise FileNotFoundError("根目录下必须包含 原始数据 和 数据表模板 两个文件夹")

    reports = aggregate_reports(raw_dir, args.month, template_dir)
    report_rows = dict(zip(TEMPLATE_FILES, reports[:6], strict=True))
    for report_name, rows in report_rows.items():
        template_path = template_dir / TEMPLATE_FILES[report_name]
        if not template_path.is_file():
            raise FileNotFoundError(f"缺少模板：{template_path}")
        output_path = output_dir / f"{OUTPUT_FILE_STEMS[report_name]}_{args.created_date}.xlsx"
        write_template_rows(template_path, output_path, rows)
        print(f"已生成 {output_path.name}：{len(rows)} 行")

    store_path = output_dir / f"{OUTPUT_FILE_STEMS['店铺数据计算']}_{args.created_date}.xlsx"
    write_template_rows(template_dir / "店铺统计表.xlsx", store_path, reports[6])
    print(f"已生成 {store_path.name}：{len(reports[6])} 行")
    personal_path = output_dir / f"售后个人数据统计表_{args.created_date}.xlsx"
    write_personal_report(template_dir / "售后个人数据统计表.xlsx", personal_path, reports[7])
    print(f"已生成 {personal_path.name}：{len(reports[7])} 行")

    error_path = output_dir / "数据登记错误.xlsx"
    write_data_entry_errors(error_path)
    print(f"已生成 {error_path.name}：{len(DATA_ENTRY_ERRORS)} 行")
    failed_marks = apply_error_marks(raw_dir)
    if failed_marks:
        print("以下原始表被占用，已完成计算和错误清单导出，待关闭后再次运行以标红：" + "、".join(failed_marks))


if __name__ == "__main__":
    main()

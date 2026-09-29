# -*- coding: utf-8 -*-
"""
影刀 RPA · Python 模块 —— 售后月度数据计算

【用途】
读取「原始数据」文件夹下的原始表，按「数据表模板」的模板表头生成 7 张月度计算表，
输出到「数据计算」文件夹。业务口径与原脚本 generate_monthly_reports.py 完全一致，
仅做"影刀化"改造：去掉命令行参数、接入影刀日志、参数/返回值适配「调用模块」。

【影刀接入步骤】
1. 打开影刀设计器 → 新建应用 → 右上角「新建 Python 模块」，模块名建议 report。
2. 将本文件全部内容覆盖粘贴进该模块并保存。
3. 右侧「更多」→「Python 包管理」，确认 openpyxl 已安装（影刀多数版本已内置）；
   若未安装，在其中执行 pip 安装 openpyxl。
4. 在可视化流程中插入「调用模块」指令，选择本模块的 generate_report 函数，
   按顺序绑定下列 4 个参数（前 3 个为文本，最后 1 个为布尔）：
   1）month        文本，统计月份，格式 YYYY-MM，例如 2026-09；留空则取当前月。
   2）created_date 文本，输出文件名里的创建日期，格式 YYYYMMDD，例如 20260910；留空则取当天。
   3）root         文本，数据根目录，即同时包含「原始数据」「数据表模板」「数据计算」的文件夹；
                   留空则自动探测（模块所在目录 → 当前工作目录 → 桌面\售后数据计算）。
   4）overwrite    布尔，同名文件是否覆盖；建议传 True。留空按 True 处理。
5. 「调用模块」的返回值绑定到一个流程变量（如 结果），再用「打印日志」输出，
   或在后续指令中读取 结果["message"] / 结果["files"] / 结果["file_count"]。

【返回值结构】（字典，便于在影刀里直接取值）
- ok          是否成功，布尔
- month       实际使用的统计月份
- created_date 实际使用的创建日期
- output_dir  输出目录绝对路径
- files       本次生成的文件绝对路径列表
- file_count  生成的文件数量
- row_count   生成的数据总行数
- message     一句话结果摘要，可直接放进「打印日志」

【运行前提】
- root 目录下必须同时存在「原始数据」和「数据表模板」两个文件夹；
- 「原始数据」内需包含脚本读取的全部原始表（店铺规范、工厂数据、退货数据、补发数据、
  三个品牌的编码表与日销表、配件编码、配件日销）；
- 「数据表模板」内需包含 6 个模板：总数据计算、小组数据计算、大组数据计算、
  工厂数据计算、车型数据计算、配件数据计算；「店铺数据计算」为代码内新建，不依赖模板。

【本机调试】
想在命令行单独验证，可在文件末尾临时加上并运行：
    generate_report("2026-09", "20260910", r"C:\\Users\\1\\Desktop\\售后数据计算")
（不要写 if __name__ == "__main__" 分支，避免影刀以独立模块运行时重复执行。）
"""

from __future__ import annotations

import os
import re
import traceback
from collections import defaultdict
from copy import copy
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Any

# 影刀内置库：这里的 print 会写入影刀「运行日志」面板；本机直接跑时自动降级为内置 print
try:
    from xbot import print
except ImportError:  # pragma: no cover - 仅本机调试时走到
    pass

# 第三方库：若导入失败，给出影刀侧的明确处理指引
try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
except ImportError as exc:  # pragma: no cover
    raise ImportError(
        "缺少 openpyxl。请在影刀设计器右侧「更多」→「Python 包管理」中安装 openpyxl 后重试。"
    ) from exc


# ---------------------------------------------------------------------------
# 常量配置
# ---------------------------------------------------------------------------

RAW_DIR_NAME = "原始数据"
TEMPLATE_DIR_NAME = "数据表模板"
OUTPUT_DIR_NAME = "数据计算"

TEMPLATE_FILES = {
    "总数据计算": "总数据计算.xlsx",
    "小组数据计算": "小组数据计算.xlsx",
    "大组数据计算": "大组数据计算.xlsx",
    "工厂数据计算": "工厂数据计算.xlsx",
    "车型数据计算": "车型数据计算.xlsx",
    "配件数据计算": "配件数据计算.xlsx",
}

MODULE_TAG = "【售后月度数据计算】"


# ---------------------------------------------------------------------------
# 基础工具
# ---------------------------------------------------------------------------

def clean(value: Any) -> str:
    """返回去空白的展示/键字符串，且不把 None 变成 'None'。"""
    if value is None:
        return ""
    return str(value).strip()


def number(value: Any) -> Decimal:
    """把表格里以字符串、空值或数字形式存在的数值统一转成 Decimal。"""
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
    """金额和数量按常规四舍五入输出为整数。"""
    return int(Decimal(value).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def ratio(numerator: Decimal, denominator: Decimal) -> float:
    return float(numerator / denominator) if denominator else 0.0


def month_key(year: int, month: int) -> str:
    return f"{year:04d}-{month:02d}"


def parse_month(value: Any, fallback_year: int) -> str | None:
    """将常见 Excel/中文日期形式解析为 YYYY-MM。

    补发数据目前记录的是 09月09日 这类不带年份的日期，因此年份由所选报表月份提供。
    """
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


def as_bool(value: Any, default: bool = True) -> bool:
    """兼容影刀布尔值与文本参数（"True"/"False"/"否"）。"""
    if value is None or value == "":
        return default
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() not in ("false", "0", "no", "n", "否", "不")


# ---------------------------------------------------------------------------
# 原始数据读取
# ---------------------------------------------------------------------------

def read_records(path: Path) -> list[dict[str, Any]]:
    """把活动工作表读成以表头为键的记录列表，使用缓存值。"""
    workbook = load_workbook(path, data_only=True, read_only=True)
    if not workbook.worksheets:
        raise ValueError(f"原始数据文件没有工作表：{path.name}")
    sheet = workbook.active
    headers = [clean(cell) for cell in next(sheet.iter_rows(min_row=1, max_row=1, values_only=True))]
    if not any(headers):
        raise ValueError(f"原始数据文件没有表头：{path.name}")
    records: list[dict[str, Any]] = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if not any(value is not None and clean(value) != "" for value in row):
            continue
        record = {headers[index]: value for index, value in enumerate(row) if index < len(headers) and headers[index]}
        records.append(record)
    return records


def read_tuple_rows(path: Path) -> list[tuple[Any, ...]]:
    workbook = load_workbook(path, data_only=True, read_only=True)
    if not workbook.worksheets:
        raise ValueError(f"原始数据文件没有工作表：{path.name}")
    return list(workbook.active.iter_rows(min_row=2, values_only=True))


# ---------------------------------------------------------------------------
# 指标容器
# ---------------------------------------------------------------------------

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
    """已确认的分组规则：A1/A2 -> A，B1/B2 -> B，C1/C2 -> C。"""
    match = re.fullmatch(r"售后([ABC])\d+组", clean(small_group))
    return f"售后{match.group(1)}组" if match else clean(small_group)


# ---------------------------------------------------------------------------
# 写表
# ---------------------------------------------------------------------------

def write_template_rows(template_path: Path, output_path: Path, rows: list[dict[str, Any]]) -> None:
    """复制模板，并按模板表头填充其第一个工作表。"""
    workbook = load_workbook(template_path)
    if not workbook.worksheets:
        raise ValueError(f"模板没有工作表：{template_path.name}")
    sheet = workbook.active
    headers = [clean(cell.value) for cell in sheet[1]]
    if not any(headers):
        raise ValueError(f"模板没有表头：{template_path.name}")

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

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)


def write_store_report(output_path: Path, rows: list[dict[str, Any]]) -> None:
    """按已确认的新版「店铺数据计算」结构，从店铺原始数据生成报表。"""
    headers = [
        "店铺名称", "售后组", "大组", "运营组", "配件组", "平台", "只销配件店铺",
        "营业额", "订单量", "配件店铺订单量", "退货单量", "签收退货单量",
        "补发次数", "配件店铺补发次数", "补偿金额", "退货率", "签收退货率",
        "补偿率", "补发率", "配件店铺补发率", "日期", "判断月份", "售后率",
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
            cell = sheet.cell(row_index, column_index, record[header])
            if "率" in header:
                cell.number_format = "0.00%"
    for column_index, header in enumerate(headers, start=1):
        width = max(12, min(28, len(header) * 2 + 4))
        sheet.column_dimensions[chr(64 + column_index) if column_index <= 26 else "A"].width = width
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)


# ---------------------------------------------------------------------------
# 核心聚合
# ---------------------------------------------------------------------------

def code_maps(raw_dir: Path) -> tuple[dict[str, dict[str, str]], dict[str, set[str]]]:
    """返回中文/英文到车型的映射，以及全部已知车型名称。"""
    file_by_brand = {
        "永久": "永久编码.xlsx",
        "菲利普": "菲利普编码.xlsx",
        "大众": "大众编码.xlsx",
    }
    maps: dict[str, dict[str, str]] = {}
    all_models: dict[str, set[str]] = {"永久": set(), "菲利普": set(), "大众": set()}
    for brand, filename in file_by_brand.items():
        chinese: dict[str, str] = {}
        english: dict[str, str] = {}
        for row in read_tuple_rows(raw_dir / filename):
            if len(row) < 3:
                continue
            model = clean(row[2])
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


def selected_records(
    records: list[dict[str, Any]], date_column: str, report_month: str, year: int
) -> list[dict[str, Any]]:
    return [record for record in records if parse_month(record.get(date_column), year) == report_month]


def is_signed_return(record: dict[str, Any]) -> bool:
    """判断是否为已签收退货，排除截回、拦截和拒收订单。"""
    reason = clean(record.get("退货退款原因"))
    return not any(
        keyword in reason for keyword in ("截回", "拦截", "拒收")
    )


def aggregate_reports(raw_dir: Path, report_month: str) -> tuple[
    list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]],
    list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]],
]:
    """生成总表、小组、大组、工厂、车型、配件、店铺 7 类报表记录。"""
    report_year = int(report_month[:4])
    compact_month = report_month.replace("-", "")
    shops = read_records(raw_dir / "店铺规范.xlsx")
    factories = read_records(raw_dir / "工厂数据.xlsx")
    returns = selected_records(read_records(raw_dir / "退货数据.xlsx"), "登记时间", report_month, report_year)
    # 「补发数据」与「补偿数据」已经拆分：前者统计全部补发事件，后者只统计
    # 标记为参与计算的补偿金额，不能再用补偿记录数量代替补发次数。
    resends = selected_records(read_records(raw_dir / "补发数据.xlsx"), "登记时间", report_month, report_year)
    compensations = [
        row for row in selected_records(read_records(raw_dir / "补偿数据.xlsx"), "登记时间", report_month, report_year)
        if clean(row.get("是否不参与计算判断")) == "否"
    ]

    shop_by_name = {clean(row.get("店铺名称")): row for row in shops if clean(row.get("店铺名称"))}
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
        return name if name in shop_by_name else shop_name_by_key.get(shop_key(name), name)
    part_only_shops = {
        name for name, row in shop_by_name.items() if clean(row.get("只销配件店铺"))
    }

    small: defaultdict[str, dict[str, Decimal]] = defaultdict(metric)
    store_metrics: defaultdict[str, dict[str, Decimal]] = defaultdict(metric)
    total = metric()
    for row in shops:
        name = clean(row.get("店铺名称"))
        group = clean(row.get("实际负责")) or clean(row.get("售后组"))
        revenue = number(row.get("当日真实营业额"))
        orders = number(row.get("车子销量"))
        part_orders = number(row.get("配件销量")) if name in part_only_shops else Decimal("0")
        add_metrics(total, revenue=revenue, orders=orders, part_orders=part_orders)
        if name:
            add_metrics(store_metrics[name], revenue=revenue, orders=orders, part_orders=part_orders)
        if group:
            add_metrics(small[group], revenue=revenue, orders=orders, part_orders=part_orders)

    for row in returns:
        name = resolved_shop_name(clean(row.get("店铺")))
        group = clean(row.get("售后组"))
        signed = Decimal("1") if is_signed_return(row) else Decimal("0")
        add_metrics(total, returns=Decimal("1"), signed_returns=signed)
        if name:
            add_metrics(store_metrics[name], returns=Decimal("1"), signed_returns=signed)
        if group:
            add_metrics(small[group], returns=Decimal("1"), signed_returns=signed)

    for row in resends:
        name = resolved_shop_name(clean(row.get("店铺")))
        group = clean(row.get("售后组")) or clean(shop_by_name.get(name, {}).get("实际负责")) or clean(shop_by_name.get(name, {}).get("售后组"))
        part_resend = Decimal("1") if name in part_only_shops else Decimal("0")
        add_metrics(total, resends=Decimal("1"), part_resends=part_resend)
        if name:
            add_metrics(store_metrics[name], resends=Decimal("1"), part_resends=part_resend)
        if group:
            add_metrics(small[group], resends=Decimal("1"), part_resends=part_resend)

    for row in compensations:
        name = clean(row.get("店铺"))
        group = clean(row.get("售后组")) or clean(shop_by_name.get(name, {}).get("实际负责")) or clean(shop_by_name.get(name, {}).get("售后组"))
        compensation = number(row.get("补偿金额"))
        add_metrics(total, compensation=compensation)
        if name:
            add_metrics(store_metrics[name], compensation=compensation)
        if group:
            add_metrics(small[group], compensation=compensation)

    large: defaultdict[str, dict[str, Decimal]] = defaultdict(metric)
    for group, values in small.items():
        add_metrics(large[large_group_name(group)], **values)

    def group_row(group: str, values: dict[str, Decimal], include_large: bool) -> dict[str, Any]:
        base = {
            "售后组": group,
            "营业额": excel_number(values["revenue"]),
            "订单量": excel_number(values["orders"]),
            "配件店铺订单量": excel_number(values["part_orders"]),
            "退货单量": excel_number(values["returns"]),
            "店铺退货": excel_number(values["returns"]),
            "签收退货单量": excel_number(values["signed_returns"]),
            "补发次数": excel_number(values["resends"]),
            "配件店铺补发次数": excel_number(values["part_resends"]),
            "补偿金额": excel_number(values["compensation"]),
            "退货率": ratio(values["returns"], values["orders"]),
            "签收退货率": ratio(values["signed_returns"], values["orders"]),
            "补偿率": ratio(values["compensation"], values["revenue"]),
            "补发率": ratio(values["resends"], values["orders"]),
            "配件店铺补发率": ratio(values["part_resends"], values["part_orders"]),
            "售后率": ratio(values["returns"] + values["resends"], values["orders"]),
            "日期": report_month,
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

    factory_metrics: defaultdict[str, dict[str, Decimal]] = defaultdict(metric)
    for row in factories:
        name = clean(row.get("工厂"))
        if not name:
            continue
        add_metrics(
            factory_metrics[name],
            # 优先使用原始字段，避免引用源表已处理的汇总列。
            revenue=number(row.get("营业额", row.get("营业额汇总"))),
            orders=number(row.get("订单量", row.get("订单量汇总"))),
        )
    for row in returns:
        name = clean(row.get("工厂"))
        if name:
            add_metrics(factory_metrics[name], returns=Decimal("1"))
    for row in resends:
        name = clean(row.get("工厂匹配")) or clean(row.get("工厂")) or "未匹配工厂"
        add_metrics(factory_metrics[name], resends=Decimal("1"))
    for row in compensations:
        name = clean(row.get("工厂")) or "未匹配工厂"
        add_metrics(factory_metrics[name], compensation=number(row.get("补偿金额")))
    factory_rows = []
    for factory, values in factory_metrics.items():
        if not any(values.values()):
            continue
        factory_rows.append({
            "工厂名称": factory,
            "工厂": factory,
            "营业额": excel_number(values["revenue"]),
            "订单量": excel_number(values["orders"]),
            "退货单量": excel_number(values["returns"]),
            "补发次数": excel_number(values["resends"]),
            "补发次数 （除万里牛）": excel_number(values["resends"]),
            "补偿金额": excel_number(values["compensation"]),
            "退货率": ratio(values["returns"], values["orders"]),
            "补偿率": ratio(values["compensation"], values["revenue"]),
            "补发率": ratio(values["resends"], values["orders"]),
            "日期": report_month,
            "月份": compact_month,
        })
    factory_rows.sort(key=lambda row: (-number(row["营业额"]), clean(row["工厂"])))

    maps, known_by_brand = code_maps(raw_dir)
    known_vehicle_models = set().union(*known_by_brand.values())
    vehicle_metrics: defaultdict[str, dict[str, Decimal]] = defaultdict(metric)
    sales_specs = [
        # “销售额（汇总）”已逐行取整，车型层应汇总原始“销售额”后再取整。
        ("永久日销.xlsx", "编码表车型", "销量", "销售额", "销量", "销售额"),
        ("菲利普日销.xlsx", "型号", "销量", "销售额", "销量", "销售额"),
        ("大众日销.xlsx", "车型", "销量", "销售额", "销量", "销售额"),
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

    for row in returns:
        brand = clean(row.get("品牌"))
        model = ""
        if brand == "永久":
            model = maps["永久_英文"].get(clean(row.get("永久车型英文编码")), "") or maps["永久_中文"].get(clean(row.get("永久中文编码")), "")
        elif brand == "菲利普":
            model = maps["菲利普_英文"].get(clean(row.get("菲利普车型英文编码")), "") or maps["菲利普_中文"].get(clean(row.get("菲利普中文编码")), "")
        elif brand == "大众":
            model = maps["大众_中文"].get(clean(row.get("大众中文编码")), "")
        if model:
            add_metrics(vehicle_metrics[model], returns=Decimal("1"))

    part_code_map = {
        clean(row[1]): clean(row[0])
        for row in read_tuple_rows(raw_dir / "配件编码.xlsx")
        if len(row) >= 2 and clean(row[0]) and clean(row[1])
    }
    part_metrics: defaultdict[str, dict[str, Decimal]] = defaultdict(metric)
    for row in read_records(raw_dir / "配件日销.xlsx"):
        part = clean(row.get("标题"))
        quantity = number(row.get("销量"))
        revenue = number(row.get("销售额"))
        if part and (quantity or revenue):
            add_metrics(part_metrics[part], orders=quantity, revenue=revenue)

    model_by_english = {
        code: model for key, mapping in maps.items() if key.endswith("_英文") for code, model in mapping.items()
    }

    def add_after_sales_detail(row: dict[str, Any], *, resend: bool = False, compensation: bool = False) -> None:
        """将补发/补偿记录归集到车型或配件。"""
        named_model = clean(row.get("车型"))
        english_code = clean(row.get("英文编码整合")) or clean(row.get("英文编码"))
        vehicle_model = named_model if named_model in known_vehicle_models else model_by_english.get(english_code, "")
        amount = number(row.get("补偿金额")) if compensation else Decimal("0")
        if vehicle_model:
            add_metrics(vehicle_metrics[vehicle_model], **({"resends": Decimal("1")} if resend else {}), **({"compensation": amount} if compensation else {}))
            return
        part = part_code_map.get(english_code, "") or named_model or clean(row.get("中文编码")) or "未匹配车型或配件"
        add_metrics(part_metrics[part], **({"resends": Decimal("1")} if resend else {}), **({"compensation": amount} if compensation else {}))

    for row in resends:
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
            "日期": report_month,
            "月份": report_month,
        }

    vehicle_rows = [detail_row(name, values, "车型") for name, values in vehicle_metrics.items() if any(values.values())]
    part_rows = [detail_row(name, values, "配件") for name, values in part_metrics.items() if any(values.values())]
    vehicle_rows.sort(key=lambda row: (-number(row["营业额"]), clean(row["车型"])))
    part_rows.sort(key=lambda row: (-number(row["销售额"]), clean(row["配件型号名称"])))

    store_rows = []
    for name, shop in shop_by_name.items():
        values = store_metrics[name]
        group = clean(shop.get("实际负责")) or clean(shop.get("售后组"))
        store_rows.append({
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
            "补发次数": excel_number(values["resends"]),
            "配件店铺补发次数": excel_number(values["part_resends"]),
            "补偿金额": excel_number(values["compensation"]),
            "退货率": ratio(values["returns"], values["orders"]),
            "签收退货率": ratio(values["signed_returns"], values["orders"]),
            "补偿率": ratio(values["compensation"], values["revenue"]),
            "补发率": ratio(values["resends"], values["orders"]),
            "配件店铺补发率": ratio(values["part_resends"], values["part_orders"]),
            "日期": report_month,
            "判断月份": compact_month,
            "售后率": ratio(values["returns"] + values["resends"], values["orders"]),
        })
    store_rows.sort(key=lambda row: (-number(row["营业额"]), clean(row["店铺名称"])))

    total_row = [{
        "已确认登记表无误": "已确认",
        "营业额": excel_number(total["revenue"]),
        "订单量": excel_number(total["orders"]),
        "退货单量": excel_number(total["returns"]),
        "补发次数": excel_number(total["resends"]),
        "补偿金额": excel_number(total["compensation"]),
        "退货率": ratio(total["returns"], total["orders"]),
        "补偿率": ratio(total["compensation"], total["revenue"]),
        "补发率": ratio(total["resends"], total["orders"]),
        "万里牛补发": 0,
        "订单补发": excel_number(total["resends"]),
        "数据日期": compact_month,
    }]
    return total_row, small_rows, large_rows, factory_rows, vehicle_rows, part_rows, store_rows


# ---------------------------------------------------------------------------
# 影刀入口
# ---------------------------------------------------------------------------

def default_root() -> Path:
    """未显式传入 root 时按优先级探测数据根目录。"""
    candidates: list[Path] = []
    try:
        candidates.append(Path(__file__).resolve().parent)
    except NameError:
        pass
    candidates.append(Path.cwd())
    candidates.append(Path.home() / "Desktop" / "售后数据计算")
    for candidate in candidates:
        if (candidate / RAW_DIR_NAME).is_dir() and (candidate / TEMPLATE_DIR_NAME).is_dir():
            return candidate
    return candidates[-1]


def _generate_report(month: str, created_date: str, root: str, overwrite: bool) -> dict[str, Any]:
    report_month = clean(month) or datetime.now().strftime("%Y-%m")
    if not re.fullmatch(r"20\d{2}-(0[1-9]|1[0-2])", report_month):
        raise ValueError(f"month 必须是 YYYY-MM 格式，例如 2026-09；当前传入：{month!r}")

    created = clean(created_date) or datetime.now().strftime("%Y%m%d")
    if not re.fullmatch(r"20\d{2}(0[1-9]|1[0-2])([0-2]\d|3[01])", created):
        raise ValueError(f"created_date 必须是 YYYYMMDD 格式，例如 20260910；当前传入：{created_date!r}")

    root_path = Path(clean(root)).expanduser().resolve() if clean(root) else default_root().resolve()
    raw_dir = root_path / RAW_DIR_NAME
    template_dir = root_path / TEMPLATE_DIR_NAME
    output_dir = root_path / OUTPUT_DIR_NAME
    if not raw_dir.is_dir() or not template_dir.is_dir():
        raise FileNotFoundError(
            f"根目录 {root_path} 下必须同时包含「{RAW_DIR_NAME}」和「{TEMPLATE_DIR_NAME}」两个文件夹"
        )

    print(f"{MODULE_TAG} 开始计算，统计月份 {report_month}，创建日期 {created}，根目录 {root_path}")

    reports = aggregate_reports(raw_dir, report_month)
    report_rows = dict(zip(TEMPLATE_FILES, reports[:6]))

    files: list[str] = []
    row_count = 0
    for report_name, rows in report_rows.items():
        template_path = template_dir / TEMPLATE_FILES[report_name]
        if not template_path.is_file():
            raise FileNotFoundError(f"缺少模板：{template_path}")
        output_path = output_dir / f"{report_name}_{created}.xlsx"
        if output_path.exists() and not overwrite:
            print(f"{MODULE_TAG} 已跳过（同名文件已存在）：{output_path.name}")
            continue
        write_template_rows(template_path, output_path, rows)
        files.append(str(output_path))
        row_count += len(rows)
        print(f"{MODULE_TAG} 已生成 {output_path.name}：{len(rows)} 行")

    store_path = output_dir / f"店铺数据计算_{created}.xlsx"
    if store_path.exists() and not overwrite:
        print(f"{MODULE_TAG} 已跳过（同名文件已存在）：{store_path.name}")
    else:
        write_store_report(store_path, reports[6])
        files.append(str(store_path))
        row_count += len(reports[6])
        print(f"{MODULE_TAG} 已生成 {store_path.name}：{len(reports[6])} 行")

    message = f"共生成 {len(files)} 个文件、{row_count} 行数据，输出目录：{output_dir}"
    print(f"{MODULE_TAG} 完成，{message}")
    return {
        "ok": True,
        "month": report_month,
        "created_date": created,
        "output_dir": str(output_dir),
        "files": files,
        "file_count": len(files),
        "row_count": row_count,
        "message": message,
    }


def generate_report(month: str = "", created_date: str = "", root: str = "", overwrite: Any = True) -> dict[str, Any]:
    """影刀「调用模块」入口：4 个参数全部可选，返回一个字典。

    :param month:        统计月份 YYYY-MM，留空取当前月
    :param created_date: 文件创建日期 YYYYMMDD，留空取当天
    :param root:         数据根目录，留空自动探测
    :param overwrite:    同名文件是否覆盖，默认 True
    """
    try:
        return _generate_report(
            month=clean(month),
            created_date=clean(created_date),
            root=clean(root),
            overwrite=as_bool(overwrite, True),
        )
    except Exception:
        print(f"{MODULE_TAG} 执行失败，错误详情：")
        print(traceback.format_exc())
        raise


def main(args: Any = None) -> dict[str, Any]:
    """影刀 Python 模块的默认入口：args 为可视化流程传入的参数字典。"""
    args = args or {}
    if not isinstance(args, dict):
        args = {}
    return generate_report(
        month=args.get("month", ""),
        created_date=args.get("created_date", ""),
        root=args.get("root", ""),
        overwrite=args.get("overwrite", True),
    )

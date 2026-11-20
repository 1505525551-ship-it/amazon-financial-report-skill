# -*- coding: utf-8 -*-
"""明细和汇总报告整理：配对 csv 与 summary，整理成一份 xlsx 并核对。"""

import argparse
import os
import re
import sys
import time

try:
    import pdfplumber  # noqa: F401  仅用于提前确认依赖存在
    import openpyxl
    from openpyxl.styles import Alignment, Font
    from openpyxl.styles import DEFAULT_FONT
    from openpyxl.utils import get_column_letter
except ImportError as error:
    sys.stderr.write("缺少依赖：%s\n" % error)
    sys.stderr.write("请先安装：pip install openpyxl pdfplumber\n")
    raise SystemExit(1)

from common import (
    ConfigError,
    NUMBER_FORMAT_FILE,
    cell_text,
    load_field_rules,
    load_number_formats,
    load_summary_rules,
    match_behavior,
    parse_amount,
    period_code,
    prepare_rules,
    resolve_number_format,
)
from detail_report import process_detail
from field_reconcile import (
    CONSISTENT as FIELD_CONSISTENT,
    FIELD_HEADERS,
    build_field_rows,
    load_reconcile_mapping,
)
from summary_report import CATEGORIES, parse_summary

FONT_NAME = "微软雅黑"
DEFAULT_FONT.name = FONT_NAME
DEFAULT_FONT.size = 11

TOLERANCE = 0.01
OUTPUT_DIR_NAME = "整理结果"
FAIL_COLOR = "FF0000"
OK_COLOR = "00703C"

AMOUNT_COLUMNS = (
    "product_sales", "product_sales_tax", "shipping_credits", "shipping_credits_tax",
    "gift_wrap_credits", "gift_wrap_credits_tax", "regulatory_fee", "tax_on_regulatory_fee",
    "promotional_rebates", "promotional_rebates_tax", "sales_tax_collected",
    "low_value_goods", "marketplace_withheld_tax", "selling_fees", "fba_fees",
    "points_fees", "other_transaction_fees", "other", "total",
)
INTEGER_COLUMNS = ("quantity",)

CHECK_HEADERS = ["文件名", "期间", "币种", "明细报告total", "收入", "费用", "转账", "税",
                 "debt", "平台应收净额", "差异", "汇总和明细报告核对",
                 "汇总报告核对", "日期范围报告核对"]
DETAIL_HEADERS = ["文件名", "店铺显示名", "法定名称", "报表期间起", "报表期间止", "币种",
                  "分类", "字段原文", "中文标准字段", "借方", "贷方", "净额"]

DEBT_CHINESE = "信用卡扣款及其他债务回收"
DEBT_KEYS = (
    "charges to credit card",
    "cargos a tarjeta de crédito",
    "cobranças no cartão de crédito",
    "クレジットカード",
    "kreditkarte",
    "carte de crédit",
    "carta di credito",
)


def parse_args(argv):
    parser = argparse.ArgumentParser(
        description="整理亚马逊日期范围报告与 summary，输出核对结果",
    )
    parser.add_argument("--input", required=True, help="放 csv 与 pdf 的文件夹")
    parser.add_argument("--output", default="", help="输出的 xlsx 路径或目录")
    parser.add_argument("--config", default="", help="配置文件夹，默认用技能自带配置")
    parser.add_argument("--tolerance", type=float, default=TOLERANCE, help="核对容差，默认 0.01")
    return parser.parse_args(argv)


def default_config_dir():
    here = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(os.path.dirname(here), "config")


def collect_batches(root):
    """每个含 csv 或 pdf 的目录算一批，每批各出一份工作簿。"""
    batches = []
    for current, dirs, files in os.walk(root):
        dirs[:] = sorted(d for d in dirs
                         if not d.startswith(".") and d != OUTPUT_DIR_NAME)
        detail_files = []
        summary_files = []
        for name in sorted(files):
            lowered = name.lower()
            path = os.path.join(current, name)
            if lowered.endswith(".csv"):
                detail_files.append(path)
            elif lowered.endswith(".pdf"):
                summary_files.append(path)
        if detail_files or summary_files:
            batches.append((current, detail_files, summary_files))
    batches.sort(key=lambda item: item[0])
    return batches


def pair_files(detail_files, summary_files):
    """同名（去扩展名）视为同一组。"""
    pairs = {}
    for path in detail_files:
        stem = os.path.splitext(os.path.basename(path))[0].strip()
        pairs.setdefault(stem, {})["detail"] = path
    for path in summary_files:
        stem = os.path.splitext(os.path.basename(path))[0].strip()
        pairs.setdefault(stem, {})["summary"] = path
    return dict(sorted(pairs.items(), key=lambda item: item[0]))


def resolve_output_path(raw, batch_dir, input_dir, stamp, single_batch):
    """每批一份工作簿，文件名带上该批的目录名，避免收集到一起时分不清。"""
    same_as_root = os.path.abspath(batch_dir) == os.path.abspath(input_dir)
    label = "" if same_as_root else os.path.basename(os.path.abspath(batch_dir))
    name = "明细和汇总报告整理_%s%s.xlsx" % (label + "_" if label else "", stamp)

    if not raw:
        target = os.path.join(batch_dir, OUTPUT_DIR_NAME, name)
    elif raw.lower().endswith(".xlsx") and single_batch:
        target = os.path.abspath(raw)
    else:
        base = raw
        if raw.lower().endswith(".xlsx"):
            base = os.path.dirname(os.path.abspath(raw))
        target = os.path.join(os.path.abspath(base), name)

    parent = os.path.dirname(target)
    if parent:
        os.makedirs(parent, exist_ok=True)
    return target


def prepare_detail(record, rules, number_formats, rule_matcher, behavior_cache):
    """补齐平台行为，并把金额列按该币种的数值格式转成数字。"""
    currency = record.get("currency", "")
    decimal_sep, thousands_sep = resolve_number_format(currency, number_formats)

    rows = record["rows"]
    rules_list = rules["behavior_rules"]
    for row in rows:
        row["behavior"] = match_behavior(
            row.get("type", ""), row.get("description", ""),
            rules_list, rule_matcher, behavior_cache,
        )
    for row in rows:
        for column in AMOUNT_COLUMNS:
            if column in row:
                row[column] = parse_amount(row[column], decimal_sep, thousands_sep)
        for column in INTEGER_COLUMNS:
            if column in row:
                row[column] = parse_amount(row[column], decimal_sep, thousands_sep)
    total = 0.0
    for row in rows:
        total += row.get("total", 0.0) or 0.0

    # 日期范围报告表内核对：每行 product_sales..other 之和与 total 的差额，取绝对值累加
    columns = list(rules["columns"])
    span = None
    if "product_sales" in columns and "other" in columns:
        span = (columns.index("product_sales"), columns.index("other") + 1)
    internal_diff = 0.0
    if span:
        for row in rows:
            subtotal = 0.0
            for column in columns[span[0]:span[1]]:
                subtotal += row.get(column, 0.0) or 0.0
            internal_diff += abs(round(subtotal - (row.get("total", 0.0) or 0.0), 2))
    return rows, round(total, 2), round(internal_diff, 2)


def fill_sheet(sheet, headers, rows, widths, amount_indexes,
               text_indexes=(), freeze="A2", color_rule=None):
    """把表头、数据和格式写进给定工作表。"""
    sheet.append(headers)
    for cell in sheet[1]:
        cell.font = Font(name=FONT_NAME, size=11, bold=True)
        cell.alignment = Alignment(vertical="center")

    for row in rows:
        sheet.append(row)

    for column, width in zip(range(1, len(headers) + 1), widths):
        sheet.column_dimensions[get_column_letter(column)].width = width
    for index in amount_indexes:
        letter = get_column_letter(index)
        for cell in sheet[letter][1:]:
            cell.number_format = "#,##0.00"
    for index in text_indexes:
        letter = get_column_letter(index)
        for cell in sheet[letter][1:]:
            cell.number_format = "@"

    if color_rule:
        color_rule(sheet)

    sheet.freeze_panes = freeze
    return sheet


def write_simple_sheet(workbook, title, headers, rows, widths, amount_indexes,
                       text_indexes=(), freeze="A2", color_rule=None):
    sheet = workbook.create_sheet(title)
    return fill_sheet(sheet, headers, rows, widths, amount_indexes,
                      text_indexes, freeze, color_rule)


CHECK_RESULT_COLUMNS = (12, 13, 14)  # L 表间核对、M 汇总报告核对、N 日期范围报告核对
CONSISTENT = "核对一致"


def write_check_sheet(sheet, rows, tolerance):
    def color_rule(target):
        for row_index in range(2, len(rows) + 2):
            for column in CHECK_RESULT_COLUMNS:
                cell = target.cell(row_index, column)
                if cell.value == "":
                    continue
                color = OK_COLOR if cell.value == CONSISTENT else FAIL_COLOR
                cell.font = Font(name=FONT_NAME, size=11, color=color)

    return fill_sheet(
        sheet, CHECK_HEADERS, rows,
        [30, 10, 8, 15, 13, 13, 13, 11, 12, 15, 11, 22, 20, 20],
        amount_indexes=(4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14),
        text_indexes=(2,),
        color_rule=color_rule,
    )


def write_detail_sheet(workbook, rules, rows_by_file):
    columns = list(rules["columns"])
    labels = list(rules["labels"])

    sheet = workbook.create_sheet("明细总表")
    for index, name in enumerate(columns, start=1):
        cell = sheet.cell(1, index, name)
        cell.font = Font(name=FONT_NAME, size=11, bold=True)
    for index, label in enumerate(labels, start=1):
        cell = sheet.cell(2, index, label)
        cell.font = Font(name=FONT_NAME, size=11, bold=True)

    row_number = 3
    for rows in rows_by_file:
        for row in rows:
            for index, column in enumerate(columns, start=1):
                sheet.cell(row_number, index, row.get(column, ""))
            row_number += 1

    for index, column in enumerate(columns, start=1):
        letter = get_column_letter(index)
        if column in AMOUNT_COLUMNS:
            sheet.column_dimensions[letter].number_format = "#,##0.00"

    widths = {"website": 30, "behavior": 26, "date_time": 20, "settlement_id": 14,
              "type": 22, "order_id": 21, "sku": 26, "description": 46,
              "tax_collection_model": 20, "transaction_status": 14,
              "transaction_release_date": 22}
    for index, column in enumerate(columns, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = widths.get(column, 13)

    sheet.freeze_panes = "A3"
    return sheet


def debt_credit(summary_result):
    """转账分类里「信用卡扣款及其他债务回收」的贷方金额。"""
    total = 0.0
    for item in summary_result.get("items", []):
        if item["category"] != "转账":
            continue
        label = re.sub(r"\s+", " ", item["label"]).strip().lower()
        if item["chinese"] == DEBT_CHINESE or any(key in label for key in DEBT_KEYS):
            total += item["credit"]
    return round(total, 2)


def build_check_row(stem, pair, detail_result, summary_result, tolerance):
    """组装一行核对结果，返回 (行数据, 状态)。状态取 ok / fail / skip。"""
    currency = ""
    period_start = ""
    period_end = ""
    if summary_result:
        currency = summary_result.get("currency", "")
        period_start = summary_result.get("period_start", "")
        period_end = summary_result.get("period_end", "")
    if not currency and detail_result:
        currency = detail_result.get("currency", "")

    detail_ok = bool(detail_result and detail_result["status"] == "ok")
    summary_ok = bool(summary_result and summary_result["status"] == "ok")

    detail_total = detail_result["detail_total"] if detail_ok else None
    four = [None, None, None, None]
    summary_total = None
    debt = None
    receivable = None
    if summary_ok:
        four = [summary_result["totals"].get(name) for name in CATEGORIES]
        summary_total = round(sum(four), 2)
        debt = debt_credit(summary_result)
        receivable = round((four[0] or 0.0) + (four[1] or 0.0) + debt, 2)

    diff = None
    cross = ""           # L 表间核对
    summary_check = ""   # M 汇总报告表内核对
    detail_check = ""    # N 日期范围报告表内核对

    # N：只要明细读出来了就能算，与 summary 无关
    if detail_result:
        if detail_ok:
            value = detail_result.get("internal_diff", 0.0)
            detail_check = CONSISTENT if value <= tolerance else value
        elif detail_result["status"] == "skip":
            detail_check = "无数据行"
        else:
            detail_check = "明细解析失败"

    # M：汇总报告内部三道校验的差额绝对值合计
    if summary_result:
        if not summary_ok:
            summary_check = "汇总解析失败"
        elif summary_result.get("internal_missing"):
            summary_check = "核对异常"
        else:
            value = summary_result.get("internal_diff", 0.0)
            summary_check = CONSISTENT if value <= tolerance else value

    # L：明细 total 合计 vs summary 四项合计
    if not pair.get("summary"):
        cross = "仅明细"
    elif not pair.get("detail"):
        cross = "仅summary"
    elif not summary_ok:
        cross = "汇总解析失败"
    elif not detail_ok:
        cross = "无数据行" if detail_result and detail_result["status"] == "skip" else "明细解析失败"
    else:
        diff = round((detail_total or 0.0) - (summary_total or 0.0), 2)
        cross = CONSISTENT if abs(diff) <= tolerance else abs(diff)

    if cross in ("仅明细", "仅summary"):
        state = "skip"
    elif cross == CONSISTENT and summary_check in (CONSISTENT, "") and detail_check in (CONSISTENT, ""):
        state = "ok"
    else:
        state = "fail"

    return [
        stem,
        period_code(period_start, period_end),
        currency,
        detail_total if detail_total is not None else "",
        four[0] if four[0] is not None else "",
        four[1] if four[1] is not None else "",
        four[2] if four[2] is not None else "",
        four[3] if four[3] is not None else "",
        debt if debt is not None else "",
        receivable if receivable is not None else "",
        diff if diff is not None else "",
        cross,
        summary_check,
        detail_check,
    ], state


def build_detail_rows(stem, summary_result):
    rows = []
    display_name = summary_result.get("display_name", "")
    legal_name = summary_result.get("legal_name", "")
    for item in summary_result["items"]:
        rows.append([
            stem, display_name, legal_name,
            summary_result.get("period_start", ""), summary_result.get("period_end", ""),
            summary_result.get("currency", ""), item["category"], item["label"],
            item["chinese"], item["debit"], item["credit"],
            round(item["debit"] + item["credit"], 2),
        ])
    return rows


def collect_unmatched(rows_by_file):
    """平台行为没匹配上的 type、出现次数和一条示例描述。"""
    seen = {}
    for rows in rows_by_file:
        for row in rows:
            if row.get("behavior"):
                continue
            key = cell_text(row.get("type", "")) or "空 type"
            entry = seen.setdefault(key, {"count": 0, "sample": cell_text(row.get("description", ""))})
            entry["count"] += 1
    return seen


def print_batch_report(batch_dir, stats, failures, unmapped, unmatched, output_path, single):
    if not single:
        print("")
        print("---- %s ----" % os.path.basename(os.path.abspath(batch_dir)))
    print("配对情况：%d 组（明细+summary），%d 份仅明细，%d 份仅 summary"
          % (stats["paired"], stats["detail_only"], stats["summary_only"]))
    print("明细整理：成功 %d 份，跳过 %d 份，失败 %d 份，共 %d 行"
          % (stats["detail_ok"], stats["detail_skip"], stats["detail_fail"], stats["detail_rows"]))
    print("summary 解析：成功 %d 份，失败 %d 份" % (stats["summary_ok"], stats["summary_fail"]))
    print("三列核对：通过 %d，未通过 %d" % (stats["check_ok"], stats["check_fail"]))
    print("输出文件：%s" % output_path)

    if failures:
        print("---- 需要处理 ----")
        for name, reason in failures:
            print("  %s：%s" % (name, reason))
    if unmapped:
        print("---- 未匹配到中文的字段（可补进 summary_report_field.csv）----")
        for (category, label), count in unmapped.items():
            print("  %s / %s  出现 %d 次" % (category, label, count))
    if unmatched:
        print("---- 平台行为没匹配上的交易（可补进 detail_report_behavior.csv）----")
        for name, entry in unmatched.items():
            print("  %s  出现 %d 次 | 示例描述：%s" % (name, entry["count"], entry["sample"]))


def process_batch(batch_dir, detail_files, summary_files, field_rules, summary_rules,
                  number_formats, rule_matcher, behavior_cache, tolerance, reconcile_mapping):
    """跑完一批文件，返回这一批的统计与结果。"""
    pairs = pair_files(detail_files, summary_files)

    check_rows = []
    detail_rows_by_file = []
    detail_rows_map = {}
    summary_results = {}
    field_rows = []
    field_stats = {"consistent": 0, "diff": 0, "unmapped": 0}
    failures = []
    unmapped = {}

    stats = {"paired": 0, "detail_only": 0, "summary_only": 0,
             "detail_ok": 0, "detail_skip": 0, "detail_fail": 0, "detail_rows": 0,
             "summary_ok": 0, "summary_fail": 0, "check_ok": 0, "check_fail": 0}

    for stem, pair in pairs.items():
        detail_result = None
        if pair.get("detail"):
            detail_result = process_detail(pair["detail"], field_rules)
            if detail_result["status"] == "ok":
                rows, total, internal_diff = prepare_detail(detail_result, field_rules, number_formats,
                                                            rule_matcher, behavior_cache)
                detail_result["rows"] = rows
                detail_result["detail_total"] = total
                detail_result["internal_diff"] = internal_diff
                detail_rows_by_file.append(rows)
                detail_rows_map[stem] = rows
                stats["detail_ok"] += 1
                stats["detail_rows"] += len(rows)
            elif detail_result["status"] == "skip":
                stats["detail_skip"] += 1
                failures.append((stem + ".csv", detail_result["message"]))
            else:
                stats["detail_fail"] += 1
                failures.append((stem + ".csv", detail_result["message"]))

        summary_result = None
        if pair.get("summary"):
            summary_result = parse_summary(pair["summary"], summary_rules,
                                           field_rules["currency_mapping"], number_formats)
            if summary_result["status"] == "ok":
                summary_results[stem] = summary_result
                stats["summary_ok"] += 1
                for item in summary_result["items"]:
                    if not item["chinese"]:
                        key = (item["category"], item["label"])
                        unmapped[key] = unmapped.get(key, 0) + 1
            else:
                stats["summary_fail"] += 1
                failures.append((stem + ".pdf", summary_result.get("message", "解析失败")))

        row, state = build_check_row(stem, pair, detail_result, summary_result, tolerance)
        check_rows.append(row)

        has_detail = bool(pair.get("detail"))
        has_summary = bool(pair.get("summary"))
        if has_detail and has_summary:
            stats["paired"] += 1
        elif has_detail:
            stats["detail_only"] += 1
        else:
            stats["summary_only"] += 1

        if state == "ok":
            stats["check_ok"] += 1
        elif state == "fail":
            stats["check_fail"] += 1
            reasons = [str(value) for value in (row[11], row[12], row[13])
                       if value != CONSISTENT and value != ""]
            failures.append((stem, "；".join(reasons)))

    for stem, summary_result in summary_results.items():
        rows, one_stats = build_field_rows(stem, summary_result,
                                           detail_rows_map.get(stem, []),
                                           reconcile_mapping, tolerance)
        field_rows.extend(rows)
        for key in field_stats:
            field_stats[key] += one_stats[key]

    return (stats, check_rows, detail_rows_by_file, summary_results,
            failures, unmapped, field_rows, field_stats)


def write_field_sheet(workbook, field_rows):
    """summary 字段 vs 明细字段的逐项核对。"""
    def color_rule(sheet):
        for index, row in enumerate(field_rows, start=2):
            cell = sheet.cell(index, 10)
            if cell.value == FIELD_CONSISTENT:
                cell.font = Font(name=FONT_NAME, size=11, color=OK_COLOR)
            elif cell.value == "差额":
                cell.font = Font(name=FONT_NAME, size=11, color=FAIL_COLOR)

    return write_simple_sheet(
        workbook, "字段核对", FIELD_HEADERS, field_rows,
        [8, 30, 8, 46, 26, 14, 14, 12, 50, 10],
        amount_indexes=(6, 7, 8),
        color_rule=color_rule,
    )


def write_todo_sheet(workbook, unmapped, unmatched):
    """有待补项时才加这一页，数据干净时不出现。"""
    if not unmapped and not unmatched:
        return None
    headers = ["类型", "分类", "原文 / type", "出现次数", "示例描述", "补到哪里"]
    rows = []
    for (category, label), count in unmapped.items():
        rows.append(["未匹配中文字段", category, label, count, "",
                     "summary_report_field.csv"])
    for name, entry in unmatched.items():
        rows.append(["未匹配平台行为", "明细 type", name, entry["count"], entry["sample"],
                     "detail_report_behavior.csv"])
    return write_simple_sheet(
        workbook, "待补清单", headers, rows,
        [16, 10, 56, 10, 46, 40],
        amount_indexes=(),
        text_indexes=(4,),
    )


def save_batch(output_path, field_rules, check_rows, summary_results,
               detail_rows_by_file, tolerance, unmapped, unmatched, field_rows):
    workbook = openpyxl.Workbook()
    check_sheet = workbook.active
    check_sheet.title = "校验核对"
    write_check_sheet(check_sheet, check_rows, tolerance)

    detail_rows = []
    for stem, result in summary_results.items():
        detail_rows.extend(build_detail_rows(stem, result))
    write_simple_sheet(
        workbook, "Summary全字段", DETAIL_HEADERS, detail_rows,
        [30, 18, 46, 12, 12, 8, 8, 56, 26, 14, 14, 14],
        amount_indexes=(10, 11, 12),
        text_indexes=(4, 5),
    )
    write_field_sheet(workbook, field_rows)
    write_detail_sheet(workbook, field_rules, detail_rows_by_file)
    write_todo_sheet(workbook, unmapped, unmatched)
    workbook.save(output_path)


def main(argv):
    args = parse_args(argv)
    input_dir = os.path.abspath(args.input)
    if not os.path.isdir(input_dir):
        print("输入目录不存在：%s" % input_dir)
        return 1

    config_dir = args.config or default_config_dir()
    try:
        field_rules = load_field_rules(config_dir)
        summary_rules = load_summary_rules(config_dir)
        number_formats = load_number_formats(os.path.join(config_dir, NUMBER_FORMAT_FILE))
        reconcile_mapping = load_reconcile_mapping(config_dir)
    except ConfigError as error:
        print("配置读取失败：%s" % error)
        return 1

    batches = collect_batches(input_dir)
    if not batches:
        print("目录里没有 csv 或 pdf：%s" % input_dir)
        return 1

    single = len(batches) == 1
    rule_matcher = prepare_rules(field_rules["behavior_rules"])
    behavior_cache = {}
    stamp = time.strftime("%Y%m%d_%H%M%S")

    print("")
    print("==== 明细和汇总报告整理 ====")
    if single:
        print("输入目录：%s" % input_dir)
    else:
        print("输入目录：%s" % input_dir)
        print("共 %d 个文件夹，每个文件夹各出一份工作簿" % len(batches))

    total_fail = 0
    total_saved = 0
    for batch_dir, detail_files, summary_files in batches:
        (stats, check_rows, detail_rows_by_file, summary_results, failures, unmapped,
         field_rows, field_stats) = process_batch(
            batch_dir, detail_files, summary_files, field_rules, summary_rules,
            number_formats, rule_matcher, behavior_cache, args.tolerance, reconcile_mapping)

        unmatched = collect_unmatched(detail_rows_by_file)
        output_path = resolve_output_path(args.output, batch_dir, input_dir, stamp, single)
        try:
            save_batch(output_path, field_rules, check_rows, summary_results,
                       detail_rows_by_file, args.tolerance, unmapped, unmatched, field_rows)
        except Exception as error:
            print("保存失败：%s（%s）" % (error, output_path))
            return 1

        print_batch_report(batch_dir, stats, failures, unmapped,
                           unmatched, output_path, single)
        print("字段核对：一致 %d 项，差额 %d 项，未配置 %d 项"
              % (field_stats["consistent"], field_stats["diff"], field_stats["unmapped"]))
        total_fail += stats["check_fail"]
        total_saved += 1

    if not single:
        print("")
        print("---- 汇总 ----")
        print("共 %d 个文件夹，生成 %d 份工作簿，核对失败 %d 组"
              % (len(batches), total_saved, total_fail))

    return 2 if total_fail else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))

# -*- coding: utf-8 -*-
"""亚马逊 summary（pdf）解析：四大字段合计 + 全部字段借贷方明细。"""

import re

import pdfplumber

from common import (
    cell_text,
    detect_currency_from_text,
    normalize_key,
    parse_amount,
    parse_report_period,
    resolve_number_format,
)

CATEGORIES = ("收入", "费用", "转账", "税")
SUBTOTAL_WORDS = {
    "subtotals", "subtotal", "subtotales", "subtotais",
    "小計", "部分合計", "subtotal geral", "totales", "totais",
}

ROW_TOLERANCE = 4.0
NUMBER_MAX_GAP = 8.0
SUMMARY_BAND = (95.0, 160.0)
BLOCK_BOTTOM = 545.0
TOTAL_LABEL_X = (0.0, 160.0)
TOTAL_NUMBER_X = 690.0
LEFT_LABEL_X = (0.0, 200.0)
LEFT_NUMBER_X = (200.0, 396.0)
LEFT_DEBIT_RIGHT = 340.0
RIGHT_LABEL_X = (400.0, 640.0)
RIGHT_NUMBER_X = (640.0, 10000.0)
RIGHT_DEBIT_RIGHT = 715.0

HEADER_BAND = (10.0, 52.0)
NAME_LABEL_X = 380.0
NAME_SPLIT_X = 480.0
DISPLAY_KEYS = (
    "display name", "nombre para mostrar", "nome de exibição", "表示名",
    "anzeigename", "nom d'affichage", "görünen ad", "nazwa wyświetlana",
    "visningsnamn", "weergavenaam", "nome visualizzato",
)
LEGAL_KEYS = (
    "legal name", "nombre legal", "nome jurídico", "正式名称",
    "rechtlicher name", "raison sociale", "yasal ad", "nazwa prawna",
    "juridiskt namn", "juridische naam", "ragione sociale",
)


class LabelResolver:
    """PDF 原文 -> 中文标准字段，支持被截断的名称按前缀匹配。"""

    def __init__(self, entries):
        self.exact = {}
        self.prefixes = []
        for category, key, chinese in entries:
            self.exact.setdefault((category, key), chinese)
            self.exact.setdefault(("*", key), chinese)
            self.prefixes.append((key, chinese))
        self.prefixes.sort(key=lambda item: -len(item[0]))

    def resolve(self, category, label):
        key = normalize_key(label)
        chinese = self.exact.get((category, key)) or self.exact.get(("*", key))
        if chinese:
            return chinese
        for prefix, value in self.prefixes:
            if len(prefix) >= 15 and key.startswith(prefix):
                return value
        return ""


def _page_tokens(page):
    """页面文字块：(左边界, 上边界, 右边界, 文本)。坐标原点在左上角，单位 pt。"""
    return [(word["x0"], word["top"], word["x1"], word["text"])
            for word in page.extract_words()]


def _rows(tokens, x_range, y_range):
    """按纵坐标把指定横坐标范围内的文字归成行。左右两栏单独处理，互不干扰。"""
    picked = [token for token in tokens
              if x_range[0] <= token[0] < x_range[1] and y_range[0] <= token[1] < y_range[1]]
    picked.sort(key=lambda token: (token[1], token[0]))
    rows = []
    for token in picked:
        if rows and abs(token[1] - rows[-1]["y"]) < ROW_TOLERANCE:
            rows[-1]["tokens"].append(token)
        else:
            rows.append({"y": token[1], "tokens": [token]})
    for row in rows:
        row["tokens"].sort(key=lambda token: token[0])
    return rows


def _row_text(row):
    return " ".join(token[3] for token in row["tokens"]).strip()


def _is_number(text):
    return bool(re.search(r"[0-9]", text))


def _nearest_number(y, numbers):
    """离标签最近的那个金额，避免跨到上下相邻行。"""
    best_gap, best_text = NUMBER_MAX_GAP, None
    for token in numbers:
        gap = abs(token[1] - y)
        if gap <= best_gap:
            best_gap, best_text = gap, token[3]
    return best_text


def _normalize_label(text):
    return re.sub(r"[\s:：]+", " ", text).strip().lower()


def _header_names(tokens):
    """页首右上角的 Display name 与 Legal name。"""
    rows = _rows(tokens, (NAME_LABEL_X, RIGHT_NUMBER_X[1]), HEADER_BAND)
    display = ""
    legal = ""
    ordered = []
    for row in rows:
        label = " ".join(token[3] for token in row["tokens"] if token[0] < NAME_SPLIT_X).strip()
        value = " ".join(token[3] for token in row["tokens"] if token[0] >= NAME_SPLIT_X).strip()
        if not label or not value:
            continue
        ordered.append(value)
        key = _normalize_label(label)
        if not legal and any(word in key for word in LEGAL_KEYS):
            legal = value
        elif not display and any(word in key for word in DISPLAY_KEYS):
            display = value
    if not display and ordered:
        display = ordered[0]
    if not legal and len(ordered) > 1:
        legal = ordered[1]
    return display, legal


def _summary_totals(tokens, keywords):
    """页首 Totals 区：左标签 + 右侧金额。"""
    totals = {}
    rows = _rows(tokens, TOTAL_LABEL_X, (SUMMARY_BAND[0], SUMMARY_BAND[1] + 6))
    numbers = [token for token in tokens
               if token[0] > TOTAL_NUMBER_X and _is_number(token[3])]
    for row in rows:
        text = _row_text(row).lower()
        for index, keyword in enumerate(keywords):
            if text != keyword.strip().lower():
                continue
            number = _nearest_number(row["y"], numbers)
            if number is not None:
                totals.setdefault(index, number)
    return totals


def _block_headers(tokens, keywords):
    """四个区块标题：本栏范围内只有一个词，且与关键词完全一致。"""
    headers = []
    for is_left in (True, False):
        x_range = LEFT_LABEL_X if is_left else RIGHT_LABEL_X
        rows = _rows(tokens, x_range, (SUMMARY_BAND[1], BLOCK_BOTTOM + 60))
        for row in rows:
            if len(row["tokens"]) != 1:
                continue
            text = _row_text(row).lower()
            for index, keyword in enumerate(keywords):
                if text == keyword.strip().lower():
                    headers.append((row["y"], is_left, index))
    headers.sort(key=lambda item: (item[0], not item[1]))
    return headers


def _split_amounts(numbers, is_left):
    """按列右边界区分借方与贷方。"""
    debit, credit = [], []
    for token in sorted(numbers, key=lambda item: item[0]):
        right_edge = token[2]
        if right_edge < (LEFT_DEBIT_RIGHT if is_left else RIGHT_DEBIT_RIGHT):
            debit.append(token[3])
        else:
            credit.append(token[3])
    return " ".join(debit).strip(), " ".join(credit).strip()


def _block_rows(tokens, y_min, y_max, is_left):
    """区块内逐项取标签与借贷金额，折行标签自动合并到同一项。"""
    label_x = LEFT_LABEL_X if is_left else RIGHT_LABEL_X
    number_x = LEFT_NUMBER_X if is_left else RIGHT_NUMBER_X

    label_rows = _rows(tokens, label_x, (y_min, y_max))
    if not label_rows:
        return []
    # 左栏的小计字样落在金额列内，要单独捞出来当标签行
    extras = [token for token in tokens
              if token[3].strip().lower() in SUBTOTAL_WORDS
              and number_x[0] <= token[0] < number_x[1]
              and y_min <= token[1] < y_max]
    for token in extras:
        if not any(abs(row["y"] - token[1]) < ROW_TOLERANCE for row in label_rows):
            label_rows.append({"y": token[1], "tokens": [token]})
    label_rows.sort(key=lambda row: row["y"])
    for row in label_rows:
        row["numbers"] = []

    numbers = [token for token in tokens
               if number_x[0] <= token[0] < number_x[1]
               and y_min <= token[1] < y_max and _is_number(token[3])]
    for token in numbers:
        nearest = min(range(len(label_rows)),
                      key=lambda index: abs(label_rows[index]["y"] - token[1]))
        if abs(label_rows[nearest]["y"] - token[1]) <= NUMBER_MAX_GAP:
            label_rows[nearest]["numbers"].append(token)

    rows = []
    pending = []
    for row in label_rows:
        text = _row_text(row)
        if not text or set(text) <= set("=- "):
            continue
        if not row["numbers"]:
            if rows:
                rows[-1]["label"] = (rows[-1]["label"] + " " + text).strip()
            else:
                pending.append(text)
            continue
        debit, credit = _split_amounts(row["numbers"], is_left)
        label = " ".join(pending + [text]).strip()
        pending = []
        rows.append({
            "label": label,
            "subtotal": text.lower() in SUBTOTAL_WORDS,
            "debit": debit,
            "credit": credit,
        })
    if pending and rows:
        rows[-1]["label"] = (rows[-1]["label"] + " " + " ".join(pending)).strip()
    return rows


def parse_summary(path, summary_rules, currency_mapping, number_formats):
    """解析一份 summary，返回四大字段合计与全部字段明细。"""
    keywords_map = summary_rules["keywords"]
    resolver = LabelResolver(summary_rules["label_map"])

    document = pdfplumber.open(path)
    try:
        pages = document.pages
        if not pages:
            return {"status": "fail", "message": "pdf 没有页面"}
        first_page_text = pages[0].extract_text() or ""
        first_text = re.sub(r"\s+", " ", first_page_text)
        currency = detect_currency_from_text(first_text, currency_mapping)
        if not currency or currency not in keywords_map:
            label = currency or "未知"
            return {"status": "fail", "message": "未能识别 summary 的币种（结果：%s）" % label}
        keywords = keywords_map[currency]
        decimal_sep, thousands_sep = resolve_number_format(currency, number_formats)

        period_start, period_end = parse_report_period(first_page_text)
        display_name, legal_name = _header_names(_page_tokens(pages[0]))

        totals = {}
        debits = {}
        credits = {}
        items = []
        warnings = []
        internal_diff = 0.0
        internal_missing = False

        for page in pages:
            tokens = _page_tokens(page)
            if page.page_number == 1:
                totals = _summary_totals(tokens, keywords)

            headers = _block_headers(tokens, keywords)
            for position, (start_y, is_left, category) in enumerate(headers):
                end_y = BLOCK_BOTTOM
                for next_y, next_left, _next in headers[position + 1:]:
                    if next_left == is_left:
                        end_y = min(next_y, BLOCK_BOTTOM)
                        break

                rows = _block_rows(tokens, start_y + 1.5, end_y, is_left)
                debit_sum = 0.0
                credit_sum = 0.0
                subtotal_debit = None
                subtotal_credit = None
                for row in rows:
                    if row["subtotal"]:
                        subtotal_debit = parse_amount(row["debit"], decimal_sep, thousands_sep)
                        subtotal_credit = parse_amount(row["credit"], decimal_sep, thousands_sep)
                        continue
                    debit = parse_amount(row["debit"], decimal_sep, thousands_sep)
                    credit = parse_amount(row["credit"], decimal_sep, thousands_sep)
                    debit_sum += debit
                    credit_sum += credit
                    items.append({
                        "category": CATEGORIES[category],
                        "label": row["label"],
                        "chinese": resolver.resolve(CATEGORIES[category], row["label"]),
                        "debit": debit,
                        "credit": credit,
                    })

                if subtotal_debit is None:
                    warnings.append("%s 区块没有找到小计行" % CATEGORIES[category])
                    internal_missing = True
                    subtotal_debit, subtotal_credit = debit_sum, credit_sum
                elif (abs(subtotal_debit - debit_sum) > 0.011
                      or abs(subtotal_credit - credit_sum) > 0.011):
                    warnings.append("%s 区块小计与明细合计不符（小计 %.2f，明细 %.2f）"
                                    % (CATEGORIES[category],
                                       subtotal_debit + subtotal_credit,
                                       debit_sum + credit_sum))
                    internal_diff += (abs(subtotal_debit - debit_sum)
                                      + abs(subtotal_credit - credit_sum))
                debits[CATEGORIES[category]] = subtotal_debit
                credits[CATEGORIES[category]] = subtotal_credit

        total_values = []
        for index, name in enumerate(CATEGORIES):
            value = parse_amount(totals.get(index, ""), decimal_sep, thousands_sep)
            total_values.append(value)
            block_sum = debits.get(name, 0.0) + credits.get(name, 0.0)
            if abs(block_sum - value) > 0.011:
                warnings.append("%s 合计与区块借贷合计不符（%.2f 对 %.2f）"
                                % (name, value, block_sum))
                internal_diff += abs(value - block_sum)

        if len(totals) < 4:
            warnings.append("页首四大字段没有全部识别到")
            internal_missing = True

        return {
            "status": "ok",
            "currency": currency,
            "display_name": display_name,
            "legal_name": legal_name,
            "period_start": period_start,
            "period_end": period_end,
            "totals": dict(zip(CATEGORIES, total_values)),
            "debits": debits,
            "credits": credits,
            "items": items,
            "internal_diff": round(internal_diff, 2),
            "internal_missing": internal_missing,
            "warnings": warnings,
        }
    finally:
        document.close()

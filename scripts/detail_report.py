# -*- coding: utf-8 -*-
"""日期范围报告（csv 明细）整理。"""

import csv
import os

from common import (
    DEFAULT_HEADER_KEYWORDS,
    ENCODINGS,
    cell_text,
    detect_currency_from_text,
    normalize_datetime,
)

HEAD_LINE_LIMIT = 40
MIN_HEADER_COMMAS = 5
HEADER_KEY_FIELDS = ("date_time", "type", "order_id")


def read_head_lines(path, limit=HEAD_LINE_LIMIT):
    """读出文件头部若干行，用来判断编码、货币和表头位置。"""
    for encoding in ENCODINGS:
        try:
            with open(path, "r", encoding=encoding) as handle:
                lines = []
                for _ in range(limit):
                    line = handle.readline()
                    if not line:
                        break
                    lines.append(line.rstrip("\r\n"))
            return lines, encoding
        except UnicodeDecodeError:
            continue
        except OSError:
            break
    return [], "utf-8"


def detect_currency(lines, mapping):
    """货币说明句优先，关键词兜底，全部忽略大小写。"""
    return detect_currency_from_text(" ".join(lines[:15]), mapping)


def find_header_row(lines, keywords, min_commas=MIN_HEADER_COMMAS):
    """找同时包含各关键字段名、逗号数达标的行。"""
    lowered_keys = [key.lower() for key in keywords if key]
    for index, raw in enumerate(lines):
        line = raw.lstrip("\ufeff")
        if line.count(",") < min_commas:
            continue
        lowered = line.lower()
        if lowered_keys and all(key in lowered for key in lowered_keys):
            return index
    return 0


def build_header_keywords(field_map):
    """用当前币种的原文列名做表头识别关键词。"""
    keywords = []
    for name in HEADER_KEY_FIELDS:
        source = cell_text(field_map.get(name, ""))
        if source:
            keywords.append(source)
    return keywords or list(DEFAULT_HEADER_KEYWORDS)


def read_csv_rows(path, encoding, header_row):
    """跳过说明行后按 csv 规则读取，支持字段内换行。"""
    with open(path, "r", encoding=encoding, newline="") as handle:
        reader = csv.reader(handle)
        for _ in range(header_row):
            next(reader, None)
        header = next(reader, None)
        if header is None:
            return [], []
        header = [cell_text(name) for name in header]
        rows = []
        for row in reader:
            if not row or all(not cell.strip() for cell in row):
                continue
            rows.append(row)
    return header, rows


def build_column_lookup(header):
    """原文列名 -> 列下标，大小写不敏感。"""
    lookup = {}
    for index, name in enumerate(header):
        key = name.lower()
        if key and key not in lookup:
            lookup[key] = index
    return lookup


def process_detail(path, rules):
    """读取一份日期范围报告，返回标准字段行。"""
    website = os.path.splitext(os.path.basename(path))[0].strip()
    lines, encoding = read_head_lines(path)
    if not lines:
        return {"website": website, "currency": "", "rows": [], "status": "fail",
                "message": "文件读取失败或为空"}

    currency = detect_currency(lines, rules["currency_mapping"])
    if not currency or currency not in rules["currency_maps"]:
        label = currency or "未知"
        return {"website": website, "currency": label, "rows": [], "status": "fail",
                "message": "未能识别币种或该币种没有字段映射（识别结果：%s）" % label}

    field_map = rules["currency_maps"][currency]
    keywords = build_header_keywords(field_map)
    header_row = find_header_row(lines, keywords)

    try:
        header, raw_rows = read_csv_rows(path, encoding, header_row)
    except Exception as error:
        return {"website": website, "currency": currency, "rows": [], "status": "fail",
                "message": "读取 csv 失败：%s" % error}

    if not header:
        return {"website": website, "currency": currency, "rows": [], "status": "fail",
                "message": "没有读到表头行"}
    if not raw_rows:
        return {"website": website, "currency": currency, "rows": [], "status": "skip",
                "message": "没有数据行"}

    lookup = build_column_lookup(header)
    columns = rules["columns"]
    prepared = []
    for column in columns:
        if column == "website":
            prepared.append(("const", website))
        elif column == "currency":
            prepared.append(("const", currency))
        elif column == "behavior":
            prepared.append(("empty", ""))
        else:
            source = cell_text(field_map.get(column, "")).lower()
            prepared.append(("column", lookup.get(source)))

    date_index = None
    release_index = None
    for column, source in zip(columns, prepared):
        if source[0] != "column":
            continue
        if column == "date_time":
            date_index = source[1]
        elif column == "transaction_release_date":
            release_index = source[1]

    rows = []
    for raw in raw_rows:
        record = {}
        for column, (kind, value) in zip(columns, prepared):
            if kind == "const":
                record[column] = value
            elif kind == "column":
                if value is not None and value < len(raw):
                    record[column] = cell_text(raw[value])
                else:
                    record[column] = ""
            else:
                record[column] = ""
        if date_index is not None:
            record["date_time"] = normalize_datetime(record.get("date_time", ""))
        if release_index is not None:
            normalized = normalize_datetime(record.get("transaction_release_date", ""))
            if normalized:
                record["transaction_release_date"] = normalized
        rows.append(record)

    return {"website": website, "currency": currency, "rows": rows, "status": "ok",
            "message": ""}

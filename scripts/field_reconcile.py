# -*- coding: utf-8 -*-
"""summary 字段与日期范围报告明细的字段级核对。

每个 summary 费用项按配置的取数规则回明细里汇总，比两边金额。
"""

import os

from common import cell_text, period_code, read_config_rows

MAPPING_FILE = "report_field_mapping.csv"
CONSISTENT = "核对一致"
UNMAPPED = "未配置"
FIELD_HEADERS = ["期间", "文件名", "分类", "summary字段", "中文标准字段", "summary金额",
                 "明细金额", "差异", "日期范围报告明细的取数规则", "核对结果"]


def load_reconcile_mapping(config_dir):
    """读映射表：summary 字段 -> 明细的金额列、行为、符号条件。"""
    path = os.path.join(config_dir, MAPPING_FILE)
    rules = []
    for row in read_config_rows(path)[1:]:
        field = cell_text(row[0]) if row else ""
        if not field:
            continue
        rules.append({
            "field": field,
            "category": cell_text(row[1]) if len(row) > 1 else "",
            "currency": cell_text(row[2]).upper() if len(row) > 2 else "",
            "columns": [item.strip() for item in cell_text(row[3]).split("|") if item.strip()],
            "behaviors": [item.strip() for item in cell_text(row[4]).split("|") if item.strip()],
            "types": [item.strip().lower() for item in cell_text(row[5]).split("|") if item.strip()],
            "keywords": [item.strip().lower() for item in cell_text(row[6]).split("|") if item.strip()],
            "condition": (cell_text(row[7]).lower() if len(row) > 7 else "any") or "any",
            "group": cell_text(row[8]) if len(row) > 8 else "",
            "note": cell_text(row[9]) if len(row) > 9 else "",
        })
    return rules


def _behavior_match(behavior, patterns):
    """行为名匹配，支持结尾通配符。"""
    if not patterns:
        return True
    for pattern in patterns:
        if pattern == behavior:
            return True
        if pattern.endswith("*") and behavior.startswith(pattern[:-1]):
            return True
    return False


def detail_sum(rows, rule):
    """按规则把明细行累加成 summary 口径的金额。"""
    total = 0.0
    for row in rows:
        if not _behavior_match(row.get("behavior", ""), rule["behaviors"]):
            continue
        if rule["types"] and cell_text(row.get("type", "")).lower() not in rule["types"]:
            continue
        if rule["keywords"]:
            description = cell_text(row.get("description", "")).lower()
            if not any(word in description for word in rule["keywords"]):
                continue
        value = 0.0
        for column in rule["columns"]:
            value += row.get(column, 0.0) or 0.0
        condition = rule["condition"]
        if condition == "pos" and value <= 0:
            continue
        if condition == "neg" and value >= 0:
            continue
        if condition.startswith("zero:"):
            if abs(row.get(condition.split(":", 1)[1], 0.0) or 0.0) > 0.001:
                continue
        if condition.startswith("nonzero:"):
            if abs(row.get(condition.split(":", 1)[1], 0.0) or 0.0) <= 0.001:
                continue
        total += value
    return round(total, 2)


def _rule_text(rules):
    if isinstance(rules, dict):
        rules = [rules]
    texts = []
    for rule in rules:
        if not rule["columns"]:
            texts.append("—")
            continue
        text = "%s / %s" % ("+".join(rule["columns"]),
                            "|".join(rule["behaviors"]) if rule["behaviors"] else "全部行为")
        condition = rule["condition"]
        if condition == "pos":
            text += " / 正数"
        elif condition == "neg":
            text += " / 负数"
        elif condition.startswith("zero:"):
            text += " / %s=0" % condition.split(":", 1)[1]
        elif condition.startswith("nonzero:"):
            text += " / %s≠0" % condition.split(":", 1)[1]
        texts.append(text)
    return " + ".join(texts)


def _pick_rules(rules, category, field, currency):
    """同一 summary 字段可以有多条规则，按 type 或行为分流后相加。

    币种专属规则存在时只用专属规则，否则用全部通用规则；两条规则的过滤条件
    必须互斥，否则同一笔明细会被加两次。
    """
    matched = [rule for rule in rules
               if rule["field"] == field and rule["category"] == category]
    specific = [rule for rule in matched if rule["currency"] and rule["currency"] == currency]
    if specific:
        return specific
    return [rule for rule in matched if not rule["currency"]]


def build_field_rows(stem, summary_result, detail_rows, mapping, tolerance):
    """一组一份，按 summary 字段逐项核对，同一核对组的多行合并成一行。

    金额为 0 且没有映射的字段直接跳过，避免整表被零值字段淹没；
    末尾追加一行合计，用来看重分类差异是否互抵。
    """
    currency = summary_result.get("currency", "")
    period = period_code(summary_result.get("period_start", ""),
                         summary_result.get("period_end", ""))
    buckets = []
    for item in summary_result.get("items", []):
        category, label, chinese = item["category"], item["label"], item["chinese"]
        net = round(item["debit"] + item["credit"], 2)
        rule = None
        rules = _pick_rules(mapping, category, chinese, currency) if chinese else []
        if rules:
            rule = rules[0]
        if rule is None:
            if abs(net) > 0.004:
                buckets.append({"rule": None, "category": category, "labels": [label],
                                "chinese": [chinese or "（未翻译）"], "net": net})
            continue
        key = (category, rule["group"] or rule["field"])
        same = next((bucket for bucket in buckets if bucket.get("key") == key), None)
        if same:
            same["labels"].append(label)
            if chinese not in same["chinese"]:
                same["chinese"].append(chinese)
            same["net"] = round(same["net"] + net, 2)
        else:
            buckets.append({"key": key, "rule": rule, "rules": rules, "category": category,
                            "labels": [label], "chinese": [chinese], "net": net})

    rows = []
    stats = {"consistent": 0, "diff": 0, "unmapped": 0}
    checked_net = 0.0
    checked_detail = 0.0
    for bucket in buckets:
        label_text = " / ".join(bucket["labels"])
        chinese_text = " / ".join(bucket["chinese"])
        if bucket["rule"] is None:
            rows.append([period, stem, bucket["category"], label_text, chinese_text,
                         bucket["net"], "", "", "—", UNMAPPED])
            stats["unmapped"] += 1
            continue
        total = round(sum(detail_sum(detail_rows, item) for item in bucket["rules"]), 2)
        diff = round(bucket["net"] - total, 2)
        result = CONSISTENT if abs(diff) <= tolerance else "差额"
        stats["consistent" if result == CONSISTENT else "diff"] += 1
        checked_net += bucket["net"]
        checked_detail += total
        rows.append([period, stem, bucket["category"], label_text, chinese_text,
                     bucket["net"], total, diff, _rule_text(bucket["rules"]), result])

    if rows:
        diff = round(checked_net - checked_detail, 2)
        rows.append([period, stem, "合计", "已核对字段合计", "——",
                     round(checked_net, 2), round(checked_detail, 2), diff, "——",
                     CONSISTENT if abs(diff) <= tolerance else "差额"])
    return rows, stats

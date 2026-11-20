# -*- coding: utf-8 -*-
"""字段核对：映射、取数、合并组与合计行的契约测试。"""

import os
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(SKILL, "scripts"))

from field_reconcile import (  # noqa: E402
    CONSISTENT,
    build_field_rows,
    detail_sum,
    load_reconcile_mapping,
)


def rule(**kwargs):
    base = {"field": "", "category": "", "currency": "", "columns": [],
            "behaviors": [], "types": [], "keywords": [], "condition": "any",
            "group": "", "note": ""}
    base.update(kwargs)
    return base


MAPPING = [
    rule(field="商品销售", category="收入", columns=["product_sales"], behaviors=["商品销售"]),
    rule(field="成交费退款", category="费用", columns=["selling_fees"], behaviors=["销售退款"], group="退款组"),
    rule(field="退款管理费", category="费用", columns=["selling_fees"], behaviors=["销售退款"], group="退款组"),
    rule(field="赔偿", category="收入", columns=["total"], behaviors=["调整-*"], condition="pos"),
    rule(field="扣款", category="费用", columns=["total"], behaviors=["调整-*"], condition="neg"),
    rule(field="日元服务费", category="费用", currency="JPY", columns=["other"], behaviors=["平台服务费-*"]),
    rule(field="通用服务费", category="费用", columns=["other"], behaviors=["平台服务费-*"]),
]

ROWS = [
    {"behavior": "商品销售", "product_sales": 100.0, "selling_fees": 0.0, "total": 100.0, "other": 0.0},
    {"behavior": "销售退款", "product_sales": -20.0, "selling_fees": -5.0, "total": -25.0, "other": 0.0},
    {"behavior": "调整-客户退货", "product_sales": 0.0, "selling_fees": 0.0, "total": 7.0, "other": 0.0},
    {"behavior": "调整-一般调整", "product_sales": 0.0, "selling_fees": 0.0, "total": -3.0, "other": 0.0},
    {"behavior": "平台服务费-订阅费（月服务费）", "product_sales": 0.0, "selling_fees": 0.0, "total": -4.0, "other": -4.0},
]


def summary(items, currency="USD", period=("2026-01-01", "2026-01-31")):
    return {"currency": currency, "period_start": period[0], "period_end": period[1], "items": [
        {"category": c, "label": label, "chinese": label, "debit": debit, "credit": credit}
        for c, label, debit, credit in items
    ]}


class DetailSumTest(unittest.TestCase):
    def test_column_and_behavior(self):
        self.assertEqual(detail_sum(ROWS, MAPPING[0]), 100.0)

    def test_wildcard_behavior(self):
        self.assertEqual(detail_sum(ROWS, MAPPING[3]), 7.0)

    def test_positive_and_negative_filters(self):
        self.assertEqual(detail_sum(ROWS, MAPPING[3]), 7.0)
        self.assertEqual(detail_sum(ROWS, MAPPING[4]), -3.0)

    def test_type_and_description_filters(self):
        by_type = rule(columns=["total"], types=["order"])
        self.assertEqual(detail_sum(ROWS, by_type), 0.0)
        ROWS.append({"type": "Order", "description": "Cost of advertising", "behavior": "商品销售",
                     "product_sales": 0.0, "selling_fees": 0.0, "total": 10.0, "other": 0.0})
        hit = rule(columns=["total"], types=["order"], keywords=["cost of advertising"])
        self.assertEqual(detail_sum(ROWS, hit), 10.0)
        miss = rule(columns=["total"], types=["order"], keywords=["subscription"])
        self.assertEqual(detail_sum(ROWS, miss), 0.0)
        ROWS.pop()

    def test_unknown_behavior_returns_zero(self):
        self.assertEqual(detail_sum(ROWS, rule(columns=["total"], behaviors=["不存在"])), 0.0)


class FieldRowsTest(unittest.TestCase):
    def test_group_merges_two_summary_rows(self):
        data = summary([
            ("费用", "成交费退款", -3.0, 0.0),
            ("费用", "退款管理费", -2.0, 0.0),
        ])
        rows, stats = build_field_rows("demo", data, ROWS, MAPPING, 0.01)
        merged = [r for r in rows if " / " in str(r[3])]
        self.assertEqual(len(merged), 1)
        self.assertEqual(merged[0][5], -5.0)
        self.assertEqual(merged[0][6], -5.0)
        self.assertEqual(merged[0][9], CONSISTENT)
        self.assertEqual(stats["consistent"], 1)

    def test_zero_unmapped_field_is_skipped(self):
        data = summary([("收入", "没有映射的零值字段", 0.0, 0.0)])
        rows, stats = build_field_rows("demo", data, ROWS, MAPPING, 0.01)
        self.assertEqual(stats["unmapped"], 0)
        self.assertEqual(rows, [])

    def test_total_row_present(self):
        data = summary([("收入", "商品销售", 0.0, 100.0)])
        rows, _stats = build_field_rows("demo", data, ROWS, MAPPING, 0.01)
        total = rows[-1]
        self.assertEqual(total[0], "202601")
        self.assertEqual(total[1], "demo")
        self.assertEqual(total[2], "合计")
        self.assertEqual(total[5], 100.0)
        self.assertEqual(total[7], 0.0)
        self.assertEqual(total[9], CONSISTENT)

    def test_diff_is_flagged(self):
        data = summary([("收入", "商品销售", 0.0, 99.0)])
        rows, stats = build_field_rows("demo", data, ROWS, MAPPING, 0.01)
        self.assertEqual(rows[0][9], "差额")
        self.assertEqual(rows[0][7], -1.0)
        self.assertEqual(stats["diff"], 1)


class MultiRuleTest(unittest.TestCase):
    """同一字段挂多条规则时，按行为分流后相加，且币种专属优先。"""

    def test_same_field_sums_all_rules(self):
        mapping = [
            rule(field="配送费", category="费用", columns=["fba_fees"], behaviors=["商品销售"]),
            rule(field="配送费", category="费用", columns=["fba_fees"],
                 behaviors=["FBA费用-评级与转售计划收费"]),
        ]
        rows = [
            {"behavior": "商品销售", "fba_fees": -10.0, "total": -10.0},
            {"behavior": "FBA费用-评级与转售计划收费", "fba_fees": -2.0, "total": -2.0},
        ]
        data = summary([("费用", "配送费", -12.0, 0.0)])
        out, _stats = build_field_rows("demo", data, rows, mapping, 0.01)
        self.assertEqual(out[0][6], -12.0)
        self.assertEqual(out[0][9], CONSISTENT)
        self.assertIn("+", out[0][8])

    def test_currency_specific_rule_wins(self):
        mapping = [
            rule(field="服务费", category="费用", columns=["other"], behaviors=["平台服务费-*"]),
            rule(field="服务费", category="费用", currency="JPY", columns=["total"],
                 behaviors=["平台服务费-*"]),
        ]
        rows = [{"behavior": "平台服务费-广告费", "other": -4.0, "total": -9.0}]
        jpy, _stats = build_field_rows("demo", summary([("费用", "服务费", -9.0, 0.0)], "JPY"),
                                       rows, mapping, 0.01)
        self.assertEqual(jpy[0][6], -9.0)
        usd, _stats = build_field_rows("demo", summary([("费用", "服务费", -4.0, 0.0)], "USD"),
                                       rows, mapping, 0.01)
        self.assertEqual(usd[0][6], -4.0)


class MappingConfigTest(unittest.TestCase):
    def test_mapping_loads(self):
        mapping = load_reconcile_mapping(os.path.join(SKILL, "config"))
        self.assertGreaterEqual(len(mapping), 20)
        for item in mapping:
            self.assertTrue(item["field"])
            self.assertTrue(item["category"])
        fields = {item["field"] for item in mapping}
        self.assertIn("FBA商品销售收入", fields)
        self.assertIn("转账至银行账户", fields)

    def test_currency_specific_rule_exists(self):
        mapping = load_reconcile_mapping(os.path.join(SKILL, "config"))
        jpy = [item for item in mapping if item["field"] == "平台服务费-订阅费"]
        self.assertTrue(jpy)
        self.assertEqual(jpy[0]["currency"], "JPY")


if __name__ == "__main__":
    unittest.main()

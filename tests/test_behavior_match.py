# -*- coding: utf-8 -*-
"""平台行为匹配的契约测试：type 与关键词必须同时满足，按表序命中即停。"""

import os
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(SKILL, "scripts"))

from common import (  # noqa: E402
    detect_currency_from_text,
    load_behavior_rules,
    load_currency_mapping,
    load_field_rules,
    load_number_formats,
    load_summary_rules,
    match_behavior,
    parse_amount,
    prepare_rules,
    resolve_number_format,
    rule_matches,
)


RULES = [
    ("仓储费-退回", ["FBA Inventory Fee - Reversal"], ["FBA storage fee"]),
    ("仓储费", ["FBA Inventory Fee"], ["FBA storage fee"]),
    ("其他库存费", ["FBA Inventory Fee"], []),
    ("广告费", ["Service Fee"], ["Cost of advertising"]),
    ("服务费兜底", ["Service Fee"], []),
    ("纯关键词规则", [], ["Special Marker"]),
]


class BehaviorMatchTest(unittest.TestCase):
    def setUp(self):
        self.prepared = prepare_rules(RULES)

    def match(self, type_text, description):
        return match_behavior(type_text, description, RULES, self.prepared)

    def test_both_columns_must_match(self):
        """type 命中但关键词命中不了细分规则时，落到同 type 的兜底行。"""
        self.assertEqual(self.match("Service Fee", "Subscription"), "服务费兜底")

    def test_type_hit_without_keyword_falls_through(self):
        """type 命中、关键词不命中，且后面没有兜底行时，不得强行归类。"""
        self.assertEqual(self.match("FBA Inventory Fee - Reversal", "FBA Long-Term Storage Fee"), "")

    def test_keyword_hit_without_type_is_rejected(self):
        """关键词命中但 type 不符，不得命中该规则。"""
        self.assertEqual(self.match("Order", "Cost of advertising"), "")

    def test_type_only_rule_matches_by_type(self):
        """只写 type 的兜底行，描述为空也能命中。"""
        self.assertEqual(self.match("FBA Inventory Fee", ""), "其他库存费")

    def test_keyword_only_rule_matches_by_keyword(self):
        """只写关键词的规则，不限制 type。"""
        self.assertEqual(self.match("Order", "has a special marker inside"), "纯关键词规则")

    def test_first_match_wins(self):
        """两条规则都满足时，取表里靠上的那条。"""
        self.assertEqual(self.match("FBA Inventory Fee", "FBA storage fee"), "仓储费")

    def test_case_insensitive(self):
        self.assertEqual(self.match("SERVICE FEE", "cost of ADVERTISING"), "广告费")

    def test_no_rule_matches(self):
        self.assertEqual(self.match("Order", "nothing here"), "")

    def test_empty_rules(self):
        self.assertEqual(match_behavior("Order", "x", [], []), "")

    def test_cache_returns_same_result(self):
        cache = {}
        first = match_behavior("Service Fee", "Subscription", RULES, self.prepared, cache)
        second = match_behavior("Service Fee", "Subscription", RULES, self.prepared, cache)
        self.assertEqual(first, second)
        self.assertEqual(len(cache), 1)

    def test_blank_cells_are_wildcards(self):
        self.assertTrue(rule_matches(set(), [], "anything", "anything"))
        self.assertTrue(rule_matches(set(), ["needle"], "x", "a needle b"))
        self.assertFalse(rule_matches(set(), ["needle"], "x", "a neddle b"))


class NumberFormatTest(unittest.TestCase):
    def test_default_format(self):
        self.assertEqual(parse_amount("1,234.56"), 1234.56)
        self.assertEqual(parse_amount("(1,234.56)"), -1234.56)
        self.assertEqual(parse_amount("1,234.56", ".", ","), 1234.56)

    def test_european_format(self):
        """欧洲大陆格式：小数逗号、千分位点号。"""
        self.assertEqual(parse_amount("1.234,56", ",", "."), 1234.56)
        self.assertEqual(parse_amount("-3.031,00", ",", "."), -3031.00)
        self.assertEqual(parse_amount("(6033,17)", ",", "."), -6033.17)

    def test_resolve_by_currency(self):
        formats = load_number_formats(os.path.join(SKILL, "config", "common_number_mapping.csv"))
        self.assertEqual(resolve_number_format("BRL", formats), (",", "."))
        self.assertEqual(resolve_number_format("USD", formats), (".", ","))
        self.assertEqual(resolve_number_format("EUR", formats), (".", ","))


class RuleConfigTest(unittest.TestCase):
    """真实配置的体检：每个 csv 都能读出内容，规则有条件，句子映射非空。"""

    CONFIG_DIR = os.path.join(SKILL, "config")

    def test_every_rule_has_a_condition(self):
        rules = load_behavior_rules(os.path.join(self.CONFIG_DIR, "detail_report_behavior.csv"))
        self.assertTrue(rules, "行为规则表为空")
        for name, types, keywords in rules:
            with self.subTest(rule=name):
                self.assertTrue(types or keywords, "规则 %s 既没写 type 也没写关键词" % name)

    def test_field_rules_load(self):
        try:
            field_rules = load_field_rules(self.CONFIG_DIR)
        except Exception as error:
            self.fail("明细配置读取失败：%s" % error)
        self.assertTrue(field_rules["columns"])
        self.assertTrue(field_rules["currency_maps"])
        self.assertTrue(field_rules["currency_mapping"]["sentences"], "货币说明句映射为空")

    def test_summary_rules_load(self):
        try:
            summary_rules = load_summary_rules(self.CONFIG_DIR)
        except Exception as error:
            self.fail("summary 配置读取失败：%s" % error)
        self.assertTrue(summary_rules["keywords"])
        self.assertTrue(summary_rules["label_map"])

    def test_currency_detection_by_sentence(self):
        """货币识别走通用说明句，csv 与 pdf 两侧文字一致，忽略大小写。"""
        mapping = load_currency_mapping(os.path.join(self.CONFIG_DIR, "common_currency_mapping.csv"))
        cases = [
            ("Account activity from Mar 1, 2026 00:00 PST through Mar 31, 2026 ALL AMOUNTS IN USD, unless specified", "USD"),
            ("Actividad de la cuenta ... TODOS LOS IMPORTES EN MXN, a menos que se especifique", "MXN"),
            ("Atividade de conta ... TODAS AS QUANTIAS EM BRL, a menos que especificado", "BRL"),
            ("... までのアカウント履歴 指定のない場合、単位は円", "JPY"),
        ]
        for text, code in cases:
            with self.subTest(currency=code):
                self.assertEqual(detect_currency_from_text(text, mapping), code)

    def test_currency_mapping_shape(self):
        mapping = load_currency_mapping(os.path.join(self.CONFIG_DIR, "common_currency_mapping.csv"))
        self.assertTrue(mapping["sentences"])
        self.assertTrue(mapping["keywords"])
        codes = {code for _text, code in mapping["sentences"]}
        self.assertEqual(codes, {"USD", "CAD", "MXN", "BRL", "JPY"})

    def test_keyword_fallback(self):
        """说明句不完整时，靠关键词兜底。"""
        mapping = load_currency_mapping(os.path.join(self.CONFIG_DIR, "common_currency_mapping.csv"))
        self.assertEqual(detect_currency_from_text("Amounts in usd unless noted", mapping), "USD")
        self.assertEqual(detect_currency_from_text("Atividade ... em brl", mapping), "BRL")


if __name__ == "__main__":
    unittest.main()

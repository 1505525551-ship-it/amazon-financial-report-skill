# -*- coding: utf-8 -*-
"""配置加载与通用工具：CSV 配置、通用货币识别、数值格式、日期标准化、行为匹配。

所有配置匹配一律忽略大小写。
"""

import csv
import os
import re
from datetime import datetime

BEHAVIOR_FILE = "detail_report_behavior.csv"
FIELD_FILE = "detail_report_field.csv"
CURRENCY_FILE = "common_currency_mapping.csv"
NUMBER_FORMAT_FILE = "common_number_mapping.csv"
KEYWORD_FILE = "summary_report_keyword.csv"
LABEL_FILE = "summary_report_field.csv"

CURRENCY_PATTERN = re.compile(r"^[A-Z]{3}$")
DEFAULT_HEADER_KEYWORDS = ("date/time", "type", "order id")
DEFAULT_FORMAT = (".", ",")  # 未列出的币种按中英美格式解析
ENCODINGS = ("utf-8-sig", "utf-8", "gbk", "latin-1")
CONFIG_ENCODINGS = ("utf-8-sig", "gbk")

MONTHS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
    "ene": 1, "abr": 4, "ago": 8, "dic": 12, "set": 9, "out": 10, "dez": 12,
    "fev": 2, "mai": 5, "mrz": 3, "okt": 10,
}

TIMEZONE_PATTERN = re.compile(
    r"\s*(?:PDT|PST|EDT|EST|CDT|CST|MDT|MST|GMT[+-]\d{1,2}(?::\d{2})?|UTC[+-]?\d*"
    r"|CET|CEST|BST|JST|BRT|ART|AEST|AEDT)\s*$",
    re.IGNORECASE,
)
AMPM_PATTERN = re.compile(r"\b([ap])\.?\s*m\.?", re.IGNORECASE)

PATTERN_DAY_FIRST = re.compile(
    r"^(\d{1,2})\s+([A-Za-z]{3,4})\.?\s+(\d{4})\s+(\d{1,2}):(\d{2}):(\d{2})(?:\s*(AM|PM))?$"
)
PATTERN_MONTH_FIRST = re.compile(
    r"^([A-Za-z]{3,4})\.?\s+(\d{1,2}),?\s+(\d{4})\s+(\d{1,2}):(\d{2}):(\d{2})(?:\s*(AM|PM))?$"
)
PATTERN_ISO = re.compile(r"^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2}):(\d{2})")


class ConfigError(Exception):
    """配置文件缺失或结构不符。"""


def cell_text(value):
    """统一转成去掉首尾空白的纯文本。"""
    if value is None:
        return ""
    return str(value).replace("\ufeff", "").strip()


def normalize_key(text):
    """用于对照表匹配的归一化文本：忽略大小写与所有空白。"""
    return re.sub(r"\s+", "", cell_text(text)).lower()


def normalize_sentence(text):
    """货币说明句匹配用的归一化：忽略大小写，空白压成一个空格。"""
    return re.sub(r"\s+", " ", cell_text(text)).lower()


def split_list(value):
    """逗号分隔的配置值转列表，去空项。"""
    return [item.strip() for item in cell_text(value).split(",") if item.strip()]


def read_config_rows(path):
    """读一个配置 csv，返回去掉全空行的二维列表。

    配置约定 UTF-8；Excel 另存可能变成 GBK，这里按 utf-8-sig -> gbk 顺序回退。
    """
    if not os.path.isfile(path):
        raise ConfigError("配置文件不存在：%s" % path)
    for encoding in CONFIG_ENCODINGS:
        try:
            with open(path, "r", encoding=encoding, newline="") as handle:
                rows = [row for row in csv.reader(handle) if any(cell_text(cell) for cell in row)]
            break
        except UnicodeDecodeError:
            continue
        except OSError as error:
            raise ConfigError("无法读取 %s：%s" % (path, error))
    else:
        raise ConfigError("配置文件编码无法识别（需要 UTF-8）：%s" % path)
    if not rows:
        raise ConfigError("配置文件是空的：%s" % path)
    return rows


# ---------------------------------------------------------------- 明细规则


def load_behavior_rules(path):
    """behavior,type_exact,description_keywords -> (名称, type 列表, 关键词列表)。"""
    rules = []
    for row in read_config_rows(path)[1:]:
        name = cell_text(row[0]) if row else ""
        if not name:
            continue
        types = split_list(row[1]) if len(row) > 1 else []
        keywords = split_list(row[2]) if len(row) > 2 else []
        rules.append((name, types, keywords))
    return rules


def load_field_sheet(path):
    """标准字段、中文名、各币种原文列名。"""
    rows = read_config_rows(path)
    header = [cell_text(value) for value in rows[0]]
    currency_columns = {}
    for index in range(2, len(header)):
        code = header[index].strip().upper()
        if CURRENCY_PATTERN.match(code):
            currency_columns[code] = index

    columns = []
    labels = []
    currency_maps = {code: {} for code in currency_columns}
    for row in rows[1:]:
        name = cell_text(row[0]) if row else ""
        if not name:
            continue
        columns.append(name)
        labels.append(cell_text(row[1]) if len(row) > 1 else "")
        for code, index in currency_columns.items():
            currency_maps[code][name] = cell_text(row[index]) if index < len(row) else ""
    return columns, labels, currency_maps


def load_currency_mapping(path):
    """货币说明句与兜底关键词 -> 币种，csv 与 pdf 通用，忽略大小写与多余空白。"""
    mapping = {"sentences": [], "keywords": []}
    for row in read_config_rows(path)[1:]:
        sentence = cell_text(row[0]) if len(row) > 0 else ""
        keyword = cell_text(row[1]) if len(row) > 1 else ""
        code = cell_text(row[2]).upper() if len(row) > 2 else ""
        if not code:
            continue
        if sentence:
            mapping["sentences"].append((sentence, code))
        if keyword:
            mapping["keywords"].append((keyword, code))
    return mapping


def detect_currency_from_text(text, mapping):
    """先按货币说明句匹配，没命中再用关键词兜底，全部忽略大小写。"""
    flattened = normalize_sentence(text)
    for sentence, code in mapping["sentences"]:
        if normalize_sentence(sentence) in flattened:
            return code
    for keyword, code in mapping["keywords"]:
        if keyword and normalize_sentence(keyword) in flattened:
            return code
    return ""


def period_code(period_start, period_end):
    """报表期间写成 yyyymm，优先用起始日期。"""
    for value in (period_start, period_end):
        if value and len(value) >= 7:
            return value[:4] + value[5:7]
    return ""


# ---------------------------------------------------------------- 通用数值格式


def load_number_formats(path):
    """数值格式,小数分隔符,千分位分隔符,适用币种 -> {币种: (小数, 千分位)}。"""
    formats = {}
    for row in read_config_rows(path)[1:]:
        name = cell_text(row[0]) if row else ""
        decimal = cell_text(row[1]) if len(row) > 1 else "."
        thousands = cell_text(row[2]) if len(row) > 2 else ","
        codes = cell_text(row[3]) if len(row) > 3 else ""
        if not name:
            continue
        for code in re.split(r"[;；,，\s]+", codes):
            if code:
                formats[code.upper()] = (decimal or ".", thousands or ",")
    return {"formats": formats, "default": DEFAULT_FORMAT}


def resolve_number_format(currency, number_formats):
    """取该币种的 (小数分隔符, 千分位分隔符)，未列出用中英美格式。"""
    return number_formats["formats"].get(cell_text(currency).upper(), number_formats["default"])


def parse_amount(value, decimal_sep=".", thousands_sep=","):
    """金额文本转浮点数，先去掉千分位分隔符，再把小数分隔符转成点号。"""
    text = cell_text(value).replace("\u00a0", " ")
    if not text or text.lower() in ("nan", "none", "--", "-"):
        return 0.0
    negative = text.startswith("(") and text.endswith(")")
    if negative:
        text = text[1:-1]
    text = re.sub(r"[^0-9,.\-]", "", text)
    if not text or text in ("-", ".", ","):
        return 0.0
    if thousands_sep and thousands_sep != decimal_sep:
        text = text.replace(thousands_sep, "")
    if decimal_sep != ".":
        text = text.replace(decimal_sep, ".")
    try:
        number = float(text)
    except ValueError:
        return 0.0
    return -number if negative else number


# ---------------------------------------------------------------- summary 规则


def load_keywords(path):
    """币种 + 收入/费用/转账/税 四个关键词。"""
    result = {}
    for row in read_config_rows(path)[1:]:
        code = cell_text(row[0]).upper()
        if not CURRENCY_PATTERN.match(code):
            continue
        words = [cell_text(row[i]) if len(row) > i else "" for i in range(1, 5)]
        if all(words):
            result[code] = words
    return result


def load_label_map(path):
    """分类、原文、中文标准字段。"""
    entries = []
    for row in read_config_rows(path)[1:]:
        category = cell_text(row[0]) if len(row) > 0 else ""
        original = cell_text(row[1]) if len(row) > 1 else ""
        chinese = cell_text(row[2]) if len(row) > 2 else ""
        if original and chinese:
            entries.append((category, normalize_key(original), chinese))
    return entries


# ---------------------------------------------------------------- 组装


def load_field_rules(config_dir):
    """明细侧配置：行为规则、字段映射、币种识别。"""
    rules = load_behavior_rules(os.path.join(config_dir, BEHAVIOR_FILE))
    columns, labels, currency_maps = load_field_sheet(os.path.join(config_dir, FIELD_FILE))
    currency_mapping = load_currency_mapping(os.path.join(config_dir, CURRENCY_FILE))

    if not columns:
        raise ConfigError("%s 没有读到字段行" % FIELD_FILE)
    if not currency_maps:
        raise ConfigError("%s 没有读到任何币种列" % FIELD_FILE)

    return {
        "columns": columns,
        "labels": labels,
        "currency_maps": currency_maps,
        "behavior_rules": rules,
        "currency_mapping": currency_mapping,
    }


def load_summary_rules(config_dir):
    """summary 侧配置：关键字、中文对照。数值格式走 load_number_formats，两侧共用。"""
    keywords = load_keywords(os.path.join(config_dir, KEYWORD_FILE))
    label_map = load_label_map(os.path.join(config_dir, LABEL_FILE))
    if not keywords:
        raise ConfigError("%s 没有读到任何币种" % KEYWORD_FILE)
    return {"keywords": keywords, "label_map": label_map}


# ---------------------------------------------------------------- 日期标准化


def normalize_datetime(value):
    """多语言日期时间统一成 YYYY-MM-DD HH:MM:SS，识别不了返回空。"""
    text = cell_text(value).replace("\u00a0", " ")
    if not text:
        return ""
    text = AMPM_PATTERN.sub(lambda match: match.group(1).upper() + "M", text)
    text = TIMEZONE_PATTERN.sub("", text).strip()
    text = re.sub(r"\b(?:de|del|da|do)\b", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"\s+", " ", text).strip().rstrip(".")

    match = PATTERN_ISO.match(text)
    if match:
        year, month, day, hour, minute, second = match.groups()
        return _format_datetime(int(year), int(month), int(day), int(hour), int(minute), int(second))

    match = PATTERN_DAY_FIRST.match(text)
    if match:
        day, month_name, year, hour, minute, second, ampm = match.groups()
        return _build_datetime(year, month_name, day, hour, minute, second, ampm)

    match = PATTERN_MONTH_FIRST.match(text)
    if match:
        month_name, day, year, hour, minute, second, ampm = match.groups()
        return _build_datetime(year, month_name, day, hour, minute, second, ampm)

    return ""


def _build_datetime(year, month_name, day, hour, minute, second, ampm):
    month = MONTHS.get(str(month_name).lower()[:3])
    if not month:
        return ""
    return _format_datetime(int(year), month, int(day), int(hour), int(minute), int(second), ampm)


def _format_datetime(year, month, day, hour, minute, second, ampm=None):
    if ampm:
        hour = hour % 12 + (12 if str(ampm).upper() == "PM" else 0)
    try:
        moment = datetime(year, month, day, hour, minute, second)
    except ValueError:
        return ""
    return moment.strftime("%Y-%m-%d %H:%M:%S")


def parse_report_period(text):
    """从 summary 页首取报表期间起止日期。"""
    matches = re.findall(r"([A-Z][a-z]{2,3})\s+(\d{1,2}),\s*(\d{4})", text)
    dates = []
    for month_name, day, year in matches[:4]:
        month = MONTHS.get(month_name.lower()[:3])
        if not month:
            continue
        try:
            dates.append(datetime(int(year), month, int(day)).strftime("%Y-%m-%d"))
        except ValueError:
            continue
    if len(dates) >= 2:
        return dates[0], dates[1]
    if len(dates) == 1:
        return dates[0], dates[0]

    cjk = re.findall(r"(\d{4})年(\d{1,2})月(\d{1,2})日", text)
    if cjk:
        values = ["%s-%02d-%02d" % (int(y), int(m), int(d)) for y, m, d in cjk[:2]]
        if len(values) == 2:
            return values[0], values[1]
        return values[0], values[0]
    return "", ""


# ---------------------------------------------------------------- 行为匹配


def prepare_rules(rules):
    """把规则预编译成匹配用的形式：type 转小写集合，关键词转小写列表。

    规则三元组为 (behavior 名称, type 列表, 关键词列表)，空列表表示该列不设限制。
    """
    prepared = []
    for name, types, keywords in rules:
        prepared.append((
            name,
            {str(item).strip().lower() for item in types if str(item).strip()},
            [str(item).strip().lower() for item in keywords if str(item).strip()],
        ))
    return prepared


def rule_matches(types, keywords, type_text, description_lower):
    """判断单条规则是否命中。

    契约：type 列与关键词列是「且」的关系。
    - type 列非空 -> type_text 必须精确命中其中之一（大小写不敏感）
    - 关键词列非空 -> 描述必须包含其中至少一个关键词（大小写不敏感）
    - 某一列为空 -> 该列不设限制；两列都空 -> 该规则匹配一切（只应作为最后兜底）
    """
    if types and type_text not in types:
        return False
    if keywords and not any(word in description_lower for word in keywords):
        return False
    return True


def match_behavior(type_value, description_value, rules, prepared, cache=None):
    """按表格从上往下匹配，返回第一条 type 与关键词同时满足的规则名称。

    匹配顺序固定为配置表行序，一旦命中立即停止，后面的规则不再参与。
    两列都填的规则必须同时命中；只填一列的规则只校验该列，这是兜底行的用法；
    两列都空的行会匹配任何交易，只能放在表末。

    参数：
        type_value: csv 里的 type 原文
        description_value: csv 里的 description 原文
        rules: 原始规则列表，仅用于判断是否为空
        prepared: prepare_rules 的产物，实际参与匹配
        cache: 可选的字典，按 (type, description) 缓存结果

    返回：
        命中的 behavior 名称；一条都没命中时返回空字符串。
    """
    if not rules:
        return ""
    type_text = cell_text(type_value).strip().lower()
    description_text = cell_text(description_value)
    key = (type_text, description_text)
    if cache is not None and key in cache:
        return cache[key]

    lowered = description_text.lower()
    chosen = ""
    for name, types, keywords in prepared:
        if rule_matches(types, keywords, type_text, lowered):
            chosen = name
            break
    if cache is not None:
        cache[key] = chosen
    return chosen

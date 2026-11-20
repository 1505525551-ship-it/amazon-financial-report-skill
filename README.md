# amazon-financial-report-skill

把亚马逊后台导出的**日期范围报告（csv 明细）**与 **summary（pdf 汇总）**配成一组，合成一份带核对结果的工作簿。

- 明细按标准字段拍平，补一列平台行为分类
- 新增"字段核对"页：summary 每个费用项回明细按列与行为汇总，逐项比金额
- summary 拆成一行一个费用项，原文与中文标准字段并列
- 按"明细 total 合计 = 收入 + 费用 + 转账 + 税"核对，容差默认 0.01
- 没进配置的费用项、没匹配上行为的交易，单独列进"待补清单"页

完整流程与已知坑见 [SKILL.md](SKILL.md)，输出列含义见 [references/output-guide.md](references/output-guide.md)。

## 安装

需要 Python 3.12 及以上。

```bash
pip install -r requirements.txt
```

依赖两个库：openpyxl（写 xlsx）与 pdfplumber（读 pdf）。两者都是 MIT，可随项目公开分发。

## 用法

```bash
python scripts/organize.py --input "放 csv 和 pdf 的文件夹"
```

| 参数 | 说明 |
| --- | --- |
| `--input` | 必填，csv 与 pdf 所在目录，会递归找 |
| `--output` | 选填，xlsx 路径或目录；默认写到每个目录下的 `整理结果/` |
| `--config` | 选填，配置目录，默认用 `config/` |
| `--tolerance` | 选填，核对容差，默认 0.01 |

配对规则：csv 与 pdf 去掉扩展名后同名，才算同一家店同一期。目录下按月份分子文件夹时，每个子文件夹各出一份工作簿。

## 目录

```
amazon-financial-report-skill/
├── SKILL.md                    主流程与约束
├── requirements.txt
├── config/                           七个 csv 配置（UTF-8）
│   ├── detail_report_behavior.csv    平台行为规则
│   ├── detail_report_field.csv       csv 字段映射（按币种一列）
│   ├── common_currency_mapping.csv   货币说明句 -> 币种（csv/pdf 通用）
│   ├── common_number_mapping.csv     数值格式（中英美 / 欧洲大陆）
│   ├── report_field_mapping.csv      summary 字段 -> 明细取数规则
│   ├── summary_report_keyword.csv    summary 四大合计关键字
│   └── summary_report_field.csv      费用项中文对照
├── references/
│   ├── output-guide.md         输出结构与排错
│   └── setup-guide.md          环境安装
├── scripts/
│   ├── organize.py             入口
│   ├── detail_report.py        csv 解析
│   ├── summary_report.py       pdf 解析（pdfplumber）
│   └── common.py               配置加载与规则匹配
└── tests/
    └── test_behavior_match.py  行为匹配契约测试
```

## 配置

七个 csv（UTF-8），改表不用改代码，重跑即生效。文件名与首行列名不能改；所有匹配忽略大小写；Excel 编辑后另存为"CSV UTF-8"。

平台行为规则（detail_report_behavior.csv）是"且"匹配：type 列写了必须命中 type，关键词列写了必须命中描述，两列都写就必须同时命中；从第一行往下扫，命中即停。规则表每个 type 的最后要留一行不写关键词的兜底。

币种识别（common_currency_mapping.csv）走货币说明句，csv 与 pdf 是同一句话，句子没命中时用同表的关键词列兜底；数值格式（common_number_mapping.csv）分中英美与欧洲大陆两种，未列出的币种按中英美格式解析。

## 自测

```bash
python -m unittest discover -s tests -p "test_*.py"
```

## 许可

作者：lulukeepgoing

代码与配置以 MIT 发布，见 [LICENSE](LICENSE)。第三方依赖：openpyxl（MIT）、pdfplumber（MIT）、pdfminer.six（MIT）、Pillow（MIT-CMU）、pypdfium2（Apache-2.0 / BSD-3-Clause）。

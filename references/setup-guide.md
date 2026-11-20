# 环境配置指南

这个技能全在本机跑，不联网、不登录、不碰任何账号凭据。要装的东西也就三样：Python 和两个库。

## 装什么

三个东西，缺一个都跑不起来。官方入口：

- Python 官网 https://www.python.org/
- Python 文档 https://www.python.org/doc/
- openpyxl 包页 https://pypi.org/project/openpyxl/
- pdfplumber 包页 https://pypi.org/project/pdfplumber/

## 需要的环境

| 项 | 要求 | 用途 | 许可 |
| --- | --- | --- | --- |
| Python | 3.12 及以上 | 运行整理脚本 | PSF |
| openpyxl | 3.1 及以上 | 写 xlsx 结果；配置读 csv 用标准库 | MIT |
| pdfplumber | 0.11 及以上 | 解析 summary pdf 的文字与坐标 | MIT |

解析 pdf 曾经用 PyMuPDF，它读取快但走 AGPL，代码公开分发会牵连整个项目，所以换成了 MIT 的 pdfplumber。代价是解析慢一些，两百多份 summary 大约多花一两分钟。

## 安装步骤

先看一眼 Python 版本，低于 3.12 得先升级：

```bash
python --version
```

然后装两个库：

```bash
pip install openpyxl pdfplumber
```

国内下载慢的话换清华镜像：

```bash
pip install -i https://pypi.tuna.tsinghua.edu.cn/simple openpyxl pdfplumber
```

装完验证一下：

```bash
python -c "import openpyxl, pdfplumber; print('ok')"
```

打印 `ok` 就说明能用了。

## 装没装好怎么判断

跑一句 import 就知道，打印 `ok` 就是齐了：

```bash
python -c "import openpyxl, pdfplumber; print('ok')"
```

整理脚本启动时也会自己检查一次，缺库会直接提示安装命令并退出，不会跑到一半崩掉。

## 常见问题

| 现象 | 原因 | 处理 |
| --- | --- | --- |
| `ModuleNotFoundError: openpyxl` | 未装库，或 pip 与 python 不是同一套环境 | 用 `python -m pip install openpyxl pdfplumber` 确保装进当前解释器 |
| `ModuleNotFoundError: pdfplumber` | 同上 | 同上 |
| 解析 pdf 报 `No module named pdfminer` | pdfplumber 装了一半 | 重装：`pip install --force-reinstall pdfplumber` |
| 脚本能跑但中文乱码 | 终端编码不是 UTF-8 | 在 cmd 执行 `chcp 65001`，或用 Windows Terminal |


## 安全边界

- 只从官方 PyPI 或可信镜像装包，来源不明的 whl 不要
- 依赖全部为宽松许可：openpyxl（MIT）、pdfplumber（MIT）、pdfminer.six（MIT）、Pillow（MIT-CMU）、pypdfium2（Apache-2.0 / BSD-3-Clause），可随技能一起公开分发
- 技能不收集、不外传任何数据；不读取凭据文件、环境变量和历史登录态
- 只读取 `--input` 指定目录下的 csv 与 pdf，只把结果写入该目录下的 `整理结果` 或 `--output` 指定位置
- 不修改、不移动、不删除源文件
- 如需撤销：删除技能目录即可，本机不会留下其他文件



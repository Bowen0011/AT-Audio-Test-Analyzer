# 🔊 AT Audio Test 数据分析工具 v4.2

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey.svg)]()

**单文件 GUI 桌面工具** — 解析产线 Audio Test 站别的行级 TSV 日志，按 SN 去重取最终结果，分站别统计良率+UPH，生成自包含 HTML 交互报告 + PNG 图表 + CSV 明细。

> A single-file desktop GUI tool for AT production line log analysis. Deduplicates by SN (final result per test item), provides per-station yield + UPH breakdown, and generates self-contained interactive HTML reports.

---

## 🆕 v4.2 更新

- **📈 7天数据对比** — 新增多日对比模式，勾选后可解析父目录下各日期子目录
- **🔘 站别切换图表** — HTML 报告中通过按钮切换各站别的7天趋势图（默认显示 AT1）
- **📊 组合图表** — 柱状图显示每日测试数 + 折线显示良率趋势，双Y轴一目了然
- **🎯 无色编码** — 折线数据点按良率自动着色（≥97%绿 / 95-97%黄 / <95%红）

## 🆕 v4.1 更新

- **HTML 图表自包含** — 改用 Chart.js 交互图表，不再依赖外部 PNG 文件。一个 HTML 拷走即看
- **UPH 分析** — 新增整体 UPH + 分站别平均 UPH + 每小时产出分布（按 SN 去重）
- **图表效果增强** — 阴影、悬浮 tooltip、柱状图失败原因文字标注
- **文件重命名** — 主文件改为 `at_analyzer.py`，删除旧版本备份

---

## ✨ 功能特性

- 🖥️ **独立单文件** — 仅依赖 matplotlib，双击即用
- 📂 **灵活数据源** — 支持 .zip / 文件夹 / 单个 .xls .txt
- 🔒 **加密兼容** — 公司加密 .xls 另存 .txt 即可分析
- 🔄 **SN 去重** — 每 SN 每项测试取最终结果，杜绝虚高 FAIL
- 📊 **5 标签页** — 站别统计 / 失败原因 / 失败SN / 全部SN(分站) + UPH 卡片
- 🔘 **7天对比** — 勾选后支持多日数据对比，按钮切换站别趋势图，默认 AT1
- 🧵 **后台线程** — 解析不冻结界面，实时进度条
- 📄 **HTML 报告** — Chart.js 交互图表（自包含，无需外部文件），支持阴影/悬浮/失败原因标注
- 📈 **PNG 图表** — 站别良率 / 失败原因 / 各站别Top3高频失败项
- 💾 **CSV 导出** — 全量明细，UTF-8 BOM，Excel 直接打开
- ⚡ **UPH 分析** — 整体吞吐量 + 分站别平均 UPH + 每小时产出（按 SN 去重）
- ⚠️ **站别检测** — 自动识别 AT01-AT07，向上逐级查找不依赖路径深度
- 🎨 **良率着色** — ≥97% 绿 / 95-97% 黄 / <95% 红
- 🔤 **通用 SN 检测** — 不依赖特定 SN 前缀，适配所有产品型号

---

## 🚀 快速开始

### 安装

```bash
pip install matplotlib
```

### 运行

```cmd
python at_analyzer.py
```

### 使用流程

1. 点击 **📁 数据源 → 选择** 选取 ZIP 文件
2. 点击 **▶ 开始分析**
3. 等待进度条完成，标签页实时显示结果 + UPH 数据
4. 点击 **📄 打开报告** 查看 HTML 完整报告

---

## 🔒 公司加密文件处理

```
右键文件 → 另存为 → 文本文件(制表符分隔)(*.txt)
→ 用本工具选择该 .txt 文件即可
```

工具自动探测 GBK/GB2312/GB18030/UTF-8 编码，不依赖 SN 前缀特征。

---

## 📊 数据格式

文件为 **GBK 编码 Tab 分隔值**，支持两种格式：

### Format B (行级，当前主力)
| 字段 | 说明 | 示例 |
|------|------|------|
| Time | 测试时间 | 2026/6/16 9:31:20 |
| SN | 设备序列号 | BKVXUN2669H00347 |
| TestChNum | 测试通道 | 右主2气密性 / 获取左1_SPAValue1 |
| TestName | 子测试项 | PA / FR / THD / Rub&Buzz |
| Result | 结果 | Pass / Fail |
| Channel | 通道 | Left / Right |

每 SN 约 40 项测试（SPA×4 + 标压/EQ/气密/MIC ×36），每项可能多次重测。

### Format A (9行块，兼容旧格式)
旧版产线日志格式，每 SN 9 行一组，自动识别。

---

## 📁 输出文件

| 文件 | 说明 |
|------|------|
| `report.html` | 自包含 HTML 交互报告（Chart.js，无需 PNG 附件） |
| `chart_station_yield.png` | 各站别良率柱状图 |
| `chart_failure_reasons.png` | 失败原因分布 |
| `chart_sn_fail_detail.png` | 各站别Top3高频失败项 |
| `detail.csv` | 全量测试明细（勾选后生成） |

### 良率着色标准

| 良率 | 颜色 |
|------|------|
| ≥ 97% | 🟢 绿色 (达标) |
| 95% ~ 97% | 🟡 黄色 (预警) |
| < 95% | 🔴 红色 (不达标) |

---

## 🛠️ 技术栈

| 层级 | 技术 |
|------|------|
| GUI | Tkinter (Python 内置) |
| 图表 | Matplotlib + Chart.js |
| 报告 | HTML + Chart.js (自包含，无需外部文件) |
| 编码 | GBK/GB2312/GB18030/UTF-8 自动探测 |
| 解析 | ZipFile 流式读取，不落盘 |
| 线程 | threading（后台解析，UI 不冻结） |

---

## 📝 Changelog

### v4.2 (2026-06-22)
- **7天数据对比**：勾选"📅 7天对比模式"，选择包含日期子目录的父文件夹
- **站别切换图表**：HTML 报告自动生成切换按钮，点击不同站别查看趋势
- **组合图表**：柱状图（测试数）+ 折线（良率），双 Y 轴，数据点自动着色
- **目录结构**：parent/2026-06-15/…, parent/2026-06-16/… 自动识别

### v4.1 (2026-06-17)
- **HTML 图表自包含**：改用 Chart.js 交互图表，不再依赖外部 PNG
- **UPH 分析**：整体 UPH + 分站别平均 UPH + 每小时产出分布
- **图表效果**：阴影插件 + 悬浮 tooltip + Top3 失败原因文字标注
- **文件整理**：主文件重命名为 `at_analyzer.py`，删除旧版备份

### v4.0 (2026-06-16)
- 合并 v1 + v3：统一 Format A/B 解析器
- 通用编码检测：不再依赖 SN 前缀
- 站别路径检测：向上逐级查找 ATxx
- ZIP 直读流：消除 Windows 不确定性

---

## 📝 License

MIT © 2026

---

*Made with ❤️ for production line engineers.*

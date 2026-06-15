# 🔊 AT Audio Test 数据分析工具 v3.2

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey.svg)]()

**单文件 GUI 桌面工具** — 解析产线 Audio Test 站别的行级 TSV 日志，按 SN 去重取最终结果，分站别统计良率，生成 HTML 报告 + PNG 图表 + CSV 明细。

> A single-file desktop GUI tool for AT production line log analysis. Deduplicates by SN (final result per test item), provides per-station yield breakdown, and generates interactive HTML reports.

---

## 📸 界面预览

```
┌──────────────────────────────────────────────────────┐
│  🔊  AT Audio Test 分析工具 v3.2                     │
├─────────────────────┬────────────────────────────────┤
│                     │  ┌──────┬──────┬──────┬──────┐  │
│ 📁 数据源            │  │ 815  │ 786  │  29  │96.4% │  │
│ [____________] [选择]│  │总SN  │ PASS │ FAIL │ 良率 │  │
│ 📦ZIP 📁文件夹 📄文件│  └──────┴──────┴──────┴──────┘  │
│                     │                                │
│ 📂 输出目录           │  ┌─ 📊站别 ──┬─ 🔍失败原因 ──┐  │
│ [____________] [选择]│  │AT01 99.1%✅│右主2气密性 FR │  │
│                     │  │AT02 96.0%🟡│右主2标压 FR   │  │
│ ⚙️ 输出选项           │  │AT03 91.2%🔴│右主2EQ FR     │  │
│ ☑ PNG ☑ HTML ☑ CSV  │  │AT04 96.0%🟡│MIC2气密差值 FR│  │
│ ☑ 自动打开            │  │AT05 97.7%✅│...           │  │
│                     │  │AT07 98.9%✅│              │  │
│ [▶ 开始分析]         │  └───────────┴──────────────┘  │
│                     │                                │
│                     │  📋 全部SN — 按站别分tab查看每台│
│                     │  🔴 失败SN — 29台不良明细       │
│                     │                                │
│                     │  [📄打开报告] [📂输出目录]       │
├─────────────────────┴────────────────────────────────┤
│ ✅ 完成！总数815台 PASS=786 FAIL=29 良率96.4%        │
└──────────────────────────────────────────────────────┘
```

---

## ✨ 功能特性

- 🖥️ **独立单文件** — 仅依赖 matplotlib，双击 .bat 即用
- 📂 **灵活数据源** — 支持 .zip / 文件夹 / 单个 .xls .txt
- 🔒 **加密兼容** — 公司加密 .xls 另存 .txt 即可分析
- 🔄 **SN 去重** — 每 SN 每项测试取最终结果，杜绝虚高 FAIL
- 📊 **4 标签页** — 站别统计 / 失败原因 / 失败SN / 全部SN(分站)
- 🧵 **后台线程** — 解析不冻结界面，实时进度条
- 📄 **HTML 报告** — 嵌入 3 张图表 + 站别/失败SN表格
- 📈 **PNG 图表** — 站别良率 / 失败原因 / 各站别Top3高频失败项
- 💾 **CSV 导出** — 全量明细，UTF-8 BOM，Excel 直接打开
- ⚠️ **站别检测** — 自动识别 AT01-AT07，少于 6 站弹窗警告
- 🎨 **良率着色** — ≥97% 绿 / 95-97% 黄 / <95% 红

---

## 🚀 快速开始

### 安装

```bash
pip install matplotlib
```

### 运行

```cmd
rem Windows: 双击 双击运行.bat
rem 或
python AT音频测试分析工具_v3.py
```

### 使用流程

1. 点击 **📁 数据源 → 选择** 选取 ZIP 文件
2. 点击 **▶ 开始分析**
3. 等待进度条完成，四个标签页实时显示结果
4. 点击 **📄 打开报告** 查看 HTML 完整报告

---

## 🔒 公司加密文件处理

```
右键文件 → 另存为 → 文本文件(制表符分隔)(*.txt)
→ 用本工具选择该 .txt 文件即可
```

工具自动探测 GBK/GB2312/GB18030/UTF-8 编码。

---

## 📊 数据格式

文件为 **GBK 编码 Tab 分隔值**，行级日志格式：

| 字段 | 说明 | 示例 |
|------|------|------|
| Time | 测试时间 | 2026/6/13 10:05:19 |
| SN | 设备序列号 | BLKVUN2665H00375 |
| TestChNum | 测试通道 | 右主2气密性 / 获取左1_SPAValue1 |
| TestName | 子测试项 | PA / FR / THD / Rub&Buzz |
| Result | 结果 | Pass / Fail |
| Channel | 通道 | Left / Right |

每 SN 约 40 项测试（SPA×4 + 标压/EQ/气密/MIC ×36），每项可能多次重测。

---

## 📁 输出文件

| 文件 | 说明 |
|------|------|
| `report.html` | 交互式 HTML 汇总报告 |
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
| 图表 | Matplotlib |
| 报告 | HTML + CSS (自包含) |
| 编码 | GBK/GB2312/GB18030/UTF-8 自动探测 |
| 解析 | ZipFile 流式读取，不落盘 |
| 线程 | threading（后台解析，UI 不冻结） |

---

## 📝 Changelog

### v3.2
- 重写解析引擎：ZIP 直接读流，消除 os.walk 不确定性
- 新增 📋全部SN 标签页（按站别分 tab）+ 🔍失败原因标签页
- 各站别Top3高频失败项图表：横轴=站别，每站3柱标注失败项
- 站别缺失弹窗警告 + 解析日志
- 良率阈值 97%/95%（绿/黄/红）
- 字段映射修正：TestChNum(字段2) + TestName(字段3)

### v3.1
- SN 级去重：每项取最终结果，消除虚假 86% FAIL
- 周期边界检测（后续改为 SN 去重）

### v2
- 初版：GUI + 图表 + HTML 报告

---

## 📝 License

MIT © 2026

---

*Made with ❤️ for production line engineers.*

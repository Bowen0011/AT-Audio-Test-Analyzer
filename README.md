# 🔊 AT Audio Test 数据分析工具

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey.svg)]()

**单文件 GUI 桌面工具** — 解析产线 Audio Test 站别的 TSV 测试数据，按 SN / 站别 / 时序多维度分析直通率、不良分布，一键生成 HTML 报告 + PNG 图表 + CSV 明细。

> A single-file desktop GUI tool for analyzing Audio Test production line data. Parse TSV test logs, visualize pass/fail rates by SN and station, and generate interactive HTML reports.

---

## 📸 界面预览

```
┌──────────────────────────────────────────────────────┐
│  🔊  AT Audio Test 数据分析工具           v2.0 单文件版 │
├─────────────────────┬────────────────────────────────┤
│                     │  ┌──────┬──────┬──────┬──────┐  │
│ 📁 数据源            │  │ 858  │ 786  │  72  │91.6% │  │
│ [____________] [选择]│  │总测试 │ PASS │ FAIL │直通率│  │
│ 📦ZIP 📁文件夹 📄文件│  └──────┴──────┴──────┴──────┘  │
│                     │                                │
│ 📂 输出目录           │  ┌─ 各站别统计(颜色标记) ────┐  │
│ [____________] [选择]│  │ AT01 121 120  1   0.8%   │  │
│                     │  │ AT02 110 102  8   7.3%   │  │
│ ⚙️ 输出选项           │  │ AT03 154 118 36  23.4%⚠ │  │
│ ☑ PNG图表 ☑ HTML报告 │  │ AT04 155 147  8   5.2%   │  │
│ ☑ CSV明细 ☑ 自动打开  │  │ AT05 133 124  9   6.8%   │  │
│                     │  └──────────────────────────┘  │
│ [▶ 开始分析]         │                                │
│                     │  ┌─ 跨站重复失败SN ──────────┐  │
│                     │  │ BLKVUN2665H00116  4站     │  │
│                     │  │ BLKVUN2665H00427  3站     │  │
│                     │  └──────────────────────────┘  │
│                     │                                │
│                     │  [📄打开报告] [📂输出目录] [💾CSV]│
├─────────────────────┴────────────────────────────────┤
│ ✅ 解析完成: 858 条, 6 个站别      ████████████ 100%  │
└──────────────────────────────────────────────────────┘
```

---

## ✨ 功能特性

- 🖥️ **独立单文件** — 零依赖其他文件，双击即用
- 📂 **灵活数据源** — 支持 .zip 压缩包 / 文件夹 / 单个 .xls .txt
- 🔒 **加密兼容** — 公司加密的 .xls 另存为 .txt 即可正常分析
- 📊 **实时预览** — Summary 卡片 + 站别表格(颜色标记不良率) + 失败SN矩阵
- 🧵 **后台线程** — 解析不冻结界面，实时进度条
- 📄 **HTML 报告** — 6 张嵌入图表 + 可排序表格 + 完整失败SN清单
- 📈 **6 张 PNG 图表** — 饼图/柱状图/折线图/箱线图/热力图
- 💾 **CSV 导出** — 858+ 字段完整明细，UTF-8 BOM 编码，Excel 直接打开

---

## 🚀 快速开始

### 安装

```bash
# 唯一依赖
pip install matplotlib
```

### 运行

```bash
# 双击运行（Windows）
AT音频测试分析工具.py

# 或命令行
python AT音频测试分析工具.py
```

### 使用流程

1. 点击 **📁 数据源 → 选择** 选取 ZIP 文件或文件夹
2. 点击 **📂 输出目录 → 选择** 指定报告保存位置
3. 点击 **▶ 开始分析**
4. 等待进度条完成，UI 实时显示结果
5. 点击 **📄 打开 HTML 报告** 在浏览器查看完整报告

---

## 🔒 公司加密文件处理

如果公司电脑自动加密 `.xls` 文件：

```
右键文件 → 另存为 → 选择「文本文件(制表符分隔)(*.txt)」
→ 用本工具选择该 .txt 文件 或 所在文件夹
```

工具自动探测 GBK/GB2312/GB18030/UTF-8 编码，无需手动转换。

---

## 📊 数据格式说明

文件实际是 **GBK 编码的 Tab 分隔值 (TSV)**，非真正的 Excel 格式。

每条测试记录固定 9 行，包含 **2000+ 列** 音频测量参数：

| 测试项 | 通道 | 说明 |
|--------|------|------|
| FR（频率响应） | 左主1/2、右主1/2、MIC1/2 | 200Hz ~ 10kHz |
| THD（总谐波失真） | 同上 | 非线性失真 |
| Rub & Buzz | 同上 | 异音/杂音检测 |
| 标压 (SPL) | 同上 | 额定声压级 |
| 气密性 | 同上 | 声学密封性 |
| SPA 参数 | 左/右通道 | 喇叭 Thiele-Small 参数 |

---

## 📁 输出文件

| 文件 | 说明 |
|------|------|
| `report.html` | 交互式 HTML 汇总报告 |
| `chart_overall_pie.png` | 整体直通率饼图 |
| `chart_station_bars.png` | 各站别 PASS/FAIL 堆叠柱状图 |
| `chart_hourly_rate.png` | 时段直通率趋势折线图 |
| `chart_cycle_time.png` | 节拍时间箱线图 |
| `chart_fail_dist.png` | 失败 SN 跨站分布 |
| `chart_fail_matrix.png` | 失败 SN × 站别 热力图 |
| `records_detail.csv` | 测试明细 CSV（勾选后生成） |

---

## 🛠️ 技术栈

| 层级 | 技术 |
|------|------|
| GUI | Tkinter (Python 内置) |
| 图表 | Matplotlib |
| 报告 | HTML + CSS (自包含) |
| 编码 | GBK/GB2312/GB18030/UTF-8 自动探测 |
| 线程 | threading（后台解析，UI 不冻结） |

---

## 📝 License

MIT © 2026

---

## 🤝 贡献

欢迎提 Issue 和 PR。

常见扩展方向：
- 添加规格限对比分析（超限自动标记）
- 导出 Excel (.xlsx) 格式报告
- 支持多天数据对比
- 邮件自动发送报告

---

*Made with ❤️ for production line engineers.*

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AT Audio Test 数据分析工具 v4.1
===============================
支持 Format A (9行块) 和 Format B (行级40项) 两种日志格式。
取每项最终结果判定SN良率。

输出: 直观汇总 + 分站别统计 + 失败原因 + UPH分析 + HTML交互报告 + 图表
- 自动识别编码（GBK/UTF-8等），不依赖SN前缀
- 按SN去重取各项最终结果，精准判定良率
- HTML报告使用Chart.js交互图表（无需外部PNG），支持阴影/悬浮效果
- 新增UPH（Units Per Hour）分析：整体+分站别+每小时分布
- 颜色阈值: ≥97%绿 / 95-97%黄 / <95%红
"""

import os, sys, csv, json, zipfile, threading, datetime, warnings, webbrowser, subprocess
from pathlib import Path
from collections import defaultdict, Counter
from tempfile import TemporaryDirectory

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from matplotlib.font_manager import FontProperties

warnings.filterwarnings("ignore")

_CJK = ["SimHei","Microsoft YaHei","PingFang SC","Noto Sans CJK SC","WenQuanYi Micro Hei","STHeiti","SimSun"]
_ENCS = ["gbk","gb2312","gb18030","utf-8","utf-16","latin-1"]
C = {"pass":"#4CAF50","fail":"#F44336","warn":"#FF9800","accent":"#2196F3","bar_pass":"#81C784","bar_fail":"#E57373"}
STYLE = {"bg":"#f0f2f5","fg":"#1a1a2e","card_bg":"#ffffff","accent":"#2196F3","accent_hover":"#1976D2",
         "success":"#4CAF50","danger":"#F44336","warning":"#FF9800","text_secondary":"#78909c",
         "input_bg":"#ffffff","tree_bg":"#ffffff","tree_sel":"#E3F2FD","progress_bg":"#e0e0e0"}

_font = None
def _get_font():
    global _font
    if _font: return _font
    av = set(f.name for f in matplotlib.font_manager.fontManager.ttflist)
    for n in _CJK:
        if n in av: _font = FontProperties(family=n); break
    if not _font: _font = FontProperties()
    plt.rcParams["font.family"] = _font.get_name()
    plt.rcParams["axes.unicode_minus"] = False
    return _font

# ═══════════════════════════════════════════════
# 解析器
# ═══════════════════════════════════════════════
def _detect_station_from_path(filepath):
    """从文件路径中提取站别名(ATxx)，向上逐级查找"""
    p = Path(filepath)
    for parent in [p] + list(p.parents):
        name = parent.name
        if name.startswith("AT") and len(name) <= 6 and name[2:].isdigit():
            return name
    return os.path.basename(os.path.dirname(filepath))  # fallback

def _detect_enc_from_bytes(raw):
    """从原始字节检测编码"""
    for e in _ENCS:
        try:
            if "SN" in raw[:8192].decode(e, errors="replace"): return e
        except: continue
    return "gbk"

# 列映射: Time[0] SN[1] TestChNum[2] TestName[3] Unit[4] Result[5] Channel[6] Value[7]...
def parse_bytes(raw, station):
    """从原始字节解析记录，需要传入已知的站别名"""
    enc = _detect_enc_from_bytes(raw)
    text = raw.decode(enc, errors="replace")
    records = []
    for line in text.split("\n"):
        line = line.strip()
        if not line or line.startswith("Time"): continue
        flds = line.split("\t")
        if len(flds) < 8: continue
        time_str = flds[0].strip()
        sn = flds[1].strip().strip("'\"")
        test_ch = flds[2].strip()
        test_name = flds[3].strip()
        result = flds[5].strip()
        if not result or result not in ("Pass","Fail","pass","fail"): continue

        val_str = flds[7].strip() if len(flds) > 7 else ""
        value = None
        if val_str:
            try: value = float(val_str)
            except: pass

        try: t = datetime.datetime.strptime(time_str, "%Y/%m/%d %H:%M:%S")
        except: t = None

        records.append({
            "sn": sn, "test_ch": test_ch, "test_name": test_name,
            "display": f"{test_ch} {test_name}".strip(),
            "result": result, "value": value, "station": station, "time": t,
        })
    return records

def parse_source(src, cb=None):
    all_recs = []; skipped = []; station_files = defaultdict(int)

    def _add_file(station, raw):
        """处理单个文件的内存数据"""
        station_files[station] += 1
        try:
            recs = parse_bytes(raw, station)
            all_recs.extend(recs)
        except Exception as e:
            skipped.append(f"{station}: {e}")
            import traceback; traceback.print_exc()

    if src.lower().endswith(".zip"):
        if cb: cb("解析ZIP中...", src)
        with zipfile.ZipFile(src, "r") as zf:
            for info in zf.infolist():
                if info.is_dir(): continue
                fn = info.filename
                if not fn.lower().endswith((".xls", ".txt")): continue
                st = _detect_station_from_path(fn)
                with zf.open(info) as f:
                    _add_file(st, f.read())
    elif os.path.isdir(src):
        if cb: cb("扫描文件夹...", src)
        for root, _, files in os.walk(src):
            for fn in files:
                if not fn.lower().endswith((".xls", ".txt")): continue
                fp = os.path.join(root, fn)
                st = _detect_station_from_path(fp)
                try:
                    with open(fp, "rb") as f: raw = f.read()
                    _add_file(st, raw)
                except Exception as e:
                    skipped.append(f"{fn}: {e}")
    else:
        # 单文件
        fp = src
        st = _detect_station_from_path(fp)
        try:
            with open(fp, "rb") as f: raw = f.read()
            _add_file(st, raw)
        except Exception as e:
            skipped.append(f"{os.path.basename(fp)}: {e}")

    return all_recs, skipped, dict(station_files)

# ═══════════════════════════════════════════════
# 分析器 — 按SN去重，取最终结果
# ═══════════════════════════════════════════════
def analyze(records):
    if not records: return {"error": "无记录"}

    # 按SN分组
    sn_data = defaultdict(list)
    for r in records:
        sn_data[r["sn"]].append(r)

    # 跟踪每个站别的时间范围（用于UPH计算）
    station_times = defaultdict(list)  # st -> [time, ...]
    for r in records:
        if r.get("time"):
            station_times[r["station"]].append(r["time"])

    # 每个SN: 按 (test_ch, test_name) 分组，取最终结果
    sn_results = {}   # sn -> {station, passed: bool, failed_items: [display], total_items: int}
    station_stats = defaultdict(lambda: {"total":0, "pass":0, "fail":0, "failed_sns":[]})
    failure_counter = Counter()  # 失败原因计数（display名）
    channel_failure = Counter()  # 按通道

    for sn, items in sn_data.items():
        items.sort(key=lambda x: x["time"] or datetime.datetime.min)
        station = items[0]["station"]

        # 按 display 分组，取最后一次结果
        by_display = defaultdict(list)
        for item in items:
            by_display[item["display"]].append(item)

        failed_displays = []
        for disp, entries in by_display.items():
            last = entries[-1]
            if last["result"] in ("Fail", "fail"):
                failed_displays.append(disp)
                failure_counter[disp] += 1
                channel_failure[last["test_ch"]] += 1

        sn_passed = len(failed_displays) == 0
        sn_results[sn] = {
            "station": station,
            "passed": sn_passed,
            "failed_items": failed_displays,
            "total_items": len(by_display),
        }
        station_stats[station]["total"] += 1
        if sn_passed:
            station_stats[station]["pass"] += 1
        else:
            station_stats[station]["fail"] += 1
            station_stats[station]["failed_sns"].append({
                "sn": sn,
                "failed": failed_displays,
                "total": len(by_display),
            })

    total_sn = len(sn_results)
    pass_sn = sum(1 for v in sn_results.values() if v["passed"])
    fail_sn = total_sn - pass_sn

    # 失败SN列表
    fail_list = sorted(
        [{"sn": sn, "station": v["station"], "failed": v["failed_items"], "total": v["total_items"]}
         for sn, v in sn_results.items() if not v["passed"]],
        key=lambda x: -len(x["failed"])
    )
    # 全部SN列表(按站别分组)
    all_sn_by_station = defaultdict(list)
    for sn, v in sn_results.items():
        all_sn_by_station[v["station"]].append({
            "sn": sn, "passed": v["passed"],
            "failed": v["failed_items"], "total": v["total_items"],
        })

    # ═══ UPH 计算 ═══
    uph_data = {}
    # 整体UPH
    all_times = [t for tl in station_times.values() for t in tl]
    if all_times:
        overall_hours = (max(all_times) - min(all_times)).total_seconds() / 3600
        uph_data["overall"] = {
            "total_sn": total_sn,
            "hours": round(overall_hours, 2),
            "uph": round(total_sn / overall_hours, 1) if overall_hours > 0 else 0,
            "start": min(all_times).strftime("%H:%M"),
            "end": max(all_times).strftime("%H:%M"),
        }
        # 每小时产出分布（按SN去重：每台设备每小时只计一次）
        hourly_sn = defaultdict(int)
        for sn, items in sn_data.items():
            times_sorted = sorted(it["time"] for it in items if it.get("time"))
            if times_sorted:
                hourly_sn[times_sorted[0].strftime("%H:00")] += 1
        uph_data["hourly"] = dict(sorted(hourly_sn.items()))
    else:
        uph_data["overall"] = {"total_sn": total_sn, "hours": 0, "uph": 0, "start": "-", "end": "-"}
        uph_data["hourly"] = {}

    # 分站别UPH
    uph_data["stations"] = {}
    for st in sorted(station_stats.keys()):
        times = station_times.get(st, [])
        if times:
            hours = (max(times) - min(times)).total_seconds() / 3600
            sn_count = station_stats[st]["total"]
            uph_data["stations"][st] = {
                "sn": sn_count,
                "hours": round(hours, 2),
                "uph": round(sn_count / hours, 1) if hours > 0 else 0,
                "start": min(times).strftime("%H:%M"),
                "end": max(times).strftime("%H:%M"),
            }
        else:
            uph_data["stations"][st] = {
                "sn": station_stats[st]["total"], "hours": 0, "uph": 0,
                "start": "-", "end": "-",
            }

    return {
        "total_sn": total_sn,
        "pass_sn": pass_sn,
        "fail_sn": fail_sn,
        "yield_rate": pass_sn/total_sn*100 if total_sn else 0,
        "station_stats": dict(station_stats),
        "failure_counter": dict(failure_counter),
        "channel_failure": dict(channel_failure),
        "fail_list": fail_list,
        "all_sn_by_station": dict(all_sn_by_station),
        "total_raw_rows": len(records),
        "uph": uph_data,
    }

# ═══════════════════════════════════════════════
# 图表
# ═══════════════════════════════════════════════
def chart_station_yield(a, out):
    _get_font()
    sts = sorted(a["station_stats"].keys())
    yields = [a["station_stats"][s]["pass"]/a["station_stats"][s]["total"]*100 for s in sts]
    totals = [a["station_stats"][s]["total"] for s in sts]
    fig, ax = plt.subplots(figsize=(max(8,len(sts)*1.5), 5))
    bars = ax.bar(sts, yields, color=[C["pass"] if y>=97 else C["warn"] if y>=95 else C["fail"] for y in yields])
    ax.set_ylabel("良率 (%)", fontsize=12)
    ax.set_title("各站别良率", fontsize=14, fontweight="bold")
    for bar, y, t in zip(bars, yields, totals):
        ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+0.3, f"{y:.1f}%\n({t}台)", ha="center", fontsize=9, fontweight="bold")
    ax.set_ylim(min(yields)-5, 102)
    ax.axhline(y=95, color="#ccc", linestyle="--", linewidth=1)
    plt.tight_layout(); fig.savefig(out, dpi=150, bbox_inches="tight"); plt.close(fig)

def chart_failure_reasons(a, out):
    _get_font()
    fc = a["failure_counter"]
    if not fc: return
    items = sorted(fc.items(), key=lambda x: -x[1])[:12]
    labels = [k[:25] for k,_ in items]
    counts = [v for _,v in items]
    fig, ax = plt.subplots(figsize=(10, max(5, len(items)*0.4)))
    colors = [C["fail"]]*len(items)
    bars = ax.barh(range(len(items)), counts, color=colors, height=0.6)
    ax.set_yticks(range(len(items)))
    ax.set_yticklabels(labels, fontsize=9)
    ax.set_xlabel("失败SN数", fontsize=12)
    ax.set_title("失败原因分布 (按SN去重)", fontsize=14, fontweight="bold")
    ax.invert_yaxis()
    for bar, c in zip(bars, counts):
        ax.text(bar.get_width()+0.1, bar.get_y()+bar.get_height()/2, str(c), va="center", fontsize=9, fontweight="bold")
    ax.xaxis.set_major_locator(mticker.MaxNLocator(integer=True))
    plt.tight_layout(); fig.savefig(out, dpi=150, bbox_inches="tight"); plt.close(fig)

def chart_sn_fail_detail(a, out):
    """各站别Top3高频失败项（分组柱状图）"""
    _get_font()
    ss = a["station_stats"]
    if not ss: return

    # 按站别统计失败项
    station_fails = {}
    for st in sorted(ss.keys()):
        counter = Counter()
        for sn_info in ss[st].get("failed_sns", []):
            for item in sn_info["failed"]:
                counter[item] += 1
        station_fails[st] = counter

    stations = sorted(station_fails.keys())
    n = len(stations)
    colors = ["#E53935", "#FB8C00", "#1E88E5"]
    bar_w = 0.22

    fig, ax = plt.subplots(figsize=(max(10, n*1.6), 6))
    max_val = 1

    for pos in range(3):
        vals = []; labels = []
        for st in stations:
            items = station_fails[st].most_common(3)
            if pos < len(items):
                vals.append(items[pos][1])
                labels.append(items[pos][0][:14])
                max_val = max(max_val, items[pos][1])
            else:
                vals.append(0); labels.append("")

        x_pos = [i + pos*bar_w for i in range(n)]
        bars = ax.bar(x_pos, vals, bar_w, color=colors[pos],
                       edgecolor="white", linewidth=0.3,
                       label=["#1最多","#2","#3"][pos])
        for i, (bar, lbl, v) in enumerate(zip(bars, labels, vals)):
            if v > 0:
                ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+0.3,
                        lbl, ha="center", va="bottom", fontsize=7, rotation=45, color="#333")

    ax.set_xticks([i + bar_w for i in range(n)])
    ax.set_xticklabels(stations, fontsize=11, fontweight="bold")
    ax.set_ylabel("失败SN数", fontsize=12)
    ax.set_title("各站别 Top3 高频失败项", fontsize=14, fontweight="bold")
    ax.legend(fontsize=8, ncol=3, loc="upper right")
    ax.yaxis.set_major_locator(mticker.MaxNLocator(integer=True))
    ax.set_ylim(0, max_val*1.5)
    plt.tight_layout(); fig.savefig(out, dpi=150, bbox_inches="tight"); plt.close(fig)

# ═══════════════════════════════════════════════
# HTML (Chart.js 自包含交互报告)
# ═══════════════════════════════════════════════
_HTML = r"""<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>AT Audio Test 分析报告</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.min.js">
</script><style>
:root{--bg:#f0f2f5;--card:#fff;--fg:#1a1a2e;--accent:#2196F3;--green:#4CAF50;
--red:#F44336;--orange:#FF9800;--muted:#78909c;--border:#e0e0e0}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,"Microsoft YaHei","PingFang SC",sans-serif;
background:var(--bg);color:var(--fg);padding:16px;line-height:1.5}
.container{max-width:1280px;margin:0 auto}
h1{text-align:center;color:#1a237e;font-size:26px;margin-bottom:4px;letter-spacing:1px}
.subtitle{text-align:center;color:var(--muted);font-size:12px;margin-bottom:24px}
/* 卡片 */
.summary{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:12px;margin-bottom:20px}
.card{background:var(--card);border-radius:12px;padding:16px 14px;text-align:center;
box-shadow:0 2px 12px rgba(0,0,0,0.06);transition:transform .2s,box-shadow .2s;cursor:default}
.card:hover{transform:translateY(-2px);box-shadow:0 6px 20px rgba(0,0,0,0.12)}
.card .v{font-size:32px;font-weight:800;letter-spacing:-1px}
.card .l{font-size:11px;color:var(--muted);margin-top:4px;text-transform:uppercase;letter-spacing:.5px}
.card.g .v{color:var(--green)}.card.r .v{color:var(--red)}
.card.b .v{color:var(--accent)}.card.o .v{color:var(--orange)}
.card .sub{font-size:11px;color:var(--muted);margin-top:3px}
/* 区块 */
.sec{background:var(--card);border-radius:12px;padding:20px;margin-bottom:16px;
box-shadow:0 2px 12px rgba(0,0,0,0.06)}
.sec h2{color:#37474f;font-size:16px;margin-bottom:14px;border-left:4px solid var(--accent);
padding-left:10px;display:flex;align-items:center;gap:8px}
.sec .chart-wrap{position:relative;width:100%;max-height:420px}
.sec .chart-wrap canvas{width:100%!important}
/* 表格 */
.tbl-wrap{overflow-x:auto}
table{width:100%;border-collapse:collapse;font-size:12px}
th{background:#263238;color:#fff;padding:9px 10px;text-align:left;font-weight:600;white-space:nowrap}
td{padding:7px 10px;border-bottom:1px solid #eceff1}
tr:hover td{background:#f5f7fa}
tr.clickable{cursor:pointer}
.badge{display:inline-block;padding:2px 9px;border-radius:10px;font-size:11px;font-weight:700;color:#fff}
.bg{background:var(--green)}.br{background:var(--red)}.bo{background:var(--orange)}
code{background:#f0f0f0;padding:1px 5px;border-radius:3px;font-size:11px}
/* UPH 特殊样式 */
.uph-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:12px;margin-bottom:14px}
.uph-card{background:#f8faff;border:1px solid #e3edf7;border-radius:10px;padding:12px 16px}
.uph-card .st{font-weight:700;color:var(--accent);font-size:13px}
.uph-card .val{font-size:20px;font-weight:800;color:#0d47a1;margin:4px 0}
.uph-card .detail{font-size:11px;color:var(--muted)}
/* 响应式 */
@media(max-width:768px){.summary{grid-template-columns:repeat(2,1fr)}}
</style></head><body><div class="container">
<h1>🔊 AT Audio Test 分析报告</h1>
<p class="subtitle">$REPORT_TIME$ | 数据源: $SOURCE_INFO$ | 按SN去重取各项最终结果</p>

<!-- ═══ 总览卡片 ═══ -->
<div class="summary">
<div class="card b"><div class="v">$TOTAL_SN$</div><div class="l">测试总数(台)</div></div>
<div class="card g"><div class="v">$PASS_SN$</div><div class="l">PASS</div></div>
<div class="card r"><div class="v">$FAIL_SN$</div><div class="l">FAIL</div></div>
<div class="card $YIELD_CLASS$"><div class="v">$YIELD_RATE$%</div><div class="l">良率</div></div>
<div class="card b"><div class="v">$UPH_VALUE$</div><div class="l">整体UPH(台/时)</div><div class="sub">$UPH_HOURS$h · $UPH_START$~$UPH_END$</div></div>
</div>

<!-- ═══ UPH 分站别 ═══ -->
<div class="sec"><h2>⚡ UPH 分析</h2>
<div class="uph-row" id="uphCards">$UPH_CARDS$</div>
<div class="chart-wrap"><canvas id="chartUphStations"></canvas></div></div>

<!-- ═══ 各站别良率 ═══ -->
<div class="sec"><h2>📊 各站别良率</h2>
<div class="chart-wrap"><canvas id="chartStationYield"></canvas></div></div>

<!-- ═══ 各站别统计表 ═══ -->
<div class="sec"><h2>📋 各站别统计</h2><div class="tbl-wrap"><table>
<tr><th>站别</th><th>测试数</th><th>PASS</th><th>FAIL</th><th>良率</th><th>UPH</th><th>时间范围</th></tr>
$STATION_TABLE_ROWS$</table></div></div>

<!-- ═══ 失败原因 ═══ -->
<div class="sec"><h2>🔍 失败原因分布 (按SN去重)</h2>
<div class="chart-wrap"><canvas id="chartFailureReasons"></canvas></div></div>

<!-- ═══ 各站别高频失败项 ═══ -->
<div class="sec"><h2>📊 各站别 Top3 高频失败项</h2>
<div class="chart-wrap"><canvas id="chartSnFailDetail"></canvas></div></div>

<!-- ═══ 失败SN明细 ═══ -->
<div class="sec"><h2>🔴 失败SN明细</h2><div class="tbl-wrap"><table>
<tr><th>SN</th><th>站别</th><th>失败项数</th><th>失败测试项</th></tr>
$SN_TABLE_ROWS$</table></div></div>

</div>

<script>
// ═══ Chart.js 阴影插件 ═══
const shadowPlugin = {
  id: 'shadow',
  beforeDatasetsDraw(chart) {
    const {ctx} = chart;
    ctx.save();
    ctx.shadowColor = 'rgba(0,0,0,0.15)';
    ctx.shadowBlur = 8;
    ctx.shadowOffsetX = 2;
    ctx.shadowOffsetY = 3;
  },
  afterDatasetsDraw(chart) {
    chart.ctx.restore();
  }
};

// 通用配置
Chart.defaults.font.family = '-apple-system,"Microsoft YaHei",sans-serif';
Chart.defaults.plugins.tooltip.backgroundColor = 'rgba(30,30,30,0.9)';
Chart.defaults.plugins.tooltip.titleFont = {size:13,weight:'bold'};
Chart.defaults.plugins.tooltip.bodyFont = {size:12};
Chart.defaults.plugins.tooltip.padding = 12;
Chart.defaults.plugins.tooltip.cornerRadius = 8;

const barOptions = {
  responsive:true, maintainAspectRatio:false,
  plugins: {legend:{display:false}},
  scales: {y:{beginAtZero:true,grid:{color:'#e8e8e8'},ticks:{font:{size:11}}},
           x:{grid:{display:false},ticks:{font:{size:11,weight:'bold'}}}},
};

// ═══ 各站别良率 ═══
(function(){
  const d = $CHART_STATION_YIELD$;
  const ctx = document.getElementById('chartStationYield').getContext('2d');
  new Chart(ctx, {
    type:'bar', plugins:[shadowPlugin],
    data:{
      labels: d.labels,
      datasets:[{
        data: d.data,
        backgroundColor: d.colors,
        borderColor: 'rgba(0,0,0,0.06)',
        borderWidth:1, borderRadius:8, borderSkipped:false,
      }]
    },
    options:{
      responsive:true, maintainAspectRatio:false,
      plugins:{
        legend:{display:false},
        tooltip:{callbacks:{label:ctx=>ctx.raw.toFixed(1)+'% ('+d.totals[ctx.dataIndex]+'台)'}}
      },
      scales:{
        y:{beginAtZero:false,min:Math.max(0,Math.min(...d.data)-8),max:102,
           grid:{color:'#e8e8e8'},ticks:{callback:v=>v+'%',font:{size:11}}},
        x:{grid:{display:false},ticks:{font:{size:12,weight:'bold'}}}
      }
    }
  });
})();

// ═══ 失败原因分布 ═══
$CHART_FAILURE_REASONS_BLOCK$

// ═══ 各站别 Top3 高频失败项 ═══
$CHART_SN_FAIL_DETAIL_BLOCK$

// ═══ UPH 分时段 ═══
$CHART_UPH_STATIONS_BLOCK$
</script></body></html>"""

def _build_failure_reasons_chart_js(a):
    """生成失败原因图表的 JavaScript"""
    fc = a.get("failure_counter", {})
    if not fc:
        return "// 无失败数据"
    items = sorted(fc.items(), key=lambda x: -x[1])[:12]
    labels = [k[:28] for k, _ in items]
    data = [v for _, v in items]
    return f"""(function(){{
  const d = {{labels:{json.dumps(labels)},data:{json.dumps(data)}}};
  if(!d.labels.length) return;
  const ctx = document.getElementById('chartFailureReasons').getContext('2d');
  new Chart(ctx, {{
    type:'bar', plugins:[shadowPlugin], options:{{indexAxis:'y',...barOptions}},
    data:{{
      labels: d.labels,
      datasets:[{{
        data: d.data,
        backgroundColor: d.data.map((_,i)=>['#F44336','#E53935','#EF5350','#E57373','#EF9A9A','#FFCDD2','#FF8A80','#FF5252','#D32F2F','#C62828','#B71C1C','#F44336'][i]),
        borderColor:'rgba(0,0,0,0.04)',borderWidth:1,borderRadius:6,
      }}]
    }}
  }});
}})();"""

def _build_sn_fail_detail_chart_js(a):
    """生成各站别Top3高频失败项的 JavaScript（带失败原因标注）"""
    ss = a.get("station_stats", {})
    if not ss:
        return "// 无失败数据"
    stations = sorted(ss.keys())
    colors = ["#E53935", "#FB8C00", "#1E88E5"]
    datasets = []
    for pos, color in enumerate(colors):
        vals = []
        lbls = []
        for st in stations:
            counter = Counter()
            for sn_info in ss[st].get("failed_sns", []):
                for item in sn_info["failed"]:
                    counter[item] += 1
            items = counter.most_common(3)
            if pos < len(items):
                vals.append(items[pos][1])
                lbls.append(items[pos][0][:14])
            else:
                vals.append(0)
                lbls.append("")
        datasets.append({
            "label": ["#1最多", "#2", "#3"][pos],
            "data": vals,
            "backgroundColor": color,
            "borderColor": "rgba(255,255,255,0.6)",
            "borderWidth": 1,
            "borderRadius": 5,
            "itemLabels": lbls,  # 失败原因名称
        })

    return f"""(function(){{
  const datasets = {json.dumps(datasets)};
  const labels = {json.dumps(stations)};
  if(!labels.length) return;

  // 柱状图标签插件
  const barLabelPlugin = {{
    id: 'barLabels',
    afterDatasetsDraw(chart) {{
      const ctx = chart.ctx;
      chart.data.datasets.forEach((ds, di) => {{
        const meta = chart.getDatasetMeta(di);
        if(!meta) return;
        meta.data.forEach((bar, i) => {{
          const label = ds.itemLabels ? ds.itemLabels[i] : '';
          if(!label || ds.data[i] === 0) return;
          ctx.save();
          ctx.fillStyle = '#333';
          ctx.font = 'bold 8px -apple-system,"Microsoft YaHei",sans-serif';
          ctx.textAlign = 'center';
          ctx.translate(bar.x, bar.y - 6);
          ctx.rotate(-0.7);
          ctx.fillText(label, 0, 0);
          ctx.restore();
        }});
      }});
    }}
  }};

  const ctx = document.getElementById('chartSnFailDetail').getContext('2d');
  new Chart(ctx, {{
    type:'bar', plugins:[shadowPlugin, barLabelPlugin],
    data:{{labels, datasets}},
    options:{{
      responsive:true, maintainAspectRatio:false,
      layout:{{padding:{{top:45}}}},
      plugins:{{legend:{{position:'top',labels:{{font:{{size:10}},padding:15,usePointStyle:true}}}}}},
      scales:{{
        y:{{beginAtZero:true,grid:{{color:'#e8e8e8'}},ticks:{{font:{{size:11}},stepSize:1}}}},
        x:{{grid:{{display:false}},ticks:{{font:{{size:12,weight:'bold'}}}}}}
      }}
    }}
  }});
}})();"""

def _build_uph_stations_chart_js(a):
    """生成总体UPH分时段图表（每小时实际测试产品数量总数）"""
    hourly = a.get("uph", {}).get("hourly", {})
    if not hourly:
        return "// 无UPH数据"
    labels = list(hourly.keys())
    data = list(hourly.values())
    total = sum(data)
    return f"""(function(){{
  const labels = {json.dumps(labels)};
  const data = {json.dumps(data)};
  if(!labels.length) return;
  const ctx = document.getElementById('chartUphStations').getContext('2d');
  new Chart(ctx, {{
    type:'bar', plugins:[shadowPlugin],
    data:{{
      labels,
      datasets:[{{
        label:'测试台数',
        data,
        backgroundColor: '#2196F3',
        borderColor:'rgba(0,0,0,0.05)',borderWidth:1,borderRadius:6,
      }}]
    }},
    options:{{
      responsive:true, maintainAspectRatio:false,
      plugins:{{
        legend:{{display:false}},
        tooltip:{{callbacks:{{label:ctx=>ctx.raw+' 台'}}}}
      }},
      scales:{{
        y:{{beginAtZero:true,grid:{{color:'#e8e8e8'}},ticks:{{font:{{size:11}},callback:v=>v+'台'}}}},
        x:{{grid:{{display:false}},ticks:{{font:{{size:11,weight:'bold'}}}}}}
      }}
    }}
  }});
}})();"""

def make_html(a, out_dir, out_path, source_info=""):
    """生成自包含 HTML 报告（Chart.js 图表，无需外部PNG）"""
    # Station table rows
    st_rows = []
    uph = a.get("uph", {}).get("stations", {})
    for s in sorted(a["station_stats"].keys()):
        d = a["station_stats"][s]
        r = d["pass"]/d["total"]*100 if d["total"] else 0
        cls = "bg" if r>=97 else "bo" if r>=95 else "br"
        su = uph.get(s, {})
        st_rows.append(
            f"<tr><td><strong>{s}</strong></td><td>{d['total']}</td>"
            f"<td>{d['pass']}</td><td>{d['fail']}</td>"
            f"<td><span class='badge {cls}'>{r:.1f}%</span></td>"
            f"<td>{su.get('uph','-')}</td><td style='font-size:11px;color:var(--muted)'>{su.get('start','-')}~{su.get('end','-')}</td></tr>"
        )

    # SN table rows
    sn_rows = []
    for s in a["fail_list"]:
        items = ", ".join(s["failed"][:5])
        sn_rows.append(
            f"<tr><td><code>{s['sn']}</code></td><td>{s['station']}</td>"
            f"<td>{len(s['failed'])}/{s['total']}</td><td style='font-size:11px'>{items}</td></tr>"
        )

    # UPH cards
    uph_cards = []
    for st in sorted(uph.keys()):
        su = uph[st]
        uph_cards.append(
            f"<div class='uph-card'><div class='st'>{st}</div>"
            f"<div class='val'>{su['uph']} <span style='font-size:12px;font-weight:400;color:var(--muted)'>台/时</span></div>"
            f"<div class='detail'>{su['sn']}台 · {su['hours']}h · {su['start']}~{su['end']}</div></div>"
        )

    yc = "g" if a["yield_rate"]>=97 else "o" if a["yield_rate"]>=95 else "r"

    # Station yield chart data
    sts = sorted(a["station_stats"].keys())
    sy_data = [a["station_stats"][s]["pass"]/a["station_stats"][s]["total"]*100 for s in sts]
    sy_colors = ["#4CAF50" if y>=97 else "#FF9800" if y>=95 else "#F44336" for y in sy_data]
    sy_totals = [a["station_stats"][s]["total"] for s in sts]
    station_yield_json = json.dumps({"labels": sts, "data": sy_data, "colors": sy_colors, "totals": sy_totals})

    # UPH overall
    uo = a.get("uph", {}).get("overall", {})

    # Build
    html = _HTML
    html = html.replace("$REPORT_TIME$", datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    html = html.replace("$SOURCE_INFO$", source_info or "-")
    html = html.replace("$TOTAL_SN$", str(a["total_sn"]))
    html = html.replace("$PASS_SN$", str(a["pass_sn"]))
    html = html.replace("$FAIL_SN$", str(a["fail_sn"]))
    html = html.replace("$YIELD_RATE$", f"{a['yield_rate']:.1f}")
    html = html.replace("$YIELD_CLASS$", yc)
    html = html.replace("$UPH_VALUE$", str(uo.get("uph", "-")))
    html = html.replace("$UPH_HOURS$", str(uo.get("hours", "-")))
    html = html.replace("$UPH_START$", str(uo.get("start", "-")))
    html = html.replace("$UPH_END$", str(uo.get("end", "-")))
    html = html.replace("$UPH_CARDS$", "".join(uph_cards))
    html = html.replace("$STATION_TABLE_ROWS$", "".join(st_rows))
    html = html.replace("$SN_TABLE_ROWS$", "".join(sn_rows) if sn_rows else "<tr><td colspan='4' style='text-align:center;color:var(--green)'>✅ 全部通过！</td></tr>")
    html = html.replace("$CHART_STATION_YIELD$", station_yield_json)
    html = html.replace("$CHART_FAILURE_REASONS_BLOCK$", _build_failure_reasons_chart_js(a))
    html = html.replace("$CHART_SN_FAIL_DETAIL_BLOCK$", _build_sn_fail_detail_chart_js(a))
    html = html.replace("$CHART_UPH_STATIONS_BLOCK$", _build_uph_stations_chart_js(a))

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    return out_path

# ═══════════════════════════════════════════════
# GUI
# ═══════════════════════════════════════════════
def _open_with_default(path):
    if not path or not os.path.exists(path): return
    try:
        if sys.platform=="win32": os.startfile(path)
        elif sys.platform=="darwin": subprocess.run(["open",path])
        else: subprocess.run(["xdg-open",path])
    except: pass

def _run_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    class App:
        def __init__(self, root):
            self.root = root
            root.title("AT Audio Test 分析工具 v4.1")
            root.geometry("1050x700"); root.minsize(850,550)
            root.configure(bg=STYLE["bg"])
            self.src = tk.StringVar()
            self.out = tk.StringVar(value=os.path.join(os.getcwd(),"at_report_v3"))
            self.status = tk.StringVar(value="就绪 — 选择数据源（ZIP/文件夹），点开始分析")
            self.prog = tk.DoubleVar(value=0)
            self.opt_charts = tk.BooleanVar(value=True)
            self.opt_html = tk.BooleanVar(value=True)
            self.opt_csv = tk.BooleanVar(value=True)
            self.opt_open = tk.BooleanVar(value=True)
            self.records = None; self.analysis = None; self.html_path = None; self.running = False
            self._build()
            root.update_idletasks()
            w,h = root.winfo_width(), root.winfo_height()
            sw,sh = root.winfo_screenwidth(), root.winfo_screenheight()
            root.geometry(f"+{(sw-w)//2}+{(sh-h)//2}")

        def _build(self):
            # Header
            hdr = tk.Frame(self.root, bg=STYLE["accent"], height=48)
            hdr.pack(fill="x"); hdr.pack_propagate(False)
            tk.Label(hdr, text="🔊 AT Audio Test 分析工具 v4.1", bg=STYLE["accent"], fg="white",
                     font=("Microsoft YaHei",15,"bold")).pack(side="left", padx=20, pady=10)

            main = tk.Frame(self.root, bg=STYLE["bg"])
            main.pack(fill="both", expand=True, padx=15, pady=(10,0))
            left = tk.Frame(main, bg=STYLE["bg"]); left.pack(side="left", fill="y", padx=(0,10))

            # Source
            f1 = tk.LabelFrame(left, text="📁 数据源", bg=STYLE["card_bg"], fg=STYLE["fg"],
                               font=("Microsoft YaHei",11,"bold"), padx=12, pady=10, relief="groove", bd=1)
            f1.pack(fill="x", pady=(0,8))
            r1 = tk.Frame(f1, bg=STYLE["card_bg"]); r1.pack(fill="x")
            self.se = tk.Entry(r1, textvariable=self.src, font=("Consolas",10), bg=STYLE["input_bg"], relief="solid", bd=1)
            self.se.pack(side="left", fill="x", expand=True, ipady=3)
            tk.Button(r1, text="📂 选择", command=self._pick_src, bg=STYLE["accent"], fg="white",
                      font=("Microsoft YaHei",9), relief="flat", padx=12, pady=3, cursor="hand2").pack(side="left", padx=(6,0))
            r2 = tk.Frame(f1, bg=STYLE["card_bg"]); r2.pack(fill="x", pady=(6,0))
            for t,c in [("📦 ZIP","zip"),("📁 文件夹","dir"),("📄 单文件","file")]:
                tk.Button(r2, text=t, command=lambda m=c: self._pick_src(m), bg="#ECEFF1", fg=STYLE["fg"],
                          font=("Microsoft YaHei",9), relief="flat", padx=10, pady=2, cursor="hand2").pack(side="left", padx=(0,6))

            # Output
            f2 = tk.LabelFrame(left, text="📂 输出目录", bg=STYLE["card_bg"], fg=STYLE["fg"],
                               font=("Microsoft YaHei",11,"bold"), padx=12, pady=10, relief="groove", bd=1)
            f2.pack(fill="x", pady=(0,8))
            r = tk.Frame(f2, bg=STYLE["card_bg"]); r.pack(fill="x")
            self.oe = tk.Entry(r, textvariable=self.out, font=("Consolas",10), bg=STYLE["input_bg"], relief="solid", bd=1)
            self.oe.pack(side="left", fill="x", expand=True, ipady=3)
            tk.Button(r, text="📂 选择", command=self._pick_out, bg=STYLE["accent"], fg="white",
                      font=("Microsoft YaHei",9), relief="flat", padx=12, pady=3, cursor="hand2").pack(side="left", padx=(6,0))

            # Options
            f3 = tk.LabelFrame(left, text="⚙️ 输出", bg=STYLE["card_bg"], fg=STYLE["fg"],
                               font=("Microsoft YaHei",11,"bold"), padx=12, pady=10, relief="groove", bd=1)
            f3.pack(fill="x", pady=(0,8))
            for v,t in [(self.opt_charts,"PNG图表"),(self.opt_html,"HTML报告"),(self.opt_csv,"CSV明细"),(self.opt_open,"自动打开")]:
                tk.Checkbutton(f3, text=t, variable=v, bg=STYLE["card_bg"], font=("Microsoft YaHei",10),
                               activebackground=STYLE["card_bg"], selectcolor=STYLE["card_bg"]).pack(anchor="w", pady=1)

            # Run
            self.btn = tk.Button(left, text="▶  开始分析", command=self._start, bg=STYLE["accent"], fg="white",
                                  font=("Microsoft YaHei",12,"bold"), relief="flat", padx=30, pady=10, cursor="hand2")
            self.btn.pack(fill="x", pady=(5,0))

            # Right panel
            right = tk.Frame(main, bg=STYLE["bg"]); right.pack(side="left", fill="both", expand=True)

            # Summary cards
            cards = tk.Frame(right, bg=STYLE["bg"]); cards.pack(fill="x", pady=(0,6))
            self.cards = {}
            for i,(k,label,color) in enumerate([
                ("total","测试总数","#2196F3"),("pass","PASS","#4CAF50"),
                ("fail","FAIL","#F44336"),("yield","良率","#FF9800"),
                ("uph","UPH","#0D47A1")]):
                cd = tk.Frame(cards, bg=STYLE["card_bg"], relief="solid", bd=1, padx=8, pady=8)
                cd.grid(row=0, column=i, padx=3, sticky="nsew"); cards.grid_columnconfigure(i, weight=1)
                self.cards[k] = {
                    "v": tk.Label(cd, text="—", bg=STYLE["card_bg"], fg=color, font=("Consolas",20,"bold")),
                    "l": tk.Label(cd, text=label, bg=STYLE["card_bg"], fg=STYLE["text_secondary"], font=("Microsoft YaHei",9))
                }
                self.cards[k]["v"].pack(); self.cards[k]["l"].pack()

            # Notebook tabs
            nb = ttk.Notebook(right); nb.pack(fill="both", expand=True)

            # Tab1: Station stats
            t1 = tk.Frame(nb, bg=STYLE["card_bg"]); nb.add(t1, text="📊 站别统计")
            self.st_tree = ttk.Treeview(t1, columns=("st","total","pass","fail","yield"), show="headings", height=14)
            self.st_tree.heading("st", text="站别"); self.st_tree.column("st", width=80)
            self.st_tree.heading("total", text="测试数"); self.st_tree.column("total", width=70, anchor="center")
            self.st_tree.heading("pass", text="PASS"); self.st_tree.column("pass", width=70, anchor="center")
            self.st_tree.heading("fail", text="FAIL"); self.st_tree.column("fail", width=70, anchor="center")
            self.st_tree.heading("yield", text="良率"); self.st_tree.column("yield", width=80, anchor="center")
            self.st_tree.pack(fill="both", expand=True)
            for tag,c in [("good","#4CAF50"),("warn","#FF9800"),("bad","#F44336")]:
                self.st_tree.tag_configure(tag, foreground=c)

            # Tab2: Failure reasons
            t2 = tk.Frame(nb, bg=STYLE["card_bg"]); nb.add(t2, text="🔍 失败原因")
            self.fr_tree = ttk.Treeview(t2, columns=("reason","count","pct"), show="headings", height=14)
            self.fr_tree.heading("reason", text="失败测试项"); self.fr_tree.column("reason", width=280)
            self.fr_tree.heading("count", text="失败SN数"); self.fr_tree.column("count", width=90, anchor="center")
            self.fr_tree.heading("pct", text="占比"); self.fr_tree.column("pct", width=80, anchor="center")
            self.fr_tree.pack(fill="both", expand=True)

            # Tab3: Failed SNs
            t3 = tk.Frame(nb, bg=STYLE["card_bg"]); nb.add(t3, text="🔴 失败SN")
            self.sn_tree = ttk.Treeview(t3, columns=("sn","st","cnt","items"), show="headings", height=14)
            self.sn_tree.heading("sn", text="SN"); self.sn_tree.column("sn", width=170)
            self.sn_tree.heading("st", text="站别"); self.sn_tree.column("st", width=60, anchor="center")
            self.sn_tree.heading("cnt", text="失败项"); self.sn_tree.column("cnt", width=70, anchor="center")
            self.sn_tree.heading("items", text="失败详情"); self.sn_tree.column("items", width=350)
            self.sn_tree.pack(fill="both", expand=True)
            self.sn_tree.tag_configure("critical", foreground="#F44336")

            # Tab4: All SNs by station
            t4 = tk.Frame(nb, bg=STYLE["card_bg"]); nb.add(t4, text="📋 全部SN")
            # Use a sub-notebook for per-station tabs
            self.all_nb = ttk.Notebook(t4)
            self.all_nb.pack(fill="both", expand=True)
            self.all_trees = {}  # station -> treeview

            # Bottom
            btns = tk.Frame(right, bg=STYLE["bg"]); btns.pack(fill="x", pady=(6,0))
            self.br = tk.Button(btns, text="📄 打开报告", command=self._open_r, bg=STYLE["success"], fg="white",
                                 font=("Microsoft YaHei",10), relief="flat", padx=16, pady=5, state="disabled")
            self.br.pack(side="left", padx=(0,6))
            self.bf = tk.Button(btns, text="📂 输出目录", command=self._open_f, bg="#ECEFF1", fg=STYLE["fg"],
                                 font=("Microsoft YaHei",10), relief="flat", padx=16, pady=5, state="disabled")
            self.bf.pack(side="left", padx=(0,6))

            # Status bar
            sbar = tk.Frame(self.root, bg=STYLE["card_bg"], height=30)
            sbar.pack(fill="x", side="bottom", pady=(8,0)); sbar.pack_propagate(False)
            tk.Label(sbar, textvariable=self.status, bg=STYLE["card_bg"], fg=STYLE["text_secondary"],
                     font=("Microsoft YaHei",9), anchor="w").pack(side="left", fill="x", padx=12, pady=4)
            self.pb = ttk.Progressbar(sbar, variable=self.prog, mode="determinate", length=180)
            self.pb.pack(side="right", padx=12, pady=4)

        def _pick_src(self, mode=None):
            if mode=="zip": p = filedialog.askopenfilename(title="选ZIP", filetypes=[("ZIP","*.zip")])
            elif mode=="dir": p = filedialog.askdirectory(title="选文件夹")
            elif mode=="file": p = filedialog.askopenfilename(title="选文件", filetypes=[("数据","*.xls;*.txt")])
            else: p = filedialog.askopenfilename(title="选数据源", filetypes=[("支持","*.zip;*.xls;*.txt")]) or filedialog.askdirectory(title="或选文件夹")
            if p: self.src.set(p); self.status.set(f"已选: {os.path.basename(p)}")

        def _pick_out(self):
            p = filedialog.askdirectory(title="输出目录", initialdir=self.out.get())
            if p: self.out.set(p)

        def _start(self):
            s = self.src.get().strip()
            if not s: messagebox.showwarning("提示","请先选择数据源"); return
            if not os.path.exists(s): messagebox.showerror("错误","路径不存在"); return
            if self.running: return
            self.running = True
            self.btn.configure(text="⏳ 分析中...", state="disabled", bg=STYLE["text_secondary"])
            self.prog.set(5); self.status.set("解析中..."); self._clear()
            threading.Thread(target=self._run, args=(s,), daemon=True).start()

        def _run(self, s):
            try:
                self.root.after(0, lambda: self.prog.set(10))
                recs, skipped, st_files = parse_source(s, cb=lambda ph,dt: self.root.after(0, lambda: self.status.set(ph)))
                if not recs: raise ValueError("无有效记录")
                self.records = recs
                self.root.after(0, lambda: self.prog.set(40))
                sinfo = " ".join(f"{st}({cnt})" for st,cnt in sorted(st_files.items()))
                n_stations = len(st_files)
                warn_msg = ""
                if n_stations < 6:
                    missing = [f"AT0{i}" for i in range(1,8) if f"AT0{i}" not in st_files]
                    warn_msg = f" ⚠️ 仅检测到{n_stations}个站别，缺少: {', '.join(missing)}"
                self.root.after(0, lambda: self.status.set(f"解析 {len(recs)}行 | 站别: {sinfo}{warn_msg}"))
                if n_stations < 6:
                    self.root.after(0, lambda: messagebox.showwarning("警告", f"仅检测到 {n_stations} 个站别，请检查数据源是否完整。\n缺少: {', '.join(missing)}"))
                if skipped:
                    sk = "; ".join(skipped[:5])
                    self.root.after(0, lambda: self.status.set(f"{self.status.get()} | 跳过: {sk}"))
                a = analyze(recs); self.analysis = a
                uo = a.get("uph", {}).get("overall", {})
                self.root.after(0, lambda: self.prog.set(60))
                self.root.after(0, lambda: self._show(a))
                od = self.out.get(); os.makedirs(od, exist_ok=True)

                if self.opt_charts.get():
                    self.root.after(0, lambda: self.prog.set(70))
                    _get_font()
                    chart_station_yield(a, os.path.join(od, "chart_station_yield.png"))
                    chart_failure_reasons(a, os.path.join(od, "chart_failure_reasons.png"))
                    chart_sn_fail_detail(a, os.path.join(od, "chart_sn_fail_detail.png"))
                if self.opt_csv.get():
                    self.root.after(0, lambda: self.prog.set(85))
                    self._write_csv(os.path.join(od, "detail.csv"))
                if self.opt_html.get():
                    self.root.after(0, lambda: self.prog.set(95))
                    self.html_path = make_html(a, od, os.path.join(od, "report.html"), os.path.basename(s))
                self.root.after(0, lambda: self.prog.set(100))
                msg = f"✅ 完成！总数{a['total_sn']}台 PASS={a['pass_sn']} FAIL={a['fail_sn']} 良率{a['yield_rate']:.1f}% UPH={uo.get('uph','-')}"
                self.root.after(0, lambda: self.status.set(msg))
                self.root.after(0, self._done)
            except Exception as e:
                import traceback; traceback.print_exc()
                self.root.after(0, lambda: self.status.set(f"❌ {e}"))
                self.root.after(0, lambda: self.prog.set(0))
                self.root.after(0, self._err)

        def _write_csv(self, p):
            if not self.records: return
            ks = ["sn","test_ch","test_name","display","result","value","station","time"]
            with open(p, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.DictWriter(f, fieldnames=ks, extrasaction="ignore"); w.writeheader()
                for r in self.records:
                    row = {k: r.get(k,"") for k in ks}
                    if row.get("time"): row["time"] = str(row["time"])
                    w.writerow(row)

        def _show(self, a):
            self.cards["total"]["v"].configure(text=str(a["total_sn"]))
            self.cards["pass"]["v"].configure(text=str(a["pass_sn"]))
            self.cards["fail"]["v"].configure(text=str(a["fail_sn"]))
            self.cards["yield"]["v"].configure(text=f"{a['yield_rate']:.1f}%")
            uo = a.get("uph", {}).get("overall", {})
            self.cards["uph"]["v"].configure(text=str(uo.get("uph", "-")))
            for t in [self.st_tree, self.fr_tree, self.sn_tree]:
                for i in t.get_children(): t.delete(i)
            # Clear all_sn sub-tabs
            for st, (frame, tree) in list(self.all_trees.items()):
                self.all_nb.forget(frame)
            self.all_trees.clear()

            # Station
            for s in sorted(a["station_stats"].keys()):
                d = a["station_stats"][s]
                r = d["pass"]/d["total"]*100 if d["total"] else 0
                tag = "good" if r>=97 else "warn" if r>=95 else "bad"
                self.st_tree.insert("", "end", values=(s, d["total"], d["pass"], d["fail"], f"{r:.1f}%"), tags=(tag,))

            # Failure reasons
            total_fail = a["fail_sn"]
            for reason, cnt in sorted(a["failure_counter"].items(), key=lambda x: -x[1]):
                pct = cnt/total_fail*100 if total_fail else 0
                self.fr_tree.insert("", "end", values=(reason, cnt, f"{pct:.1f}%"))

            # Failed SNs
            for s in a["fail_list"]:
                items = ", ".join(s["failed"][:5])
                self.sn_tree.insert("", "end", values=(s["sn"], s["station"], f"{len(s['failed'])}/{s['total']}", items))

            # All SNs by station
            for st in sorted(a["all_sn_by_station"].keys()):
                sns = sorted(a["all_sn_by_station"][st], key=lambda x: x["sn"])
                frame = tk.Frame(self.all_nb, bg=STYLE["card_bg"])
                tree = ttk.Treeview(frame, columns=("sn","result","fail_items"), show="headings", height=12)
                tree.heading("sn", text="SN"); tree.column("sn", width=170)
                tree.heading("result", text="结果"); tree.column("result", width=60, anchor="center")
                tree.heading("fail_items", text="失败项"); tree.column("fail_items", width=400)
                tree.pack(fill="both", expand=True)
                tree.tag_configure("pass", foreground="#4CAF50")
                tree.tag_configure("fail", foreground="#F44336")
                for sn in sns:
                    tag = "pass" if sn["passed"] else "fail"
                    result = "✅ PASS" if sn["passed"] else "❌ FAIL"
                    fails = ", ".join(sn["failed"][:5]) if sn["failed"] else ""
                    tree.insert("", "end", values=(sn["sn"], result, fails), tags=(tag,))
                self.all_nb.add(frame, text=f"{st}({len(sns)})")
                self.all_trees[st] = (frame, tree)

        def _clear(self):
            for k in self.cards: self.cards[k]["v"].configure(text="—")
            for t in [self.st_tree, self.fr_tree, self.sn_tree]:
                for i in t.get_children(): t.delete(i)
            for st, (frame, tree) in list(self.all_trees.items()):
                self.all_nb.forget(frame)
            self.all_trees.clear()
            self.html_path = None
            self.br.configure(state="disabled"); self.bf.configure(state="disabled")

        def _done(self):
            self.running = False
            self.btn.configure(text="▶  重新分析", state="normal", bg=STYLE["accent"])
            self.br.configure(state="normal"); self.bf.configure(state="normal")
            if self.opt_open.get() and self.html_path: self._open_r()

        def _err(self):
            self.running = False
            self.btn.configure(text="▶  开始分析", state="normal", bg=STYLE["accent"])

        def _open_r(self):
            if self.html_path and os.path.exists(self.html_path): _open_with_default(self.html_path)
            else: messagebox.showinfo("提示","报告不存在")

        def _open_f(self):
            o = self.out.get()
            if os.path.exists(o): _open_with_default(o)
            else: messagebox.showinfo("提示","目录不存在")

    root = tk.Tk()
    try:
        from ctypes import windll; windll.shcore.SetProcessDpiAwareness(1)
    except: pass
    App(root); root.mainloop()

def main(): _run_gui()
if __name__ == "__main__": main()

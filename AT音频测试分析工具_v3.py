#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AT Audio Test 数据分析工具 v3.1 — 周期去重版
=============================================
每台SN会被重复测试多次（~19个周期），每次周期从 FR 开始：
  FR → Frequency(信息行) → Rub&Buzz → THD

本工具按测试周期去重，输出：
  - 周期级统计（PASS/FAIL 周期数）
  - SN级最终结果（任一周期间FAIL=该SN不良）
  - HTML报告 + PNG图表 + CSV明细

运行: python AT音频测试分析工具_v3.py
"""

import os, sys, re, csv, zipfile, threading, datetime, warnings, webbrowser, subprocess
from pathlib import Path
from collections import defaultdict, Counter
from tempfile import TemporaryDirectory

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from matplotlib.font_manager import FontProperties

warnings.filterwarnings("ignore")

# ═══════════════════════════════════════════════
# 常量
# ═══════════════════════════════════════════════
_CJK = ["SimHei","Microsoft YaHei","PingFang SC","Noto Sans CJK SC",
        "WenQuanYi Micro Hei","STHeiti","SimSun"]
_ENCS = ["gbk","gb2312","gb18030","utf-8","utf-16","latin-1"]
C = {"pass":"#4CAF50","fail":"#F44336","bar_pass":"#81C784","bar_fail":"#E57373",
     "warn":"#FF9800","accent":"#2196F3"}
STYLE = {"bg":"#f0f2f5","fg":"#1a1a2e","card_bg":"#ffffff","accent":"#2196F3",
         "accent_hover":"#1976D2","success":"#4CAF50","danger":"#F44336",
         "warning":"#FF9800","text_secondary":"#78909c","input_bg":"#ffffff",
         "tree_bg":"#ffffff","tree_sel":"#E3F2FD","progress_bg":"#e0e0e0"}
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

def _detect_enc(fp):
    for e in _ENCS:
        try:
            with open(fp,"rb") as f: d=f.read(4096)
            if "SN" in d.decode(e,errors="replace"): return e
        except: continue
    return "gbk"

# ═══════════════════════════════════════════════
# 解析器 — 识别 FR 为周期边界
# ═══════════════════════════════════════════════
def parse_file(filepath):
    """解析 TC661_data.xls。每行: Time\\tSN\\tTestChNum\\tTestName\\tUnit\\tResult\\tChannel\\tValue\\tUpper\\tLower
    返回 [{sn, test_name, result, value, upper, lower, station, time}]
    'Frequency' 行为信息行(result为空)，保留用于上下文但不参与判等。
    """
    enc = _detect_enc(filepath)
    p = Path(filepath)
    station = p.parent.parent.parent.parent.name  # ATxx
    if not station.startswith("AT"):
        station = p.parent.parent.parent.name

    records = []
    with open(filepath,"rb") as f: raw = f.read()
    text = raw.decode(enc, errors="replace")
    for line in text.split("\n"):
        line = line.strip()
        if not line or line.startswith("Time"): continue
        flds = line.split("\t")
        if len(flds) < 7: continue
        time_str = flds[0].strip()
        sn = flds[1].strip().strip("'\"")
        test_name = flds[3].strip()
        result = flds[5].strip()

        try:
            t = datetime.datetime.strptime(time_str, "%Y/%m/%d %H:%M:%S")
        except: continue

        # Parse value/upper/lower if present
        val_str = flds[7].strip() if len(flds) > 7 else ""
        value = None; upper = None; lower = None
        if val_str:
            try:
                value = float(val_str)
                upper = float(flds[8].strip()) if len(flds)>8 and flds[8].strip() else None
                lower = float(flds[9].strip()) if len(flds)>9 and flds[9].strip() else None
            except: pass

        records.append({
            "sn": sn, "test_name": test_name, "result": result,
            "value": value, "upper": upper, "lower": lower,
            "station": station, "time": t,
        })
    return records

def parse_source(src, cb=None):
    """统一入口，返回 (records, skipped)"""
    all_recs = []; skipped = []
    if src.endswith(".zip"):
        if cb: cb("解压中...",src)
        with TemporaryDirectory() as td:
            with zipfile.ZipFile(src,"r") as zf: zf.extractall(td)
            for root,_,files in os.walk(td):
                for fn in files:
                    if fn.lower().endswith((".xls",".txt")):
                        fp = os.path.join(root,fn)
                        try: all_recs.extend(parse_file(fp))
                        except Exception as e: skipped.append(f"{fn}: {e}")
    elif os.path.isdir(src):
        if cb: cb("扫描中...",src)
        for root,_,files in os.walk(src):
            for fn in files:
                if fn.lower().endswith((".xls",".txt")):
                    fp = os.path.join(root,fn)
                    try: all_recs.extend(parse_file(fp))
                    except Exception as e: skipped.append(f"{fn}: {e}")
    else:
        all_recs = parse_file(src)
    return all_recs, skipped

# ═══════════════════════════════════════════════
# 分析器 — 按测试周期去重
# ═══════════════════════════════════════════════
def analyze(records):
    """按FR为周期边界，去重统计"""
    if not records: return {"error": "无记录"}

    # Step 1: 按SN分组并排序
    sn_data = defaultdict(list)
    for r in records:
        sn_data[r["sn"]].append(r)
    for sn in sn_data:
        sn_data[sn].sort(key=lambda x: x["time"])

    # Step 2: 按FR边界切分成周期
    all_cycles = []  # [{sn, station, items: [{test_name, result}], pass: bool}]
    sn_cycles = defaultdict(list)

    for sn, entries in sn_data.items():
        cycles = []; current = []
        for r in entries:
            if r["test_name"] == "FR":
                if current: cycles.append(current)
                current = [r]
            elif current:
                current.append(r)
        if current: cycles.append(current)

        for cyc in cycles:
            # 判定：有Result且为Fail才算失败(忽略Frequency这类info行)
            fails = [r for r in cyc if r["result"] and r["result"] not in ("Pass","pass","")]
            cycle_pass = len(fails) == 0
            failed_items = list(set(r["test_name"] for r in fails))

            cycle_info = {
                "sn": sn,
                "station": cyc[0]["station"],
                "pass": cycle_pass,
                "items": [r["test_name"] for r in cyc],
                "failed_items": failed_items,
                "time": cyc[0]["time"],
            }
            all_cycles.append(cycle_info)
            sn_cycles[sn].append(cycle_info)

    # Step 3: 周期级统计
    total_cycles = len(all_cycles)
    fail_cycles = [c for c in all_cycles if not c["pass"]]
    pass_cycles = total_cycles - len(fail_cycles)

    # Per station (cycle-level)
    by_station = defaultdict(lambda: {"cycles":0,"pass":0,"fail":0})
    for c in all_cycles:
        s = c["station"]
        by_station[s]["cycles"] += 1
        if c["pass"]: by_station[s]["pass"] += 1
        else: by_station[s]["fail"] += 1

    # Per test item (in failed cycles only, deduplicated per cycle)
    by_test = defaultdict(lambda: {"total_cycles":0,"fail_cycles":0})
    # Count which items appear in which cycles
    for c in all_cycles:
        seen_in_cycle = set()
        for item in c["items"]:
            if not item: continue  # skip blanks
            if item not in seen_in_cycle:
                by_test[item]["total_cycles"] += 1
                seen_in_cycle.add(item)
    for c in fail_cycles:
        for item in c["failed_items"]:
            if not item: continue
            by_test[item]["fail_cycles"] += 1

    # Step 4: SN级最终结果
    sn_pass = 0; sn_fail = 0
    sn_fail_list = []
    for sn, cycles in sn_cycles.items():
        final_pass = all(c["pass"] for c in cycles)
        if final_pass:
            sn_pass += 1
        else:
            sn_fail += 1
        fail_count = sum(1 for c in cycles if not c["pass"])
        if fail_count > 0:
            all_failed = set()
            for c in cycles:
                if not c["pass"]:
                    all_failed.update(c["failed_items"])
            sn_fail_list.append({
                "sn": sn, "fail_cycles": fail_count, "total_cycles": len(cycles),
                "rate": fail_count/len(cycles)*100,
                "failed_items": sorted(all_failed),
                "station": cycles[0]["station"],
            })

    sn_fail_list.sort(key=lambda x: -x["fail_cycles"])

    return {
        "total_raw": len(records),         # 原始行数
        "total_cycles": total_cycles,       # 去重后周期数
        "pass_cycles": pass_cycles,
        "fail_cycles": len(fail_cycles),
        "pass_rate": pass_cycles/total_cycles*100 if total_cycles else 0,
        "fail_rate": len(fail_cycles)/total_cycles*100 if total_cycles else 0,
        "unique_sns": len(sn_data),
        "sn_pass": sn_pass, "sn_fail": sn_fail,
        "sn_pass_rate": sn_pass/len(sn_data)*100 if sn_data else 0,
        "by_station": dict(by_station),
        "by_test": dict(by_test),
        "sn_fail_list": sn_fail_list,
        "avg_cycles_per_sn": total_cycles/len(sn_data) if sn_data else 0,
    }

# ═══════════════════════════════════════════════
# 图表
# ═══════════════════════════════════════════════
def _chart_test_fail_bars(a, out):
    _get_font()
    tests = sorted(a["by_test"].keys(), key=lambda t: -a["by_test"][t]["fail_cycles"])
    labels = [t[:20] for t in tests]
    fails = [a["by_test"][t]["fail_cycles"] for t in tests]
    totals = [a["by_test"][t]["total_cycles"] for t in tests]
    rates = [f/t*100 if t else 0 for f,t in zip(fails,totals)]

    fig,ax = plt.subplots(figsize=(max(8,len(tests)*0.5),5))
    bars = ax.bar(range(len(tests)), fails, color=[C["fail"] if r>5 else C["warn"] if r>1 else C["bar_pass"] for r in rates])
    ax.set_xticks(range(len(tests))); ax.set_xticklabels(labels, rotation=45, ha="right", fontsize=9)
    ax.set_ylabel("失败周期数",fontsize=12)
    ax.set_title("各测试项失败周期数",fontsize=14,fontweight="bold")
    for i,(f,r) in enumerate(zip(fails,rates)):
        ax.text(i,f+0.5,f"{r:.1f}%",ha="center",va="bottom",fontsize=8,fontweight="bold",color=C["fail"] if r>5 else "#555")
    ax.yaxis.set_major_locator(mticker.MaxNLocator(integer=True))
    plt.tight_layout(); fig.savefig(out,dpi=150,bbox_inches="tight"); plt.close(fig)

def _chart_station_fail_bars(a, out):
    _get_font()
    sts = sorted(a["by_station"].keys())
    fails = [a["by_station"][s]["fail"] for s in sts]
    totals = [a["by_station"][s]["cycles"] for s in sts]
    rates = [f/t*100 if t else 0 for f,t in zip(fails,totals)]
    fig,ax = plt.subplots(figsize=(max(8,len(sts)*1.5),5))
    bars = ax.bar(sts, rates, color=[C["fail"] if r>5 else C["warn"] if r>1 else C["bar_pass"] for r in rates])
    ax.set_ylabel("周期不良率 (%)",fontsize=12)
    ax.set_title("各站别测试周期不良率",fontsize=14,fontweight="bold")
    for bar,r in zip(bars,rates):
        ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+0.1, f"{r:.1f}%", ha="center",fontsize=10,fontweight="bold")
    ax.set_ylim(0,max(rates)*1.3+1)
    plt.tight_layout(); fig.savefig(out,dpi=150,bbox_inches="tight"); plt.close(fig)

def _chart_sn_fail_dist(a, out):
    _get_font()
    sl = a["sn_fail_list"]
    if not sl: return
    fail_counts = [s["fail_cycles"] for s in sl]
    dist = Counter(fail_counts)
    cats = sorted(dist.keys())
    fig,ax = plt.subplots(figsize=(8,5))
    ax.bar([f"{c}周期失败" for c in cats], [dist[c] for c in cats],
           color=[C["fail"] if c>5 else C["warn"] if c>1 else C["bar_pass"] for c in cats])
    ax.set_ylabel("SN数量",fontsize=12)
    ax.set_title(f"失败SN中失败周期数分布（共{len(sl)}个不良SN）",fontsize=14,fontweight="bold")
    ax.yaxis.set_major_locator(mticker.MaxNLocator(integer=True))
    plt.tight_layout(); fig.savefig(out,dpi=150,bbox_inches="tight"); plt.close(fig)

# ═══════════════════════════════════════════════
# HTML 报告
# ═══════════════════════════════════════════════
_HTML = """<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>AT Audio Test 数据分析报告 v3.1</title><style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:-apple-system,"Microsoft YaHei","PingFang SC",sans-serif;background:#f5f7fa;color:#333;padding:20px}}
.container{{max-width:1200px;margin:0 auto}}
h1{{text-align:center;color:#1a237e;margin-bottom:6px;font-size:28px}}
.subtitle{{text-align:center;color:#78909c;margin-bottom:30px;font-size:14px}}
.summary-cards{{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:15px;margin-bottom:30px}}
.card{{background:white;border-radius:10px;padding:20px;text-align:center;box-shadow:0 2px 8px rgba(0,0,0,0.08)}}
.card .value{{font-size:36px;font-weight:700}}
.card .label{{font-size:13px;color:#78909c;margin-top:4px}}
.card.pass .value{{color:#4CAF50}}.card.fail .value{{color:#F44336}}
.card.warn .value{{color:#FF9800}}.card.info .value{{color:#2196F3}}
.chart-section{{background:white;border-radius:10px;padding:20px;margin-bottom:20px;box-shadow:0 2px 8px rgba(0,0,0,0.08)}}
.chart-section h2{{color:#37474f;margin-bottom:15px;font-size:18px;border-left:4px solid #2196F3;padding-left:12px}}
.chart-section img{{max-width:100%;height:auto;display:block;margin:0 auto}}
table{{width:100%;border-collapse:collapse;font-size:13px}}
th{{background:#263238;color:white;padding:10px 8px;text-align:left;font-weight:600}}
td{{padding:8px;border-bottom:1px solid #eceff1}}tr:hover td{{background:#f5f5f5}}
.badge{{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600;color:white}}
.badge-pass{{background:#4CAF50}}.badge-fail{{background:#F44336}}.badge-warn{{background:#FF9800}}
.note{{background:#FFF3E0;border-left:4px solid #FF9800;padding:12px 16px;margin-bottom:20px;border-radius:4px;font-size:13px;color:#E65100}}
</style></head><body><div class="container">
<h1>🔊 AT Audio Test 数据分析报告 v3.1</h1>
<p class="subtitle">生成时间: {report_time}  |  数据来源: {source_info}  |  去重方式: FR周期边界</p>
<div class="note">
📌 <b>去重说明：</b>每条SN进行多次测试周期（以FR起始为一个周期），每个周期含 FR→Frequency→Rub&Buzz→THD 四项。
下表为<b>周期级</b>统计。每个SN平均 {avg_cycles:.0f} 个周期，{sn_fail}台SN至少一个周期失败即为SN不良。
</div>
<div class="summary-cards">
<div class="card info"><div class="value">{total_cycles}</div><div class="label">测试周期(去重)</div></div>
<div class="card pass"><div class="value">{pass_cycles}</div><div class="label">PASS周期</div></div>
<div class="card fail"><div class="value">{fail_cycles}</div><div class="label">FAIL周期</div></div>
<div class="card {rate_class}"><div class="value">{pass_rate:.1f}%</div><div class="label">周期合格率</div></div>
<div class="card info"><div class="value">{unique_sns}</div><div class="label">测试SN数</div></div>
<div class="card {sn_class}"><div class="value">{sn_pass_rate:.1f}%</div><div class="label">SN良率</div></div>
</div>
<div class="chart-section"><h2>📊 各测试项(失败周期数)</h2><img src="chart_test_fails.png"></div>
<div class="chart-section"><h2>📊 各站别周期不良率</h2><img src="chart_station_fails.png"></div>
<div class="chart-section"><h2>📊 失败SN周期数分布</h2><img src="chart_sn_dist.png"></div>
<div class="chart-section"><h2>📋 各测试项周期统计</h2><table>
<tr><th>测试项</th><th>出现周期数</th><th>失败周期数</th><th>周期不良率</th></tr>{test_table}</table></div>
<div class="chart-section"><h2>📋 各站别周期统计</h2><table>
<tr><th>站别</th><th>总周期</th><th>PASS</th><th>FAIL</th><th>不良率</th></tr>{station_table}</table></div>
<div class="chart-section"><h2>🔴 失败SN (按失败周期数排序)</h2><table>
<tr><th>SN</th><th>站别</th><th>失败周期/总周期</th><th>失败率</th><th>失败测试项</th></tr>{sn_table}</table></div>
</div></body></html>"""

def make_html(a, out_dir, out_path, source_info=""):
    _get_font()
    # Test table
    test_rows=[]
    for t in sorted(a["by_test"].keys(), key=lambda t:-a["by_test"][t]["fail_cycles"]):
        d=a["by_test"][t]; total=d["total_cycles"]; fail=d["fail_cycles"]
        rate=fail/total*100 if total else 0
        rc="badge-fail" if rate>5 else "badge-warn" if rate>1 else "badge-pass"
        test_rows.append(f"<tr><td>{t}</td><td>{total}</td>"
                         f"<td>{fail}</td><td><span class='badge {rc}'>{rate:.1f}%</span></td></tr>")
    # Station table
    st_rows=[]
    for s in sorted(a["by_station"].keys()):
        d=a["by_station"][s]; total=d["cycles"]; fail=d["fail"]
        rate=fail/total*100 if total else 0
        rc="badge-fail" if rate>5 else "badge-warn" if rate>1 else "badge-pass"
        st_rows.append(f"<tr><td><strong>{s}</strong></td><td>{total}</td><td>{total-fail}</td>"
                       f"<td>{fail}</td><td><span class='badge {rc}'>{rate:.1f}%</span></td></tr>")
    # SN table
    sn_rows=[]
    for s in a["sn_fail_list"][:80]:
        items=", ".join(s["failed_items"][:4])
        rc="badge-fail" if s["rate"]>20 else "badge-warn"
        sn_rows.append(f"<tr><td><code>{s['sn']}</code></td><td>{s['station']}</td>"
                       f"<td>{s['fail_cycles']}/{s['total_cycles']}</td>"
                       f"<td><span class='badge {rc}'>{s['rate']:.1f}%</span></td>"
                       f"<td style='font-size:11px'>{items}</td></tr>")

    rate_cls = "pass" if a["pass_rate"]>=95 else "warn" if a["pass_rate"]>=85 else "fail"
    sn_cls = "pass" if a["sn_pass_rate"]>=95 else "warn" if a["sn_pass_rate"]>=90 else "fail"

    html = _HTML.format(
        report_time=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        source_info=source_info,
        total_cycles=a["total_cycles"], pass_cycles=a["pass_cycles"],
        fail_cycles=a["fail_cycles"], pass_rate=a["pass_rate"],
        rate_class=rate_cls,
        unique_sns=a["unique_sns"],
        sn_pass=a["sn_pass"], sn_fail=a["sn_fail"],
        sn_pass_rate=a["sn_pass_rate"], sn_class=sn_cls,
        avg_cycles=a["avg_cycles_per_sn"],
        test_table="".join(test_rows), station_table="".join(st_rows),
        sn_table="".join(sn_rows),
    )
    with open(out_path,"w",encoding="utf-8") as f: f.write(html)
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
            self.root=root
            root.title("AT Audio Test 数据分析工具 v3.1 — 周期去重")
            root.geometry("1050x700"); root.minsize(850,550)
            root.configure(bg=STYLE["bg"])
            self.src=tk.StringVar()
            self.out=tk.StringVar(value=os.path.join(os.getcwd(),"at_report_v3"))
            self.status=tk.StringVar(value="就绪 — 请选择数据源")
            self.prog=tk.DoubleVar(value=0)
            self.opt_charts=tk.BooleanVar(value=True); self.opt_html=tk.BooleanVar(value=True)
            self.opt_csv=tk.BooleanVar(value=True); self.opt_open=tk.BooleanVar(value=True)
            self.records=None; self.analysis=None; self.html_path=None; self.running=False
            self._build()
            root.update_idletasks()
            w,h=root.winfo_width(),root.winfo_height()
            sw,sh=root.winfo_screenwidth(),root.winfo_screenheight()
            root.geometry(f"+{(sw-w)//2}+{(sh-h)//2}")

        def _build(self):
            hdr=tk.Frame(self.root,bg=STYLE["accent"],height=52)
            hdr.pack(fill="x"); hdr.pack_propagate(False)
            tk.Label(hdr,text="🔊 AT Audio Test 数据分析工具 v3.1",bg=STYLE["accent"],fg="white",
                     font=("Microsoft YaHei",15,"bold")).pack(side="left",padx=20,pady=10)

            main=tk.Frame(self.root,bg=STYLE["bg"]); main.pack(fill="both",expand=True,padx=15,pady=(10,0))
            left=tk.Frame(main,bg=STYLE["bg"]); left.pack(side="left",fill="y",padx=(0,10))

            # Source
            f1=tk.LabelFrame(left,text="📁 数据源",bg=STYLE["card_bg"],fg=STYLE["fg"],
                             font=("Microsoft YaHei",11,"bold"),padx=12,pady=10,relief="groove",bd=1)
            f1.pack(fill="x",pady=(0,8))
            r1=tk.Frame(f1,bg=STYLE["card_bg"]); r1.pack(fill="x")
            self.se=tk.Entry(r1,textvariable=self.src,font=("Consolas",10),bg=STYLE["input_bg"],relief="solid",bd=1)
            self.se.pack(side="left",fill="x",expand=True,ipady=3)
            tk.Button(r1,text="📂 选择",command=self._pick_src,bg=STYLE["accent"],fg="white",
                      font=("Microsoft YaHei",9),relief="flat",padx=12,pady=3,cursor="hand2").pack(side="left",padx=(6,0))
            r2=tk.Frame(f1,bg=STYLE["card_bg"]); r2.pack(fill="x",pady=(6,0))
            for t,c in [("📦 ZIP","zip"),("📁 文件夹","dir"),("📄 单文件","file")]:
                tk.Button(r2,text=t,command=lambda m=c: self._pick_src(m),bg="#ECEFF1",fg=STYLE["fg"],
                          font=("Microsoft YaHei",9),relief="flat",padx=10,pady=2,cursor="hand2").pack(side="left",padx=(0,6))

            # Output
            f2=tk.LabelFrame(left,text="📂 输出目录",bg=STYLE["card_bg"],fg=STYLE["fg"],
                             font=("Microsoft YaHei",11,"bold"),padx=12,pady=10,relief="groove",bd=1)
            f2.pack(fill="x",pady=(0,8))
            r=tk.Frame(f2,bg=STYLE["card_bg"]); r.pack(fill="x")
            self.oe=tk.Entry(r,textvariable=self.out,font=("Consolas",10),bg=STYLE["input_bg"],relief="solid",bd=1)
            self.oe.pack(side="left",fill="x",expand=True,ipady=3)
            tk.Button(r,text="📂 选择",command=self._pick_out,bg=STYLE["accent"],fg="white",
                      font=("Microsoft YaHei",9),relief="flat",padx=12,pady=3,cursor="hand2").pack(side="left",padx=(6,0))

            # Options
            f3=tk.LabelFrame(left,text="⚙️ 输出",bg=STYLE["card_bg"],fg=STYLE["fg"],
                             font=("Microsoft YaHei",11,"bold"),padx=12,pady=10,relief="groove",bd=1)
            f3.pack(fill="x",pady=(0,8))
            for v,t in [(self.opt_charts,"PNG图表"),(self.opt_html,"HTML报告"),(self.opt_csv,"CSV明细"),(self.opt_open,"自动打开")]:
                tk.Checkbutton(f3,text=t,variable=v,bg=STYLE["card_bg"],font=("Microsoft YaHei",10),
                               activebackground=STYLE["card_bg"],selectcolor=STYLE["card_bg"]).pack(anchor="w",pady=1)

            # Run button
            self.btn=tk.Button(left,text="▶  开始分析",command=self._start,bg=STYLE["accent"],fg="white",
                                font=("Microsoft YaHei",12,"bold"),relief="flat",padx=30,pady=10,cursor="hand2")
            self.btn.pack(fill="x",pady=(5,0))

            # Right panel
            right=tk.Frame(main,bg=STYLE["bg"]); right.pack(side="left",fill="both",expand=True)
            cards=tk.Frame(right,bg=STYLE["bg"]); cards.pack(fill="x",pady=(0,6))
            self.cards={}
            for i,(k,l,c) in enumerate([("cycles","测试周期","#2196F3"),("pass_cycles","PASS周期","#4CAF50"),
                                         ("fail_cycles","FAIL周期","#F44336"),("sn_pass","SN良率","#FF9800")]):
                cd=tk.Frame(cards,bg=STYLE["card_bg"],relief="solid",bd=1,padx=8,pady=8)
                cd.grid(row=0,column=i,padx=3,sticky="nsew"); cards.grid_columnconfigure(i,weight=1)
                self.cards[k]={"v":tk.Label(cd,text="—",bg=STYLE["card_bg"],fg=c,font=("Consolas",20,"bold")),
                               "l":tk.Label(cd,text=l,bg=STYLE["card_bg"],fg=STYLE["text_secondary"],font=("Microsoft YaHei",9))}
                self.cards[k]["v"].pack(); self.cards[k]["l"].pack()

            # Tables
            n=ttk.Notebook(right); n.pack(fill="both",expand=True)
            # Tab1: test items
            t1=tk.Frame(n,bg=STYLE["card_bg"]); n.add(t1,text="📊 测试项(周期)")
            self.test_tree=ttk.Treeview(t1,columns=("test","cycles","fail","rate"),show="headings",height=14)
            self.test_tree.heading("test",text="测试项"); self.test_tree.column("test",width=180)
            self.test_tree.heading("cycles",text="周期数"); self.test_tree.column("cycles",width=70,anchor="center")
            self.test_tree.heading("fail",text="失败周期"); self.test_tree.column("fail",width=70,anchor="center")
            self.test_tree.heading("rate",text="不良率"); self.test_tree.column("rate",width=70,anchor="center")
            self.test_tree.pack(fill="both",expand=True)
            for tag,color in [("high","#F44336"),("mid","#FF9800"),("low","#4CAF50")]:
                self.test_tree.tag_configure(tag,foreground=color)

            # Tab2: SN failures
            t2=tk.Frame(n,bg=STYLE["card_bg"]); n.add(t2,text="🔴 失败SN")
            self.sn_tree=ttk.Treeview(t2,columns=("sn","station","fail_cycles","failed_items"),show="headings",height=14)
            self.sn_tree.heading("sn",text="SN"); self.sn_tree.column("sn",width=170)
            self.sn_tree.heading("station",text="站别"); self.sn_tree.column("station",width=60,anchor="center")
            self.sn_tree.heading("fail_cycles",text="失败/总周期"); self.sn_tree.column("fail_cycles",width=90,anchor="center")
            self.sn_tree.heading("failed_items",text="失败测试项"); self.sn_tree.column("failed_items",width=250)
            self.sn_tree.pack(fill="both",expand=True)
            self.sn_tree.tag_configure("critical",foreground="#F44336",font=("Consolas",10,"bold"))

            # Bottom buttons
            btns=tk.Frame(right,bg=STYLE["bg"]); btns.pack(fill="x",pady=(6,0))
            self.br=tk.Button(btns,text="📄 打开报告",command=self._open_r,bg=STYLE["success"],fg="white",
                               font=("Microsoft YaHei",10),relief="flat",padx=16,pady=5,state="disabled")
            self.br.pack(side="left",padx=(0,6))
            self.bf=tk.Button(btns,text="📂 输出目录",command=self._open_f,bg="#ECEFF1",fg=STYLE["fg"],
                               font=("Microsoft YaHei",10),relief="flat",padx=16,pady=5,state="disabled")
            self.bf.pack(side="left",padx=(0,6))

            # Status bar
            sbar=tk.Frame(self.root,bg=STYLE["card_bg"],height=30); sbar.pack(fill="x",side="bottom",pady=(8,0))
            sbar.pack_propagate(False)
            tk.Label(sbar,textvariable=self.status,bg=STYLE["card_bg"],fg=STYLE["text_secondary"],
                     font=("Microsoft YaHei",9),anchor="w").pack(side="left",fill="x",padx=12,pady=4)
            self.pb=ttk.Progressbar(sbar,variable=self.prog,mode="determinate",length=180)
            self.pb.pack(side="right",padx=12,pady=4)

        def _pick_src(self, mode=None):
            if mode=="zip": p=filedialog.askopenfilename(title="选ZIP",filetypes=[("ZIP","*.zip")])
            elif mode=="dir": p=filedialog.askdirectory(title="选文件夹")
            elif mode=="file": p=filedialog.askopenfilename(title="选文件",filetypes=[("数据","*.xls;*.txt")])
            else: p=filedialog.askopenfilename(title="选数据源",filetypes=[("支持","*.zip;*.xls;*.txt")]) or filedialog.askdirectory(title="或选文件夹")
            if p: self.src.set(p); self.status.set(f"已选: {os.path.basename(p)}")

        def _pick_out(self):
            p=filedialog.askdirectory(title="输出目录",initialdir=self.out.get())
            if p: self.out.set(p)

        def _start(self):
            s=self.src.get().strip()
            if not s: messagebox.showwarning("提示","请先选择数据源"); return
            if not os.path.exists(s): messagebox.showerror("错误","路径不存在"); return
            if self.running: return
            self.running=True; self.btn.configure(text="⏳ 分析中...",state="disabled",bg=STYLE["text_secondary"])
            self.prog.set(5); self.status.set("解析中..."); self._clear()
            threading.Thread(target=self._run,args=(s,),daemon=True).start()

        def _run(self,s):
            try:
                self.root.after(0,lambda:self.prog.set(10))
                recs,skipped=parse_source(s,cb=lambda ph,dt:self.root.after(0,lambda:self.status.set(ph)))
                if not recs: raise ValueError("无有效记录")
                self.records=recs
                self.root.after(0,lambda:self.prog.set(40))
                self.root.after(0,lambda:self.status.set(f"解析: {len(recs)}行, 按周期去重分析中..."))
                a=analyze(recs); self.analysis=a
                self.root.after(0,lambda:self.prog.set(60))
                self.root.after(0,lambda:self._show(a))
                od=self.out.get(); os.makedirs(od,exist_ok=True)
                if self.opt_charts.get():
                    self.root.after(0,lambda:self.prog.set(70))
                    self.root.after(0,lambda:self.status.set("图表中..."))
                    _get_font()
                    _chart_test_fail_bars(a,os.path.join(od,"chart_test_fails.png"))
                    _chart_station_fail_bars(a,os.path.join(od,"chart_station_fails.png"))
                    _chart_sn_fail_dist(a,os.path.join(od,"chart_sn_dist.png"))
                if self.opt_csv.get():
                    self.root.after(0,lambda:self.prog.set(85))
                    self._write_csv(os.path.join(od,"records_detail.csv"))
                if self.opt_html.get():
                    self.root.after(0,lambda:self.prog.set(95))
                    self.root.after(0,lambda:self.status.set("HTML报告中..."))
                    self.html_path=make_html(a,od,os.path.join(od,"report.html"),os.path.basename(s))
                self.root.after(0,lambda:self.prog.set(100))
                self.root.after(0,lambda:self.status.set(f"✅ 完成！{a['total_cycles']}周期 PASS={a['pass_rate']:.1f}% SN良率={a['sn_pass_rate']:.1f}%"))
                self.root.after(0,self._done)
            except Exception as e:
                import traceback; traceback.print_exc()
                self.root.after(0,lambda:self.status.set(f"❌ {e}"))
                self.root.after(0,lambda:self.prog.set(0))
                self.root.after(0,self._err)

        def _write_csv(self,p):
            if not self.records: return
            ks=["sn","test_name","result","value","upper","lower","station","time"]
            with open(p,"w",newline="",encoding="utf-8-sig") as f:
                w=csv.DictWriter(f,fieldnames=ks,extrasaction="ignore"); w.writeheader()
                for r in self.records:
                    row={k:r.get(k,"") for k in ks}
                    if row.get("time"): row["time"]=str(row["time"])
                    w.writerow(row)

        def _show(self,a):
            self.cards["cycles"]["v"].configure(text=str(a["total_cycles"]))
            self.cards["pass_cycles"]["v"].configure(text=str(a["pass_cycles"]))
            self.cards["fail_cycles"]["v"].configure(text=str(a["fail_cycles"]))
            self.cards["sn_pass"]["v"].configure(text=f"{a['sn_pass_rate']:.1f}%")
            for t in [self.test_tree,self.sn_tree]:
                for i in t.get_children(): t.delete(i)
            # Test items (cycle-level)
            for tn in sorted(a["by_test"],key=lambda t:-a["by_test"][t]["fail_cycles"]):
                d=a["by_test"][tn]; total=d["total_cycles"]; fail=d["fail_cycles"]
                rate=fail/total*100 if total else 0
                tag="high" if rate>5 else "mid" if rate>1 else "low"
                self.test_tree.insert("","end",values=(tn,total,fail,f"{rate:.1f}%"),tags=(tag,))
            # SN failures
            for s in a["sn_fail_list"][:80]:
                tag="critical" if s["rate"]>20 else ""
                self.sn_tree.insert("","end",values=(
                    s["sn"],s["station"],
                    f"{s['fail_cycles']}/{s['total_cycles']}",
                    ", ".join(s["failed_items"][:4]),
                ),tags=(tag,) if tag else())

        def _clear(self):
            for k in self.cards: self.cards[k]["v"].configure(text="—")
            for t in [self.test_tree,self.sn_tree]:
                for i in t.get_children(): t.delete(i)
            self.html_path=None
            self.br.configure(state="disabled"); self.bf.configure(state="disabled")

        def _done(self):
            self.running=False; self.btn.configure(text="▶  重新分析",state="normal",bg=STYLE["accent"])
            self.br.configure(state="normal"); self.bf.configure(state="normal")
            if self.opt_open.get() and self.html_path: self._open_r()

        def _err(self):
            self.running=False; self.btn.configure(text="▶  开始分析",state="normal",bg=STYLE["accent"])

        def _open_r(self):
            if self.html_path and os.path.exists(self.html_path): _open_with_default(self.html_path)
            else: messagebox.showinfo("提示","报告不存在")

        def _open_f(self):
            o=self.out.get()
            if os.path.exists(o): _open_with_default(o)
            else: messagebox.showinfo("提示","目录不存在")

    root=tk.Tk()
    try:
        from ctypes import windll; windll.shcore.SetProcessDpiAwareness(1)
    except: pass
    App(root); root.mainloop()

def main(): _run_gui()
if __name__=="__main__": main()

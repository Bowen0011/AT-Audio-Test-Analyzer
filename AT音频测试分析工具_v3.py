#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AT Audio Test 数据分析工具 v3.2
===============================
每SN进行40项测试(SPA×4 + 标压/EQ/气密/MIC×36)，每项可能多次重测。
取每项最终结果判定SN良率。

输出: 直观汇总 + 分站别统计 + 失败原因 + HTML报告 + 图表
"""

import os, sys, csv, zipfile, threading, datetime, warnings, webbrowser, subprocess
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
    bars = ax.bar(sts, yields, color=[C["pass"] if y>=95 else C["warn"] if y>=90 else C["fail"] for y in yields])
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

def chart_channel_failure(a, out):
    _get_font()
    cf = a["channel_failure"]
    if not cf: return
    items = sorted(cf.items(), key=lambda x: -x[1])[:10]
    labels = [k[:20] for k,_ in items]
    counts = [v for _,v in items]
    fig, ax = plt.subplots(figsize=(8,5))
    ax.bar(range(len(items)), counts, color=C["fail"])
    ax.set_xticks(range(len(items)))
    ax.set_xticklabels(labels, rotation=30, ha="right", fontsize=9)
    ax.set_ylabel("失败次数", fontsize=12)
    ax.set_title("失败通道分布", fontsize=14, fontweight="bold")
    ax.yaxis.set_major_locator(mticker.MaxNLocator(integer=True))
    for i, c in enumerate(counts):
        ax.text(i, c+0.2, str(c), ha="center", fontsize=9, fontweight="bold")
    plt.tight_layout(); fig.savefig(out, dpi=150, bbox_inches="tight"); plt.close(fig)

# ═══════════════════════════════════════════════
# HTML
# ═══════════════════════════════════════════════
_HTML = """<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>AT Audio Test 分析报告</title><style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:-apple-system,"Microsoft YaHei",sans-serif;background:#f5f7fa;color:#333;padding:20px}}
.container{{max-width:1200px;margin:0 auto}}
h1{{text-align:center;color:#1a237e;margin-bottom:6px;font-size:26px}}
.subtitle{{text-align:center;color:#78909c;margin-bottom:30px;font-size:13px}}
.summary{{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:15px;margin-bottom:30px}}
.card{{background:white;border-radius:10px;padding:18px;text-align:center;box-shadow:0 2px 8px rgba(0,0,0,0.08)}}
.card .v{{font-size:34px;font-weight:700}}.card .l{{font-size:12px;color:#78909c;margin-top:4px}}
.card.g .v{{color:#4CAF50}}.card.r .v{{color:#F44336}}.card.b .v{{color:#2196F3}}.card.o .v{{color:#FF9800}}
.sec{{background:white;border-radius:10px;padding:20px;margin-bottom:18px;box-shadow:0 2px 8px rgba(0,0,0,0.08)}}
.sec h2{{color:#37474f;margin-bottom:15px;font-size:17px;border-left:4px solid #2196F3;padding-left:12px}}
.sec img{{max-width:100%;height:auto;display:block;margin:0 auto}}
table{{width:100%;border-collapse:collapse;font-size:13px}}
th{{background:#263238;color:white;padding:9px 8px;text-align:left;font-weight:600}}
td{{padding:7px 8px;border-bottom:1px solid #eceff1}}tr:hover td{{background:#f5f5f5}}
.badge{{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600;color:white}}
.bg{{background:#4CAF50}}.br{{background:#F44336}}.bo{{background:#FF9800}}
</style></head><body><div class="container">
<h1>🔊 AT Audio Test 分析报告</h1>
<p class="subtitle">{report_time} | 数据: {source_info} | 每SN取各项最终结果</p>
<div class="summary">
<div class="card b"><div class="v">{total_sn}</div><div class="l">测试总数(台)</div></div>
<div class="card g"><div class="v">{pass_sn}</div><div class="l">PASS</div></div>
<div class="card r"><div class="v">{fail_sn}</div><div class="l">FAIL</div></div>
<div class="card {yc}"><div class="v">{yield_rate:.1f}%</div><div class="l">良率</div></div>
</div>
<div class="sec"><h2>📊 各站别良率</h2><img src="chart_station_yield.png"></div>
<div class="sec"><h2>📋 各站别统计</h2><table>
<tr><th>站别</th><th>测试数</th><th>PASS</th><th>FAIL</th><th>良率</th></tr>{station_table}</table></div>
<div class="sec"><h2>📊 失败原因</h2><img src="chart_failure_reasons.png"></div>
<div class="sec"><h2>📊 失败通道</h2><img src="chart_channel_failure.png"></div>
<div class="sec"><h2>🔴 失败SN明细</h2><table>
<tr><th>SN</th><th>站别</th><th>失败项数</th><th>失败测试项</th></tr>{sn_table}</table></div>
</div></body></html>"""

def make_html(a, out_dir, out_path, source_info=""):
    # Station table
    st_rows = []
    for s in sorted(a["station_stats"].keys()):
        d = a["station_stats"][s]
        r = d["pass"]/d["total"]*100 if d["total"] else 0
        cls = "bg" if r>=95 else "bo" if r>=90 else "br"
        st_rows.append(f"<tr><td><strong>{s}</strong></td><td>{d['total']}</td>"
                       f"<td>{d['pass']}</td><td>{d['fail']}</td><td><span class='badge {cls}'>{r:.1f}%</span></td></tr>")
    # SN table
    sn_rows = []
    for s in a["fail_list"]:
        items = ", ".join(s["failed"][:5])
        sn_rows.append(f"<tr><td><code>{s['sn']}</code></td><td>{s['station']}</td>"
                       f"<td>{len(s['failed'])}/{s['total']}</td><td style='font-size:11px'>{items}</td></tr>")

    yc = "g" if a["yield_rate"]>=95 else "o" if a["yield_rate"]>=90 else "r"
    html = _HTML.format(
        report_time=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        source_info=source_info,
        total_sn=a["total_sn"], pass_sn=a["pass_sn"], fail_sn=a["fail_sn"],
        yield_rate=a["yield_rate"], yc=yc,
        station_table="".join(st_rows), sn_table="".join(sn_rows),
    )
    with open(out_path, "w", encoding="utf-8") as f: f.write(html)
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
            root.title("AT Audio Test 分析工具 v3.2")
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
            tk.Label(hdr, text="🔊 AT Audio Test 分析工具 v3.2", bg=STYLE["accent"], fg="white",
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
                ("fail","FAIL","#F44336"),("yield","良率","#FF9800")]):
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
                self.root.after(0, lambda: self.prog.set(60))
                self.root.after(0, lambda: self._show(a))
                od = self.out.get(); os.makedirs(od, exist_ok=True)

                if self.opt_charts.get():
                    self.root.after(0, lambda: self.prog.set(70))
                    _get_font()
                    chart_station_yield(a, os.path.join(od, "chart_station_yield.png"))
                    chart_failure_reasons(a, os.path.join(od, "chart_failure_reasons.png"))
                    chart_channel_failure(a, os.path.join(od, "chart_channel_failure.png"))
                if self.opt_csv.get():
                    self.root.after(0, lambda: self.prog.set(85))
                    self._write_csv(os.path.join(od, "detail.csv"))
                if self.opt_html.get():
                    self.root.after(0, lambda: self.prog.set(95))
                    self.html_path = make_html(a, od, os.path.join(od, "report.html"), os.path.basename(s))
                self.root.after(0, lambda: self.prog.set(100))
                msg = f"✅ 完成！总数{a['total_sn']}台 PASS={a['pass_sn']} FAIL={a['fail_sn']} 良率{a['yield_rate']:.1f}%"
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
                tag = "good" if r>=95 else "warn" if r>=90 else "bad"
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

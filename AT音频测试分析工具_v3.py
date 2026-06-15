#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AT Audio Test 数据分析工具 v3 — 行级日志格式
=============================================
解析 AT 站别导出的 TC661_data.xls (TSV行级格式)。
每行 = 一条测试项，自带测量值+上限+下限，可直接判等。

输出: HTML报告 + PNG图表 + CSV明细
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
            with open(fp,"rb") as f:
                d=f.read(4096)
            if "SN" in d.decode(e,errors="replace"): return e
        except: continue
    return "gbk"

# ═══════════════════════════════════════════════
# 解析器 — 行级 TSV
# ═══════════════════════════════════════════════
def parse_file(filepath):
    """解析 TC661_data.xls，每行一条测试项
    返回 [{sn, test_name, channel, value, upper, lower, pass, station}]
    """
    enc = _detect_enc(filepath)
    # Station from path: .../ATxx/TestData/.../TC661_data.xls → ATxx
    p = Path(filepath)
    station = p.parent.parent.parent.parent.name  # ATxx
    if not station.startswith("AT"):
        station = p.parent.parent.parent.name  # fallback

    records = []
    with open(filepath,"rb") as f:
        raw = f.read()
    text = raw.decode(enc, errors="replace")
    for line in text.split("\n"):
        line = line.strip()
        if not line or line.startswith("Time"): continue
        flds = line.split("\t")
        if len(flds) < 9: continue
        val_str = flds[7].strip()
        if not val_str: continue  # skip info rows
        sn = flds[1].strip().strip("'\"")
        test_name = flds[2].strip()
        channel = flds[6].strip()
        try:
            v = float(val_str)
            upper = float(flds[8].strip()) if flds[8].strip() else None
            lower = float(flds[9].strip()) if len(flds)>9 and flds[9].strip() else None
        except: continue
        lo = min(upper,lower) if upper is not None and lower is not None else None
        hi = max(upper,lower) if upper is not None and lower is not None else None
        passed = True
        if lo is not None and hi is not None:
            passed = lo <= v <= hi
        records.append({
            "sn":sn,"test_name":test_name,"channel":channel,
            "value":v,"upper":upper,"lower":lower,
            "pass":passed,"station":station,"file":filepath,
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
# 分析器
# ═══════════════════════════════════════════════
def analyze(records):
    if not records: return {"error":"无记录"}
    total = len(records)
    failed = [r for r in records if not r["pass"]]
    passed = total - len(failed)

    # Per-station
    by_station = defaultdict(lambda: {"total":0,"fail":0,"pass":0})
    for r in records:
        s = r["station"]
        by_station[s]["total"] += 1
        if r["pass"]: by_station[s]["pass"] += 1
        else: by_station[s]["fail"] += 1

    # Per test_name
    by_test = defaultdict(lambda: {"total":0,"fail":0})
    for r in records:
        t = r["test_name"]
        by_test[t]["total"] += 1
        if not r["pass"]: by_test[t]["fail"] += 1

    # Per SN
    by_sn = defaultdict(lambda: {"total":0,"fail":0,"failed_tests":[]})
    for r in records:
        sn = r["sn"]
        by_sn[sn]["total"] += 1
        if not r["pass"]:
            by_sn[sn]["fail"] += 1
            by_sn[sn]["failed_tests"].append(r["test_name"])

    sn_fail_list = sorted(
        [{"sn":sn,"fail":d["fail"],"total":d["total"],
          "rate":d["fail"]/d["total"]*100,"tests":d["failed_tests"][:5]}
         for sn,d in by_sn.items() if d["fail"]>0],
        key=lambda x:-x["fail"])

    # Per channel
    by_ch = defaultdict(lambda: {"total":0,"fail":0})
    for r in records:
        ch = r["channel"]
        by_ch[ch]["total"] += 1
        if not r["pass"]: by_ch[ch]["fail"] += 1

    return {
        "total":total,"pass":passed,"fail":len(failed),
        "pass_rate":passed/total*100 if total else 0,
        "fail_rate":len(failed)/total*100 if total else 0,
        "by_station":dict(by_station),
        "by_test":dict(by_test),
        "by_sn":dict(by_sn),
        "sn_fail_list":sn_fail_list,
        "by_channel":dict(by_ch),
        "unique_sns":len(by_sn),
        "unique_tests":len(by_test),
    }

# ═══════════════════════════════════════════════
# 图表
# ═══════════════════════════════════════════════
def _chart_test_fail_bars(a, out):
    _get_font()
    tests = sorted(a["by_test"].keys(), key=lambda t: -a["by_test"][t]["fail"])
    labels = [t[:20] for t in tests]
    fails = [a["by_test"][t]["fail"] for t in tests]
    totals = [a["by_test"][t]["total"] for t in tests]
    rates = [f/t*100 if t else 0 for f,t in zip(fails,totals)]

    fig,ax = plt.subplots(figsize=(max(10,len(tests)*0.6),6))
    bars = ax.bar(range(len(tests)), fails, color=[C["fail"] if r>30 else C["warn"] if r>10 else C["bar_pass"] for r in rates])
    ax.set_xticks(range(len(tests))); ax.set_xticklabels(labels, rotation=45, ha="right", fontsize=9)
    ax.set_ylabel("失败次数",fontsize=12)
    ax.set_title("各测试项失败次数",fontsize=14,fontweight="bold")
    for i,(f,r) in enumerate(zip(fails,rates)):
        ax.text(i,f+2,f"{r:.1f}%",ha="center",va="bottom",fontsize=8,fontweight="bold",color=C["fail"] if r>30 else "#555")
    ax.yaxis.set_major_locator(mticker.MaxNLocator(integer=True))
    plt.tight_layout(); fig.savefig(out,dpi=150,bbox_inches="tight"); plt.close(fig)

def _chart_station_fail_bars(a, out):
    _get_font()
    sts = sorted(a["by_station"].keys())
    fails = [a["by_station"][s]["fail"] for s in sts]
    totals = [a["by_station"][s]["total"] for s in sts]
    rates = [f/t*100 if t else 0 for f,t in zip(fails,totals)]
    fig,ax = plt.subplots(figsize=(max(8,len(sts)*1.5),5))
    bars = ax.bar(sts, rates, color=[C["fail"] if r>30 else C["warn"] if r>10 else C["bar_pass"] for r in rates])
    ax.set_ylabel("不良率 (%)",fontsize=12)
    ax.set_title("各站别测试项不良率",fontsize=14,fontweight="bold")
    for bar,r in zip(bars,rates):
        ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+1, f"{r:.1f}%", ha="center",fontsize=10,fontweight="bold")
    ax.set_ylim(0,max(rates)*1.2+5)
    plt.tight_layout(); fig.savefig(out,dpi=150,bbox_inches="tight"); plt.close(fig)

def _chart_channel_fail_pie(a, out):
    _get_font()
    chs = {ch:d["fail"] for ch,d in a["by_channel"].items() if d["fail"]>0}
    if not chs: return
    fig,ax = plt.subplots(figsize=(6,5))
    wedges,_,_ = ax.pie(list(chs.values()), labels=None, autopct="%1.1f%%",
                        colors=[C["fail"],C["warn"],C["accent"]], startangle=90)
    ax.legend(wedges,[f"{k} ({v})" for k,v in chs.items()],loc="lower center",ncol=2,frameon=False)
    ax.set_title("失败分布(按通道)",fontsize=14,fontweight="bold")
    plt.tight_layout(); fig.savefig(out,dpi=150,bbox_inches="tight"); plt.close(fig)

def _chart_sn_fail_dist(a, out):
    _get_font()
    sl = a["sn_fail_list"]
    if not sl: return
    fail_counts = [s["fail"] for s in sl]
    dist = Counter(fail_counts)
    cats = sorted(dist.keys())
    fig,ax = plt.subplots(figsize=(8,5))
    ax.bar([f"{c}项失败" for c in cats], [dist[c] for c in cats],
           color=[C["fail"] if c>10 else C["warn"] if c>5 else C["bar_pass"] for c in cats])
    ax.set_ylabel("SN数量",fontsize=12)
    ax.set_title(f"SN失败项数分布（共{len(sl)}个SN有失败）",fontsize=14,fontweight="bold")
    ax.yaxis.set_major_locator(mticker.MaxNLocator(integer=True))
    plt.tight_layout(); fig.savefig(out,dpi=150,bbox_inches="tight"); plt.close(fig)

# ═══════════════════════════════════════════════
# HTML 报告
# ═══════════════════════════════════════════════
_HTML = """<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>AT Audio Test 数据分析报告 v3</title><style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:-apple-system,"Microsoft YaHei","PingFang SC",sans-serif;background:#f5f7fa;color:#333;padding:20px}}
.container{{max-width:1200px;margin:0 auto}}
h1{{text-align:center;color:#1a237e;margin-bottom:6px;font-size:28px}}
.subtitle{{text-align:center;color:#78909c;margin-bottom:30px;font-size:14px}}
.summary-cards{{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:15px;margin-bottom:30px}}
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
</style></head><body><div class="container">
<h1>🔊 AT Audio Test 数据分析报告 v3</h1>
<p class="subtitle">生成时间: {report_time}  |  数据来源: {source_info}</p>
<div class="summary-cards">
<div class="card info"><div class="value">{total}</div><div class="label">总测试项</div></div>
<div class="card pass"><div class="value">{pass_count}</div><div class="label">PASS</div></div>
<div class="card fail"><div class="value">{fail_count}</div><div class="label">FAIL</div></div>
<div class="card {rate_class}"><div class="value">{pass_rate:.1f}%</div><div class="label">合格率</div></div>
<div class="card info"><div class="value">{unique_sns}</div><div class="label">测试SN数</div></div>
<div class="card info"><div class="value">{unique_tests}</div><div class="label">测试项目数</div></div>
</div>
<div class="chart-section"><h2>📊 各测试项失败次数</h2><img src="chart_test_fails.png"></div>
<div class="chart-section"><h2>📊 各站别不良率</h2><img src="chart_station_fails.png"></div>
<div class="chart-section"><h2>📊 失败分布(按通道)</h2><img src="chart_channel_pie.png"></div>
<div class="chart-section"><h2>📊 SN失败项数分布</h2><img src="chart_sn_dist.png"></div>
<div class="chart-section"><h2>📋 各测试项统计</h2><table>
<tr><th>测试项</th><th>总数</th><th>PASS</th><th>FAIL</th><th>不良率</th></tr>{test_table}</table></div>
<div class="chart-section"><h2>📋 各站别统计</h2><table>
<tr><th>站别</th><th>总测试项</th><th>PASS</th><th>FAIL</th><th>不良率</th></tr>{station_table}</table></div>
<div class="chart-section"><h2>🔴 失败SN Top 50</h2><table>
<tr><th>SN</th><th>失败项数</th><th>失败率</th><th>典型失败测试</th></tr>{sn_table}</table></div>
</div></body></html>"""

def make_html(a, out_dir, out_path, source_info=""):
    _get_font()
    # Test table
    test_rows=[]
    for t in sorted(a["by_test"].keys(), key=lambda t:-a["by_test"][t]["fail"]):
        d=a["by_test"][t]; total=d["total"]; fail=d["fail"]
        rate=fail/total*100 if total else 0
        rc="badge-fail" if rate>30 else "badge-warn" if rate>10 else "badge-pass"
        test_rows.append(f"<tr><td>{t}</td><td>{total}</td><td>{total-fail}</td>"
                         f"<td>{fail}</td><td><span class='badge {rc}'>{rate:.1f}%</span></td></tr>")
    # Station table
    st_rows=[]
    for s in sorted(a["by_station"].keys()):
        d=a["by_station"][s]; total=d["total"]; fail=d["fail"]
        rate=fail/total*100 if total else 0
        rc="badge-fail" if rate>30 else "badge-warn" if rate>10 else "badge-pass"
        st_rows.append(f"<tr><td><strong>{s}</strong></td><td>{total}</td><td>{total-fail}</td>"
                       f"<td>{fail}</td><td><span class='badge {rc}'>{rate:.1f}%</span></td></tr>")
    # SN table
    sn_rows=[]
    for s in a["sn_fail_list"][:50]:
        tests=", ".join(s["tests"][:4]) if s["tests"] else "—"
        rc="badge-fail" if s["rate"]>50 else "badge-warn"
        sn_rows.append(f"<tr><td><code>{s['sn']}</code></td><td>{s['fail']}/{s['total']}</td>"
                       f"<td><span class='badge {rc}'>{s['rate']:.1f}%</span></td><td style='font-size:11px'>{tests}</td></tr>")

    html = _HTML.format(
        report_time=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        source_info=source_info,
        total=a["total"], pass_count=a["pass"], fail_count=a["fail"],
        pass_rate=a["pass_rate"],
        rate_class="pass" if a["pass_rate"]>=95 else "warn" if a["pass_rate"]>=85 else "fail",
        unique_sns=a["unique_sns"], unique_tests=a["unique_tests"],
        test_table="".join(test_rows), station_table="".join(st_rows), sn_table="".join(sn_rows),
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
            root.title("AT Audio Test 数据分析工具 v3")
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
            tk=self.__class__.__dict__; ttk=ttk
            hdr=tk.Frame(self.root,bg=STYLE["accent"],height=52)
            hdr.pack(fill="x"); hdr.pack_propagate(False)
            tk.Label(hdr,text="🔊 AT Audio Test 数据分析工具 v3",bg=STYLE["accent"],fg="white",
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
            for i,(k,l,c) in enumerate([("total","总测试项","#2196F3"),("pass","PASS","#4CAF50"),
                                         ("fail","FAIL","#F44336"),("rate","合格率","#FF9800")]):
                cd=tk.Frame(cards,bg=STYLE["card_bg"],relief="solid",bd=1,padx=8,pady=8)
                cd.grid(row=0,column=i,padx=3,sticky="nsew"); cards.grid_columnconfigure(i,weight=1)
                self.cards[k]={"v":tk.Label(cd,text="—",bg=STYLE["card_bg"],fg=c,font=("Consolas",20,"bold")),
                               "l":tk.Label(cd,text=l,bg=STYLE["card_bg"],fg=STYLE["text_secondary"],font=("Microsoft YaHei",9))}
                self.cards[k]["v"].pack(); self.cards[k]["l"].pack()

            # Tables
            n=tk.Notebook(right); n.pack(fill="both",expand=True)
            # Tab1: test items
            t1=tk.Frame(n,bg=STYLE["card_bg"]); n.add(t1,text="📊 测试项统计")
            self.test_tree=ttk.Treeview(t1,columns=("test","total","fail","rate"),show="headings",height=14)
            self.test_tree.heading("test",text="测试项"); self.test_tree.column("test",width=180)
            self.test_tree.heading("total",text="总数"); self.test_tree.column("total",width=60,anchor="center")
            self.test_tree.heading("fail",text="FAIL"); self.test_tree.column("fail",width=60,anchor="center")
            self.test_tree.heading("rate",text="不良率"); self.test_tree.column("rate",width=70,anchor="center")
            self.test_tree.pack(fill="both",expand=True)
            for tag,color in [("high","#F44336"),("mid","#FF9800"),("low","#4CAF50")]:
                self.test_tree.tag_configure(tag,foreground=color)

            # Tab2: SN failures
            t2=tk.Frame(n,bg=STYLE["card_bg"]); n.add(t2,text="🔴 失败SN")
            self.sn_tree=ttk.Treeview(t2,columns=("sn","fail","rate","tests"),show="headings",height=14)
            self.sn_tree.heading("sn",text="SN"); self.sn_tree.column("sn",width=170)
            self.sn_tree.heading("fail",text="失败项"); self.sn_tree.column("fail",width=70,anchor="center")
            self.sn_tree.heading("rate",text="失败率"); self.sn_tree.column("rate",width=70,anchor="center")
            self.sn_tree.heading("tests",text="典型失败测试"); self.sn_tree.column("tests",width=250)
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
            threading.Thread(target=self._run,s=(s,),daemon=True).start()

        def _run(self,s):
            try:
                self.root.after(0,lambda:self.prog.set(10))
                recs,skipped=parse_source(s,cb=lambda ph,dt:self.root.after(0,lambda:self.status.set(ph)))
                if not recs: raise ValueError("无有效记录")
                self.records=recs
                self.root.after(0,lambda:self.prog.set(40))
                self.root.after(0,lambda:self.status.set(f"解析: {len(recs)}条, 分析中..."))
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
                    _chart_channel_fail_pie(a,os.path.join(od,"chart_channel_pie.png"))
                    _chart_sn_fail_dist(a,os.path.join(od,"chart_sn_dist.png"))
                if self.opt_csv.get():
                    self.root.after(0,lambda:self.prog.set(85))
                    self._write_csv(os.path.join(od,"records_detail.csv"))
                if self.opt_html.get():
                    self.root.after(0,lambda:self.prog.set(95))
                    self.root.after(0,lambda:self.status.set("HTML报告中..."))
                    self.html_path=make_html(a,od,os.path.join(od,"report.html"),os.path.basename(s))
                self.root.after(0,lambda:self.prog.set(100))
                self.root.after(0,lambda:self.status.set("✅ 完成！"))
                self.root.after(0,self._done)
            except Exception as e:
                import traceback; traceback.print_exc()
                self.root.after(0,lambda:self.status.set(f"❌ {e}"))
                self.root.after(0,lambda:self.prog.set(0))
                self.root.after(0,self._err)

        def _write_csv(self,p):
            if not self.records: return
            ks=["sn","test_name","channel","value","upper","lower","pass","station"]
            with open(p,"w",newline="",encoding="utf-8-sig") as f:
                w=csv.DictWriter(f,fieldnames=ks,extrasaction="ignore"); w.writeheader()
                for r in self.records: w.writerow({k:r.get(k,"") for k in ks})

        def _show(self,a):
            self.cards["total"]["v"].configure(text=str(a["total"]))
            self.cards["pass"]["v"].configure(text=str(a["pass"]))
            self.cards["fail"]["v"].configure(text=str(a["fail"]))
            self.cards["rate"]["v"].configure(text=f"{a['pass_rate']:.1f}%")
            for t in [self.test_tree,self.sn_tree]:
                for i in t.get_children(): t.delete(i)
            for tn in sorted(a["by_test"],key=lambda t:-a["by_test"][t]["fail"]):
                d=a["by_test"][tn]; total=d["total"]; fail=d["fail"]
                rate=fail/total*100 if total else 0
                tag="high" if rate>30 else "mid" if rate>10 else "low"
                self.test_tree.insert("","end",values=(tn,total,fail,f"{rate:.1f}%"),tags=(tag,))
            for s in a["sn_fail_list"][:80]:
                tag="critical" if s["rate"]>50 else ""
                self.sn_tree.insert("","end",values=(s["sn"],f"{s['fail']}/{s['total']}",f"{s['rate']:.1f}%",
                                    ", ".join(s["tests"][:4])),tags=(tag,) if tag else())

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

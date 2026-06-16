#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AT Audio Test 数据分析工具
===========================
独立单文件 GUI 工具。双击运行即可。
- 选择 .zip / 文件夹 / 单个 .xls .txt 数据源
- UI 内实时预览 Summary + 站别统计 + 失败SN
- 一键生成 HTML 报告 + PNG 图表 + CSV 明细
- 后台线程解析，界面不冻结

依赖: Python 3.8+, matplotlib
工作电脑安装: pip install matplotlib
"""

import os, sys, re, csv, json, zipfile, threading, datetime, warnings, webbrowser, subprocess
from pathlib import Path
from collections import defaultdict, Counter
from tempfile import TemporaryDirectory

# ── matplotlib ──
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from matplotlib.font_manager import FontProperties

warnings.filterwarnings("ignore", category=UserWarning)

# ═══════════════════════════════════════════════════════════════════
# 常量
# ═══════════════════════════════════════════════════════════════════

_CJK_FONTS = [
    "SimHei", "Microsoft YaHei", "PingFang SC", "Noto Sans CJK SC",
    "WenQuanYi Micro Hei", "WenQuanYi Zen Hei", "Source Han Sans SC",
    "STHeiti", "Heiti SC", "SimSun", "FangSong", "KaiTi",
]
_ENCODINGS = ["gbk", "gb2312", "gb18030", "utf-8", "utf-16", "latin-1"]
COLORS = {"pass": "#4CAF50", "fail": "#F44336", "unknown": "#9E9E9E",
          "bar_pass": "#81C784", "bar_fail": "#E57373", "highlight": "#FF9800"}

_chinese_font = None

def _get_font():
    global _chinese_font
    if _chinese_font is not None:
        return _chinese_font
    available = set(f.name for f in matplotlib.font_manager.fontManager.ttflist)
    for name in _CJK_FONTS:
        if name in available:
            _chinese_font = FontProperties(family=name)
            break
    if _chinese_font is None:
        _chinese_font = FontProperties()
    plt.rcParams["font.family"] = _chinese_font.get_name()
    plt.rcParams["axes.unicode_minus"] = False
    return _chinese_font


# ═══════════════════════════════════════════════════════════════════
# 核心解析器
# ═══════════════════════════════════════════════════════════════════

def _detect_encoding(filepath):
    for enc in _ENCODINGS:
        try:
            with open(filepath, "rb") as f:
                data = f.read(4096)
            text = data.decode(enc, errors="replace")
            if "SN" in text or "BLKVUN" in text:
                return enc
        except Exception:
            continue
    return "gbk"


def iter_records(filepath):
    """逐条产出测试记录 dict"""
    encoding = _detect_encoding(filepath)
    station_name = Path(filepath).parent.parent.name
    # 如果取到的是临时目录名（如 tmpXXXXXX），退一级
    if station_name.startswith("tmp") or len(station_name) < 2:
        station_name = Path(filepath).parent.name

    with open(filepath, "rb") as f:
        raw = f.read()
    text = raw.decode(encoding, errors="replace")
    lines = text.split("\n")

    header_cols = None
    limits_lo = []  # Row 0: upper spec limits (from col 8)
    limits_hi = []  # Row 1: lower spec limits (from col 8)
    for i, line in enumerate(lines):
        if i == 0:
            limits_lo = line.strip().split("\t")
        elif i == 1:
            limits_hi = line.strip().split("\t")
        if line.startswith("SN\t") or line.startswith("SN\r"):
            header_cols = line.strip().split("\t")
    if header_cols is None:
        raise ValueError(f"未找到表头: {filepath}")

    sn_indices = [i for i, line in enumerate(lines)
                  if re.match(r"^[A-Z0-9]{10,30}$", line.split("\t")[0].strip())
                  and not line.startswith("^")]

    if not sn_indices:
        raise ValueError(f"未找到测试记录: {filepath}")

    record_span = sn_indices[1] - sn_indices[0] if len(sn_indices) >= 2 and 5 <= sn_indices[1] - sn_indices[0] <= 15 else 9

    for sn_line_num in sn_indices:
        try:
            fields = lines[sn_line_num].strip().split("\t")
            if len(fields) < 8:
                continue
            record = {
                "sn": fields[0].strip(), "date": fields[1].strip() if len(fields) > 1 else "",
                "fixture_id": fields[2].strip() if len(fields) > 2 else "",
                "station": fields[3].strip() if len(fields) > 3 else "",
                "project": fields[4].strip() if len(fields) > 4 else "",
                "result": fields[5].strip() if len(fields) > 5 else "",
                "error_code": fields[6].strip() if len(fields) > 6 else "",
                "cycle_time": fields[7].strip() if len(fields) > 7 else "",
                "file_source": station_name,
                "file_path": filepath,
                "header_cols": header_cols,
                "_limits_lo": limits_lo,
                "_limits_hi": limits_hi,
                "_fields_raw": fields,
            }
            # SPA values (line +5)
            spa_line = sn_line_num + 5
            if spa_line < len(lines):
                sf = lines[spa_line].strip().split("\t")
                record["spa_values"] = [sf[i].strip() if i < len(sf) else "" for i in range(5)]
            else:
                record["spa_values"] = []
            # freq sweep (line +8)
            sweep_line = sn_line_num + 8
            if sweep_line < len(lines):
                record["freq_sweep"] = [s.strip() for s in lines[sweep_line].strip().split("\t")[1:] if s.strip()]
            else:
                record["freq_sweep"] = []
            # product info (line +2)
            pi_line = sn_line_num + 2
            if pi_line < len(lines) and "GETPRODUCTINFO:" in lines[pi_line]:
                record["product_info"] = lines[pi_line].split("GETPRODUCTINFO:")[-1].strip()
            else:
                record["product_info"] = ""
            yield record
        except Exception:
            continue

def parse_source(source_path, progress_callback=None):
    """统一入口：支持 zip / 文件夹 / 单文件
    progress_callback(phase, detail) 用于 GUI 进度反馈
    返回 (records, skipped_files)
    """
    all_records = []
    skipped = []

    if source_path.endswith(".zip"):
        if progress_callback: progress_callback("解压中...", source_path)
        with TemporaryDirectory() as tmpdir:
            with zipfile.ZipFile(source_path, "r") as zf:
                zf.extractall(tmpdir)
            for root, dirs, files in os.walk(tmpdir):
                for fname in files:
                    if fname.lower().endswith((".xls", ".txt")):
                        fpath = os.path.join(root, fname)
                        try:
                            all_records.extend(iter_records(fpath))
                        except Exception as e:
                            skipped.append(f"{fname}: {e}")
    elif os.path.isdir(source_path):
        if progress_callback: progress_callback("扫描文件夹...", source_path)
        for root, dirs, files in os.walk(source_path):
            for fname in files:
                if fname.lower().endswith((".xls", ".txt")):
                    fpath = os.path.join(root, fname)
                    try:
                        all_records.extend(iter_records(fpath))
                    except Exception as e:
                        skipped.append(f"{fname}: {e}")
    else:
        if progress_callback: progress_callback("读取文件...", source_path)
        for rec in iter_records(source_path):
            all_records.append(rec)
        parent = Path(source_path).parent.parent.name
        if parent.startswith("tmp") or len(parent) < 2:
            parent = Path(source_path).parent.name
        for r in all_records:
            r["file_source"] = parent

    return all_records, skipped

# ═══════════════════════════════════════════════════════════════════
# 分析引擎
# ═══════════════════════════════════════════════════════════════════

def analyze(records):
    if not records:
        return {"error": "无有效记录"}

    total = len(records)
    pass_recs = [r for r in records if r["result"] == "PASS"]
    fail_recs = [r for r in records if r["result"] != "PASS" and r["result"]]
    pass_count, fail_count = len(pass_recs), len(fail_recs)

    by_station = defaultdict(lambda: {"total": 0, "pass": 0, "fail": 0, "fail_sns": [], "cycle_times": []})
    for r in records:
        st = r["file_source"]
        by_station[st]["total"] += 1
        if r["result"] == "PASS":
            by_station[st]["pass"] += 1
        elif r["result"]:
            by_station[st]["fail"] += 1
            by_station[st]["fail_sns"].append(r["sn"])
        try:
            by_station[st]["cycle_times"].append(float(r["cycle_time"]))
        except Exception:
            pass

    by_sn = defaultdict(lambda: {"stations_failed": [], "fail_count": 0, "results": []})
    for r in records:
        sn = r["sn"]
        info = by_sn[sn]
        info["results"].append({"station": r["file_source"], "result": r["result"],
                                "error_code": r["error_code"], "cycle_time": r["cycle_time"],
                                "date": r["date"]})
        if r["result"] == "PASS":
            pass
        elif r["result"]:
            info["stations_failed"].append(r["file_source"])
            info["fail_count"] += 1

    failed_sns = sorted(
        [{"sn": sn, **{k: v for k, v in info.items()}}
         for sn, info in by_sn.items() if info["fail_count"] > 0],
        key=lambda x: x["fail_count"], reverse=True,
    )

    by_hour = defaultdict(lambda: {"total": 0, "pass": 0, "fail": 0})
    for r in records:
        try:
            hour = r["date"][8:10] if len(r["date"]) >= 10 else ""
            if hour:
                by_hour[hour]["total"] += 1
                if r["result"] == "PASS":
                    by_hour[hour]["pass"] += 1
                elif r["result"]:
                    by_hour[hour]["fail"] += 1
        except Exception:
            pass

    all_ct = []
    for r in records:
        try:
            all_ct.append(float(r["cycle_time"]))
        except Exception:
            pass
    ct_stats = {}
    if all_ct:
        sorted_ct = sorted(all_ct)
        ct_stats = {"avg": sum(all_ct) / len(all_ct), "min": min(all_ct),
                    "max": max(all_ct), "median": sorted_ct[len(all_ct) // 2]}

    products = set()
    for r in records:
        if r.get("product_info"):
            products.add(r["product_info"])

    multi_fail = [x for x in failed_sns if x["fail_count"] >= 2]

    return {
        "total": total, "pass": pass_count, "fail": fail_count,
        "pass_rate": pass_count / total * 100 if total else 0,
        "fail_rate": fail_count / total * 100 if total else 0,
        "by_station": dict(by_station),
        "failed_sns": failed_sns, "multi_fail_sns": multi_fail,
        "by_hour": dict(by_hour), "cycle_time_stats": ct_stats,
        "products": list(products), "total_stations": len(by_station),
    }


def analyze_failure_reasons(records, analysis):
    """失败原因分析 — 逐项对比测量值与规格限，定位超限测试项
    返回: {categories: [{name, count, pct, detail}], by_station: ...}
    """
    failed = [r for r in records if r["result"] != "PASS" and r["result"]]
    if not failed:
        return {"categories": [], "by_station": {}, "total_fail": 0}

    reasons = defaultdict(lambda: {"count": 0, "sns": [], "stations": set()})

    for r in failed:
        matched = False
        limits_lo = r.get("_limits_lo", [])
        limits_hi = r.get("_limits_hi", [])
        header = r.get("header_cols", [])

        # ── SPA 值超限 (cols 10-13, 格式 "<lower,upper>") ──
        spas = r.get("spa_values", [])
        if len(spas) >= 5:
            for ch in range(1, 5):
                v_str = spas[ch] if ch < len(spas) else ""
                if not v_str:
                    continue
                try:
                    v = float(v_str)
                    col_idx = 9 + ch  # cols 10-13
                    limit_str = limits_lo[col_idx] if col_idx < len(limits_lo) else ""
                    if limit_str.startswith("<") and "," in limit_str:
                        parts = limit_str.strip("<>").split(",")
                        lo = float(parts[0]) if parts[0] else 0
                        hi = float(parts[1]) if parts[1] else float("inf")
                        if v < lo or v > hi:
                            name = f"SPA-CH{ch}超限({header[col_idx][:12] if col_idx < len(header) else ''})"
                            reasons[name]["count"] += 1
                            reasons[name]["sns"].append(r["sn"])
                            reasons[name]["stations"].add(r["file_source"])
                            matched = True
                except Exception:
                    pass

        # ── 频率扫描值超限 (cols 14+ ) ──
        sweep = r.get("freq_sweep", [])
        if sweep:
            category_violations = defaultdict(int)
            for idx, s in enumerate(sweep):
                col_idx = idx + 14  # freq sweep starts at col 14 (0-indexed = index 0->col14)
                if col_idx >= len(limits_lo) or col_idx >= len(limits_hi):
                    break
                lo_str = limits_lo[col_idx].strip()
                hi_str = limits_hi[col_idx].strip()
                if not lo_str or not hi_str:
                    continue
                try:
                    lo_v = float(lo_str)
                    hi_v = float(hi_str)
                    v = float(s)
                    lower = min(lo_v, hi_v)
                    upper = max(lo_v, hi_v)
                    if v < lower or v > upper:
                        # Categorize by header name
                        hdr_name = header[col_idx] if col_idx < len(header) else f"Col{col_idx}"
                        # Extract category: "左主1标压_FR_Left_200" → "FR"
                        if "FR_" in hdr_name:
                            cat = "FR超限"
                        elif "THD_" in hdr_name:
                            cat = "THD超限"
                        elif "Rub&Buzz" in hdr_name:
                            cat = "R&B超限"
                        elif "气密性" in hdr_name:
                            cat = "气密性超限"
                        elif "MIC" in hdr_name:
                            cat = "MIC超限"
                        else:
                            cat = "频扫超限"
                        category_violations[cat] += 1
                except Exception:
                    pass

            for cat, cnt in category_violations.items():
                if cnt > 0:
                    reasons[cat]["count"] += 1
                    reasons[cat]["sns"].append(r["sn"])
                    reasons[cat]["stations"].add(r["file_source"])
                    matched = True

        # ── 电量管控 col 8 ──
        col8 = r["_fields_raw"][8].strip() if len(r["_fields_raw"]) > 8 else ""
        if col8:
            try:
                v8 = float(col8)
                limit8 = limits_lo[8] if len(limits_lo) > 8 else ""
                if limit8.startswith("<") and "," in limit8:
                    parts = limit8.strip("<>").split(",")
                    lo, hi = float(parts[0]), float(parts[1])
                    if v8 < lo or v8 > hi:
                        reasons["电量管控超限"]["count"] += 1
                        reasons["电量管控超限"]["sns"].append(r["sn"])
                        reasons["电量管控超限"]["stations"].add(r["file_source"])
                        matched = True
            except Exception:
                pass

        if not matched:
            reasons["其他原因(ErrorCode空)"]["count"] += 1
            reasons["其他原因(ErrorCode空)"]["sns"].append(r["sn"])
            reasons["其他原因(ErrorCode空)"]["stations"].add(r["file_source"])

    categories = []
    for name, data in sorted(reasons.items(), key=lambda x: -x[1]["count"]):
        categories.append({
            "name": name,
            "count": data["count"],
            "pct": data["count"] / len(failed) * 100 if failed else 0,
            "detail": ", ".join(sorted(data["stations"])) if data["stations"] else "",
            "sns": data["sns"][:10],
        })

    by_station_reasons = defaultdict(lambda: defaultdict(int))
    for name, data in reasons.items():
        for st in data["stations"]:
            by_station_reasons[st][name] += data["count"]

    return {
        "categories": categories,
        "by_station": {st: [{"reason": r, "count": c} for r, c in sorted(cnt.items(), key=lambda x: -x[1])]
                       for st, cnt in by_station_reasons.items()},
        "total_fail": len(failed),
    }


# ═══════════════════════════════════════════════════════════════════
# 图表生成
# ═══════════════════════════════════════════════════════════════════

def _make_chart_overall_pie(analysis, out_path):
    _get_font()
    fig, ax = plt.subplots(figsize=(6, 5))
    labels = ["PASS", "FAIL"]
    sizes = [analysis["pass"], analysis["fail"]]
    colors = [COLORS["pass"], COLORS["fail"]]
    wedges, texts, autotexts = ax.pie(
        sizes, labels=None, autopct="%1.1f%%", colors=colors,
        startangle=90, pctdistance=0.6, explode=(0, 0.05),
    )
    for at in autotexts:
        at.set_fontsize(13); at.set_fontweight("bold"); at.set_color("white")
    ax.legend(wedges, [f"{l}  ({s}台)" for l, s in zip(labels, sizes)],
              loc="lower center", ncol=2, frameon=False, prop={"size": 11})
    ax.set_title(f"整体直通率\nPASS率: {analysis['pass_rate']:.1f}%  |  总数: {analysis['total']}  |  FAIL: {analysis['fail']}",
                 fontsize=14, fontweight="bold", pad=20)
    plt.tight_layout()
    fig.savefig(out_path, dpi=150, bbox_inches="tight")
    plt.close(fig)


def _make_chart_station_bars(analysis, out_path):
    _get_font()
    stations = sorted(analysis["by_station"].keys())
    pass_vals = [analysis["by_station"][s]["pass"] for s in stations]
    fail_vals = [analysis["by_station"][s]["fail"] for s in stations]
    x = range(len(stations))
    fig, ax = plt.subplots(figsize=(max(8, len(stations) * 1.8), 6))
    ax.bar(x, pass_vals, 0.6, label="PASS", color=COLORS["bar_pass"], edgecolor="white")
    ax.bar(x, fail_vals, 0.6, bottom=pass_vals, label="FAIL", color=COLORS["bar_fail"], edgecolor="white")
    for i, st in enumerate(stations):
        total = pass_vals[i] + fail_vals[i]
        rate = fail_vals[i] / total * 100 if total else 0
        ax.text(i, total + 1, f"{rate:.1f}%", ha="center", va="bottom",
                fontsize=9, color=COLORS["fail"] if rate > 10 else "#555",
                fontweight="bold" if rate > 10 else "normal")
    ax.set_xticks(x); ax.set_xticklabels(stations, fontsize=11)
    ax.set_ylabel("测试数量", fontsize=12)
    ax.set_title("各站别 PASS / FAIL 分布", fontsize=14, fontweight="bold")
    ax.legend(fontsize=11, frameon=False)
    ax.yaxis.set_major_locator(mticker.MaxNLocator(integer=True))
    ax.set_ylim(0, max(p + f for p, f in zip(pass_vals, fail_vals)) * 1.15)
    plt.tight_layout()
    fig.savefig(out_path, dpi=150, bbox_inches="tight")
    plt.close(fig)


def _make_chart_hourly(analysis, out_path):
    _get_font()
    by_hour = analysis.get("by_hour", {})
    if not by_hour:
        return
    hours = sorted(by_hour.keys())
    rates = [by_hour[h]["pass"] / by_hour[h]["total"] * 100 if by_hour[h]["total"] else 0 for h in hours]
    fig, ax = plt.subplots(figsize=(10, 5))
    ax.plot(hours, rates, marker="o", linewidth=2, markersize=8, color="#2196F3",
            markerfacecolor="white", markeredgewidth=2)
    ax.fill_between(range(len(hours)), rates, alpha=0.1, color="#2196F3")
    ax.set_xlabel("时段（小时）", fontsize=12); ax.set_ylabel("直通率 (%)", fontsize=12)
    ax.set_title("时段直通率趋势", fontsize=14, fontweight="bold")
    ax.set_ylim(0, 105)
    ax.axhline(y=analysis["pass_rate"], color=COLORS["highlight"], linestyle="--",
               linewidth=1, label=f"整体直通率 {analysis['pass_rate']:.1f}%")
    ax.legend(fontsize=10, frameon=False); ax.grid(axis="y", alpha=0.3)
    plt.tight_layout()
    fig.savefig(out_path, dpi=150, bbox_inches="tight")
    plt.close(fig)


def _make_chart_cycle_time(analysis, out_path):
    _get_font()
    stations = sorted(analysis["by_station"].keys())
    data = [analysis["by_station"][s]["cycle_times"] for s in stations if analysis["by_station"][s]["cycle_times"]]
    labels = [s for s in stations if analysis["by_station"][s]["cycle_times"]]
    if not data:
        return
    fig, ax = plt.subplots(figsize=(max(8, len(labels) * 1.5), 5))
    bp = ax.boxplot(data, patch_artist=True, showmeans=True,
                    meanprops={"marker": "D", "markerfacecolor": COLORS["highlight"],
                               "markeredgecolor": "black", "markersize": 6})
    ax.set_xticklabels(labels)
    for patch in bp["boxes"]:
        patch.set_facecolor("#E3F2FD"); patch.set_edgecolor("#1976D2"); patch.set_linewidth(1.2)
    ax.set_ylabel("节拍时间 (秒)", fontsize=12)
    ax.set_title("各站别节拍时间分布", fontsize=14, fontweight="bold")
    ax.grid(axis="y", alpha=0.3)
    plt.tight_layout()
    fig.savefig(out_path, dpi=150, bbox_inches="tight")
    plt.close(fig)


def _make_chart_fail_dist(analysis, out_path):
    _get_font()
    failed = analysis.get("failed_sns", [])
    if not failed:
        fig, ax = plt.subplots(figsize=(8, 2))
        ax.text(0.5, 0.5, "无不良记录", ha="center", va="center", fontsize=16, color="#999")
        ax.axis("off")
        fig.savefig(out_path, dpi=150, bbox_inches="tight")
        plt.close(fig)
        return
    dist = Counter(sn["fail_count"] for sn in failed)
    fig, ax = plt.subplots(figsize=(8, 5))
    cats = sorted(dist.keys())
    counts = [dist[c] for c in cats]
    bars = ax.bar([f"{c}站失败" for c in cats], counts,
                  color=[COLORS["highlight"] if c >= 2 else COLORS["bar_fail"] for c in cats],
                  edgecolor="white")
    for bar, val in zip(bars, counts):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.5,
                str(val), ha="center", va="bottom", fontsize=12, fontweight="bold")
    ax.set_ylabel("SN 数量", fontsize=12)
    ax.set_title(f"失败 SN 跨站分布（共 {len(failed)} 个 SN 失败）", fontsize=14, fontweight="bold")
    ax.yaxis.set_major_locator(mticker.MaxNLocator(integer=True))
    plt.tight_layout()
    fig.savefig(out_path, dpi=150, bbox_inches="tight")
    plt.close(fig)


def _make_chart_fail_matrix(analysis, out_path):
    _get_font()
    failed = analysis.get("failed_sns", [])
    if not failed:
        return
    stations = sorted(analysis["by_station"].keys())
    top = failed[:min(30, len(failed))]
    matrix = [[1 if st in sn["stations_failed"] else 0 for st in stations] for sn in top]
    sn_labels = [sn["sn"][-8:] for sn in top]
    fig, ax = plt.subplots(figsize=(max(10, len(stations) * 1.5), max(6, len(top) * 0.35)))
    ax.imshow(matrix, cmap="RdYlGn_r", aspect="auto", vmin=0, vmax=1)
    ax.set_xticks(range(len(stations))); ax.set_xticklabels(stations, fontsize=9, rotation=45, ha="right")
    ax.set_yticks(range(len(sn_labels))); ax.set_yticklabels(sn_labels, fontsize=7, family="monospace")
    ax.set_title(f"失败 SN × 站别 热力图（前 {len(top)}）", fontsize=12, fontweight="bold")
    plt.tight_layout()
    fig.savefig(out_path, dpi=150, bbox_inches="tight")
    plt.close(fig)


def _make_chart_failure_reasons(failure_analysis, out_path):
    """失败原因分布饼图"""
    _get_font()
    categories = failure_analysis.get("categories", [])
    if not categories:
        fig, ax = plt.subplots(figsize=(6, 2))
        ax.text(0.5, 0.5, "无失败记录", ha="center", va="center", fontsize=16, color="#999")
        ax.axis("off")
        fig.savefig(out_path, dpi=150, bbox_inches="tight")
        plt.close(fig)
        return

    # Take top 8 categories, merge rest
    top = categories[:8]
    labels = [c["name"] for c in top]
    sizes = [c["count"] for c in top]
    rest = sum(c["count"] for c in categories[8:])
    if rest > 0:
        labels.append("其他")
        sizes.append(rest)

    colors_list = ["#F44336", "#FF9800", "#FF5722", "#E91E63",
                   "#9C27B0", "#673AB7", "#3F51B5", "#2196F3", "#607D8B"]

    fig, ax = plt.subplots(figsize=(8, 6))
    wedges, texts, autotexts = ax.pie(
        sizes, labels=None, autopct="%1.1f%%",
        colors=colors_list[:len(sizes)], startangle=140,
        pctdistance=0.75, explode=[0.03] * len(sizes),
    )
    for at in autotexts:
        at.set_fontsize(10); at.set_fontweight("bold")

    legend_labels = [f"{l} ({s})" for l, s in zip(labels, sizes)]
    ax.legend(wedges, legend_labels, loc="center left",
              bbox_to_anchor=(1, 0.5), frameon=False, prop={"size": 10})

    ax.set_title(
        f"失败原因分布（共 {failure_analysis['total_fail']} 次失败）",
        fontsize=14, fontweight="bold", pad=20,
    )
    plt.tight_layout()
    fig.savefig(out_path, dpi=150, bbox_inches="tight")
    plt.close(fig)


# ═══════════════════════════════════════════════════════════════════
# HTML 报告
# ═══════════════════════════════════════════════════════════════════

_HTML = """<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>AT Audio Test 数据分析报告</title><style>
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
td{{padding:8px;border-bottom:1px solid #eceff1}}
tr:hover td{{background:#f5f5f5}}
.badge{{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600;color:white}}
.badge-pass{{background:#4CAF50}}.badge-fail{{background:#F44336}}.badge-warn{{background:#FF9800}}
.note{{font-size:12px;color:#999;margin-top:6px}}
</style></head><body><div class="container">
<h1>🔊 AT Audio Test 数据分析报告</h1>
<p class="subtitle">生成时间: {report_time}  |  产品: {products}  |  数据日期: 2026-06-13</p>
<div class="summary-cards">
<div class="card info"><div class="value">{total}</div><div class="label">总测试数</div></div>
<div class="card pass"><div class="value">{pass_count}</div><div class="label">PASS</div></div>
<div class="card fail"><div class="value">{fail_count}</div><div class="label">FAIL</div></div>
<div class="card {pass_rate_class}"><div class="value">{pass_rate:.1f}%</div><div class="label">直通率</div></div>
<div class="card info"><div class="value">{total_stations}</div><div class="label">测试站别</div></div>
<div class="card info"><div class="value">{multi_fail_count}</div><div class="label">多站失败SN</div></div>
</div>
<div class="chart-section"><h2>📊 整体直通率</h2><img src="chart_overall_pie.png" alt="直通率饼图"></div>
<div class="chart-section"><h2>📊 各站别 PASS / FAIL 分布</h2><img src="chart_station_bars.png" alt="站别分布"></div>
<div class="chart-section"><h2>📊 时段直通率趋势</h2><img src="chart_hourly_rate.png" alt="时段趋势"></div>
<div class="chart-section"><h2>📊 节拍时间分布</h2><img src="chart_cycle_time.png" alt="节拍分布"></div>
<div class="chart-section"><h2>📊 不良 SN 跨站分布</h2><img src="chart_fail_dist.png" alt="不良分布"></div>
<div class="chart-section"><h2>📊 失败 SN × 站别 矩阵（前30）</h2><img src="chart_fail_matrix.png" alt="失败矩阵"></div>
<div class="chart-section"><h2>📋 各站别统计明细</h2><table>
<tr><th>站别</th><th>总数</th><th>PASS</th><th>FAIL</th><th>不良率</th><th>平均节拍(s)</th><th>最小节拍</th><th>最大节拍</th></tr>
{station_table_rows}</table></div>
<div class="chart-section"><h2>🔴 跨站重复失败 SN（{multi_fail_count} 个）</h2>
{multi_fail_table}<p class="note">⚠ 以下 SN 在多个站别均测试失败，建议优先排查。</p></div>
<div class="chart-section"><h2>🔍 失败原因透视分析</h2>
<img src="chart_failure_reasons.png" alt="失败原因分布" style="max-width:70%">
<table><tr><th>失败原因</th><th>次数</th><th>占比</th><th>涉及站别</th><th>典型SN</th></tr>
{failure_reason_rows}</table></div>
<div class="chart-section"><h2>📋 全部失败 SN 清单（{all_fail_count} 个）</h2>
{all_fail_table}</div>
</div></body></html>"""


def generate_html_report(analysis, failure_analysis, chart_dir, out_path):
    _get_font()
    # Station table
    station_rows = []
    for st in sorted(analysis["by_station"].keys()):
        d = analysis["by_station"][st]
        total = d["total"]
        rate = d["fail"] / total * 100 if total else 0
        cts = d["cycle_times"]
        avg_ct = f"{sum(cts)/len(cts):.1f}" if cts else "-"
        min_ct = f"{min(cts):.1f}" if cts else "-"
        max_ct = f"{max(cts):.1f}" if cts else "-"
        rc = "badge-pass" if rate < 3 else "badge-warn" if rate < 10 else "badge-fail"
        station_rows.append(
            f"<tr><td><strong>{st}</strong></td><td>{total}</td><td>{d['pass']}</td><td>{d['fail']}</td>"
            f"<td><span class='badge {rc}'>{rate:.1f}%</span></td>"
            f"<td>{avg_ct}</td><td>{min_ct}</td><td>{max_ct}</td></tr>"
        )

    # Multi-fail table
    multi_fail = analysis.get("multi_fail_sns", [])
    if multi_fail:
        multi_rows = []
        for sn_info in multi_fail[:50]:
            multi_rows.append(
                f"<tr><td><code>{sn_info['sn']}</code></td>"
                f"<td><span class='badge badge-fail'>{sn_info['fail_count']}站</span></td>"
                f"<td>{', '.join(sn_info['stations_failed'])}</td></tr>"
            )
        multi_table = (
            "<table><tr><th>SN</th><th>失败站数</th><th>失败站别</th></tr>"
            + "".join(multi_rows) + "</table>"
        )
    else:
        multi_table = "<p>✅ 无跨站重复失败</p>"

    # All fail table
    all_failed = analysis.get("failed_sns", [])
    if all_failed:
        fail_rows = []
        for sn_info in all_failed[:100]:
            stations = ", ".join(f"{r['station']}({r['result']})" for r in sn_info["results"])
            fail_rows.append(
                f"<tr><td><code>{sn_info['sn']}</code></td>"
                f"<td>{sn_info['fail_count']}</td><td style='font-size:11px'>{stations}</td></tr>"
            )
        all_fail_table = (
            "<table><tr><th>SN</th><th>失败次数</th><th>详情</th></tr>"
            + "".join(fail_rows) + "</table>"
        )
        if len(all_failed) > 100:
            all_fail_table += f"<p class='note'>... 共 {len(all_failed)} 条，仅显示前 100 条</p>"
    else:
        all_fail_table = "<p>✅ 无失败记录</p>"

    # Failure reason rows
    fr_categories = failure_analysis.get("categories", [])
    if fr_categories:
        fr_rows = []
        for c in fr_categories:
            sns = ", ".join(c.get("sns", [])[:3]) if c.get("sns") else "—"
            fr_rows.append(
                f"<tr><td>{c['name']}</td><td>{c['count']}</td>"
                f"<td>{c['pct']:.1f}%</td><td>{c['detail']}</td>"
                f"<td style='font-size:11px'><code>{sns}</code></td></tr>"
            )
        failure_reason_rows = "".join(fr_rows)
    else:
        failure_reason_rows = "<tr><td colspan='5'>✅ 无失败记录</td></tr>"

    html = _HTML.format(
        report_time=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        products=", ".join(analysis.get("products", ["未知"])),
        total=analysis["total"], pass_count=analysis["pass"], fail_count=analysis["fail"],
        pass_rate=analysis["pass_rate"],
        pass_rate_class="pass" if analysis["pass_rate"] >= 95 else "warn" if analysis["pass_rate"] >= 85 else "fail",
        total_stations=analysis["total_stations"],
        multi_fail_count=len(multi_fail), all_fail_count=len(all_failed),
        station_table_rows="".join(station_rows),
        multi_fail_table=multi_table, all_fail_table=all_fail_table,
        failure_reason_rows=failure_reason_rows,
    )
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    return out_path


# ═══════════════════════════════════════════════════════════════════
# 工具函数
# ═══════════════════════════════════════════════════════════════════

def open_with_default(path):
    if not path or not os.path.exists(path):
        return
    try:
        if sys.platform == "win32": os.startfile(path)
        elif sys.platform == "darwin": subprocess.run(["open", path])
        else: subprocess.run(["xdg-open", path])
    except Exception:
        pass


# ═══════════════════════════════════════════════════════════════════
# 入口（GUI 延迟加载，分析引擎无 tkinter 也可单独使用）
# ═══════════════════════════════════════════════════════════════════

# GUI 样式（模块级常量，_run_gui 引用）
STYLE = {
    "bg": "#f0f2f5", "fg": "#1a1a2e", "card_bg": "#ffffff",
    "accent": "#2196F3", "accent_hover": "#1976D2",
    "success": "#4CAF50", "danger": "#F44336", "warning": "#FF9800",
    "border": "#e0e0e0", "text_secondary": "#78909c",
    "input_bg": "#ffffff", "tree_bg": "#ffffff", "tree_sel": "#E3F2FD",
    "progress_bg": "#e0e0e0",
}


def _run_gui():
    """延迟导入 tkinter，仅 GUI 模式需要"""
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    class App:
        def __init__(self, root):
            self.root = root
            self.root.title("AT Audio Test 数据分析工具")
            self.root.geometry("1100x750")
            self.root.minsize(900, 600)
            self.root.configure(bg=STYLE["bg"])

            self.source_path = tk.StringVar()
            self.output_dir = tk.StringVar(value=os.path.join(os.getcwd(), "at_report"))
            self.status_text = tk.StringVar(value="就绪 — 请选择数据源文件或文件夹")
            self.progress_var = tk.DoubleVar(value=0)

            self.opt_charts = tk.BooleanVar(value=True)
            self.opt_html = tk.BooleanVar(value=True)
            self.opt_csv = tk.BooleanVar(value=True)
            self.opt_auto_open = tk.BooleanVar(value=True)

            self.analysis_result = None
            self.records = None
            self.report_html_path = None
            self.is_running = False

            self._apply_theme()
            self._build_ui()
            self._center_window()

        def _apply_theme(self):
            style = ttk.Style(self.root)
            try: style.theme_use("clam")
            except tk.TclError: pass
            style.configure("TProgressbar", background=STYLE["accent"], troughcolor=STYLE["progress_bg"])
            style.configure("Treeview", background=STYLE["tree_bg"], foreground=STYLE["fg"],
                            fieldbackground=STYLE["tree_bg"], font=("Consolas", 10), rowheight=26)
            style.configure("Treeview.Heading", background=STYLE["card_bg"], foreground=STYLE["fg"],
                            font=("Microsoft YaHei", 9, "bold"), padding=5)
            style.map("Treeview", background=[("selected", STYLE["tree_sel"])])

        def _center_window(self):
            self.root.update_idletasks()
            w, h = self.root.winfo_width(), self.root.winfo_height()
            sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
            self.root.geometry(f"+{(sw-w)//2}+{(sh-h)//2}")

        def _build_ui(self):
            hdr = tk.Frame(self.root, bg=STYLE["accent"], height=56)
            hdr.pack(fill="x"); hdr.pack_propagate(False)
            tk.Label(hdr, text="🔊  AT Audio Test 数据分析工具", bg=STYLE["accent"], fg="white",
                     font=("Microsoft YaHei", 16, "bold")).pack(side="left", padx=20, pady=12)
            tk.Label(hdr, text="v2.0 单文件版", bg=STYLE["accent"],
                     fg="#B0BEC5", font=("Microsoft YaHei", 9)).pack(side="right", padx=20, pady=12)

            main = tk.Frame(self.root, bg=STYLE["bg"])
            main.pack(fill="both", expand=True, padx=15, pady=(15, 0))

            left = tk.Frame(main, bg=STYLE["bg"]); left.pack(side="left", fill="y", padx=(0, 10))
            self._build_source_section(left)
            self._build_output_section(left)
            self._build_options_section(left)
            self._build_action_button(left)

            right = tk.Frame(main, bg=STYLE["bg"]); right.pack(side="left", fill="both", expand=True)
            self._build_result_section(right)

            sbar = tk.Frame(self.root, bg=STYLE["card_bg"], height=34)
            sbar.pack(fill="x", side="bottom", pady=(10, 0)); sbar.pack_propagate(False)
            tk.Label(sbar, textvariable=self.status_text, bg=STYLE["card_bg"],
                     fg=STYLE["text_secondary"], font=("Microsoft YaHei", 9), anchor="w"
                     ).pack(side="left", fill="x", padx=15, pady=6)
            self.progress_bar = ttk.Progressbar(sbar, variable=self.progress_var, mode="determinate", length=200)
            self.progress_bar.pack(side="right", padx=15, pady=6)

        def _build_source_section(self, parent):
            frm = tk.LabelFrame(parent, text="📁  数据源", bg=STYLE["card_bg"], fg=STYLE["fg"],
                                font=("Microsoft YaHei", 11, "bold"), padx=15, pady=12, relief="groove", bd=1)
            frm.pack(fill="x", pady=(0, 10))
            r1 = tk.Frame(frm, bg=STYLE["card_bg"]); r1.pack(fill="x", pady=(5, 8))
            self.src_entry = tk.Entry(r1, textvariable=self.source_path, font=("Consolas", 10),
                                       bg=STYLE["input_bg"], relief="solid", bd=1)
            self.src_entry.pack(side="left", fill="x", expand=True, ipady=4)
            tk.Button(r1, text="📂 选择...", command=self._browse_source,
                      bg=STYLE["accent"], fg="white", font=("Microsoft YaHei", 9),
                      relief="flat", padx=15, pady=4, cursor="hand2",
                      activebackground=STYLE["accent_hover"]).pack(side="left", padx=(8, 0))
            r2 = tk.Frame(frm, bg=STYLE["card_bg"]); r2.pack(fill="x")
            for txt, cmd in [("📦 选ZIP", "zip"), ("📁 选文件夹", "dir"), ("📄 选单文件", "file")]:
                tk.Button(r2, text=txt, command=lambda c=cmd: self._browse_source(c),
                          bg="#ECEFF1", fg=STYLE["fg"], font=("Microsoft YaHei", 9),
                          relief="flat", padx=12, pady=3, cursor="hand2").pack(side="left", padx=(0, 8))
            tk.Label(frm, text="支持 .zip / .xls / .txt (TSV格式)   |   加密文件请先另存为 .txt",
                     bg=STYLE["card_bg"], fg=STYLE["text_secondary"], font=("Microsoft YaHei", 8)
                     ).pack(anchor="w", pady=(8, 0))

        def _build_output_section(self, parent):
            frm = tk.LabelFrame(parent, text="📂  输出目录", bg=STYLE["card_bg"], fg=STYLE["fg"],
                                font=("Microsoft YaHei", 11, "bold"), padx=15, pady=12, relief="groove", bd=1)
            frm.pack(fill="x", pady=(0, 10))
            r = tk.Frame(frm, bg=STYLE["card_bg"]); r.pack(fill="x")
            self.out_entry = tk.Entry(r, textvariable=self.output_dir, font=("Consolas", 10),
                                       bg=STYLE["input_bg"], relief="solid", bd=1)
            self.out_entry.pack(side="left", fill="x", expand=True, ipady=4)
            tk.Button(r, text="📂 选择...", command=self._browse_output,
                      bg=STYLE["accent"], fg="white", font=("Microsoft YaHei", 9),
                      relief="flat", padx=15, pady=4, cursor="hand2",
                      activebackground=STYLE["accent_hover"]).pack(side="left", padx=(8, 0))

        def _build_options_section(self, parent):
            frm = tk.LabelFrame(parent, text="⚙️  输出选项", bg=STYLE["card_bg"], fg=STYLE["fg"],
                                font=("Microsoft YaHei", 11, "bold"), padx=15, pady=12, relief="groove", bd=1)
            frm.pack(fill="x", pady=(0, 10))
            for var, txt in [(self.opt_charts, "生成 PNG 图表（6张）"),
                             (self.opt_html, "生成 HTML 汇总报告"),
                             (self.opt_csv, "导出 CSV 明细"),
                             (self.opt_auto_open, "完成后自动打开报告")]:
                tk.Checkbutton(frm, text=txt, variable=var, bg=STYLE["card_bg"],
                               font=("Microsoft YaHei", 10), activebackground=STYLE["card_bg"],
                               selectcolor=STYLE["card_bg"]).pack(anchor="w", pady=2)

        def _build_action_button(self, parent):
            frm = tk.Frame(parent, bg=STYLE["bg"]); frm.pack(fill="x", pady=(5, 10))
            self.run_btn = tk.Button(frm, text="▶  开始分析", command=self._start_analysis,
                                      bg=STYLE["accent"], fg="white",
                                      font=("Microsoft YaHei", 12, "bold"),
                                      relief="flat", padx=30, pady=10, cursor="hand2",
                                      activebackground=STYLE["accent_hover"])
            self.run_btn.pack(fill="x", expand=True)

        def _build_result_section(self, parent):
            cards_frm = tk.Frame(parent, bg=STYLE["bg"]); cards_frm.pack(fill="x", pady=(0, 8))
            self.cards = {}
            for i, (key, label, color) in enumerate([
                ("total", "总测试数", "#2196F3"), ("pass", "PASS", "#4CAF50"),
                ("fail", "FAIL", "#F44336"), ("rate", "直通率", "#FF9800"),
            ]):
                card = tk.Frame(cards_frm, bg=STYLE["card_bg"], relief="solid", bd=1, padx=10, pady=10)
                card.grid(row=0, column=i, padx=4, sticky="nsew")
                cards_frm.grid_columnconfigure(i, weight=1)
                self.cards[key] = {
                    "val": tk.Label(card, text="—", bg=STYLE["card_bg"], fg=color,
                                    font=("Consolas", 22, "bold")),
                    "lbl": tk.Label(card, text=label, bg=STYLE["card_bg"],
                                    fg=STYLE["text_secondary"], font=("Microsoft YaHei", 9)),
                }
                self.cards[key]["val"].pack(); self.cards[key]["lbl"].pack()

            tbl_frm = tk.Frame(parent, bg=STYLE["bg"]); tbl_frm.pack(fill="both", expand=True)

            sframe = tk.LabelFrame(tbl_frm, text="📊  各站别统计", bg=STYLE["card_bg"],
                                    fg=STYLE["fg"], font=("Microsoft YaHei", 10, "bold"),
                                    padx=8, pady=8, relief="groove", bd=1)
            sframe.pack(side="left", fill="both", expand=True, padx=(0, 4))
            self.station_tree = ttk.Treeview(sframe, columns=("station", "total", "pass", "fail", "rate", "avg_ct"),
                                              show="headings", height=8)
            for col, txt, w in [("station", "站别", 70), ("total", "总数", 55), ("pass", "PASS", 55),
                                ("fail", "FAIL", 55), ("rate", "不良率", 75), ("avg_ct", "平均节拍", 75)]:
                self.station_tree.heading(col, text=txt)
                self.station_tree.column(col, width=w, anchor="center")
            self.station_tree.pack(fill="both", expand=True)
            for tag, color in [("high_fail", "#F44336"), ("mid_fail", "#FF9800"), ("low_fail", "#4CAF50")]:
                self.station_tree.tag_configure(tag, foreground=color)

            fframe = tk.LabelFrame(tbl_frm, text="🔴  跨站重复失败 SN", bg=STYLE["card_bg"],
                                    fg=STYLE["fg"], font=("Microsoft YaHei", 10, "bold"),
                                    padx=8, pady=8, relief="groove", bd=1)
            fframe.pack(side="left", fill="both", expand=True, padx=(4, 0))
            self.fail_tree = ttk.Treeview(fframe, columns=("sn", "count", "stations"),
                                           show="headings", height=8)
            self.fail_tree.heading("sn", text="SN"); self.fail_tree.column("sn", width=160)
            self.fail_tree.heading("count", text="失败站数"); self.fail_tree.column("count", width=70, anchor="center")
            self.fail_tree.heading("stations", text="失败站别"); self.fail_tree.column("stations", width=150)
            self.fail_tree.pack(fill="both", expand=True)
            self.fail_tree.tag_configure("critical", foreground="#F44336", font=("Consolas", 10, "bold"))

            # Failure reasons table (below the two above)
            rframe = tk.LabelFrame(parent, text="🔍  失败原因透视", bg=STYLE["card_bg"],
                                    fg=STYLE["fg"], font=("Microsoft YaHei", 10, "bold"),
                                    padx=8, pady=8, relief="groove", bd=1)
            rframe.pack(fill="x", pady=(4, 0))
            self.reason_tree = ttk.Treeview(rframe, columns=("reason", "count", "pct", "detail"),
                                             show="headings", height=5)
            self.reason_tree.heading("reason", text="失败原因"); self.reason_tree.column("reason", width=200)
            self.reason_tree.heading("count", text="次数"); self.reason_tree.column("count", width=60, anchor="center")
            self.reason_tree.heading("pct", text="占比"); self.reason_tree.column("pct", width=60, anchor="center")
            self.reason_tree.heading("detail", text="涉及站别"); self.reason_tree.column("detail", width=200)
            self.reason_tree.pack(fill="x")
            for tag, color in [("high", "#F44336"), ("mid", "#FF9800"), ("low", "#2196F3")]:
                self.reason_tree.tag_configure(tag, foreground=color)

            btns = tk.Frame(parent, bg=STYLE["bg"]); btns.pack(fill="x", pady=(8, 0))
            self.btn_report = tk.Button(btns, text="📄 打开 HTML 报告", command=self._open_report,
                                         bg=STYLE["success"], fg="white", font=("Microsoft YaHei", 10),
                                         relief="flat", padx=18, pady=6, cursor="hand2", state="disabled")
            self.btn_report.pack(side="left", padx=(0, 8))
            self.btn_folder = tk.Button(btns, text="📂 打开输出目录", command=self._open_folder,
                                         bg="#ECEFF1", fg=STYLE["fg"], font=("Microsoft YaHei", 10),
                                         relief="flat", padx=18, pady=6, cursor="hand2", state="disabled")
            self.btn_folder.pack(side="left", padx=(0, 8))
            self.btn_csv = tk.Button(btns, text="💾 另存 CSV", command=self._save_csv,
                                      bg="#ECEFF1", fg=STYLE["fg"], font=("Microsoft YaHei", 10),
                                      relief="flat", padx=18, pady=6, cursor="hand2", state="disabled")
            self.btn_csv.pack(side="left")

        # ── Events ──
        def _browse_source(self, mode=None):
            if mode == "zip":
                path = filedialog.askopenfilename(title="选择 AT 数据压缩包",
                    filetypes=[("ZIP 压缩包", "*.zip"), ("所有文件", "*.*")])
            elif mode == "dir":
                path = filedialog.askdirectory(title="选择数据文件夹")
            elif mode == "file":
                path = filedialog.askopenfilename(title="选择数据文件",
                    filetypes=[("数据文件", "*.xls;*.txt"), ("所有文件", "*.*")])
            else:
                path = filedialog.askopenfilename(title="选择数据源",
                    filetypes=[("支持格式", "*.zip;*.xls;*.txt"), ("所有文件", "*.*")])
                if not path:
                    path = filedialog.askdirectory(title="或者选择数据文件夹")
            if path:
                self.source_path.set(path)
                self.status_text.set(f"已选择: {os.path.basename(path)}")

        def _browse_output(self):
            path = filedialog.askdirectory(title="选择输出目录", initialdir=self.output_dir.get())
            if path:
                self.output_dir.set(path)

        def _start_analysis(self):
            source = self.source_path.get().strip()
            if not source:
                messagebox.showwarning("提示", "请先选择数据源。"); return
            if not os.path.exists(source):
                messagebox.showerror("错误", f"路径不存在:\n{source}"); return
            if self.is_running:
                messagebox.showinfo("提示", "分析正在进行中..."); return

            self.is_running = True
            self.run_btn.configure(text="⏳ 分析中...", state="disabled", bg=STYLE["text_secondary"])
            self.progress_var.set(5)
            self.status_text.set("正在解析数据文件...")
            self._clear_results()
            threading.Thread(target=self._run_thread, args=(source,), daemon=True).start()

        def _run_thread(self, source):
            try:
                self.root.after(0, lambda: self.progress_var.set(10))
                records, skipped = parse_source(source, progress_callback=lambda phase, detail:
                    self.root.after(0, lambda: self.status_text.set(f"{phase}")))
                if not records:
                    msg = "未找到有效的测试记录。请确认文件包含 BLKVUN 开头的 SN。"
                    if skipped:
                        msg += f"\n跳过的文件: {len(skipped)} 个"
                    raise ValueError(msg)
                self.records = records
                self.root.after(0, lambda: self.progress_var.set(40))
                stations = sorted(set(r["file_source"] for r in records))
                status = f"解析完成: {len(records)} 条, {len(stations)} 个站别"
                if skipped:
                    status += f"（跳过 {len(skipped)} 个文件）"
                self.root.after(0, lambda: self.status_text.set(status))
                analysis = analyze(records)
                self.analysis_result = analysis
                failure_analysis = analyze_failure_reasons(records, analysis)
                self.root.after(0, lambda: self.progress_var.set(60))
                self.root.after(0, lambda: self._show_results(analysis, failure_analysis))
                out_dir = self.output_dir.get()
                os.makedirs(out_dir, exist_ok=True)

                if self.opt_charts.get():
                    self.root.after(0, lambda: self.progress_var.set(70))
                    self.root.after(0, lambda: self.status_text.set("生成图表中..."))
                    _get_font()
                    _make_chart_overall_pie(analysis, os.path.join(out_dir, "chart_overall_pie.png"))
                    _make_chart_station_bars(analysis, os.path.join(out_dir, "chart_station_bars.png"))
                    _make_chart_hourly(analysis, os.path.join(out_dir, "chart_hourly_rate.png"))
                    _make_chart_cycle_time(analysis, os.path.join(out_dir, "chart_cycle_time.png"))
                    _make_chart_fail_dist(analysis, os.path.join(out_dir, "chart_fail_dist.png"))
                    _make_chart_fail_matrix(analysis, os.path.join(out_dir, "chart_fail_matrix.png"))
                    _make_chart_failure_reasons(failure_analysis, os.path.join(out_dir, "chart_failure_reasons.png"))

                if self.opt_csv.get():
                    self.root.after(0, lambda: self.progress_var.set(85))
                    self.root.after(0, lambda: self.status_text.set("导出 CSV 中..."))
                    self._write_csv(os.path.join(out_dir, "records_detail.csv"))

                if self.opt_html.get():
                    self.root.after(0, lambda: self.progress_var.set(95))
                    self.root.after(0, lambda: self.status_text.set("生成 HTML 报告中..."))
                    self.report_html_path = generate_html_report(analysis, failure_analysis, out_dir, os.path.join(out_dir, "report.html"))

                self.root.after(0, lambda: self.progress_var.set(100))
                self.root.after(0, lambda: self.status_text.set("✅ 分析完成！"))
                self.root.after(0, self._on_done)
            except Exception as e:
                import traceback; traceback.print_exc()
                self.root.after(0, lambda: self.status_text.set(f"❌ {str(e)[:80]}"))
                self.root.after(0, lambda: self.progress_var.set(0))
                self.root.after(0, self._on_error)
                self.root.after(0, lambda: messagebox.showerror("分析失败", str(e)))

        def _write_csv(self, path):
            if not self.records: return
            keys = ["sn", "date", "fixture_id", "station", "project",
                    "result", "error_code", "cycle_time", "product_info", "file_source"]
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.DictWriter(f, fieldnames=keys, extrasaction="ignore")
                w.writeheader()
                for r in self.records:
                    w.writerow({k: r.get(k, "") for k in keys})

        def _show_results(self, a, fr=None):
            self.cards["total"]["val"].configure(text=str(a["total"]))
            self.cards["pass"]["val"].configure(text=str(a["pass"]))
            self.cards["fail"]["val"].configure(text=str(a["fail"]))
            self.cards["rate"]["val"].configure(text=f"{a['pass_rate']:.1f}%")
            for item in self.station_tree.get_children():
                self.station_tree.delete(item)
            for st in sorted(a["by_station"].keys()):
                d = a["by_station"][st]
                total = d["total"]; rate = d["fail"] / total * 100 if total else 0
                cts = d["cycle_times"]
                avg_ct = f"{sum(cts)/len(cts):.1f}s" if cts else "-"
                tag = "high_fail" if rate >= 15 else "mid_fail" if rate >= 5 else "low_fail"
                self.station_tree.insert("", "end", values=(st, total, d["pass"], d["fail"], f"{rate:.1f}%", avg_ct), tags=(tag,))
            for item in self.fail_tree.get_children():
                self.fail_tree.delete(item)
            for sn_info in a.get("multi_fail_sns", [])[:30]:
                tag = "critical" if sn_info["fail_count"] >= 3 else ""
                self.fail_tree.insert("", "end", values=(
                    sn_info["sn"], sn_info["fail_count"], ", ".join(sn_info["stations_failed"]),
                ), tags=(tag,) if tag else ())
            if not a.get("multi_fail_sns"):
                self.fail_tree.insert("", "end", values=("✅ 无跨站重复失败", "", ""))
            # Failure reasons
            for item in self.reason_tree.get_children():
                self.reason_tree.delete(item)
            if fr:
                for c in fr.get("categories", [])[:12]:
                    pct = c["pct"]
                    tag = "high" if pct > 30 else "mid" if pct > 15 else "low"
                    self.reason_tree.insert("", "end", values=(
                        c["name"], c["count"], f"{pct:.1f}%", c["detail"],
                    ), tags=(tag,))
            else:
                self.reason_tree.insert("", "end", values=("—", "", "", ""))

        def _clear_results(self):
            for k in self.cards:
                self.cards[k]["val"].configure(text="—")
            for tree in [self.station_tree, self.fail_tree, self.reason_tree]:
                for item in tree.get_children():
                    tree.delete(item)
            self.report_html_path = None
            for btn in [self.btn_report, self.btn_folder, self.btn_csv]:
                btn.configure(state="disabled")

        def _on_done(self):
            self.is_running = False
            self.run_btn.configure(text="▶  重新分析", state="normal", bg=STYLE["accent"])
            for btn in [self.btn_report, self.btn_folder, self.btn_csv]:
                btn.configure(state="normal")
            if self.opt_auto_open.get() and self.report_html_path:
                self._open_report()

        def _on_error(self):
            self.is_running = False
            self.run_btn.configure(text="▶  开始分析", state="normal", bg=STYLE["accent"])

        def _open_report(self):
            if self.report_html_path and os.path.exists(self.report_html_path):
                open_with_default(self.report_html_path)
            else:
                messagebox.showinfo("提示", "报告文件不存在，请先运行分析。")

        def _open_folder(self):
            out = self.output_dir.get()
            if os.path.exists(out):
                open_with_default(out)
            else:
                messagebox.showinfo("提示", "输出目录不存在。")

        def _save_csv(self):
            if not self.records:
                messagebox.showinfo("提示", "暂无分析数据。"); return
            path = filedialog.asksaveasfilename(title="导出 CSV", defaultextension=".csv",
                filetypes=[("CSV", "*.csv")], initialdir=self.output_dir.get(), initialfile="records_detail.csv")
            if path:
                self._write_csv(path)
                self.status_text.set(f"CSV 已保存: {os.path.basename(path)}")
                messagebox.showinfo("成功", f"已保存到:\n{path}")

    root = tk.Tk()
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    App(root)
    root.mainloop()


def main():
    _run_gui()


if __name__ == "__main__":
    main()

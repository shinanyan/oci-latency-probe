#!/usr/bin/env python3
"""
Oracle Cloud Object Storage — TCP Latency Probe

Measures TCP handshake latency (port 443) to all 42 Oracle Cloud
Object Storage endpoints worldwide. Outputs:
  - oci_latency_report.xlsx   (9-sheet Excel workbook)
  - oci_latency_chart.html    (interactive Plotly.js line chart)

Usage:
    python oci_latency_probe.py
"""

import json
import socket
import statistics
import sys
import time
from pathlib import Path

# Fix Windows console encoding for special characters
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# ---------------------------------------------------------------------------
# Dependencies
# ---------------------------------------------------------------------------
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    import subprocess
    print("Installing openpyxl ...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter

# ===========================================================================
#  CONFIG
# ===========================================================================

PROBE_COUNT = 30
PORT = 443
TIMEOUT = 5.0

# (region_code, name_cn, name_en, host)
ENDPOINTS = [
    # -- Asia-Pacific --
    ("ap-osaka-1",       "日本中部大阪",           "Osaka, Japan",                  "objectstorage.ap-osaka-1.oraclecloud.com"),
    ("ap-tokyo-1",       "日本东部东京",           "Tokyo, Japan",                  "objectstorage.ap-tokyo-1.oraclecloud.com"),
    ("ap-seoul-1",       "韩国中部首尔",           "Seoul, South Korea",            "objectstorage.ap-seoul-1.oraclecloud.com"),
    ("ap-chuncheon-1",   "韩国北部春川",           "Chuncheon, South Korea",        "objectstorage.ap-chuncheon-1.oraclecloud.com"),
    ("ap-singapore-1",   "新加坡",                 "Singapore",                     "objectstorage.ap-singapore-1.oraclecloud.com"),
    ("ap-singapore-2",   "新加坡西",               "Singapore West",                "objectstorage.ap-singapore-2.oraclecloud.com"),
    ("ap-mumbai-1",      "印度西部孟买",           "Mumbai, India",                 "objectstorage.ap-mumbai-1.oraclecloud.com"),
    ("ap-hyderabad-1",   "印度南部海得拉巴",       "Hyderabad, India",              "objectstorage.ap-hyderabad-1.oraclecloud.com"),
    ("ap-batam-1",       "印度尼西亚巴淡",         "Batam, Indonesia",              "objectstorage.ap-batam-1.oraclecloud.com"),
    ("ap-sydney-1",      "澳大利亚东部悉尼",       "Sydney, Australia",             "objectstorage.ap-sydney-1.oraclecloud.com"),
    ("ap-melbourne-1",   "澳大利亚东南部墨尔本",   "Melbourne, Australia",          "objectstorage.ap-melbourne-1.oraclecloud.com"),
    # -- North America --
    ("us-ashburn-1",     "美国东部阿什本",         "Ashburn, USA",                  "objectstorage.us-ashburn-1.oraclecloud.com"),
    ("us-phoenix-1",     "美国西部凤凰城",         "Phoenix, USA",                  "objectstorage.us-phoenix-1.oraclecloud.com"),
    ("us-sanjose-1",     "美国西部圣何塞",         "San Jose, USA",                 "objectstorage.us-sanjose-1.oraclecloud.com"),
    ("us-chicago-1",     "美国中西部芝加哥",       "Chicago, USA",                  "objectstorage.us-chicago-1.oraclecloud.com"),
    ("ca-toronto-1",     "加拿大东南部多伦多",     "Toronto, Canada",               "objectstorage.ca-toronto-1.oraclecloud.com"),
    ("ca-montreal-1",    "加拿大东南部蒙特利尔",   "Montreal, Canada",              "objectstorage.ca-montreal-1.oraclecloud.com"),
    ("mx-monterrey-1",   "墨西哥东北部蒙特雷",     "Monterrey, Mexico",             "objectstorage.mx-monterrey-1.oraclecloud.com"),
    ("mx-queretaro-1",   "墨西哥中部克雷塔罗",     "Querétaro, Mexico",             "objectstorage.mx-queretaro-1.oraclecloud.com"),
    # -- South America --
    ("sa-saopaulo-1",    "巴西东部圣保罗",         "São Paulo, Brazil",             "objectstorage.sa-saopaulo-1.oraclecloud.com"),
    ("sa-vinhedo-1",     "巴西南部维涅杜",         "Vinhedo, Brazil",               "objectstorage.sa-vinhedo-1.oraclecloud.com"),
    ("sa-santiago-1",    "智利中部圣地亚哥",       "Santiago, Chile",               "objectstorage.sa-santiago-1.oraclecloud.com"),
    ("sa-valparaiso-1",  "智利西部瓦尔帕莱索",     "Valparaíso, Chile",             "objectstorage.sa-valparaiso-1.oraclecloud.com"),
    ("sa-bogota-1",      "哥伦比亚中部波哥大",     "Bogotá, Colombia",              "objectstorage.sa-bogota-1.oraclecloud.com"),
    # -- Europe --
    ("uk-london-1",      "英国南部伦敦",           "London, UK",                    "objectstorage.uk-london-1.oraclecloud.com"),
    ("uk-cardiff-1",     "英国西部加的夫",         "Cardiff, UK",                   "objectstorage.uk-cardiff-1.oraclecloud.com"),
    ("eu-paris-1",       "法国中部巴黎",           "Paris, France",                 "objectstorage.eu-paris-1.oraclecloud.com"),
    ("eu-marseille-1",   "法国南部马赛",           "Marseille, France",             "objectstorage.eu-marseille-1.oraclecloud.com"),
    ("eu-frankfurt-1",   "德国中部法兰克福",       "Frankfurt, Germany",            "objectstorage.eu-frankfurt-1.oraclecloud.com"),
    ("eu-zurich-1",      "瑞士北部苏黎世",         "Zurich, Switzerland",           "objectstorage.eu-zurich-1.oraclecloud.com"),
    ("eu-turin-1",       "意大利北部都灵",         "Turin, Italy",                  "objectstorage.eu-turin-1.oraclecloud.com"),
    ("eu-milan-1",       "意大利西北部米兰",       "Milan, Italy",                  "objectstorage.eu-milan-1.oraclecloud.com"),
    ("eu-madrid-1",      "西班牙中部马德里",       "Madrid, Spain",                 "objectstorage.eu-madrid-1.oraclecloud.com"),
    ("eu-madrid-3",      "西班牙中部马德里3",      "Madrid 3, Spain",               "objectstorage.eu-madrid-3.oraclecloud.com"),
    ("il-jerusalem-1",   "以色列中部耶路撒冷",     "Jerusalem, Israel",             "objectstorage.il-jerusalem-1.oraclecloud.com"),
    ("eu-stockholm-1",   "瑞典中部斯德哥尔摩",     "Stockholm, Sweden",             "objectstorage.eu-stockholm-1.oraclecloud.com"),
    ("eu-amsterdam-1",   "荷兰西北部阿姆斯特丹",   "Amsterdam, Netherlands",        "objectstorage.eu-amsterdam-1.oraclecloud.com"),
    # -- Middle East --
    ("me-dubai-1",       "阿联酋迪拜",             "Dubai, UAE",                    "objectstorage.me-dubai-1.oraclecloud.com"),
    ("me-abudhabi-1",    "阿联酋阿布扎比",         "Abu Dhabi, UAE",                "objectstorage.me-abudhabi-1.oraclecloud.com"),
    ("me-jeddah-1",      "沙特阿拉伯西部吉达",     "Jeddah, Saudi Arabia",          "objectstorage.me-jeddah-1.oraclecloud.com"),
    ("me-riyadh-1",      "沙特阿拉伯中部利雅得",   "Riyadh, Saudi Arabia",          "objectstorage.me-riyadh-1.oraclecloud.com"),
    # -- Africa --
    ("af-johannesburg-1","南非中部约翰内斯堡",     "Johannesburg, South Africa",    "objectstorage.af-johannesburg-1.oraclecloud.com"),
]

# Region grouping (order matters for display)
REGION_GROUPS = [
    ("亚太地区 (APAC)",      "Asia-Pacific",      ("ap",),           ["#3366CC", "#4477DD", "#5588EE", "#6699FF", "#77AAFF", "#88BBFF", "#99CCFF", "#AADDFF", "#BBEEFF", "#CCFFFF", "#DDEEFF"]),
    ("北美地区 (NA)",        "North America",     ("us", "ca", "mx"),["#CC3333", "#DD4444", "#EE5555", "#FF6666", "#FF7777", "#FF8888", "#FF9999", "#FFAAAA"]),
    ("南美地区 (SA)",        "South America",     ("sa",),           ["#33CC33", "#44DD44", "#55EE55", "#66FF66", "#77FF77"]),
    ("欧洲地区 (EU)",        "Europe",            ("uk", "eu", "il"),["#CC33CC", "#DD44DD", "#EE55EE", "#FF66FF", "#CC77CC", "#DD88DD", "#EE99EE", "#FFAAFF", "#CCBBCC", "#DDCCDD", "#CC33AA", "#DD44BB", "#EE55CC"]),
    ("中东地区 (ME)",        "Middle East",       ("me",),           ["#CC6600", "#DD7700", "#EE8800", "#FF9900"]),
    ("非洲地区 (AF)",        "Africa",            ("af",),           ["#996633"]),
]

# ===========================================================================
#  NETWORK PROBE
# ===========================================================================

def tcp_latency(host: str, port: int = PORT, timeout: float = TIMEOUT) -> float | None:
    """Measure TCP handshake latency (ms). Returns None on failure."""
    try:
        t0 = time.perf_counter()
        with socket.create_connection((host, port), timeout=timeout):
            return round((time.perf_counter() - t0) * 1000, 2)
    except Exception:
        return None


def probe_host(host: str, count: int = PROBE_COUNT) -> list[float]:
    """Probe host `count` times, return latencies in ms."""
    latencies = []
    for _ in range(count):
        lat = tcp_latency(host)
        if lat is not None:
            latencies.append(lat)
        time.sleep(0.05)
    return latencies


def trimmed_mean(data: list[float], trim_pct: float = 0.10) -> float | None:
    """Mean after trimming top & bottom `trim_pct` extremes."""
    if len(data) < 3:
        return None
    s = sorted(data)
    k = max(1, int(len(s) * trim_pct))
    trimmed = s[k: len(s) - k]
    return round(statistics.mean(trimmed), 2) if trimmed else None

# ===========================================================================
#  EXCEL BUILDER
# ===========================================================================

# --- Styles ----------------------------------------------------------------
def _styles():
    return dict(
        hfont=Font(name="Microsoft YaHei", bold=True, size=11, color="FFFFFF"),
        hfill=PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid"),
        halign=Alignment(horizontal="center", vertical="center", wrap_text=True),
        calign=Alignment(horizontal="center", vertical="center"),
        lalign=Alignment(horizontal="left", vertical="center"),
        border=Border(left=Side(style="thin"), right=Side(style="thin"),
                       top=Side(style="thin"), bottom=Side(style="thin")),
        green=PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),
        yellow=PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid"),
        red=PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
        title_font=Font(name="Microsoft YaHei", bold=True, size=14, color="1F4E79"),
    )


# --- Common summary sheet writer -------------------------------------------
def _write_summary_sheet(ws, title, results, st, merge_end="L"):
    """Write a summary sheet (main report or per-region)."""
    ws.merge_cells(f"A1:{merge_end}1")
    ws["A1"].value = title
    ws["A1"].font = st["title_font"]
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    headers = [
        "Region Code", "Region Name (CN)", "Region Name (EN)", "Hostname",
        "Sent", "Received", "Lost", "Loss %",
        "Min (ms)", "Max (ms)", "Avg (ms)", "Trimmed Avg (ms)",
    ]
    widths = [18, 24, 28, 48, 8, 10, 8, 10, 12, 12, 12, 18]

    for c, (h, w) in enumerate(zip(headers, widths), 1):
        cell = ws.cell(row=2, column=c, value=h)
        cell.font = st["hfont"]; cell.fill = st["hfill"]
        cell.alignment = st["halign"]; cell.border = st["border"]
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.row_dimensions[2].height = 28

    for r, rec in enumerate(results, 3):
        vals = [
            rec["region"], rec["name_cn"], rec["name_en"], rec["host"],
            rec["sent"], rec["received"], rec["lost"], rec["loss_pct"],
            rec["min_ms"], rec["max_ms"], rec["avg_ms"], rec["trimmed_avg_ms"],
        ]
        for c, val in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c, value=val if val is not None else "N/A")
            cell.border = st["border"]
            cell.alignment = st["lalign"] if c <= 4 else st["calign"]

        lp = rec["loss_pct"]
        lc = ws.cell(row=r, column=8)
        lc.fill = st["green"] if lp == 0 else (st["yellow"] if lp < 50 else st["red"])

        avg = rec["avg_ms"]
        ac = ws.cell(row=r, column=11)
        if avg is not None:
            ac.fill = st["green"] if avg < 100 else (st["yellow"] if avg < 300 else st["red"])

    ws.freeze_panes = "A3"
    if results:
        ws.auto_filter.ref = f"A2:{merge_end}{2 + len(results)}"


# --- Raw data sheet --------------------------------------------------------
def _write_raw_data_sheet(wb, results, st):
    ws = wb.create_sheet("Raw Data")
    n = PROBE_COUNT
    last_col = get_column_letter(4 + n)

    ws.merge_cells(f"A1:{last_col}1")
    ws["A1"].value = f"Raw Probe Data — {n} TCP connections per endpoint (ms)"
    ws["A1"].font = st["title_font"]
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    all_h = ["Region Code", "Region Name (CN)", "Region Name (EN)", "Hostname"] + [f"P{i}" for i in range(1, n + 1)]
    for c, h in enumerate(all_h, 1):
        cell = ws.cell(row=2, column=c, value=h)
        cell.font = st["hfont"]; cell.fill = st["hfill"]
        cell.alignment = st["halign"]; cell.border = st["border"]
    ws.column_dimensions["A"].width = 18; ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 28; ws.column_dimensions["D"].width = 48
    for c in range(5, 5 + n):
        ws.column_dimensions[get_column_letter(c)].width = 9
    ws.row_dimensions[2].height = 28

    for r, rec in enumerate(results, 3):
        for c, val in enumerate([rec["region"], rec["name_cn"], rec["name_en"], rec["host"]], 1):
            cell = ws.cell(row=r, column=c, value=val)
            cell.border = st["border"]; cell.alignment = st["lalign"]
        lats = rec.get("latencies", [])
        for i in range(n):
            val = lats[i] if i < len(lats) else None
            cell = ws.cell(row=r, column=5 + i, value=val if val is not None else "N/A")
            cell.border = st["border"]; cell.alignment = st["calign"]
            if val is not None:
                cell.fill = st["green"] if val < 100 else (st["yellow"] if val < 300 else st["red"])

    ws.freeze_panes = "E3"
    if results:
        ws.auto_filter.ref = f"A2:{last_col}{2 + len(results)}"


# --- Build entire Excel workbook -------------------------------------------
def build_excel(results: list[dict], output_path: Path):
    wb = Workbook()
    st = _styles()

    # 1. Main report
    ws = wb.active
    ws.title = "OCI TCP Latency Report"
    _write_summary_sheet(ws,
        f"Oracle Cloud Object Storage — TCP Latency Report (port {PORT}, {PROBE_COUNT} probes each)",
        results, st, merge_end="L")

    # 2-7. Per-region detail
    for sheet_name, _, prefixes, _ in REGION_GROUPS:
        region_results = [r for r in results if r["region"].split("-")[0] in prefixes]
        ws_r = wb.create_sheet(sheet_name)
        _write_summary_sheet(ws_r,
            f"{sheet_name} — Oracle Cloud Object Storage TCP Latency",
            region_results, st, merge_end="L")

    # 8. Raw data
    _write_raw_data_sheet(wb, results, st)

    # 9. Summary by Region (last)
    _write_region_summary_sheet(wb, results, st)

    wb.save(output_path)
    return output_path


def _write_region_summary_sheet(wb, results, st):
    ws = wb.create_sheet("Summary by Region")
    region_map = {
        "ap": "Asia-Pacific", "us": "North America", "ca": "North America", "mx": "North America",
        "sa": "South America",
        "uk": "Europe", "eu": "Europe", "il": "Europe",
        "me": "Middle East",
        "af": "Africa",
    }
    regions_order = ["Asia-Pacific", "North America", "South America", "Europe", "Middle East", "Africa"]

    summary = {r: [] for r in regions_order}
    for rec in results:
        grp = region_map.get(rec["region"].split("-")[0], "Other")
        if rec["avg_ms"] is not None:
            summary[grp].append(rec["avg_ms"])

    ws.merge_cells("A1:D1")
    ws["A1"].value = "Summary by Geographic Region"
    ws["A1"].font = st["title_font"]
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    for c, (h, w) in enumerate(zip(
        ["Region", "Endpoints Tested", "Best Avg (ms)", "Worst Avg (ms)"],
        [24, 20, 18, 18]
    ), 1):
        cell = ws.cell(row=2, column=c, value=h)
        cell.font = st["hfont"]; cell.fill = st["hfill"]
        cell.alignment = st["halign"]; cell.border = st["border"]
        ws.column_dimensions[get_column_letter(c)].width = w

    for r, region in enumerate(regions_order, 3):
        avgs = summary[region]
        for c, val in enumerate([
            region, len(avgs),
            round(min(avgs), 2) if avgs else "N/A",
            round(max(avgs), 2) if avgs else "N/A",
        ], 1):
            cell = ws.cell(row=r, column=c, value=val)
            cell.border = st["border"]; cell.alignment = st["calign"]

    ws.freeze_panes = "A3"

# ===========================================================================
#  INTERACTIVE HTML CHART (Plotly.js via CDN)
# ===========================================================================

def build_chart_html(results: list[dict], output_path: Path):
    """Generate a self-contained HTML file with an interactive Plotly.js chart."""

    # Build data structure for the chart
    chart_data = []
    for sheet_name, region_en, prefixes, colors in REGION_GROUPS:
        endpoints_in_region = []
        region_results = [r for r in results if r["region"].split("-")[0] in prefixes]
        for i, rec in enumerate(region_results):
            color = colors[i % len(colors)]
            endpoints_in_region.append({
                "id": rec["region"],
                "label": f"{rec['name_en']} ({rec['region']})",
                "label_cn": rec["name_cn"],
                "color": color,
                "latencies": rec.get("latencies", []),
            })
        chart_data.append({
            "region_name": sheet_name,
            "region_en": region_en,
            "endpoints": endpoints_in_region,
        })

    data_json = json.dumps(chart_data, ensure_ascii=False, indent=2)
    probe_count = PROBE_COUNT

    html = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>OCI Latency Chart — TCP Probe Results</title>
<script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
<style>
  * {{ margin: 0; padding: 0; box-sizing: border-box; }}
  body {{ font-family: "Segoe UI", "Microsoft YaHei", sans-serif; background: #f5f7fa; color: #333; }}
  .header {{ background: linear-gradient(135deg, #1a3a5c 0%, #2d6ca2 100%); color: #fff; padding: 20px 32px; }}
  .header h1 {{ font-size: 22px; font-weight: 600; }}
  .header p {{ font-size: 13px; opacity: 0.85; margin-top: 4px; }}
  .layout {{ display: flex; height: calc(100vh - 78px); }}
  .sidebar {{ width: 320px; min-width: 320px; background: #fff; border-right: 1px solid #e0e4e8; overflow-y: auto; padding: 12px 0; }}
  .sidebar h3 {{ font-size: 13px; color: #888; text-transform: uppercase; letter-spacing: 1px; padding: 8px 16px; margin-top: 4px; }}
  .region-group {{ border-bottom: 1px solid #f0f0f0; }}
  .region-header {{ display: flex; align-items: center; padding: 10px 16px; cursor: pointer; user-select: none; font-size: 14px; font-weight: 600; transition: background 0.15s; }}
  .region-header:hover {{ background: #f8f9fb; }}
  .region-header .arrow {{ font-size: 10px; margin-right: 8px; transition: transform 0.2s; width: 14px; text-align: center; }}
  .region-header .arrow.open {{ transform: rotate(90deg); }}
  .region-header input[type=checkbox] {{ margin-right: 8px; transform: scale(1.1); accent-color: #4472C4; }}
  .region-header .badge {{ font-size: 11px; background: #e8ecf1; color: #666; border-radius: 10px; padding: 2px 8px; margin-left: auto; }}
  .endpoint-list {{ display: none; padding: 0 0 8px 0; }}
  .endpoint-list.open {{ display: block; }}
  .endpoint-row {{ display: flex; align-items: center; padding: 5px 16px 5px 48px; font-size: 12px; cursor: pointer; transition: background 0.1s; }}
  .endpoint-row:hover {{ background: #f8f9fb; }}
  .endpoint-row .color-dot {{ width: 10px; height: 10px; border-radius: 50%; margin-right: 8px; flex-shrink: 0; }}
  .endpoint-row input[type=checkbox] {{ margin-right: 8px; transform: scale(0.95); accent-color: #4472C4; }}
  .endpoint-row .cn-name {{ color: #999; margin-left: 6px; font-size: 11px; }}
  .main {{ flex: 1; display: flex; flex-direction: column; }}
  .toolbar {{ display: flex; gap: 8px; padding: 10px 16px; background: #fff; border-bottom: 1px solid #e0e4e8; align-items: center; }}
  .toolbar button {{ font-size: 12px; padding: 6px 14px; border: 1px solid #d0d4d8; border-radius: 4px; background: #fff; cursor: pointer; transition: all 0.15s; }}
  .toolbar button:hover {{ background: #f0f3f7; border-color: #b0b4b8; }}
  .toolbar button.primary {{ background: #4472C4; color: #fff; border-color: #4472C4; }}
  .toolbar button.primary:hover {{ background: #3a62a8; }}
  .chart-container {{ flex: 1; min-height: 0; }}
  #chart {{ width: 100%; height: 100%; }}
</style>
</head>
<body>
<div class="header">
  <h1>OCI Object Storage — TCP Latency Probe Chart</h1>
  <p>{probe_count} TCP connections per endpoint · port 443 · interactive legend with region/endpoint toggles</p>
</div>
<div class="layout">
  <div class="sidebar" id="sidebar">
    <h3>Toggle Regions</h3>
    <div id="region-list"></div>
  </div>
  <div class="main">
    <div class="toolbar">
      <button class="primary" onclick="selectAll()">Select All</button>
      <button onclick="deselectAll()">Deselect All</button>
      <button onclick="selectBest()">Top 5 Fastest</button>
      <button onclick="selectWorst()">Top 5 Slowest</button>
      <button onclick="resetZoom()">Reset Zoom</button>
    </div>
    <div class="chart-container"><div id="chart"></div></div>
  </div>
</div>

<script>
const CHART_DATA = {data_json};
const PROBE_COUNT = {probe_count};

// ---- State ----
const expandedRegions = new Set(CHART_DATA.map(r => r.region_name));
const seriesVisible = {{}};  // endpointId -> bool

// ---- Build sidebar ----
const regionList = document.getElementById('region-list');
CHART_DATA.forEach((region, ri) => {{
  const div = document.createElement('div');
  div.className = 'region-group';
  const isOpen = expandedRegions.has(region.region_name);

  div.innerHTML = `
    <div class="region-header" onclick="toggleRegion('${{region.region_name}}')">
      <span class="arrow${{isOpen ? ' open' : ''}}" id="arrow_${{ri}}">&#9654;</span>
      <input type="checkbox" checked onclick="toggleRegionEndpoints('${{region.region_name}}', this.checked); event.stopPropagation();">
      <span>${{region.region_name}}</span>
      <span class="badge">${{region.endpoints.length}}</span>
    </div>
    <div class="endpoint-list${{isOpen ? ' open' : ''}}" id="list_${{ri}}">
      ${{region.endpoints.map((ep, ei) => {{
        seriesVisible[ep.id] = true;
        return `
          <div class="endpoint-row" onclick="toggleEndpoint('${{ep.id}}'); event.stopPropagation();">
            <span class="color-dot" style="background:${{ep.color}}"></span>
            <input type="checkbox" checked onclick="toggleEndpoint('${{ep.id}}'); event.stopPropagation();">
            <span>${{ep.label}}</span>
            <span class="cn-name">${{ep.label_cn}}</span>
          </div>`;
      }}).join('')}}
    </div>`;
  regionList.appendChild(div);
}});

// ---- Plotly chart ----
const traces = [];
CHART_DATA.forEach(region => {{
  region.endpoints.forEach(ep => {{
    const x = Array.from({{length: ep.latencies.length}}, (_, i) => i + 1);
    const y = ep.latencies;
    traces.push({{
      x: x,
      y: y,
      type: 'scatter',
      mode: 'lines+markers',
      name: ep.label + ' (' + ep.label_cn + ')',
      line: {{ color: ep.color, width: 1.5 }},
      marker: {{ size: 3, color: ep.color }},
      legendgroup: region.region_name,
      legendgrouptitle: {{ text: region.region_name }},
      hovertemplate: '<b>%{{fullData.name}}</b><br>Probe %{{x}}<br>Latency: <b>%{{y:.1f}} ms</b><extra></extra>',
      visible: true,
      meta: {{ endpointId: ep.id, regionName: region.region_name }}
    }});
  }});
}});

const layout = {{
  xaxis: {{
    title: 'Probe Number',
    dtick: 5,
    gridcolor: '#e8e8e8',
    zeroline: false,
    range: [0.5, PROBE_COUNT + 0.5]
  }},
  yaxis: {{
    title: 'Latency (ms)',
    gridcolor: '#e8e8e8',
    zeroline: false,
    rangemode: 'tozero'
  }},
  legend: {{
    groupclick: 'toggleitem',
    itemclick: 'toggle',
    tracegroupgap: 12,
    font: {{ size: 11 }},
    x: 1.02, xanchor: 'left',
    y: 1, yanchor: 'top',
    bgcolor: 'rgba(255,255,255,0.9)',
    bordercolor: '#e0e0e0',
    borderwidth: 1
  }},
  margin: {{ l: 60, r: 20, t: 20, b: 50 }},
  paper_bgcolor: '#fff',
  plot_bgcolor: '#fafbfc',
  hovermode: 'closest',
  dragmode: 'zoom'
}};

const config = {{
  displayModeBar: true,
  modeBarButtonsToRemove: ['lasso2d', 'select2d'],
  displaylogo: false,
  responsive: true
}};

Plotly.newPlot('chart', traces, layout, config);

// ---- Interactions ----
function toggleRegion(name) {{
  const idx = CHART_DATA.findIndex(r => r.region_name === name);
  const list = document.getElementById('list_' + idx);
  const arrow = document.getElementById('arrow_' + idx);
  const isOpen = list.classList.toggle('open');
  arrow.classList.toggle('open', isOpen);
}}

function toggleRegionEndpoints(regionName, visible) {{
  const updates = [];
  traces.forEach((t, i) => {{
    if (t.meta.regionName === regionName) {{
      updates.push({{visible: visible}});
      seriesVisible[t.meta.endpointId] = visible;
    }}
  }});
  if (updates.length > 0) {{
    const traceIndices = traces.map((t, i) => t.meta.regionName === regionName ? i : -1).filter(i => i >= 0);
    Plotly.restyle('chart', {{visible: visible}}, traceIndices);
  }}
  syncCheckboxes();
}}

function toggleEndpoint(endpointId) {{
  const newVal = !seriesVisible[endpointId];
  seriesVisible[endpointId] = newVal;
  const traceIdx = traces.findIndex(t => t.meta.endpointId === endpointId);
  if (traceIdx >= 0) {{
    Plotly.restyle('chart', {{visible: newVal}}, [traceIdx]);
  }}
  syncCheckboxes();
}}

function syncCheckboxes() {{
  // Update endpoint checkboxes
  document.querySelectorAll('.endpoint-row input[type=checkbox]').forEach(cb => {{
    const row = cb.closest('.endpoint-row');
    const span = row.querySelector('span:nth-child(3)');
    if (span) {{
      const epId = traces.find(t => t.name.startsWith(span.textContent));
      if (epId) cb.checked = seriesVisible[epId.meta.endpointId] || false;
    }}
  }});
}}

function selectAll() {{
  Object.keys(seriesVisible).forEach(k => seriesVisible[k] = true);
  Plotly.restyle('chart', {{visible: true}}, traces.map((_, i) => i));
  document.querySelectorAll('input[type=checkbox]').forEach(cb => cb.checked = true);
}}

function deselectAll() {{
  Object.keys(seriesVisible).forEach(k => seriesVisible[k] = false);
  Plotly.restyle('chart', {{visible: false}}, traces.map((_, i) => i));
  document.querySelectorAll('input[type=checkbox]').forEach(cb => cb.checked = false);
}}

function selectBest() {{
  // Find 5 endpoints with lowest average latency
  const avgs = traces.map((t, i) => ({{
    idx: i, id: t.meta.endpointId,
    avg: t.y.reduce((a, b) => a + b, 0) / t.y.length
  }}));
  avgs.sort((a, b) => a.avg - b.avg);
  const top5 = new Set(avgs.slice(0, 5).map(a => a.id));
  // Hide all, show top 5
  traces.forEach(t => {{
    seriesVisible[t.meta.endpointId] = top5.has(t.meta.endpointId);
  }});
  Plotly.restyle('chart', {{visible: traces.map(t => top5.has(t.meta.endpointId))}});
  syncCheckboxes();
}}

function selectWorst() {{
  const avgs = traces.map((t, i) => ({{
    idx: i, id: t.meta.endpointId,
    avg: t.y.reduce((a, b) => a + b, 0) / t.y.length
  }}));
  avgs.sort((a, b) => b.avg - a.avg);
  const top5 = new Set(avgs.slice(0, 5).map(a => a.id));
  traces.forEach(t => {{
    seriesVisible[t.meta.endpointId] = top5.has(t.meta.endpointId);
  }});
  Plotly.restyle('chart', {{visible: traces.map(t => top5.has(t.meta.endpointId))}});
  syncCheckboxes();
}}

function resetZoom() {{
  Plotly.relayout('chart', {{'xaxis.autorange': true, 'yaxis.autorange': true}});
}}

// Sync checkboxes on legend click
document.getElementById('chart').on('plotly_legendclick', function() {{
  setTimeout(() => {{
    const visibleMap = {{}};
    traces.forEach(t => {{
      visibleMap[t.meta.endpointId] = t.visible !== false;
    }});
    Object.assign(seriesVisible, visibleMap);
    syncCheckboxes();
  }}, 100);
}});
</script>
</body>
</html>'''

    output_path.write_text(html, encoding="utf-8")
    return output_path

# ===========================================================================
#  MAIN
# ===========================================================================

def main():
    results = []
    total = len(ENDPOINTS)
    print(f"Oracle Cloud OCI Latency Probe — {total} endpoints, {PROBE_COUNT} probes each\n")

    for idx, (region, name_cn, name_en, host) in enumerate(ENDPOINTS, 1):
        print(f"[{idx:2d}/{total}] Probing {name_en} ({region}) -> {host}:{PORT}  ...", end=" ", flush=True)
        latencies = probe_host(host, count=PROBE_COUNT)
        received = len(latencies)
        lost = PROBE_COUNT - received

        if received == 0:
            print("ALL FAILED")
            results.append({
                "region": region, "name_cn": name_cn, "name_en": name_en, "host": host,
                "sent": PROBE_COUNT, "received": 0, "lost": lost, "loss_pct": 100.0,
                "min_ms": None, "max_ms": None, "avg_ms": None, "trimmed_avg_ms": None,
                "latencies": [],
            })
            continue

        min_lat = round(min(latencies), 2)
        max_lat = round(max(latencies), 2)
        avg_lat = round(statistics.mean(latencies), 2)
        trim_avg = trimmed_mean(latencies)

        print(f"min={min_lat}ms  max={max_lat}ms  avg={avg_lat}ms  trimmed_avg={trim_avg}ms")

        results.append({
            "region": region, "name_cn": name_cn, "name_en": name_en, "host": host,
            "sent": PROBE_COUNT, "received": received, "lost": lost,
            "loss_pct": round(lost / PROBE_COUNT * 100, 1),
            "min_ms": min_lat, "max_ms": max_lat, "avg_ms": avg_lat, "trimmed_avg_ms": trim_avg,
            "latencies": latencies,
        })

    # Generate Excel
    xlsx_path = Path.cwd() / "oci_latency_report.xlsx"
    build_excel(results, xlsx_path)
    print(f"\nExcel report saved to: {xlsx_path}")

    # Generate interactive HTML chart
    html_path = Path.cwd() / "oci_latency_chart.html"
    build_chart_html(results, html_path)
    print(f"Interactive chart saved to: {html_path}")

    print("\nDone!")


if __name__ == "__main__":
    main()
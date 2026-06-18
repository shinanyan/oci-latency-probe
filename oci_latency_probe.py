#!/usr/bin/env python3
"""
Oracle Cloud Object Storage — TCP Latency Probe

Measures TCP handshake latency (port 443) to all 42 Oracle Cloud
Object Storage endpoints worldwide. Outputs a formatted Excel report
with min / max / average / trimmed-average (10% extremes removed),
plus a raw-data sheet with all 30 probe values per endpoint.

Usage:
    python oci_latency_probe.py
"""

import socket
import time
import statistics
import sys
from pathlib import Path

# Fix Windows console encoding for special characters
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

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

# ---------------------------------------------------------------------------
# All 42 Oracle Cloud Object Storage public endpoints
#   (region_code, name_cn, name_en, host)
# ---------------------------------------------------------------------------
ENDPOINTS = [
    # Asia-Pacific
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
    # North America
    ("us-ashburn-1",     "美国东部阿什本",         "Ashburn, USA",                  "objectstorage.us-ashburn-1.oraclecloud.com"),
    ("us-phoenix-1",     "美国西部凤凰城",         "Phoenix, USA",                  "objectstorage.us-phoenix-1.oraclecloud.com"),
    ("us-sanjose-1",     "美国西部圣何塞",         "San Jose, USA",                 "objectstorage.us-sanjose-1.oraclecloud.com"),
    ("us-chicago-1",     "美国中西部芝加哥",       "Chicago, USA",                  "objectstorage.us-chicago-1.oraclecloud.com"),
    ("ca-toronto-1",     "加拿大东南部多伦多",     "Toronto, Canada",               "objectstorage.ca-toronto-1.oraclecloud.com"),
    ("ca-montreal-1",    "加拿大东南部蒙特利尔",   "Montreal, Canada",              "objectstorage.ca-montreal-1.oraclecloud.com"),
    ("mx-monterrey-1",   "墨西哥东北部蒙特雷",     "Monterrey, Mexico",             "objectstorage.mx-monterrey-1.oraclecloud.com"),
    ("mx-queretaro-1",   "墨西哥中部克雷塔罗",     "Querétaro, Mexico",             "objectstorage.mx-queretaro-1.oraclecloud.com"),
    # South America
    ("sa-saopaulo-1",    "巴西东部圣保罗",         "São Paulo, Brazil",             "objectstorage.sa-saopaulo-1.oraclecloud.com"),
    ("sa-vinhedo-1",     "巴西南部维涅杜",         "Vinhedo, Brazil",               "objectstorage.sa-vinhedo-1.oraclecloud.com"),
    ("sa-santiago-1",    "智利中部圣地亚哥",       "Santiago, Chile",               "objectstorage.sa-santiago-1.oraclecloud.com"),
    ("sa-valparaiso-1",  "智利西部瓦尔帕莱索",     "Valparaíso, Chile",             "objectstorage.sa-valparaiso-1.oraclecloud.com"),
    ("sa-bogota-1",      "哥伦比亚中部波哥大",     "Bogotá, Colombia",              "objectstorage.sa-bogota-1.oraclecloud.com"),
    # Europe
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
    # Middle East
    ("me-dubai-1",       "阿联酋迪拜",             "Dubai, UAE",                    "objectstorage.me-dubai-1.oraclecloud.com"),
    ("me-abudhabi-1",    "阿联酋阿布扎比",         "Abu Dhabi, UAE",                "objectstorage.me-abudhabi-1.oraclecloud.com"),
    ("me-jeddah-1",      "沙特阿拉伯西部吉达",     "Jeddah, Saudi Arabia",          "objectstorage.me-jeddah-1.oraclecloud.com"),
    ("me-riyadh-1",      "沙特阿拉伯中部利雅得",   "Riyadh, Saudi Arabia",          "objectstorage.me-riyadh-1.oraclecloud.com"),
    # Africa
    ("af-johannesburg-1","南非中部约翰内斯堡",     "Johannesburg, South Africa",    "objectstorage.af-johannesburg-1.oraclecloud.com"),
]

PROBE_COUNT = 30
PORT = 443
TIMEOUT = 5.0  # seconds per TCP connect attempt


def tcp_latency(host: str, port: int = PORT, timeout: float = TIMEOUT) -> float | None:
    """Measure TCP handshake latency to host:port in milliseconds. Returns None on failure."""
    try:
        t0 = time.perf_counter()
        with socket.create_connection((host, port), timeout=timeout):
            elapsed = (time.perf_counter() - t0) * 1000
        return round(elapsed, 2)
    except Exception:
        return None


def probe_host(host: str, count: int = PROBE_COUNT) -> list[float]:
    """Probe a host `count` times, return list of latencies in ms."""
    latencies = []
    for _ in range(count):
        lat = tcp_latency(host)
        if lat is not None:
            latencies.append(lat)
        time.sleep(0.05)
    return latencies


def trimmed_mean(data: list[float], trim_pct: float = 0.10) -> float | None:
    """
    Mean after trimming the top `trim_pct` and bottom `trim_pct` as extreme values.
    e.g. trim_pct=0.10 → remove top 10% and bottom 10%, average the middle 80%.
    """
    if len(data) < 3:
        return None
    sorted_data = sorted(data)
    n = len(sorted_data)
    k = max(1, int(n * trim_pct))
    trimmed = sorted_data[k : n - k]
    if not trimmed:
        return None
    return round(statistics.mean(trimmed), 2)


# ---------------------------------------------------------------------------
# Shared styles
# ---------------------------------------------------------------------------
def _make_styles():
    return dict(
        header_font=Font(name="Microsoft YaHei", bold=True, size=11, color="FFFFFF"),
        header_fill=PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid"),
        header_align=Alignment(horizontal="center", vertical="center", wrap_text=True),
        cell_align=Alignment(horizontal="center", vertical="center"),
        cell_align_left=Alignment(horizontal="left", vertical="center"),
        thin_border=Border(
            left=Side(style="thin"), right=Side(style="thin"),
            top=Side(style="thin"), bottom=Side(style="thin"),
        ),
        green_fill=PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),
        yellow_fill=PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid"),
        red_fill=PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
    )


# ---------------------------------------------------------------------------
# Common sheet header + data writer
# ---------------------------------------------------------------------------
def _write_summary_sheet(ws, title, results, s, merge_end="L"):
    """Write a summary-style sheet (main report or per-region)."""
    ws.merge_cells(f"A1:{merge_end}1")
    ws["A1"].value = title
    ws["A1"].font = Font(name="Microsoft YaHei", bold=True, size=14, color="1F4E79")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    headers = [
        "Region Code", "Region Name (CN)", "Region Name (EN)", "Hostname",
        "Sent", "Received", "Lost", "Loss %",
        "Min (ms)", "Max (ms)", "Avg (ms)", "Trimmed Avg (ms)",
    ]
    col_widths = [18, 24, 28, 48, 8, 10, 8, 10, 12, 12, 12, 18]

    for c, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=2, column=c, value=h)
        cell.font = s["header_font"]
        cell.fill = s["header_fill"]
        cell.alignment = s["header_align"]
        cell.border = s["thin_border"]
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
            cell.border = s["thin_border"]
            cell.alignment = s["cell_align_left"] if c <= 4 else s["cell_align"]

        # Color-code loss %
        lp = rec["loss_pct"]
        lc = ws.cell(row=r, column=8)
        if lp == 0:
            lc.fill = s["green_fill"]
        elif lp < 50:
            lc.fill = s["yellow_fill"]
        else:
            lc.fill = s["red_fill"]

        # Color-code avg latency
        avg = rec["avg_ms"]
        ac = ws.cell(row=r, column=11)
        if avg is not None:
            if avg < 100:
                ac.fill = s["green_fill"]
            elif avg < 300:
                ac.fill = s["yellow_fill"]
            else:
                ac.fill = s["red_fill"]

    ws.freeze_panes = "A3"
    if results:
        ws.auto_filter.ref = f"A2:{merge_end}{2 + len(results)}"


# ---------------------------------------------------------------------------
# Raw Data sheet — all 30 probe values per endpoint
# ---------------------------------------------------------------------------
def _write_raw_data_sheet(wb, results, s):
    ws = wb.create_sheet("Raw Data")

    n_probes = PROBE_COUNT
    last_col = get_column_letter(4 + n_probes)  # e.g. "AH" for 30 probes

    ws.merge_cells(f"A1:{last_col}1")
    ws["A1"].value = f"Raw Probe Data — {n_probes} TCP connections per endpoint (ms)"
    ws["A1"].font = Font(name="Microsoft YaHei", bold=True, size=14, color="1F4E79")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    # Headers
    base_headers = ["Region Code", "Region Name (CN)", "Region Name (EN)", "Hostname"]
    probe_headers = [f"P{i}" for i in range(1, n_probes + 1)]
    all_headers = base_headers + probe_headers

    for c, h in enumerate(all_headers, 1):
        cell = ws.cell(row=2, column=c, value=h)
        cell.font = s["header_font"]
        cell.fill = s["header_fill"]
        cell.alignment = s["header_align"]
        cell.border = s["thin_border"]
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 28
    ws.column_dimensions["D"].width = 48
    for c in range(5, 5 + n_probes):
        ws.column_dimensions[get_column_letter(c)].width = 9
    ws.row_dimensions[2].height = 28

    for r, rec in enumerate(results, 3):
        # Base info
        for c, val in enumerate([
            rec["region"], rec["name_cn"], rec["name_en"], rec["host"],
        ], 1):
            cell = ws.cell(row=r, column=c, value=val)
            cell.border = s["thin_border"]
            cell.alignment = s["cell_align_left"]

        # Probe values
        latencies = rec.get("latencies", [])
        for i in range(n_probes):
            val = latencies[i] if i < len(latencies) else None
            cell = ws.cell(row=r, column=5 + i, value=val if val is not None else "N/A")
            cell.border = s["thin_border"]
            cell.alignment = s["cell_align"]
            if val is not None:
                if val < 100:
                    cell.fill = s["green_fill"]
                elif val < 300:
                    cell.fill = s["yellow_fill"]
                else:
                    cell.fill = s["red_fill"]

    ws.freeze_panes = "E3"
    if results:
        ws.auto_filter.ref = f"A2:{last_col}{2 + len(results)}"


# ---------------------------------------------------------------------------
# Excel report
# ---------------------------------------------------------------------------
def build_excel(results: list[dict], output_path: Path):
    wb = Workbook()
    s = _make_styles()

    # ---- Sheet 1: Main Report ----
    ws = wb.active
    ws.title = "OCI TCP Latency Report"
    _write_summary_sheet(ws,
        f"Oracle Cloud Object Storage — TCP Latency Report (port {PORT}, {PROBE_COUNT} probes each)",
        results, s, merge_end="L")

    # ---- Sheets 2-7: Per-region detail ----
    region_groups = [
        ("亚太地区 (APAC)",      ("ap",)),
        ("北美地区 (NA)",        ("us", "ca", "mx")),
        ("南美地区 (SA)",        ("sa",)),
        ("欧洲地区 (EU)",        ("uk", "eu", "il")),
        ("中东地区 (ME)",        ("me",)),
        ("非洲地区 (AF)",        ("af",)),
    ]

    for sheet_name, prefixes in region_groups:
        region_results = [
            rec for rec in results
            if rec["region"].split("-")[0] in prefixes
        ]
        ws_r = wb.create_sheet(sheet_name)
        _write_summary_sheet(ws_r,
            f"{sheet_name} — Oracle Cloud Object Storage TCP Latency",
            region_results, s, merge_end="L")

    # ---- Sheet 8: Raw Data ----
    _write_raw_data_sheet(wb, results, s)

    # ---- Sheet 9: Summary by Region (last) ----
    ws2 = wb.create_sheet("Summary by Region")
    regions_order = ["Asia-Pacific", "North America", "South America", "Europe", "Middle East", "Africa"]
    region_map = {
        "ap": "Asia-Pacific", "us": "North America", "ca": "North America", "mx": "North America",
        "sa": "South America",
        "uk": "Europe", "eu": "Europe", "il": "Europe",
        "me": "Middle East",
        "af": "Africa",
    }

    summary = {r: [] for r in regions_order}
    for rec in results:
        grp = region_map.get(rec["region"].split("-")[0], "Other")
        if rec["avg_ms"] is not None:
            summary[grp].append(rec["avg_ms"])

    ws2.merge_cells("A1:D1")
    ws2["A1"].value = "Summary by Geographic Region"
    ws2["A1"].font = Font(name="Microsoft YaHei", bold=True, size=14, color="1F4E79")
    ws2["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 32

    for c, (h, w) in enumerate(zip(
        ["Region", "Endpoints Tested", "Best Avg (ms)", "Worst Avg (ms)"],
        [24, 20, 18, 18]
    ), 1):
        cell = ws2.cell(row=2, column=c, value=h)
        cell.font = s["header_font"]
        cell.fill = s["header_fill"]
        cell.alignment = s["header_align"]
        cell.border = s["thin_border"]
        ws2.column_dimensions[get_column_letter(c)].width = w

    for r, region in enumerate(regions_order, 3):
        avgs = summary[region]
        for c, val in enumerate([
            region,
            len(avgs),
            round(min(avgs), 2) if avgs else "N/A",
            round(max(avgs), 2) if avgs else "N/A",
        ], 1):
            cell = ws2.cell(row=r, column=c, value=val)
            cell.border = s["thin_border"]
            cell.alignment = s["cell_align"]

    ws2.freeze_panes = "A3"

    wb.save(output_path)
    return output_path


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    output_path = Path.cwd() / "oci_latency_report.xlsx"
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
                "sent": PROBE_COUNT, "received": 0, "lost": lost,
                "loss_pct": 100.0,
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

    build_excel(results, output_path)
    print(f"\nDone! Report saved to: {output_path}")


if __name__ == "__main__":
    main()
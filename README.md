[English](README.md) | [中文](README_CN.md)

# OCI Latency Probe 🔭

TCP latency measurement tool for all 42 [Oracle Cloud Object Storage](https://www.oracle.com/cloud/storage/object-storage/) public endpoints worldwide. Generates a color-coded Excel report with min / max / average / trimmed-average latency.

## Why TCP instead of ICMP?

Oracle Cloud blocks ICMP (ping) on object storage endpoints. This tool measures **TCP handshake latency to port 443 (HTTPS)**, which reflects real application-level network performance more accurately than ICMP ever would.

## Quick Start

```bash
# 1. Clone
git clone https://github.com/shinanyan/oci-latency-probe.git
cd oci-latency-probe

# 2. Install dependencies
pip install openpyxl

# 3. Run
python oci_latency_probe.py
```

The script probes all 42 endpoints (30 TCP connections each) and writes `oci_latency_report.xlsx` to the current directory.

## Requirements

- Python 3.8+
- [openpyxl](https://openpyxl.readthedocs.io/) (auto-installed if missing)

## Output

The Excel workbook contains two sheets:

### Sheet 1 — OCI TCP Latency Report

| Column | Description |
|--------|-------------|
| Region Code | Oracle region identifier (e.g. `ap-tokyo-1`) |
| Region Name | Human-readable name (e.g. 日本东部东京) |
| Hostname | Object storage endpoint |
| Sent / Received / Lost | Probe counts |
| Loss % | Packet loss rate (color-coded: 🟢 0% / 🟡 <50% / 🔴 ≥50%) |
| Min / Max (ms) | Best and worst single probe |
| Avg (ms) | Arithmetic mean of all probes |
| Trimmed Avg (ms) | Mean after removing top 10% and bottom 10% extremes |

### Sheet 2 — Summary by Region

Aggregated best/worst average latency per geographic region.

## Endpoints Covered

| Region | Endpoints |
|--------|-----------|
| Asia-Pacific | Osaka, Tokyo, Seoul, Chuncheon, Singapore ×2, Mumbai, Hyderabad, Batam, Sydney, Melbourne |
| North America | Ashburn, Phoenix, San Jose, Chicago, Toronto, Montreal, Monterrey, Querétaro |
| South America | São Paulo, Vinhedo, Santiago, Valparaíso, Bogotá |
| Europe | London, Cardiff, Paris, Marseille, Frankfurt, Zurich, Turin, Milan, Madrid ×2, Jerusalem, Stockholm, Amsterdam |
| Middle East | Dubai, Abu Dhabi, Jeddah, Riyadh |
| Africa | Johannesburg |

## Customization

Edit the constants at the top of `oci_latency_probe.py`:

```python
PROBE_COUNT = 30   # probes per endpoint
PORT = 443         # target port
TIMEOUT = 5.0      # seconds per TCP connect attempt
```

## License

MIT — see [LICENSE](LICENSE) for details.
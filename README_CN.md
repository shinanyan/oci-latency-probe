[English](README.md) | [中文](README_CN.md)

# OCI 延迟探测工具 🔭

一键测量全球 42 个 [Oracle Cloud 对象存储](https://www.oracle.com/cloud/storage/object-storage/) 公共端点的 TCP 延迟。生成带颜色标记的 Excel 报告，包含最低延迟、最高延迟、平均延迟和去除极端值后的平均延迟。

## 为什么用 TCP 而不是 ICMP？

Oracle Cloud 的对象存储端点禁用了 ICMP（ping）。本工具通过测量 **TCP 到 443 端口（HTTPS）的握手延迟**，比 ICMP 更准确地反映实际应用层网络性能。

## 快速开始

### 即时探测

```bash
git clone https://github.com/shinanyan/oci-latency-probe.git
cd oci-latency-probe
pip install openpyxl
python oci_latency_probe.py
```

脚本会对全部 42 个端点各进行 30 次 TCP 探测，生成 `oci_latency_report.xlsx` + `oci_latency_chart.html`。

### 长时间监测

持续循环探测，`Ctrl+C` 随时停止，所有已收集数据自动保存并生成报告。

```bash
python oci_monitor.py
```

输出文件：
- `oci_monitor_data.json` — 原始数据，每 5 个 cycle 自动保存（中断后可续传）
- `oci_monitor_report.xlsx` — cycle 汇总 + 各端点长期统计
- `oci_monitor_chart.html` — 时间序列图（X 轴 = cycle 时间戳，Y 轴 = 平均延迟）

## 环境要求

- Python 3.8+
- [openpyxl](https://openpyxl.readthedocs.io/)（缺失时会自动安装）

## 输出说明

Excel 工作簿包含 9 个 Sheet：

| Sheet | 说明 |
|-------|------|
| OCI TCP Latency Report | 全部 42 端点汇总 |
| 亚太地区 (APAC) ~ 非洲地区 (AF) | 六大地域各一个 Sheet |
| Raw Data | 全部 30 次探测原始数据 |
| Summary by Region | 六大区域汇总对比 |

同时生成交互式 HTML 图表（`oci_latency_chart.html`），支持区域/端点勾选、缩放、极端值裁剪等。

### 汇总表列说明

| 列名 | 说明 |
|------|------|
| Region Code | Oracle 区域标识（如 `ap-tokyo-1`） |
| Region Name | 区域中文名称（如 日本东部东京） |
| Hostname | 对象存储端点地址 |
| Sent / Received / Lost | 探测次数统计 |
| Loss % | 失败率（颜色标记：🟢 0% / 🟡 <50% / 🔴 ≥50%） |
| Min / Max (ms) | 单次探测的最低/最高延迟 |
| Avg (ms) | 所有探测的算术平均值 |
| Trimmed Avg (ms) | 去掉最高 10% 和最低 10% 极端值后的平均值 |

## 覆盖端点

| 区域 | 端点 |
|------|------|
| 亚太地区 | 大阪、东京、首尔、春川、新加坡 ×2、孟买、海得拉巴、巴淡、悉尼、墨尔本 |
| 北美地区 | 阿什本、凤凰城、圣何塞、芝加哥、多伦多、蒙特利尔、蒙特雷、克雷塔罗 |
| 南美地区 | 圣保罗、维涅杜、圣地亚哥、瓦尔帕莱索、波哥大 |
| 欧洲地区 | 伦敦、加的夫、巴黎、马赛、法兰克福、苏黎世、都灵、米兰、马德里 ×2、耶路撒冷、斯德哥尔摩、阿姆斯特丹 |
| 中东地区 | 迪拜、阿布扎比、吉达、利雅得 |
| 非洲地区 | 约翰内斯堡 |

## 自定义参数

编辑 `oci_latency_probe.py` 顶部的常量：

```python
PROBE_COUNT = 30   # 每个端点的探测次数
PORT = 443         # 目标端口
TIMEOUT = 5.0      # 每次 TCP 连接的超时秒数
```

## 开源协议

MIT — 详见 [LICENSE](LICENSE)。
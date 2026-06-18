[English](README.md) | [中文](README_CN.md)

# OCI 延迟探测工具 🔭

一键测量全球 42 个 [Oracle Cloud 对象存储](https://www.oracle.com/cloud/storage/object-storage/) 公共端点的 TCP 延迟。生成带颜色标记的 Excel 报告，包含最低延迟、最高延迟、平均延迟和去除极端值后的平均延迟。

## 为什么用 TCP 而不是 ICMP？

Oracle Cloud 的对象存储端点禁用了 ICMP（ping）。本工具通过测量 **TCP 到 443 端口（HTTPS）的握手延迟**，比 ICMP 更准确地反映实际应用层网络性能。

## 快速开始

```bash
# 1. 克隆仓库
git clone https://github.com/shinanyan/oci-latency-probe.git
cd oci-latency-probe

# 2. 安装依赖
pip install openpyxl

# 3. 运行
python oci_latency_probe.py
```

脚本会对全部 42 个端点各进行 30 次 TCP 探测，结果写入当前目录的 `oci_latency_report.xlsx`。

## 环境要求

- Python 3.8+
- [openpyxl](https://openpyxl.readthedocs.io/)（缺失时会自动安装）

## 输出说明

Excel 工作簿包含两个 Sheet：

### Sheet 1 — OCI TCP Latency Report

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

### Sheet 2 — Summary by Region

按六大地理区域汇总，展示每个区域的最佳/最差平均延迟。

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
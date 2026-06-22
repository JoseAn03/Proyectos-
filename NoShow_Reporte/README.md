# NoShow Reporte — Daily No-Show Analysis

[![Python](https://img.shields.io/badge/Python-3.x-3776AB?style=flat&logo=python&logoColor=white)](https://python.org)
[![openpyxl](https://img.shields.io/badge/openpyxl-Excel%20Automation-217346?style=flat&logo=microsoftexcel&logoColor=white)]()

Automated No-Show report generation — processes active reservations, identifies No-Show patterns, and generates a formatted report organized alphabetically by brand.

## Problem

No-Show tracking was non-existent. There was no automated way to identify which customers didn't show up, making it impossible to analyze patterns or calculate revenue loss.

## Solution

### Python Script (Recommended)

```bash
pip install openpyxl
python3 noshow_reporte.py manifiesto.xlsx NoShow_Reporte.xlsx
```

### VBA Macro (Legacy)

1. Open source Excel → `Alt + F11` → Insert → Module
2. Paste code from `NOSHOW_MACRO.bas`
3. Close → `Alt + F8` → Run `GenerarResManifest`

## Features

- **Brand Classification:** Splits reservations by pickup location (Alamo/Enterprise/National)
- **Alphabetical Sorting:** A-Z by customer name within each brand
- **Conditional Formatting:** Brand colors, letter grouping, VIP indicators
- **Executive Summary:** No-Show rates and patterns at a glance

## Files

| File | Description |
|------|-------------|
| [`noshow_reporte.py`](./noshow_reporte.py) | **Python version** (recommended) |
| [`NOSHOW_MACRO.bas`](./NOSHOW_MACRO.bas) | VBA macro — legacy version |

## Tech Stack

- **Python 3.x** with openpyxl
- **VBA** for legacy automation
- **Conditional Formatting** by brand
- **Data Classification** by location codes

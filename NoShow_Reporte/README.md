[![Python](https://img.shields.io/badge/Python-3.x-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://python.org)
[![Excel](https://img.shields.io/badge/Excel-VBA%20Macros-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)]()

# NoShow Reporte — Daily No-Show Analysis

Automated No-Show report generation — processes active reservations, identifies no-show patterns, and produces a formatted report organized by brand with alphabetical grouping.

## The Problem

No-Show tracking was non-existent. There was no automated way to identify customer no-shows or analyze patterns across brands.

## The Solution

Python and VBA scripts that automate daily No-Show reporting with brand classification, VIP detection, and conditional formatting.

### Python Script (Recommended)
```bash
pip install openpyxl
python3 noshow_reporte.py manifiesto.xlsx NoShow_Reporte.xlsx
```

### VBA Macro (Legacy)
Open Excel → `Alt + F11` → Insert Module → Paste code → `Alt + F8` → Run

## Features
- **Brand Classification** by pickup location codes
- **Alphabetical Sorting** A-Z by customer name
- **Conditional Formatting** with brand-specific colors
- **VIP/Expedia Detection** with visual indicators

## Files
| File | Description |
|------|-------------|
| `noshow_reporte.py` | Python automation script (recommended) |
| `NOSHOW_MACRO.bas` | VBA macro (legacy) |

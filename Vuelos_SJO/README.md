[![Python](https://img.shields.io/badge/Python-3.x-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://python.org)
[![Excel](https://img.shields.io/badge/Excel-VBA%20Macros-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)]()

# Vuelos SJO — International Flight Analysis

Processes flight data from **Juan Santamaria International Airport (SJO)**, filters international flights, and generates formatted reports organized by hourly time blocks.

## The Problem

FlightRadar24 data required manual filtering to separate international from domestic flights and reorganize into operational time blocks for airport coordination.

## The Solution

Python script that automatically filters, categorizes, and formats flight data into a professional report.

### Python Script (Recommended)
```bash
pip install openpyxl
python3 vuelos_sjo.py vuelos.xlsx Reporte_Vuelos_SJO.xlsx
```

### VBA Macro (Legacy)
Copy FlightRadar24 data → `Alt + F11` → Insert Module → Paste code → Run

## Features
- **Domestic Exclusion** — filters out domestic routes (Tambor, Tamarindo, Liberia, Quepos)
- **Airline Filtering** — excludes specific airlines (Sansa, LATAM, GetJet, NetJets)
- **Time Blocks** — groups flights into 6 operational windows (05:00-23:59)
- **Professional Format** — zebra striping, corporate headers, auto date stamp

## Files
| File | Description |
|------|-------------|
| `vuelos_sjo.py` | Python automation script (recommended) |
| `VUELOS_MACRO_REPORTE.bas` | VBA macro (legacy) |

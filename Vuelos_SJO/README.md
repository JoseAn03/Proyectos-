# Vuelos SJO — International Flight Analysis

[![Python](https://img.shields.io/badge/Python-3.x-3776AB?style=flat&logo=python&logoColor=white)](https://python.org)
[![openpyxl](https://img.shields.io/badge/openpyxl-Excel%20Automation-217346?style=flat&logo=microsoftexcel&logoColor=white)]()

Processes flight data from **Juan Santamaria International Airport (SJO)**, filters international flights, and generates a formatted report organized by hourly time blocks.

## Problem

Flight data from FlightRadar24 needed manual filtering to separate international from domestic flights and organize them into operational time blocks.

## Solution

### Python Script (Recommended)

```bash
pip install openpyxl
python3 vuelos_sjo.py vuelos.xlsx Reporte_Vuelos_SJO.xlsx
```

### VBA Macro (Legacy)

1. Copy FlightRadar24 data into an Excel sheet
2. `Alt + F11` → Insert → Module
3. Paste code from `VUELOS_MACRO_REPORTE.bas`
4. Close → `Alt + F8` → Run `GenerarReporteVuelos`

## Features

- **Domestic Exclusion:** Filters out domestic routes (Tambor, Tamarindo, Liberia, Quepos, etc.)
- **Airline Filtering:** Excludes specific airlines (Sansa, LATAM, GetJet, NetJets, Viva, Mas Air)
- **Hourly Blocks:** Groups flights into 6 time blocks:
  - 05:00–08:59 | 09:00–11:59 | 12:00–14:59
  - 15:00–17:59 | 18:00–20:59 | 21:00–23:59
- **Professional Format:** Zebra striping, corporate headers, automatic date stamp

## Files

| File | Description |
|------|-------------|
| [`vuelos_sjo.py`](./vuelos_sjo.py) | **Python version** (recommended) |
| [`VUELOS_MACRO_REPORTE.bas`](./VUELOS_MACRO_REPORTE.bas) | VBA macro — legacy version |

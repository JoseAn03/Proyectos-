# GamePlan Reservas — Daily Reservation Processing

[![Python](https://img.shields.io/badge/Python-3.x-3776AB?style=flat&logo=python&logoColor=white)](https://python.org)
[![openpyxl](https://img.shields.io/badge/openpyxl-Excel%20Automation-217346?style=flat&logo=microsoftexcel&logoColor=white)](https://openpyxl.readthedocs.io/)
[![VBA](https://img.shields.io/badge/VBA-Macro-217346?style=flat&logo=microsoftexcel&logoColor=white)]()

Automates the daily **Game Plan** report — processes the reservation manifest and generates a formatted Excel workbook with three brand tabs (Alamo, Enterprise, National), alphabetically sorted with conditional formatting.

## Problem

Processing 500+ daily reservations manually across 3 brands took **2+ hours** of copy-paste, sorting, and formatting. This script reduces it to **~15 minutes**.

## Solution

### Python Script (Recommended)

```bash
pip install openpyxl
python3 gameplan_reservas.py manifiesto.xlsx GamePlan_Procesado.xlsx
```

### VBA Macro (Legacy)

1. Open the source Excel file (ResManifest_.xlsx)
2. Press `Alt + F11` → Insert → Module
3. Paste code from `GAMEPLAN_MACRO_FINAL_2026.bas`
4. Close editor → press `Alt + F8` → Run `ProcesarReservas`

## What It Does

| Step | Description |
|------|-------------|
| **1. Read** | Reads the manifest from the active sheet |
| **2. Classify** | Splits by brand based on pickup location codes: |
| | - **Alamo** → SJOE71 / SJOT71 |
| | - **Enterprise** → SJOC61 / SJOT61 |
| | - **National** → SJOE01 / SJOT01 |
| **3. Sort** | Alphabetically by customer name |
| **4. Group** | A-Z letter grouping with borders and format |
| **5. Format** | Brand colors on names, indicators in ** column |
| **6. Flags** | MAIN VIP / EXPEDIA detection |

## Output Format

| Column | Content |
|--------|---------|
| A: Letra | Letter A-Z (first row of each group) — red bold |
| B: N. Reserva | Reservation number |
| C: Marca | `ALAMO//` / `ENTERPRISE//` / `NATIONAL//` |
| D: ** | `I`/`P`/`E`, `FL`, or combos like `I-FL` |
| E: Nombre | Customer name — brand colored |
| F: Clase | Vehicle class |
| G: **MAIN VIP** | `* MAIN VIP*` or `* EXPEDIA*` indicators |
| H-K: Dates/Locations | Pickup/dropoff dates and locations |
| L: Tarifa Diaria | Daily rate |
| M: Agencia de recom. | Referral agency |

## Files

| File | Description |
|------|-------------|
| [`gameplan_reservas.py`](./gameplan_reservas.py) | **Python version** (recommended) — full automation with openpyxl |
| [`GAMEPLAN_MACRO_FINAL_2026.bas`](./GAMEPLAN_MACRO_FINAL_2026.bas) | VBA macro — legacy Excel automation |
| [`prompt_ia_gameplan.md`](./prompt_ia_gameplan.md) | AI prompt — alternative Claude-assisted processing |

## Tech Stack

- **Python 3.x** with openpyxl library
- **VBA** for legacy Excel automation
- **Conditional Formatting** by brand
- **String Analysis** for VIP/Expedia detection
- **Alphanumeric Sorting** with A-Z grouping

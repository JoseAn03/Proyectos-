[![Python](https://img.shields.io/badge/Python-3.x-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://python.org)
[![Excel](https://img.shields.io/badge/Excel-VBA%20Macros-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)]()

# GamePlan Reservas — Daily Reservation Processing

Automates the daily **Game Plan** report — transforms raw reservation manifests into a formatted Excel workbook with three brand tabs, alphabetical sorting, and conditional formatting.

## The Problem

Processing 500+ daily reservations across 3 brands manually required **2+ hours** of repetitive copy-paste, sorting, and formatting.

## The Solution

A Python script that reduces the process to **~15 minutes** — an 87.5% time savings.

### Python Script (Recommended)
```bash
pip install openpyxl
python3 gameplan_reservas.py manifiesto.xlsx GamePlan_Procesado.xlsx
```

### VBA Macro (Legacy Alternative)
Open Excel → `Alt + F11` → Insert Module → Paste code → `Alt + F8` → Run

## Output Format

| Column | Content |
|--------|---------|
| **A: Letra** | A-Z grouping (red bold on first row per group) |
| **B: N. Reserva** | Reservation ID |
| **C: Marca** | Brand name with color coding |
| **D: ** | Flags (I/P/E, FL, or combined) |
| **E: Nombre** | Customer name (brand-colored) |
| **F: Clase** | Vehicle class |
| **G: MAIN VIP** | VIP or Expedia indicators |
| **H-M** | Dates, locations, rate, agency |

## Files

| File | Description |
|------|-------------|
| `gameplan_reservas.py` | Python automation script (recommended) |
| `GAMEPLAN_MACRO_FINAL_2026.bas` | VBA macro (legacy) |
| `prompt_ia_gameplan.md` | AI prompt for alternative processing |

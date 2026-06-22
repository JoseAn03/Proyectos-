# Proyectos de Automatización y Análisis de Datos

<p align="center">
  <img src="./assets/workflow_diagram.png" alt="Workflow Diagram" width="800">
</p>

<p align="center">
  <strong>Data Pipeline & Analytics</strong> — Juan Santamaria International Airport (SJO)
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Python-3.x-3776AB?style=flat&logo=python&logoColor=white" alt="Python">
  <img src="https://img.shields.io/badge/SQL-MySQL%208.0-4479A1?style=flat&logo=mysql&logoColor=white" alt="MySQL">
  <img src="https://img.shields.io/badge/Power%20BI-Desktop-F2C811?style=flat&logo=powerbi&logoColor=black" alt="Power BI">
  <img src="https://img.shields.io/badge/VBA-Excel-217346?style=flat&logo=microsoftexcel&logoColor=white" alt="VBA">
  <img src="https://img.shields.io/badge/DAX-Measures-1F4E79?style=flat" alt="DAX">
  <img src="https://img.shields.io/badge/openpyxl-Excel%20Automation-217346?style=flat" alt="openpyxl">
</p>

---

## Overview

Portfolio of real-world data analysis projects developed at **Enterprise Holdings** (Alamo, Enterprise, National) at SJO Airport. These scripts automate daily operational reports, analyze No-Show patterns, generate executive dashboards, and process international flight data — turning 2+ hours of manual Excel work into minutes of automated processing.

### Key Results

| Metric | Before | After |
|--------|--------|-------|
| Daily Game Plan generation | 2 hours manual | **15 minutes automated** |
| Reservas procesadas/día | Manual copy-paste | **500+ automated** |
| Marcas consolidadas | 3 archivos separados | **1 archivo, 3 pestañas** |
| No-Show analysis | No existía | **Reporte diario automático** |
| Executive reporting | Solo Game Plan | **Power BI dashboard + SQL queries** |

---

## Projects

| Project | Description | Tech |
|---------|-------------|------|
| [**SQL Portfolio**](./SQL_Portfolio/) | Database schema, daily Game Plan queries, No-Show analysis, executive reports, VIP/Expedia detection | MySQL, CTEs, Window Functions |
| [**GamePlan Reservas**](./GamePlan_Reservas/) | Automated daily reservation processing into formatted 3-tab Excel (Alamo/Enterprise/National) | Python (openpyxl), VBA |
| [**NoShow Reporte**](./NoShow_Reporte/) | Automated No-Show report with brand grouping and conditional formatting | Python (openpyxl), VBA |
| [**Power BI Dashboard**](./PowerBI_Dashboard/) | 4-page dashboard: Executive summary, Brand analysis, Hourly patterns, No-Show tracking | DAX, Power Query, CSV data |
| [**Vuelos SJO**](./Vuelos_SJO/) | International flight processing, airline filtering, hourly block categorization | Python (openpyxl), VBA |
| [**Guías**](./Guias/) | Documentation and AI prompts for report generation | Prompt Engineering |

---

## Technologies

| Category | Tools |
|----------|-------|
| **Languages** | Python, SQL, VBA (Excel Macros) |
| **Libraries** | openpyxl, pandas, matplotlib, numpy |
| **Databases** | MySQL 8.0, MariaDB |
| **BI & Visualization** | Power BI Desktop, DAX, Power Query |
| **Excel** | PivotTables, Conditional Formatting, VBA Macros |
| **Version Control** | Git, GitHub |

---

## Sample SQL Query

```sql
-- Daily Game Plan Report with VIP/Expedia indicators
SELECT 
    ROW_NUMBER() OVER (PARTITION BY m.nombre_marca ORDER BY r.fecha_recogida) AS letra,
    r.numero_reserva,
    CONCAT(m.nombre_marca, '//') AS marca,
    CASE 
        WHEN r.es_vip = TRUE THEN '* MAIN VIP*'
        WHEN r.agencia LIKE '%11617270%' THEN '* EXPEDIA*'
        ELSE ''
    END AS tipo_vip,
    r.nombre_cliente,
    DATE_FORMAT(r.fecha_recogida, '%H:%i') AS hora_recogida
FROM reservas r
JOIN marcas m ON r.marca_id = m.marca_id
WHERE DATE(r.fecha_recogida) = CURDATE() AND r.estado = 'activa'
ORDER BY m.nombre_marca, r.nombre_cliente;
```

---

## Quick Start

Each project has its own README with setup instructions. For the SQL portfolio specifically:

```bash
# Option 1: Online (no install)
# Go to https://dbfiddle.uk/ → MySQL 8.0 → paste portfolio_completo.sql

# Option 2: Local MySQL
mysql -u root -p < SQL_Portfolio/portfolio_completo.sql
```

---

## Contact

**Jose Andres Sequeira Hernandez**  
Data Analyst | Business Intelligence  

[![LinkedIn](https://img.shields.io/badge/LinkedIn-Connect-0A66C2?style=flat&logo=linkedin)](https://linkedin.com/in/joseandres-sequeira-hernandez-3aaa03285)  
[![Email](https://img.shields.io/badge/Email-Contact-EA4335?style=flat&logo=gmail)](mailto:chomita0317@gmail.com)  
[![GitHub](https://img.shields.io/badge/GitHub-JoseAn03-181717?style=flat&logo=github)](https://github.com/JoseAn03)

---

> *"Turning raw reservation data into actionable insights — one query at a time."*

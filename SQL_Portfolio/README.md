# SQL Portfolio - Rent a Car Reservation System

**SQL queries for daily reservation management, no-show analysis, and executive reporting at Juan Santamaria International Airport (SJO)**

## Overview

This portfolio contains SQL scripts developed for processing 500+ daily reservations across three major car rental brands: **Alamo**, **Enterprise**, and **National**. These queries automate the daily Game Plan report, analyze No-Show patterns by brand/time block, and generate executive summaries for operational decision-making.

## Scripts

| Script | Description |
|--------|-------------|
| `01_creacion_tablas_reservas.sql` | Database schema: brands, reservations, no-show log, daily reports |
| `02_analisis_reservas_diarias.sql` | Daily Game Plan, hourly block analysis, brand summaries |
| `03_analisis_no_show.sql` | No-show rates by brand, day of week, vehicle class, customer |
| `04_reportes_ejecutivos.sql` | Top agencies, weekly trends, rate distribution |
| `05_resumen_gameplan.sql` | Complete Game Plan report with VIP/Expedia indicators |
| `portfolio_completo.sql` | **All-in-one script** with sample data (ready to run) |

## Technologies

- **SQL** (MySQL / MariaDB compatible)
- **DDL**: CREATE TABLE, FOREIGN KEY, INDEX, ENUM
- **DML**: INSERT, SELECT, JOIN, GROUP BY, Window Functions
- **Analytical**: CASE WHEN, Aggregations, Subqueries, Date Functions

## Quick Start (Online)

1. Go to **https://dbfiddle.uk/**
2. Select **MySQL 8.0**
3. Paste and run `portfolio_completo.sql`
4. Execute each query section individually

## Sample Query

```sql
-- Daily Game Plan Report
SELECT 
    ROW_NUMBER() OVER (PARTITION BY m.nombre_marca ORDER BY r.fecha_recogida) AS letra,
    r.numero_reserva,
    CONCAT(m.nombre_marca, '//') AS marca,
    r.nombre_cliente,
    CASE WHEN r.es_vip = TRUE THEN '* MAIN VIP*' ELSE '' END AS tipo,
    DATE_FORMAT(r.fecha_recogida, '%H:%i') AS hora
FROM reservas r
JOIN marcas m ON r.marca_id = m.marca_id
WHERE DATE(r.fecha_recogida) = CURDATE() AND r.estado = 'activa'
ORDER BY m.nombre_marca, r.fecha_recogida;
```

## Contact

**Jose Andres Sequeira Hernandez**  
Data Analyst | Business Intelligence  
[LinkedIn](https://linkedin.com/in/joseandres-sequeira-hernandez-3aaa03285)

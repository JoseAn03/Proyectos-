# SQL Portfolio — Rent a Car Reservation System

[![MySQL](https://img.shields.io/badge/MySQL-8.0-4479A1?style=flat&logo=mysql&logoColor=white)](https://dev.mysql.com/)
[![DB Fiddle](https://img.shields.io/badge/Try%20Online-DB%20Fiddle-F37626?style=flat)](https://dbfiddle.uk/)

**SQL queries for daily reservation management, No-Show analysis, and executive reporting** at Enterprise Holdings — processing **500+ daily reservations** across Alamo, Enterprise, and National at Juan Santamaria International Airport (SJO).

## Scripts

| # | Script | Description | Concepts |
|---|--------|-------------|----------|
| 1 | [`01_creacion_tablas_reservas.sql`](./01_creacion_tablas_reservas.sql) | Database schema: brands, reservations, No-Show log | DDL, FK, ENUM, INDEX, AUTO_INCREMENT |
| 2 | [`02_analisis_reservas_diarias.sql`](./02_analisis_reservas_diarias.sql) | Daily Game Plan, hourly blocks, brand summaries | JOIN, GROUP BY, CASE WHEN, DATE_FORMAT |
| 3 | [`03_analisis_no_show.sql`](./03_analisis_no_show.sql) | No-Show rates by brand, day-of-week, vehicle class, customer | Aggregations, Subqueries, ROUND, HAVING |
| 4 | [`04_reportes_ejecutivos.sql`](./04_reportes_ejecutivos.sql) | Top agencies, weekly trends, rate distribution | Window Functions, RANK, LAG, ROLLUP |
| 5 | [`05_resumen_gameplan.sql`](./05_resumen_gameplan.sql) | Complete Game Plan with VIP/Expedia indicators | ROW_NUMBER, PARTITION BY, CONCAT, CTEs |
| 🎯 | [`portfolio_completo.sql`](./portfolio_completo.sql) | **All-in-one**: DDL + sample data + all queries (ready to run) | Everything above |

## Sample Outputs

### 1. Daily Game Plan (`02_analisis_reservas_diarias.sql`)

```sql
-- Hourly reservation volume
+----------+---------------------+-------+
| Marca    | Bloque              | Total |
+----------+---------------------+-------+
| ALAMO    | 07:00 - 08:59       |    12 |
| ALAMO    | 09:00 - 10:59       |    18 |
| ALAMO    | 11:00 - 12:59       |    15 |
| ENTERPRISE | 07:00 - 08:59      |     8 |
| NATIONAL | 09:00 - 10:59       |    10 |
+----------+---------------------+-------+
```

### 2. No-Show Analysis (`03_analisis_no_show.sql`)

```sql
-- No-Show rate by brand
+------------+------------+----------+
| Marca      | No Shows   | Tasa     |
+------------+------------+----------+
| ALAMO      |        145 |   28.43% |
| ENTERPRISE |         82 |   26.11% |
| NATIONAL   |         78 |   25.83% |
+------------+------------+----------+
```

### 3. Executive Report (`04_reportes_ejecutivos.sql`)

```sql
-- Top 5 agencies by reservation volume
+------------------+------------+-----------+
| Agencia          | Reservas   | Ingreso   |
+------------------+------------+-----------+
| Expedia          |        89 | $12,460.00 |
| Booking.com      |        67 | $10,050.00 |
| Direct (Web)     |        52 |  $7,280.00 |
| Kayak            |        41 |  $5,740.00 |
| Priceline        |        38 |  $5,320.00 |
+------------------+------------+-----------+
```

### 4. Game Plan con VIP/Expedia (`05_resumen_gameplan.sql`)

```sql
-- Daily Game Plan with indicators
+-------+--------------+----------+------+---------------+---------+--------------+
| Letra | N. Reserva   | Marca    | **   | Nombre        | Clase   | MAIN VIP     |
+-------+--------------+----------+------+---------------+---------+--------------+
| A     | ANC001234    | ALAMO//  | I    | Adams, John   | Premium | * MAIN VIP*  |
|       | ANC001567    | ALAMO//  | ⭐FL⭐ | Anderson, M. | Compact |              |
|       | ANC001890    | ALAMO//  | ⭐I-FL⭐ | Austin, S. | SUV     | * EXPEDIA*   |
| B     | ANC002345    | ENTERPRISE// | P | Baker, L.    | Sedan   |              |
+-------+--------------+----------+------+---------------+---------+--------------+
```

## Quick Start

### Option 1: Online (recommended)
1. Go to **[https://dbfiddle.uk/](https://dbfiddle.uk/)**
2. Select **MySQL 8.0**
3. Paste the contents of `portfolio_completo.sql`
4. Run the setup, then execute each query section individually

### Option 2: Local MySQL
```bash
mysql -u root -p < portfolio_completo.sql
```

### Option 3: Individual scripts
```bash
# Create schema first
mysql -u root -p < 01_creacion_tablas_reservas.sql

# Then run any analysis script
mysql -u root -p < 02_analisis_reservas_diarias.sql
```

> **Note:** `portfolio_completo.sql` includes everything — schema, sample data, and all queries in one file.

## Key SQL Concepts Used

| Concept | Examples |
|---------|----------|
| **DDL** | CREATE TABLE, ALTER TABLE, FOREIGN KEY, INDEX, ENUM |
| **Joins** | INNER JOIN, LEFT JOIN, self-joins |
| **Aggregations** | COUNT, SUM, AVG, GROUP BY, HAVING |
| **Window Functions** | ROW_NUMBER, RANK, LAG, PARTITION BY |
| **Conditional Logic** | CASE WHEN, IF, COALESCE, NULLIF |
| **Date Functions** | DATE_FORMAT, CURDATE, DATEDIFF, WEEKDAY |
| **String Functions** | CONCAT, SUBSTRING, LIKE, LOCATE |
| **Subqueries** | Scalar, correlated, EXISTS |
| **CTEs** | WITH ... AS (Common Table Expressions) |
| **Views** | CREATE VIEW for reusable reports |

---

**Contact:** [Jose An03 Sequeira](https://github.com/JoseAn03) — [LinkedIn](https://linkedin.com/in/joseandres-sequeira-hernandez-3aaa03285)

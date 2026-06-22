# SQL Portfolio — Reservation System Analysis

<p align="center">
  <img src="https://img.shields.io/badge/MySQL-8.0-4479A1?style=flat&logo=mysql&logoColor=white" alt="MySQL">
  <img src="https://img.shields.io/badge/MariaDB-10.5+-003545?style=flat&logo=mariadb&logoColor=white" alt="MariaDB">
  <img src="https://img.shields.io/badge/Pure%20SQL-No%20Dependencies-00ADD8?style=flat" alt="Pure SQL">
</p>

Analytical SQL queries for **Enterprise Holdings** (Alamo, Enterprise, National) at SJO Airport — daily Game Plan automation, No-Show analysis, and executive reporting.

---

## Scripts

| # | Script | Description | Key Concepts |
|---|--------|-------------|--------------|
| 1 | [`creacion_tablas_reservas.sql`](./01_creacion_tablas_reservas.sql) | Database schema: brands, reservations, No-Show log | DDL, FK, ENUM, INDEX |
| 2 | [`analisis_reservas_diarias.sql`](./02_analisis_reservas_diarias.sql) | Daily Game Plan, hourly blocks, brand summaries | JOIN, GROUP BY, CASE |
| 3 | [`analisis_no_show.sql`](./03_analisis_no_show.sql) | No-Show rates by brand, day, vehicle class | Subqueries, HAVING |
| 4 | [`reportes_ejecutivos.sql`](./04_reportes_ejecutivos.sql) | Top agencies, weekly trends, revenue analysis | RANK, LAG, ROLLUP |
| 5 | [`resumen_gameplan.sql`](./05_resumen_gameplan.sql) | Game Plan with VIP/Expedia indicators | ROW_NUMBER, PARTITION BY |
| All | [`portfolio_completo.sql`](./portfolio_completo.sql) | DDL + 200+ sample records + all queries | Everything above |

---

## Sample Results

**No-Show Rate by Brand**

| Brand | No Shows | Rate |
|-------|----------|------|
| ALAMO | 145 | 28.43% |
| ENTERPRISE | 82 | 26.11% |
| NATIONAL | 78 | 25.83% |

**Top 5 Agencies by Revenue**

| Agency | Reservations | Revenue |
|--------|-------------|---------|
| Expedia | 89 | $12,460 |
| Booking.com | 67 | $10,050 |
| Direct (Web) | 52 | $7,280 |
| Kayak | 41 | $5,740 |
| Priceline | 38 | $5,320 |

**Game Plan with VIP/Expedia Flags**

| Letra | N. Reserva | Brand | ** | Customer | Class | Status |
|-------|-----------|-------|-----|----------|-------|--------|
| A | ANC001234 | ALAMO// | I | Adams, John | Premium | * MAIN VIP* |
|   | ANC001567 | ALAMO// | FL | Anderson, M. | Compact | |
|   | ANC001890 | ALAMO// | I-FL | Austin, S. | SUV | * EXPEDIA* |
| B | ANC002345 | ENTERPRISE// | P | Baker, L. | Sedan | |

---

## Getting Started

**Prerequisites:** MySQL 8.0+ or MariaDB 10.5+ with no external dependencies.

```bash
# Create schema, insert sample data, and run all queries
mysql -u root -p < portfolio_completo.sql

# Or run individual scripts
mysql -u root -p < 01_creacion_tablas_reservas.sql
mysql -u root -p < 02_analisis_reservas_diarias.sql
```

---

## SQL Techniques

| Technique | Scripts | Purpose |
|-----------|---------|---------|
| DDL | 1 | Table creation, constraints, indexing |
| JOINs | 2-5 | Multi-table relationship queries |
| CASE WHEN | 2, 5 | Conditional brand/status logic |
| GROUP BY + HAVING | 2, 3 | Aggregations with filtering |
| Window Functions | 4 | RANK, LAG for trend analysis |
| ROW_NUMBER / PARTITION BY | 5 | A-Z letter grouping by brand |
| Date Functions | 2, 3 | DATE_FORMAT, DATEDIFF, WEEKDAY |
| Subqueries | 3 | Correlated subqueries for repeat analysis |
| CTEs | 5 | Readable common table expressions |
| Views | Schema | Reusable report definitions |

---

## Contact

**Jose Andres Sequeira Hernandez**  
[LinkedIn](https://linkedin.com/in/joseandres-sequeira-hernandez-3aaa03285) · [GitHub](https://github.com/JoseAn03)

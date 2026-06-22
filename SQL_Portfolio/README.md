<div align="center">
  
  [![MySQL](https://img.shields.io/badge/MySQL-8.0-4479A1?style=for-the-badge&logo=mysql&logoColor=white)](https://dev.mysql.com/)
  [![MariaDB](https://img.shields.io/badge/MariaDB-10.5+-003545?style=for-the-badge&logo=mariadb&logoColor=white)](https://mariadb.org/)
  
</div>

# SQL Portfolio — Reservation System Analysis

A collection of analytical SQL queries developed for **Enterprise Holdings** (Alamo, Enterprise, National) at Juan Santamaria International Airport (SJO). These scripts automate daily Game Plan reporting, No-Show analysis, and executive-level business intelligence.

---

## 📋 Scripts

<table>
  <thead>
    <tr>
      <th>#</th>
      <th>Script</th>
      <th>Description</th>
      <th>SQL Concepts</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td align="center">1</td>
      <td><a href="./01_creacion_tablas_reservas.sql">creacion_tablas_reservas.sql</a></td>
      <td>Database schema: brands, reservations, No-Show log, daily reports</td>
      <td><code>DDL</code> <code>FK</code> <code>ENUM</code> <code>INDEX</code></td>
    </tr>
    <tr>
      <td align="center">2</td>
      <td><a href="./02_analisis_reservas_diarias.sql">analisis_reservas_diarias.sql</a></td>
      <td>Daily Game Plan, hourly reservation blocks, brand summaries</td>
      <td><code>JOIN</code> <code>GROUP BY</code> <code>CASE</code></td>
    </tr>
    <tr>
      <td align="center">3</td>
      <td><a href="./03_analisis_no_show.sql">analisis_no_show.sql</a></td>
      <td>No-Show rates by brand, day-of-week, vehicle class, repeat customers</td>
      <td><code>Subqueries</code> <code>HAVING</code></td>
    </tr>
    <tr>
      <td align="center">4</td>
      <td><a href="./04_reportes_ejecutivos.sql">reportes_ejecutivos.sql</a></td>
      <td>Top agencies, weekly trends, rate distribution, revenue analysis</td>
      <td><code>RANK</code> <code>LAG</code> <code>ROLLUP</code></td>
    </tr>
    <tr>
      <td align="center">5</td>
      <td><a href="./05_resumen_gameplan.sql">resumen_gameplan.sql</a></td>
      <td>Complete Game Plan with VIP/Expedia indicators and customer flags</td>
      <td><code>ROW_NUMBER</code> <code>PARTITION BY</code></td>
    </tr>
    <tr>
      <td align="center">🎯</td>
      <td><a href="./portfolio_completo.sql">portfolio_completo.sql</a></td>
      <td><strong>All-in-one:</strong> DDL + 200+ sample records + all queries</td>
      <td><em>Everything above</em></td>
    </tr>
  </tbody>
</table>

---

## 📊 Sample Query Results

### Daily Game Plan — Hourly Reservation Volume

```sql
-- From: 02_analisis_reservas_diarias.sql
SELECT m.nombre_marca AS Marca,
       CASE ... END AS Bloque,
       COUNT(*) AS Total
FROM reservas r ...
GROUP BY Marca, Bloque;
```

| Marca | Bloque | Total |
|-------|--------|-------|
| ALAMO | 07:00 - 08:59 | 12 |
| ALAMO | 09:00 - 10:59 | 18 |
| ENTERPRISE | 07:00 - 08:59 | 8 |
| NATIONAL | 09:00 - 10:59 | 10 |

### No-Show Rate by Brand

| Marca | No Shows | Rate |
|-------|----------|------|
| ALAMO | 145 | 28.43% |
| ENTERPRISE | 82 | 26.11% |
| NATIONAL | 78 | 25.83% |

### Top 5 Agencies by Revenue

| Agency | Reservations | Revenue |
|--------|-------------|---------|
| Expedia | 89 | $12,460 |
| Booking.com | 67 | $10,050 |
| Direct (Web) | 52 | $7,280 |
| Kayak | 41 | $5,740 |
| Priceline | 38 | $5,320 |

### Game Plan with VIP/Expedia Indicators

| Letra | Reservation | Brand | ** | Customer | Class | Status |
|-------|-------------|-------|-----|----------|-------|--------|
| A | ANC001234 | ALAMO// | I | Adams, John | Premium | * MAIN VIP* |
| | ANC001567 | ALAMO// | FL | Anderson, M. | Compact | |
| | ANC001890 | ALAMO// | I-FL | Austin, S. | SUV | * EXPEDIA* |
| B | ANC002345 | ENTERPRISE// | P | Baker, L. | Sedan | |

---

## 🚀 Getting Started

### Prerequisites
- MySQL 8.0+ or MariaDB 10.5+
- No external dependencies — pure SQL

### Execute the Complete Portfolio

```bash
mysql -u root -p < portfolio_completo.sql
```

This single command creates the database schema, inserts 200+ sample reservation records, and runs all analytical queries.

### Run Individual Scripts

```bash
# 1. Create the database schema first
mysql -u root -p < 01_creacion_tablas_reservas.sql

# 2. Run any analysis script
mysql -u root -p < 02_analisis_reservas_diarias.sql
mysql -u root -p < 03_analisis_no_show.sql
mysql -u root -p < 04_reportes_ejecutivos.sql
mysql -u root -p < 05_resumen_gameplan.sql
```

---

## 🧠 SQL Techniques Demonstrated

<div align="center">

| Technique | Used In | Purpose |
|-----------|---------|---------|
| **DDL** | Script 1 | Table creation, constraints, indexing |
| **INNER / LEFT JOIN** | Scripts 2-5 | Multi-table relationship queries |
| **CASE WHEN** | Scripts 2, 5 | Conditional logic for brand/status flags |
| **GROUP BY + HAVING** | Scripts 2, 3 | Aggregation with filtered groups |
| **Window Functions** | Script 4 | RANK, LAG for trend analysis |
| **ROW_NUMBER / PARTITION BY** | Script 5 | A-Z letter grouping per brand |
| **Date Functions** | Scripts 2, 3 | DATE_FORMAT, DATEDIFF, WEEKDAY |
| **Subqueries** | Script 3 | Correlated subqueries for repeat customers |
| **CTEs** | Script 5 | Common Table Expressions for readability |
| **Views** | Schema | Reusable report definitions |

</div>

---

## 📬 Contact

**Jose Andres Sequeira Hernandez**

[![LinkedIn](https://img.shields.io/badge/LinkedIn-Connect-0A66C2?style=for-the-badge&logo=linkedin&logoColor=white)](https://linkedin.com/in/joseandres-sequeira-hernandez-3aaa03285)
[![Email](https://img.shields.io/badge/Email-Contact-EA4335?style=for-the-badge&logo=gmail&logoColor=white)](mailto:chomita0317@gmail.com)
[![GitHub](https://img.shields.io/badge/GitHub-JoseAn03-181717?style=for-the-badge&logo=github&logoColor=white)](https://github.com/JoseAn03)

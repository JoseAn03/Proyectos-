# 📊 Dashboard de Reservas en Power BI

Guía paso a paso para construir un dashboard profesional de análisis de reservas, usando los mismos datos que procesás en tu día a día en ANC Renta Car.

---

## 1. 📁 Datos

Los archivos CSV necesarios ya están generados en esta carpeta:

| Archivo | Contenido |
|---------|-----------|
| `reservas_ejemplo.csv` | 200 reservas con marca, fechas, tarifas, agencia, VIP, no-show |
| `no_show_ejemplo.csv` | Registros de clientes que no llegaron |
| `vuelos_ejemplo.csv` | Vuelos internacionales SJO (para cruzar con reservas) |

> 💡 Reemplazalos después con tus datos reales (solo exportá el manifiesto a CSV).

---

## 2. 🚀 Primeros pasos en Power BI Desktop

1. **Descargá Power BI Desktop** (gratis): [powerbi.microsoft.com/desktop](https://powerbi.microsoft.com/desktop)
2. Abrí el programa
3. **Obtener datos → CSV** → seleccioná los 3 archivos
4. En el panel derecho, conectá las tablas con **Modelado > Administrar relaciones**:
   - `reservas[Marca]` ↔ `no_show[Marca]`
   - `reservas[ID_Reserva]` ↔ `no_show[ID_Reserva]`

---

## 3. 📐 Diseño del Dashboard (4 páginas)

### Página 1: 📈 Resumen Ejecutivo

**Visuales:**
- **Tarjetas (4):**
  - Total Reservas (`=COUNTROWS(reservas)`)
  - Tasa No Show (`=DIVIDE(COUNTROWS(no_show), COUNTROWS(reservas))` en %)
  - Ingreso Promedio por Reserva (`=AVERAGE(reservas[Tarifa_Diaria])`)
  - Clientes VIP (`=CALCULATE(COUNTROWS(reservas), reservas[Es_VIP]="Sí")`)
- **Gráfico de barras:** Reservas por Marca (Alamo vs Enterprise vs National)
- **Gráfico de líneas:** Reservas por día del mes
- **Segmentadores:**
  - Mes (para filtrar por período)
  - Marca (Alamo / Enterprise / National)

**Layout sugerido:**
```
┌──────────────────────────────────────────────┐
│                  TÍTULO                       │
├─────────┬─────────┬─────────┬────────────────┤
│ Tarjeta1│ Tarjeta2│ Tarjeta3│   Tarjeta4     │
├─────────┴─────────┴─────────┴────────────────┤
│   Barras: Reservas x Marca   │ Línea: Tendencia│
├───────────────────────────────────────────────┤
│    Segmentadores (Mes, Marca)                 │
└───────────────────────────────────────────────┘
```

### Página 2: 🔍 Análisis por Marca

**Visuales:**
- **Gráfico de anillo:** Distribución por Marca
- **Tabla:** Top 10 clientes por marca (Nombre, Tarifa, Días)
- **Gráfico de barras apiladas:** Clase de vehículo por marca
- **Matriz:** Marca × Mes (valor = total ingresos)
- **Gráfico de columnas:** No Show por marca

**Filtro de página:** Marca (para ver una marca a la vez)

### Página 3: 🕐 Análisis Horario

**Visuales:**
- **Gráfico de columnas:** Reservas por hora del día (para ver horarios pico)
- **Gráfico de líneas:** Duración promedio (días) por día de la semana
- **Gráfico de dispersión:** Tarifa diaria vs. duración de renta
- **Mapa:** (si tenés ubicaciones) o tabla: Ubicación × No Show rate

### Página 4: ❌ No Show & Expedia

**Visuales:**
- **Gráfico de barras:** % No Show por agencia (top 5)
- **Tabla:** Clientes No Show (fecha, marca, tarifa perdida)
- **Gráfico de columnas:** No Show por día de la semana
- **Indicador:** Expedia vs No Expedia (tasa de no show comparativa)

---

## 4. 🧮 Medidas DAX (para copiar y pegar)

En Power BI, andá a **Modelado > Nueva medida** y pegalas:

```dax
Total Reservas = COUNTROWS(reservas)

Total No Show = COUNTROWS(no_show)

% No Show = 
DIVIDE([Total No Show], [Total Reservas], 0)

Tarifa Promedio = AVERAGE(reservas[Tarifa_Diaria])

Ingreso Total = SUM(reservas[Tarifa_Diaria])

Duración Promedio = AVERAGE(reservas[Duracion_Dias])

% VIP = 
DIVIDE(
    CALCULATE(COUNTROWS(reservas), reservas[Es_VIP] = "Sí"),
    [Total Reservas],
    0
)

% Expedia = 
DIVIDE(
    CALCULATE(COUNTROWS(reservas), reservas[Es_Expedia] = "Sí"),
    [Total Reservas],
    0
)

Reservas Alamo = 
CALCULATE([Total Reservas], reservas[Marca] = "Alamo")

Reservas Enterprise = 
CALCULATE([Total Reservas], reservas[Marca] = "Enterprise")

Reservas National = 
CALCULATE([Total Reservas], reservas[Marca] = "National")

No Show por Marca = 
DIVIDE(
    CALCULATE(COUNTROWS(no_show)),
    CALCULATE(COUNTROWS(reservas)),
    0
)

Ranking Clientes = 
RANKX(ALL(reservas), [Total Reservas], , DESC, Dense)
```

---

## 5. 🎨 Colores sugeridos (tu marca)

Usá estos colores en los gráficos para mantener consistencia con ANC:

| Marca | Color | Código Hex |
|-------|-------|------------|
| **Alamo** | Azul | `#0070C0` |
| **Enterprise** | Negro | `#000000` |
| **National** | Verde | `#00B050` |

**Fondo general:** Blanco o gris claro (`#F5F5F5`)
**Títulos:** Azul oscuro (`#1F4E79`), negrita, tamaño 16-18
**Tarjetas KPI:** Fondo blanco, borde sutil

---

## 6. ⚡ Tips pro

1. **Actualización automática:** Configurá el origen de datos apuntando a la carpeta donde guardás los manifiestos. Así cada que abras Power BI se refresca solo.

2. **Bookmarks:** Creá bookmarks para alternar entre vista "Resumen" y "Detalle" en la misma página.

3. **Tooltips:** En el gráfico de barras de marca, configurá un tooltip con la tabla de top clientes de esa marca.

4. **Publicar:** Si tenés licencia Pro, publicá a **Power BI Service** y compartí el link directo.

5. **Teléfono:** Creá una vista optimizada para celular (Vista > Diseño de teléfono).

---

## 7. 📁 Cómo usar tus datos reales

1. Ejecutá tu script `gameplan_reservas.py` para generar el Excel de reservas
2. Exportá cada pestaña (ALAMO, ENTERPRISE, NATIONAL) a CSV
3. Unilos con un pequeño script o con Power Query (Combinar archivos)
4. Conectá las columnas: las del CSV son las mismas del dashboard

---

¿Atorado en algún paso? Decime y te ayudo con los DAX, relaciones o visuales específicos.

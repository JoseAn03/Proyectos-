# ✈️ Vuelos SJO — Análisis de Vuelos Internacionales

Procesa datos de vuelos del Aeropuerto Juan Santamaría, filtra vuelos internacionales excluyendo rutas domésticas y ciertas aerolíneas, y genera un reporte organizado por bloques horarios.

## Archivo

| Archivo | Descripción |
|---------|-------------|
| [`VUELOS_MACRO_REPORTE.bas`](./VUELOS_MACRO_REPORTE.bas) | Macro VBA para generar reporte de vuelos internacionales |

## ¿Qué hace?

1. Lee datos de vuelos desde la hoja activa
2. Excluye vuelos domésticos (Tambor, Tamarindo, Liberia, Quepos, etc.)
3. Excluye aerolíneas específicas (Sansa, LATAM, GetJet, NetJets, Viva, Mas Air)
4. Clasifica vuelos en 6 bloques horarios:
   - 05:00–08:59 | 09:00–11:59 | 12:00–14:59
   - 15:00–17:59 | 18:00–20:59 | 21:00–23:59
5. Genera hoja "Reporte" con formato profesional:
   - Título y fecha automática
   - Encabezados con colores corporativos
   - Filas alternadas (zebra striping)
   - Bloques horarios con separadores visuales

## Cómo usar

1. Copiá los datos de FlightRadar24 en una hoja de Excel
2. Presioná `Alt + F11` → Editor de VBA
3. Insertar → Módulo
4. Pegá el código de `VUELOS_MACRO_REPORTE.bas`
5. Cerra el editor
6. Presioná `Alt + F8`, seleccioná `GenerarReporteVuelos` → Ejecutar

## Resultado

📊 Hoja "Reporte" con todos los vuelos internacionales del día, organizados por bloques horarios y listos para compartir.

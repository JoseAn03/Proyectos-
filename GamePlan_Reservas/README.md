# 📋 GamePlan Reservas — Automatización de Reporte Diario

Procesa el manifiesto diario de reservas y genera un Excel con tres pestañas (Alamo, Enterprise, National), ordenado alfabéticamente por nombre del cliente, con formato condicional por marca y colores.

## Archivos

| Archivo | Descripción |
|---------|-------------|
| [`GAMEPLAN_MACRO_FINAL_2026.bas`](./GAMEPLAN_MACRO_FINAL_2026.bas) | Macro VBA completa para procesar reservas |
| [`prompt_ia_gameplan.md`](./prompt_ia_gameplan.md) | Prompt para procesar el mismo reporte con Claude (alternativa IA) |

## ¿Qué hace?

1. Lee el manifiesto desde la hoja "Sheet"
2. Clasifica cada reserva por marca según la ubicación de recogida:
   - **Alamo** → SJOE71 / SJOT71
   - **Enterprise** → SJOC61 / SJOT61
   - **National** → SJOE01 / SJOT01
3. Ordena alfabéticamente por nombre del cliente
4. Agrupa por letra inicial (A-Z) con bordes y formato
5. Aplica colores:
   - 🔵 Alamo → azul
   - ⚫ Enterprise → negro
   - 🟢 National → verde
6. Calcula indicadores:
   - **Columna ****: `I`/`P`/`E` para clientes con `+`, `⭐FL⭐` para `~`/`*`, combinaciones como `⭐I-FL⭐`
   - **Columna MAIN VIP**: `* MAIN VIP*` o `* EXPEDIA*` según ubicación o agencia

## Cómo usar la macro VBA

1. Abrí el archivo Excel fuente (ResManifest_.xlsx)
2. Presioná `Alt + F11` → Editor de VBA
3. Insertar → Módulo
4. Pegá el código de `GAMEPLAN_MACRO_FINAL_2026.bas`
5. Cerra el editor
6. Presioná `Alt + F8`, seleccioná `ProcesarReservas` → Ejecutar

## Resultado

📊 Excel con 3 pestañas (ALAMO, ENTERPRISE, NATIONAL), cada una con 13 columnas formateadas y listas para usar.

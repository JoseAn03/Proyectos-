# 🚫 NoShow Reporte — Automatización de Reporte Diario

Genera un reporte de No Show con todas las reservas activas del día, organizadas alfabéticamente por nombre del cliente y agrupadas por marca.

## Archivo

| Archivo | Descripción |
|---------|-------------|
| [`NOSHOW_MACRO.bas`](./NOSHOW_MACRO.bas) | Macro VBA para generar el reporte No Show |

## ¿Qué hace?

1. Lee los datos del manifiesto desde la hoja "Sheet"
2. Almacena y ordena todas las reservas alfabéticamente
3. Clasifica por marca según ubicación de recogida
4. Genera una hoja "Manifiesto" con formato completo:
   - Columnas: Letra | N.º reserva | Marca | ** | Nombre | Clase | **MAIN VIP** | Fechas | Ubicaciones | Tarifa | Agencia
   - Letra inicial (A-Z) en rojo con borde superior en cada grupo
   - Colores por marca en columnas Marca y Nombre
   - Indicadores FL y Marca en columna **
   - MAIN VIP y EXPEDIA en columna **MAIN VIP**

## Cómo usar

1. Abrí el archivo Excel fuente (ResManifest_.xlsx)
2. Presioná `Alt + F11` → Editor de VBA
3. Insertar → Módulo
4. Pegá el código de `NOSHOW_MACRO.bas`
5. Cerra el editor
6. Presioná `Alt + F8`, seleccioná `GenerarResManifest` → Ejecutar

## Resultado

📊 Hoja "Manifiesto" con todas las reservas formateadas y listas para revisión.

=========================================================================
  SCRIPT: 04_reportes_ejecutivos.sql
  PROYECTO: Base de Datos de Reservas - ANC Renta Car
  DESCRIPCION: Consultas para reportes ejecutivos semanales/mensuales
  TECNOLOGIA: SQL
=========================================================================

-- -----------------------------------------------------------
-- 4.1  TOP 10 AGENCIAS DE RECOMENDACION POR VOLUMEN
-- -----------------------------------------------------------
SELECT 
    r.agencia_recomienda,
    COUNT(*) AS total_reservas,
    ROUND(SUM(r.tarifa_diaria), 2) AS ingreso_total,
    ROUND(AVG(r.tarifa_diaria), 2) AS tarifa_promedio,
    SUM(CASE WHEN r.estado = 'no_show' THEN 1 ELSE 0 END) AS no_shows
FROM reservas r
WHERE r.agencia_recomienda IS NOT NULL 
    AND r.agencia_recomienda != ''
    AND r.fecha_recogida >= DATE_SUB(CURDATE(), INTERVAL 3 MONTH)
GROUP BY r.agencia_recomienda
ORDER BY total_reservas DESC
LIMIT 10;

-- -----------------------------------------------------------
-- 4.2  TENDENCIA SEMANAL DE RESERVAS (ULTIMAS 12 SEMANAS)
-- -----------------------------------------------------------
SELECT 
    YEARWEEK(r.fecha_recogida, 1) AS semana,
    MIN(DATE(r.fecha_recogida)) AS fecha_inicio_semana,
    COUNT(*) AS reservas_semana,
    ROUND(AVG(r.tarifa_diaria), 2) AS tarifa_promedio,
    SUM(CASE WHEN r.estado = 'no_show' THEN 1 ELSE 0 END) AS no_shows_semana
FROM reservas r
WHERE r.fecha_recogida >= DATE_SUB(CURDATE(), INTERVAL 12 WEEK)
GROUP BY YEARWEEK(r.fecha_recogida, 1)
ORDER BY semana DESC;

-- -----------------------------------------------------------
-- 4.3  DISTRIBUCION DE TARIFAS POR MARCA
-- -----------------------------------------------------------
SELECT 
    m.nombre_marca,
    COUNT(*) AS total_reservas,
    ROUND(MIN(r.tarifa_diaria), 2) AS tarifa_min,
    ROUND(AVG(r.tarifa_diaria), 2) AS tarifa_promedio,
    ROUND(MAX(r.tarifa_diaria), 2) AS tarifa_max,
    ROUND(STDDEV(r.tarifa_diaria), 2) AS desviacion_estandar
FROM reservas r
JOIN marcas m ON r.marca_id = m.marca_id
WHERE r.fecha_recogida >= DATE_SUB(CURDATE(), INTERVAL 6 MONTH)
GROUP BY m.nombre_marca
ORDER BY tarifa_promedio DESC;

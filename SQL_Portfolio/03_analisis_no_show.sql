=========================================================================
  SCRIPT: 03_analisis_no_show.sql
  PROYECTO: Base de Datos de Reservas - ANC Renta Car
  DESCRIPCION: Consultas para analisis de patrones de No-Show
  TECNOLOGIA: SQL
=========================================================================

-- -----------------------------------------------------------
-- 3.1  TASA DE NO-SHOW POR MARCA (ULTIMOS 30 DIAS)
-- -----------------------------------------------------------
SELECT 
    m.nombre_marca,
    COUNT(*) AS total_reservas_30d,
    SUM(CASE WHEN r.estado = 'no_show' THEN 1 ELSE 0 END) AS total_no_show,
    ROUND(
        SUM(CASE WHEN r.estado = 'no_show' THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 2
    ) AS tasa_no_show_porcentaje
FROM reservas r
JOIN marcas m ON r.marca_id = m.marca_id
WHERE r.fecha_recogida >= DATE_SUB(CURDATE(), INTERVAL 30 DAY)
GROUP BY m.nombre_marca
ORDER BY tasa_no_show_porcentaje DESC;

-- -----------------------------------------------------------
-- 3.2  NO-SHOW POR DIA DE LA SEMANA
-- -----------------------------------------------------------
SELECT 
    DAYNAME(r.fecha_recogida) AS dia_semana,
    DAYOFWEEK(r.fecha_recogida) AS num_dia,
    COUNT(*) AS total_reservas,
    SUM(CASE WHEN r.estado = 'no_show' THEN 1 ELSE 0 END) AS no_shows,
    ROUND(
        SUM(CASE WHEN r.estado = 'no_show' THEN 1 ELSE 0 END) * 100.0 / NULLIF(COUNT(*), 0), 1
    ) AS tasa_no_show
FROM reservas r
WHERE r.fecha_recogida >= DATE_SUB(CURDATE(), INTERVAL 90 DAY)
GROUP BY DAYNAME(r.fecha_recogida), DAYOFWEEK(r.fecha_recogida)
ORDER BY num_dia;

-- -----------------------------------------------------------
-- 3.3  CLIENTES CON MAYOR HISTORIAL DE NO-SHOW
-- -----------------------------------------------------------
SELECT 
    r.nombre_cliente,
    COUNT(*) AS total_reservas,
    SUM(CASE WHEN r.estado = 'no_show' THEN 1 ELSE 0 END) AS total_no_show,
    ROUND(
        SUM(CASE WHEN r.estado = 'no_show' THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 1
    ) AS tasa_no_show
FROM reservas r
WHERE r.fecha_recogida >= DATE_SUB(CURDATE(), INTERVAL 6 MONTH)
GROUP BY r.nombre_cliente
HAVING COUNT(*) >= 3 AND total_no_show >= 1
ORDER BY tasa_no_show DESC
LIMIT 20;

-- -----------------------------------------------------------
-- 3.4  NO-SHOW POR CLASE DE VEHICULO
-- -----------------------------------------------------------
SELECT 
    r.clase_vehiculo,
    COUNT(*) AS total,
    SUM(CASE WHEN r.estado = 'no_show' THEN 1 ELSE 0 END) AS no_show,
    ROUND(
        SUM(CASE WHEN r.estado = 'no_show' THEN 1 ELSE 0 END) * 100.0 / NULLIF(COUNT(*), 0), 1
    ) AS porcentaje
FROM reservas r
WHERE r.fecha_recogida >= DATE_SUB(CURDATE(), INTERVAL 90 DAY)
GROUP BY r.clase_vehiculo
ORDER BY porcentaje DESC;

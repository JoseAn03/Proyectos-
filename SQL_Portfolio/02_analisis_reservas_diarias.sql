=========================================================================
  SCRIPT: 02_analisis_reservas_diarias.sql
  PROYECTO: Base de Datos de Reservas - ANC Renta Car
  DESCRIPCION: Consultas de analisis diario de reservas por marca
  TECNOLOGIA: SQL
=========================================================================

-- -----------------------------------------------------------
-- 2.1  RESUMEN DIARIO DE RESERVAS POR MARCA
-- -----------------------------------------------------------
SELECT 
    m.nombre_marca,
    DATE(r.fecha_recogida) AS fecha,
    COUNT(*) AS total_reservas,
    ROUND(AVG(r.tarifa_diaria), 2) AS tarifa_promedio,
    ROUND(SUM(r.tarifa_diaria), 2) AS ingreso_estimado,
    SUM(CASE WHEN r.estado = 'no_show' THEN 1 ELSE 0 END) AS total_no_show,
    ROUND(
        SUM(CASE WHEN r.estado = 'no_show' THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 1
    ) AS porcentaje_no_show
FROM reservas r
JOIN marcas m ON r.marca_id = m.marca_id
WHERE DATE(r.fecha_recogida) = CURDATE()
GROUP BY m.nombre_marca, DATE(r.fecha_recogida)
ORDER BY m.nombre_marca;

-- -----------------------------------------------------------
-- 2.2  PROXIMAS RECOGIDAS DEL DIA (GAME PLAN)
-- -----------------------------------------------------------
SELECT 
    r.numero_reserva,
    m.nombre_marca AS marca,
    r.nombre_cliente,
    r.clase_vehiculo,
    TIME(r.fecha_recogida) AS hora_recogida,
    TIME(r.fecha_entrega) AS hora_entrega,
    r.tarifa_diaria,
    CASE 
        WHEN r.es_vip = TRUE THEN '* MAIN VIP*'
        WHEN r.agencia_recomienda LIKE '%11617270%' THEN '* EXPEDIA*'
        ELSE ''
    END AS tipo_reserva,
    CASE 
        WHEN r.es_flight = TRUE THEN '⭐FL⭐'
        WHEN r.es_vip = TRUE THEN '⭐VIP⭐'
        ELSE ''
    END AS indicador
FROM reservas r
JOIN marcas m ON r.marca_id = m.marca_id
WHERE DATE(r.fecha_recogida) = CURDATE()
    AND r.estado = 'activa'
ORDER BY m.nombre_marca, r.fecha_recogida;

-- -----------------------------------------------------------
-- 2.3  RESERVAS POR BLOQUE HORARIO
-- -----------------------------------------------------------
SELECT 
    m.nombre_marca,
    CASE 
        WHEN TIME(r.fecha_recogida) BETWEEN '06:00:00' AND '09:59:59' THEN '06:00-09:59'
        WHEN TIME(r.fecha_recogida) BETWEEN '10:00:00' AND '13:59:59' THEN '10:00-13:59'
        WHEN TIME(r.fecha_recogida) BETWEEN '14:00:00' AND '17:59:59' THEN '14:00-17:59'
        WHEN TIME(r.fecha_recogida) BETWEEN '18:00:00' AND '23:59:59' THEN '18:00-23:59'
        ELSE '00:00-05:59'
    END AS bloque_horario,
    COUNT(*) AS cantidad_reservas
FROM reservas r
JOIN marcas m ON r.marca_id = m.marca_id
WHERE DATE(r.fecha_recogida) = CURDATE()
GROUP BY m.nombre_marca, bloque_horario
ORDER BY m.nombre_marca, bloque_horario;

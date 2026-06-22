=========================================================================
  SCRIPT: 05_resumen_gameplan.sql
  PROYECTO: Base de Datos de Reservas - ANC Renta Car
  DESCRIPCION: Consulta principal para generar el reporte Game Plan diario
  TECNOLOGIA: SQL
=========================================================================

-- -----------------------------------------------------------
-- 5.1  GAME PLAN COMPLETO DEL DIA
-- Combina toda la info necesaria para el reporte operativo diario
-- con formato similar al que usamos en Excel
-- -----------------------------------------------------------
SELECT 
    ROW_NUMBER() OVER (PARTITION BY m.nombre_marca ORDER BY r.fecha_recogida) AS letra,
    r.numero_reserva,
    m.nombre_marca AS marca,
    CASE 
        WHEN r.es_flight = TRUE THEN CONCAT('⭐', 
            CASE m.nombre_marca 
                WHEN 'ALAMO' THEN 'I'
                WHEN 'ENTERPRISE' THEN 'P'
                WHEN 'NATIONAL' THEN 'E'
            END, '-FL⭐')
        WHEN r.es_vip = TRUE THEN 
            CASE m.nombre_marca 
                WHEN 'ALAMO' THEN '⭐I⭐'
                WHEN 'ENTERPRISE' THEN '⭐P⭐'
                WHEN 'NATIONAL' THEN '⭐E⭐'
            END
        ELSE ''
    END AS indicador,
    r.nombre_cliente,
    r.clase_vehiculo,
    CASE 
        WHEN r.agencia_recomienda LIKE '%11617270%' THEN '* EXPEDIA*'
        WHEN r.es_vip = TRUE THEN '* MAIN VIP*'
        ELSE ''
    END AS tipo_reserva,
    DATE_FORMAT(r.fecha_recogida, '%Y-%m-%d') AS fecha_recogida,
    DATE_FORMAT(r.fecha_entrega, '%Y-%m-%d') AS fecha_entrega,
    TIME(r.fecha_recogida) AS hora_recogida,
    TIME(r.fecha_entrega) AS hora_entrega,
    r.ubicacion_recogida,
    r.ubicacion_entrega,
    CONCAT('$', FORMAT(r.tarifa_diaria, 2)) AS tarifa_diaria,
    r.agencia_recomienda
FROM reservas r
JOIN marcas m ON r.marca_id = m.marca_id
WHERE DATE(r.fecha_recogida) = CURDATE()
    AND r.estado = 'activa'
ORDER BY m.nombre_marca, r.fecha_recogida;

-- -----------------------------------------------------------
-- 5.2  RESUMEN EJECUTIVO - TOTALES DEL DIA
-- -----------------------------------------------------------
SELECT 
    m.nombre_marca,
    COUNT(*) AS total_reservas_activas,
    ROUND(SUM(r.tarifa_diaria), 2) AS ingreso_estimado_total,
    COUNT(DISTINCT r.clase_vehiculo) AS tipos_vehiculo_distintos,
    ROUND(AVG(r.tarifa_diaria), 2) AS tarifa_promedio
FROM reservas r
JOIN marcas m ON r.marca_id = m.marca_id
WHERE DATE(r.fecha_recogida) = CURDATE()
    AND r.estado = 'activa'
GROUP BY m.nombre_marca
ORDER BY m.nombre_marca;

-- -----------------------------------------------------------
-- 5.3  CLIENTES VIP Y EXPEDIA DEL DIA
-- -----------------------------------------------------------
SELECT 
    r.numero_reserva,
    m.nombre_marca,
    r.nombre_cliente,
    r.clase_vehiculo,
    TIME(r.fecha_recogida) AS hora,
    CASE 
        WHEN r.es_vip = TRUE THEN 'MAIN VIP'
        WHEN r.agencia_recomienda LIKE '%11617270%' THEN 'EXPEDIA'
        ELSE 'REGULAR'
    END AS categoria
FROM reservas r
JOIN marcas m ON r.marca_id = m.marca_id
WHERE DATE(r.fecha_recogida) = CURDATE()
    AND (r.es_vip = TRUE OR r.agencia_recomienda LIKE '%11617270%')
ORDER BY categoria, m.nombre_marca;

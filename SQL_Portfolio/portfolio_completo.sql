=========================================================================
  PORTAFOLIO SQL - JOSE SEQUEIRA
  TEMA: Gestion de Reservas - Renta de Autos
  NOTA: Este script se puede correr en MySQL, MariaDB o SQLite
        (en SQLite cambiar INT por INTEGER y AUTO_INCREMENT por 
         AUTOINCREMENT)
=========================================================================

=========================================================================
  PARTE 1: CREACION DE TABLAS (DDL)
=========================================================================

CREATE DATABASE IF NOT EXISTS rentacar_sjo;
USE rentacar_sjo;

-- Tabla de marcas
CREATE TABLE marcas (
    marca_id INT PRIMARY KEY AUTO_INCREMENT,
    nombre_marca VARCHAR(50) NOT NULL,
    codigo_ubicacion VARCHAR(10) NOT NULL
);

-- Tabla de reservas
CREATE TABLE reservas (
    reserva_id INT PRIMARY KEY AUTO_INCREMENT,
    numero_reserva VARCHAR(20) NOT NULL UNIQUE,
    marca_id INT NOT NULL,
    nombre_cliente VARCHAR(100) NOT NULL,
    clase_vehiculo VARCHAR(50),
    tarifa_diaria DECIMAL(10,2),
    fecha_recogida DATETIME NOT NULL,
    fecha_entrega DATETIME NOT NULL,
    ubicacion_recogida VARCHAR(20),
    ubicacion_entrega VARCHAR(20),
    agencia_recomienda VARCHAR(100),
    es_vip BOOLEAN DEFAULT FALSE,
    es_flight BOOLEAN DEFAULT FALSE,
    estado ENUM('activa','completada','no_show','cancelada') DEFAULT 'activa',
    FOREIGN KEY (marca_id) REFERENCES marcas(marca_id)
);

-- Tabla de no-shows
CREATE TABLE no_show_log (
    no_show_id INT PRIMARY KEY AUTO_INCREMENT,
    reserva_id INT NOT NULL,
    fecha_no_show DATETIME NOT NULL,
    hora_bloque VARCHAR(20),
    FOREIGN KEY (reserva_id) REFERENCES reservas(reserva_id)
);

=========================================================================
  PARTE 2: INSERCION DE DATOS DE EJEMPLO
=========================================================================

-- Marcas
INSERT INTO marcas (nombre_marca, codigo_ubicacion) VALUES
('ALAMO', 'SJOE71'),
('ENTERPRISE', 'SJOC61'),
('NATIONAL', 'SJOE01');

-- Reservas (datos realistas basados en operacion diaria SJO)
INSERT INTO reservas (numero_reserva, marca_id, nombre_cliente, clase_vehiculo, tarifa_diaria, fecha_recogida, fecha_entrega, ubicacion_recogida, ubicacion_entrega, agencia_recomienda, es_vip, es_flight, estado) VALUES
('RES-001', 1, 'Carlos Mendez', 'Compacto', 45.00, '2026-06-21 08:30:00', '2026-06-24 08:30:00', 'SJOE71', 'SJOE71', 'Expedia 11617270', FALSE, TRUE, 'activa'),
('RES-002', 1, 'Maria Rodriguez', 'SUV', 85.00, '2026-06-21 09:00:00', '2026-06-23 09:00:00', 'SJOE71', 'SJOE71', 'Booking.com', FALSE, FALSE, 'activa'),
('RES-003', 2, 'John Smith', 'Intermedio', 55.00, '2026-06-21 07:45:00', '2026-06-22 07:45:00', 'SJOC61', 'SJOC61', 'Expedia 11617270', TRUE, TRUE, 'activa'),
('RES-004', 2, 'Ana Jimenez', 'Estandar', 50.00, '2026-06-21 10:30:00', '2026-06-25 10:30:00', 'SJOC61', 'SJOC61', 'Directo', FALSE, FALSE, 'activa'),
('RES-005', 3, 'Pedro Castro', 'Premium', 120.00, '2026-06-21 06:15:00', '2026-06-21 18:00:00', 'SJOE01', 'SJOE01', 'Despegar', TRUE, FALSE, 'activa'),
('RES-006', 1, 'Laura Gomez', 'Compacto', 42.00, '2026-06-21 14:00:00', '2026-06-22 14:00:00', 'SJOE71', 'SJOE71', 'Kayak', FALSE, TRUE, 'activa'),
('RES-007', 2, 'Roberto Diaz', 'Camioneta', 95.00, '2026-06-21 11:30:00', '2026-06-24 11:30:00', 'SJOC61', 'SJOC61', 'Expedia 11617270', FALSE, FALSE, 'no_show'),
('RES-008', 3, 'Sofia Herrera', 'Ejecutivo', 75.00, '2026-06-21 16:00:00', '2026-06-23 16:00:00', 'SJOE01', 'SJOE01', 'Hotel Barcelo', FALSE, FALSE, 'activa'),
('RES-009', 1, 'David Cordero', 'SUV', 88.00, '2026-06-21 05:30:00', '2026-06-22 05:30:00', 'SJOE71', 'SJOE71', 'Priceline', TRUE, TRUE, 'activa'),
('RES-010', 2, 'Marta Solis', 'Estandar', 48.00, '2026-06-21 12:00:00', '2026-06-22 12:00:00', 'SJOC61', 'SJOE71', 'Directo', FALSE, FALSE, 'activa'),
('RES-011', 1, 'Andres Mora', 'Compacto', 40.00, '2026-06-20 08:00:00', '2026-06-21 08:00:00', 'SJOE71', 'SJOE71', 'Expedia 11617270', FALSE, FALSE, 'no_show'),
('RES-012', 3, 'Karla Rojas', 'Intermedio', 52.00, '2026-06-20 15:00:00', '2026-06-23 15:00:00', 'SJOE01', 'SJOE01', 'Booking.com', FALSE, FALSE, 'completada'),
('RES-013', 2, 'Michael Brown', 'Premium', 130.00, '2026-06-20 09:30:00', '2026-06-22 09:30:00', 'SJOC61', 'SJOC61', 'Amex FHR', TRUE, FALSE, 'completada'),
('RES-014', 1, 'Cristina Vega', 'Estandar', 46.00, '2026-06-19 11:00:00', '2026-06-21 11:00:00', 'SJOE71', 'SJOE71', 'Directo', FALSE, FALSE, 'no_show'),
('RES-015', 3, 'Luis Sandoval', 'Camioneta', 92.00, '2026-06-19 07:00:00', '2026-06-22 07:00:00', 'SJOE01', 'SJOE01', 'Enterprise', TRUE, FALSE, 'completada');

-- No-show logs
INSERT INTO no_show_log (reserva_id, fecha_no_show, hora_bloque) VALUES
(7, '2026-06-21 11:30:00', '10:00-13:59'),
(11, '2026-06-21 08:00:00', '06:00-09:59'),
(14, '2026-06-21 11:00:00', '10:00-13:59');

=========================================================================
  PARTE 3: CONSULTAS - GAME PLAN DIARIO
=========================================================================

-- 3.1 GAME PLAN - Reservas activas del dia de hoy
-- Esta es la consulta principal que genera el reporte operativo
SELECT 
    ROW_NUMBER() OVER (PARTITION BY m.nombre_marca ORDER BY r.fecha_recogida) AS letra,
    r.numero_reserva AS 'No. Reserva',
    CONCAT(m.nombre_marca, '//') AS marca,
    CASE 
        WHEN r.es_flight = TRUE AND r.es_vip = TRUE THEN 
            CONCAT('⭐', LEFT(m.nombre_marca, 1), '-FL⭐')
        WHEN r.es_flight = TRUE THEN CONCAT('⭐', LEFT(m.nombre_marca, 1), '-FL⭐')
        WHEN r.es_vip = TRUE THEN CONCAT('⭐', LEFT(m.nombre_marca, 1), '⭐')
        ELSE ''
    END AS indicador,
    r.nombre_cliente,
    r.clase_vehiculo,
    CASE 
        WHEN r.es_vip = TRUE THEN '* MAIN VIP*'
        WHEN r.agencia_recomienda LIKE '%11617270%' THEN '* EXPEDIA*'
        ELSE ''
    END AS tipo,
    DATE_FORMAT(r.fecha_recogida, '%m/%d/%Y') AS fecha_recogida,
    DATE_FORMAT(r.fecha_entrega, '%m/%d/%Y') AS fecha_entrega,
    TIME_FORMAT(r.fecha_recogida, '%H:%i') AS hora_recogida,
    TIME_FORMAT(r.fecha_entrega, '%H:%i') AS hora_entrega,
    r.ubicacion_recogida,
    r.ubicacion_entrega,
    CONCAT('$', FORMAT(r.tarifa_diaria, 2)) AS tarifa,
    r.agencia_recomienda
FROM reservas r
JOIN marcas m ON r.marca_id = m.marca_id
WHERE DATE(r.fecha_recogida) = CURDATE()
    AND r.estado = 'activa'
ORDER BY m.nombre_marca, r.fecha_recogida;

=========================================================================
  PARTE 4: CONSULTAS - ANALISIS DE NO-SHOW
=========================================================================

-- 4.1 Tasa de No-Show por marca
SELECT 
    m.nombre_marca,
    COUNT(*) AS total_reservas,
    SUM(CASE WHEN r.estado = 'no_show' THEN 1 ELSE 0 END) AS no_shows,
    ROUND(
        SUM(CASE WHEN r.estado = 'no_show' THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 1
    ) AS '% No-Show'
FROM reservas r
JOIN marcas m ON r.marca_id = m.marca_id
GROUP BY m.nombre_marca;

-- 4.2 No-Show por bloque horario
SELECT 
    CASE 
        WHEN TIME(r.fecha_recogida) BETWEEN '06:00:00' AND '09:59:59' THEN 'Madrugada (06-09)'
        WHEN TIME(r.fecha_recogida) BETWEEN '10:00:00' AND '13:59:59' THEN 'Mediodia (10-13)'
        WHEN TIME(r.fecha_recogida) BETWEEN '14:00:00' AND '17:59:59' THEN 'Tarde (14-17)'
        ELSE 'Noche (18-23)'
    END AS bloque,
    COUNT(*) AS total_reservas,
    SUM(CASE WHEN r.estado = 'no_show' THEN 1 ELSE 0 END) AS no_shows,
    ROUND(
        SUM(CASE WHEN r.estado = 'no_show' THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 1
    ) AS tasa_no_show
FROM reservas r
GROUP BY bloque
ORDER BY tasa_no_show DESC;

-- 4.3 Clientes con historial de No-Show
SELECT 
    r.nombre_cliente,
    COUNT(*) AS reservas_totales,
    SUM(CASE WHEN r.estado = 'no_show' THEN 1 ELSE 0 END) AS no_shows,
    ROUND(
        SUM(CASE WHEN r.estado = 'no_show' THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 0
    ) AS 'tasa_%'
FROM reservas r
GROUP BY r.nombre_cliente
HAVING no_shows > 0
ORDER BY tasa_% DESC;

=========================================================================
  PARTE 5: CONSULTAS - REPORTES EJECUTIVOS
=========================================================================

-- 5.1 Resumen del dia por marca
SELECT 
    m.nombre_marca,
    COUNT(*) AS reservas_activas,
    ROUND(SUM(r.tarifa_diaria), 2) AS ingreso_estimado,
    ROUND(AVG(r.tarifa_diaria), 2) AS tarifa_promedio,
    COUNT(DISTINCT r.clase_vehiculo) AS clases_distintas
FROM reservas r
JOIN marcas m ON r.marca_id = m.marca_id
WHERE DATE(r.fecha_recogida) = CURDATE()
    AND r.estado = 'activa'
GROUP BY m.nombre_marca;

-- 5.2 Top 5 agencias de recomendacion
SELECT 
    r.agencia_recomienda,
    COUNT(*) AS reservas,
    ROUND(SUM(r.tarifa_diaria), 2) AS ingreso
FROM reservas r
WHERE r.agencia_recomienda IS NOT NULL AND r.agencia_recomienda != ''
GROUP BY r.agencia_recomienda
ORDER BY reservas DESC
LIMIT 5;

-- 5.3 Reservas VIP y Expedia
SELECT 
    r.numero_reserva,
    LEFT(m.nombre_marca, 1) AS marca,
    r.nombre_cliente,
    r.clase_vehiculo,
    TIME_FORMAT(r.fecha_recogida, '%H:%i') AS hora,
    CASE 
        WHEN r.es_vip = TRUE AND r.agencia_recomienda LIKE '%11617270%' THEN 'VIP + EXPEDIA'
        WHEN r.es_vip = TRUE THEN '* MAIN VIP*'
        WHEN r.agencia_recomienda LIKE '%11617270%' THEN '* EXPEDIA*'
    END AS categoria
FROM reservas r
JOIN marcas m ON r.marca_id = m.marca_id
WHERE (r.es_vip = TRUE OR r.agencia_recomienda LIKE '%11617270%')
ORDER BY categoria, m.nombre_marca;

=========================================================================
  PARTE 6: CONSULTAS - ESTADISTICAS AVANZADAS
=========================================================================

-- 6.1 Distribucion de tarifas por clase de vehiculo
SELECT 
    r.clase_vehiculo,
    COUNT(*) AS cantidad,
    ROUND(MIN(r.tarifa_diaria), 2) AS tarifa_min,
    ROUND(AVG(r.tarifa_diaria), 2) AS tarifa_promedio,
    ROUND(MAX(r.tarifa_diaria), 2) AS tarifa_max,
    ROUND(AVG(r.tarifa_diaria) - MIN(r.tarifa_diaria), 2) AS rango
FROM reservas r
GROUP BY r.clase_vehiculo
ORDER BY tarifa_promedio DESC;

-- 6.2 Duracion promedio de renta por marca
SELECT 
    m.nombre_marca,
    COUNT(*) AS total,
    ROUND(AVG(TIMESTAMPDIFF(HOUR, r.fecha_recogida, r.fecha_entrega) / 24), 1) AS dias_promedio,
    ROUND(SUM(r.tarifa_diaria * TIMESTAMPDIFF(DAY, r.fecha_recogida, r.fecha_entrega)), 2) AS ingreso_total
FROM reservas r
JOIN marcas m ON r.marca_id = m.marca_id
WHERE r.estado IN ('activa', 'completada')
GROUP BY m.nombre_marca;

-- 6.3 Tasa de conversion (activas vs no-show)
SELECT 
    r.estado,
    COUNT(*) AS cantidad,
    ROUND(COUNT(*) * 100.0 / (SELECT COUNT(*) FROM reservas), 1) AS porcentaje
FROM reservas r
GROUP BY r.estado
ORDER BY cantidad DESC;

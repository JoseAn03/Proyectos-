=========================================================================
  SCRIPT: 01_creacion_tablas_reservas.sql
  PROYECTO: Base de Datos de Reservas - ANC Renta Car
  DESCRIPCION: Creacion de tablas para gestion de reservas diarias
  TECNOLOGIA: SQL (MySQL / PostgreSQL compatible)
=========================================================================

-- -----------------------------------------------------------
-- TABLA: marcas (Alamo, Enterprise, National)
-- -----------------------------------------------------------
CREATE TABLE marcas (
    marca_id INT PRIMARY KEY AUTO_INCREMENT,
    nombre_marca VARCHAR(50) NOT NULL,
    codigo_ubicacion VARCHAR(10) NOT NULL,
    UNIQUE KEY uk_marca (nombre_marca, codigo_ubicacion)
);

-- -----------------------------------------------------------
-- TABLA: reservas
-- -----------------------------------------------------------
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
    estado ENUM('activa', 'completada', 'no_show', 'cancelada') DEFAULT 'activa',
    fecha_creacion TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (marca_id) REFERENCES marcas(marca_id),
    INDEX idx_fecha_recogida (fecha_recogida),
    INDEX idx_estado (estado),
    INDEX idx_marca_fecha (marca_id, fecha_recogida)
);

-- -----------------------------------------------------------
-- TABLA: no_show_log (registro de no-shows)
-- -----------------------------------------------------------
CREATE TABLE no_show_log (
    no_show_id INT PRIMARY KEY AUTO_INCREMENT,
    reserva_id INT NOT NULL,
    fecha_no_show DATETIME NOT NULL,
    hora_bloque VARCHAR(20),
    motivo VARCHAR(200),
    FOREIGN KEY (reserva_id) REFERENCES reservas(reserva_id),
    INDEX idx_fecha_ns (fecha_no_show)
);

-- -----------------------------------------------------------
-- TABLA: reportes_diarios
-- -----------------------------------------------------------
CREATE TABLE reportes_diarios (
    reporte_id INT PRIMARY KEY AUTO_INCREMENT,
    fecha_reporte DATE NOT NULL,
    marca_id INT NOT NULL,
    total_reservas INT DEFAULT 0,
    total_no_show INT DEFAULT 0,
    total_activas INT DEFAULT 0,
    total_canceladas INT DEFAULT 0,
    ingresos_estimados DECIMAL(12,2) DEFAULT 0,
    generado_por VARCHAR(50) DEFAULT 'script_automatico',
    fecha_generacion TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (marca_id) REFERENCES marcas(marca_id),
    UNIQUE KEY uk_reporte_diario (fecha_reporte, marca_id)
);

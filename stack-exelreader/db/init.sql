--CREAR BASE DE DATOS
CREATE DATABASE `agenda` COLLATE 'utf16_spanish_ci';

--ANNADIR TABLA PERSONAS
CREATE TABLE personas (
  `nombre` varchar(100) NOT NULL,
  `apellidos` varchar(300) NOT NULL,
  `email` varchar(100),
  `telefono` varchar(12),
  `genero` enum('FEMENINO','MASCULINO','NEUTRO','OTRO') NOT NULL,
  `id` int NOT NULL AUTO_INCREMENT PRIMARY KEY
) ENGINE='InnoDB';

--AGREGAR PERSONA
INSERT INTO `personas` (`nombre`, `apellidos`, `email`, `telefono`, `genero`, `id`)
VALUES ('Antonio', 'Martinez Colmenero', 'antonio23@gmail.com', '677744201', '\"MASCULINO\",', '0');
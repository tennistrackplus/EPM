-- ============================================================
-- DRACO PLANNING · Script de creación del esquema de control
-- ============================================================
-- Válido tanto para BigQuery (GoogleSQL) como para Snowflake: la
-- sintaxis usada aquí (CREATE SCHEMA, CREATE TABLE, tipos STRING/
-- INTEGER/BOOLEAN/DATE/DATETIME/TIMESTAMP) es común a ambos motores.
-- Sustituye {PROJECT} por tu Project ID de GCP, o simplemente omite
-- esa parte del nombre en Snowflake (usa tu base de datos, ej. DRACO).
--
-- Nota: la propia app ejecuta este bootstrap automáticamente la
-- primera vez que conectas (ver js/schema.js), así que normalmente
-- no hace falta correr este script a mano — se deja aquí como
-- referencia / para revisión por un DBA.
-- ============================================================

-- 1. Dataset de control -----------------------------------------------------
CREATE SCHEMA IF NOT EXISTS `{PROJECT}.DRACO_CONTROL`
OPTIONS (
  description = 'Dataset de control de Draco Planning (proyectos, dimensiones, cubos)'
);

-- 2. Tabla de proyectos ------------------------------------------------------
CREATE TABLE IF NOT EXISTS `{PROJECT}.DRACO_CONTROL.PROYECTOS` (
  PROYECTO_ID     STRING NOT NULL,   -- identificador único (uuid)
  PROYECTO        STRING NOT NULL,   -- nombre del proyecto (modelo)
  DESCRIPCION     STRING,
  DATASET         STRING NOT NULL,   -- dataset físico DRACO_<PROYECTO> creado en BigQuery
  USUARIO         STRING,
  FECHA_CREACION  TIMESTAMP
);

-- 3. Tabla de dimensiones ----------------------------------------------------
CREATE TABLE IF NOT EXISTS `{PROJECT}.DRACO_CONTROL.DIMENSIONES` (
  DIMENSION_ID        STRING NOT NULL,
  PROYECTO_ID         STRING NOT NULL,   -- FK -> PROYECTOS.PROYECTO_ID
  DIMENSION           STRING NOT NULL,   -- nombre de la dimensión
  DESCRIPCION         STRING,
  TABLA               STRING NOT NULL,   -- tabla física DRACO_<DIMENSION>
  CAMPOS_JSON         STRING,            -- definición de campos (nombre/tipo/clave) en JSON
  USUARIO             STRING,
  FECHA_CREACION      TIMESTAMP,
  FECHA_MODIFICACION  TIMESTAMP
);

-- 4. Tabla de cubos -----------------------------------------------------------
CREATE TABLE IF NOT EXISTS `{PROJECT}.DRACO_CONTROL.CUBOS` (
  CUBO_ID             STRING NOT NULL,
  PROYECTO_ID         STRING NOT NULL,   -- FK -> PROYECTOS.PROYECTO_ID
  CUBOS               STRING NOT NULL,   -- nombre del cubo
  DESCRIPCION         STRING,
  TABLA               STRING NOT NULL,   -- tabla física DRACO_<CUBO>
  CAMPOS_JSON         STRING,
  USUARIO             STRING,
  FECHA_CREACION      TIMESTAMP,
  FECHA_MODIFICACION  TIMESTAMP
);

-- ============================================================
-- Nota: las tablas físicas de cada dimensión/cubo (DRACO_<nombre>)
-- se crean dinámicamente desde la app, dentro del dataset propio
-- de cada proyecto (DRACO_<PROYECTO>), con las columnas que el
-- usuario defina en el diseñador de campos.
-- ============================================================

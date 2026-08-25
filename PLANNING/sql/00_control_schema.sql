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
  CAMPOS_JSON         STRING,            -- { dimensions:[...], measures:[...] } en JSON
  USUARIO             STRING,
  FECHA_CREACION      TIMESTAMP,
  FECHA_MODIFICACION  TIMESTAMP,
  MODELO_YAML_PATH    STRING,            -- ruta del modelo semántico dentro del stage (ver js/semantic-model.js)
  MODELO_YAML_FECHA   TIMESTAMP          -- última vez que se generó/subió el YAML
);

-- 4bis. Migración para instalaciones ya existentes (idempotente, mismo
--       comando en BigQuery y Snowflake):
-- ALTER TABLE `{PROJECT}.DRACO_CONTROL.CUBOS` ADD COLUMN IF NOT EXISTS MODELO_YAML_PATH STRING;
-- ALTER TABLE `{PROJECT}.DRACO_CONTROL.CUBOS` ADD COLUMN IF NOT EXISTS MODELO_YAML_FECHA TIMESTAMP;

-- 5. Tabla de jerarquías -------------------------------------------------------
CREATE TABLE IF NOT EXISTS `{PROJECT}.DRACO_CONTROL.JERARQUIAS` (
  JERARQUIA_ID        STRING NOT NULL,
  DIMENSION_ID        STRING NOT NULL,   -- FK -> DIMENSIONES.DIMENSION_ID
  PROYECTO_ID         STRING NOT NULL,
  JERARQUIA           STRING NOT NULL,   -- nombre de la jerarquía
  NIVELES_JSON        STRING,            -- array ordenado de atributos, nivel superior primero
  USUARIO             STRING,
  FECHA_CREACION      TIMESTAMP,
  FECHA_MODIFICACION  TIMESTAMP
);

-- 6. Interfaces (cargas de datos) ---------------------------------------------
CREATE TABLE IF NOT EXISTS `{PROJECT}.DRACO_CONTROL.INTERFACES` (
  INTERFAZ_ID           STRING NOT NULL,
  PROYECTO_ID           STRING NOT NULL,   -- FK -> PROYECTOS.PROYECTO_ID
  INTERFAZ              STRING NOT NULL,   -- nombre corto, sin espacios (se usa como nombre de función Python)
  TIPO                  STRING NOT NULL,   -- 'TABLA' | 'FICHERO'
  ORIGEN                STRING,            -- nombre de la tabla origen, o tipo de fichero (csv/xlsx/json/fixed)
  CONECTOR              STRING,            -- 'bigquery' | 'snowflake' (solo cuando TIPO = 'TABLA')
  CUBO_ID               STRING NOT NULL,   -- FK -> CUBOS.CUBO_ID (destino)
  MAPPING_MODE          STRING,            -- 'VISUAL' | 'CODIGO'
  MAPPING_CODE          STRING,            -- código Python cuando MAPPING_MODE = 'CODIGO'
  INPUT_TRANSFORM_CODE  STRING,
  OUTPUT_TRANSFORM_CODE STRING,
  USUARIO               STRING,
  FECHA_CREACION        TIMESTAMP,
  FECHA_MODIFICACION    TIMESTAMP
);

-- 7. Características del origen de cada interfaz (separadores, encoding...) --
CREATE TABLE IF NOT EXISTS `{PROJECT}.DRACO_CONTROL.INTERFACES_VALUES` (
  PROYECTO_ID     STRING NOT NULL,
  INTERFAZ_ID     STRING NOT NULL,   -- FK -> INTERFACES.INTERFAZ_ID
  CARACTERISTICA  STRING NOT NULL,   -- ej. SEPARADOR_MILES, SEPARADOR_DECIMAL, SEPARADOR_CAMPO, CODIFICACION...
  VALOR           STRING
);

-- 8. Campos de entrada esperados por cada interfaz ----------------------------
CREATE TABLE IF NOT EXISTS `{PROJECT}.DRACO_CONTROL.INTERFACES_INPUT` (
  PROYECTO_ID  STRING NOT NULL,
  INTERFAZ_ID  STRING NOT NULL,   -- FK -> INTERFACES.INTERFAZ_ID
  CAMPO        STRING NOT NULL,
  TIPO         STRING,            -- STRING/INTEGER/FLOAT/NUMERIC/BOOLEAN/DATE/DATETIME/TIMESTAMP
  ORDEN        INTEGER
);

-- 9. Filtros sobre los campos de entrada --------------------------------------
CREATE TABLE IF NOT EXISTS `{PROJECT}.DRACO_CONTROL.INTERFACES_INPUT_FILTERS` (
  PROYECTO_ID  STRING NOT NULL,
  INTERFAZ_ID  STRING NOT NULL,   -- FK -> INTERFACES.INTERFAZ_ID
  CAMPO        STRING NOT NULL,
  TIPO         STRING NOT NULL,   -- 'VALOR' (constante) | 'VARIABLE'
  VALOR        STRING             -- valor constante, o nombre de la variable
);

-- 10. Mapeo campo a campo (origen -> cubo destino) ----------------------------
CREATE TABLE IF NOT EXISTS `{PROJECT}.DRACO_CONTROL.INTERFACES_MAPPING` (
  PROYECTO_ID    STRING NOT NULL,
  INTERFAZ_ID    STRING NOT NULL,  -- FK -> INTERFACES.INTERFAZ_ID
  CAMPO_DESTINO  STRING NOT NULL,  -- id del campo del cubo (colId de dimensión o identificador de medida)
  TIPO           STRING NOT NULL,  -- 'CAMPO' | 'CONSTANTE' | 'VARIABLE' | 'FORMULA' | 'FUNCION'
  VALOR          STRING,           -- nombre del campo origen / valor constante / nombre variable / expresión
  CODIGO         STRING            -- código Python, solo cuando TIPO = 'FUNCION'
);

-- 11. Flujos de carga (procesos) -----------------------------------------------
CREATE TABLE IF NOT EXISTS `{PROJECT}.DRACO_CONTROL.FLUJOS` (
  FLUJO_ID            STRING NOT NULL,
  PROYECTO_ID         STRING NOT NULL,   -- FK -> PROYECTOS.PROYECTO_ID
  FLUJO               STRING NOT NULL,   -- nombre del proceso
  TIPO                STRING NOT NULL,   -- 'AUTOMATICO' | 'MANUAL'
  SCHEDULE_JSON       STRING,            -- planificación (solo AUTOMATICO), en JSON
  SCREEN_TITLE        STRING,            -- título de la pantalla de variables (solo MANUAL)
  USUARIO             STRING,
  FECHA_CREACION      TIMESTAMP,
  FECHA_MODIFICACION  TIMESTAMP
);

-- 12. Cadena de interfaces de cada flujo, en orden -----------------------------
CREATE TABLE IF NOT EXISTS `{PROJECT}.DRACO_CONTROL.FLUJOS_INTERFACES` (
  PROYECTO_ID  STRING NOT NULL,
  FLUJO_ID     STRING NOT NULL,   -- FK -> FLUJOS.FLUJO_ID
  PASO_ID      STRING NOT NULL,   -- id del paso dentro de la cadena
  INTERFAZ_ID  STRING NOT NULL,   -- FK -> INTERFACES.INTERFAZ_ID
  ORDEN        INTEGER
);

-- 13. Asignación de variables por paso de la cadena (fichero/filtro/mapeo) -----
CREATE TABLE IF NOT EXISTS `{PROJECT}.DRACO_CONTROL.FLUJOS_INTERFACES_TARGETS` (
  PROYECTO_ID  STRING NOT NULL,
  FLUJO_ID     STRING NOT NULL,
  PASO_ID      STRING NOT NULL,   -- FK -> FLUJOS_INTERFACES.PASO_ID
  GRUPO        STRING NOT NULL,   -- 'FILE' | 'FILTER' | 'MAPPING'
  CLAVE        STRING NOT NULL,   -- clave dentro del grupo (ruta, nombre de campo...)
  TIPO         STRING NOT NULL,   -- 'CONSTANTE' | 'VARIABLE'
  VALOR        STRING             -- valor constante, o nombre de la variable de pantalla
);

-- 14. Bloques de la pantalla de variables (solo flujos manuales) --------------
CREATE TABLE IF NOT EXISTS `{PROJECT}.DRACO_CONTROL.FLUJOS_SCREEN_BLOCKS` (
  PROYECTO_ID  STRING NOT NULL,
  FLUJO_ID     STRING NOT NULL,
  BLOQUE_ID    STRING NOT NULL,
  TIPO         STRING NOT NULL,   -- 'VARIABLE' | 'FRAME' | 'TEXTO' | 'SKIP' (espacio en blanco) | 'ULINE' (línea separadora)
  ORDEN        INTEGER,
  TITULO       STRING,            -- solo FRAME
  CONTENIDO    STRING             -- solo TEXTO
);

-- 15. Variables de la pantalla (sueltas o dentro de un frame) -----------------
CREATE TABLE IF NOT EXISTS `{PROJECT}.DRACO_CONTROL.FLUJOS_SCREEN_VARIABLES` (
  PROYECTO_ID  STRING NOT NULL,
  FLUJO_ID     STRING NOT NULL,
  VARIABLE_ID  STRING NOT NULL,
  BLOQUE_ID    STRING NOT NULL,   -- FK -> FLUJOS_SCREEN_BLOCKS.BLOQUE_ID (el bloque VARIABLE o el FRAME contenedor)
  NOMBRE       STRING NOT NULL,   -- nombre técnico
  ETIQUETA     STRING,
  TIPO         STRING,
  SELECT_MODE  STRING,            -- 'unico' | 'rango' | 'multiple' | 'cualquiera' (estilo select-options SAP; de momento solo informativo)
  ORDEN        INTEGER
);

-- 16. Ejecuciones de flujo (cabecera) ------------------------------------------
CREATE TABLE IF NOT EXISTS `{PROJECT}.DRACO_CONTROL.FLUJOS_RUNS` (
  RUN_ID          STRING NOT NULL,
  PROYECTO_ID     STRING NOT NULL,
  FLUJO_ID        STRING NOT NULL,   -- FK -> FLUJOS.FLUJO_ID
  ESTADO          STRING NOT NULL,   -- 'PENDIENTE' | 'EN_CURSO' | 'OK' | 'ERROR'
  VARIABLES_JSON  STRING,            -- variables de pantalla usadas en esta ejecución
  MENSAJE         STRING,            -- error global, si lo hay
  USUARIO         STRING,
  FECHA_INICIO    TIMESTAMP,
  FECHA_FIN       TIMESTAMP
);

-- 17. Ejecuciones de flujo (estado por paso, para el monitor) -------------------
CREATE TABLE IF NOT EXISTS `{PROJECT}.DRACO_CONTROL.FLUJOS_RUN_STEPS` (
  RUN_ID       STRING NOT NULL,   -- FK -> FLUJOS_RUNS.RUN_ID
  PASO_ID      STRING NOT NULL,   -- FK -> FLUJOS_INTERFACES.PASO_ID
  ORDEN        INTEGER,
  INTERFAZ_ID  STRING,
  ESTADO       STRING NOT NULL,   -- 'PENDIENTE' | 'EN_CURSO' | 'OK' | 'ERROR'
  FILAS        INTEGER,
  MENSAJE      STRING,
  FECHA_INICIO TIMESTAMP,
  FECHA_FIN    TIMESTAMP
);

-- 18. Buffer de subida de ficheros a Snowflake (solo Snowflake) ---------------
-- El navegador no puede ejecutar el comando PUT nativo de Snowflake (requiere
-- un driver/cliente local), así que sube el fichero troceado en base64 por
-- SQL normal (INSERT) y un stored procedure Python (ver
-- sql/02_snowflake_file_upload.sql) lo reensambla y lo escribe en el stage
-- con session.file.put_stream(). Esta tabla es solo el buffer temporal: cada
-- fila se borra en cuanto el procedure termina de reensamblar el fichero.
CREATE TABLE IF NOT EXISTS `{PROJECT}.DRACO_CONTROL.FILE_UPLOAD_CHUNKS` (
  UPLOAD_ID    STRING NOT NULL,   -- uuid generado en el navegador por subida
  CHUNK_INDEX  INTEGER NOT NULL,  -- orden del trozo (0-based)
  CHUNK_B64    STRING NOT NULL,   -- contenido del trozo en base64
  FECHA_CARGA  TIMESTAMP
);

-- ============================================================
-- Nota: las tablas físicas de cada dimensión/cubo (DRACO_<nombre>)
-- se crean dinámicamente desde la app, dentro del dataset propio
-- de cada proyecto (DRACO_<PROYECTO>), con las columnas que el
-- usuario defina en el diseñador de campos.
-- ============================================================

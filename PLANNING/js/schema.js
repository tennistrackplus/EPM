/**
 * ============================================================
 * DRACO PLANNING — ESQUEMA DE CONTROL
 * ============================================================
 * Crea (si no existen) el dataset/esquema de control y las tablas
 * maestras de Draco en el motor activo (BigQuery o Snowflake):
 *   - DRACO_CONTROL.PROYECTOS
 *   - DRACO_CONTROL.DIMENSIONES
 *   - DRACO_CONTROL.CUBOS
 */
const DracoSchema = {

    ddl(table) {
        const t = Provider.qualifyControl(table);
        const ddl = {
            PROYECTOS: `
                CREATE TABLE IF NOT EXISTS ${t} (
                    PROYECTO_ID STRING NOT NULL,
                    PROYECTO STRING NOT NULL,
                    DESCRIPCION STRING,
                    DATASET STRING NOT NULL,
                    USUARIO STRING,
                    FECHA_CREACION TIMESTAMP
                )`,
            DIMENSIONES: `
                CREATE TABLE IF NOT EXISTS ${t} (
                    DIMENSION_ID STRING NOT NULL,
                    PROYECTO_ID STRING NOT NULL,
                    DIMENSION STRING NOT NULL,
                    DESCRIPCION STRING,
                    TABLA STRING NOT NULL,
                    CAMPOS_JSON STRING,
                    USUARIO STRING,
                    FECHA_CREACION TIMESTAMP,
                    FECHA_MODIFICACION TIMESTAMP
                )`,
            CUBOS: `
                CREATE TABLE IF NOT EXISTS ${t} (
                    CUBO_ID STRING NOT NULL,
                    PROYECTO_ID STRING NOT NULL,
                    CUBOS STRING NOT NULL,
                    DESCRIPCION STRING,
                    TABLA STRING NOT NULL,
                    CAMPOS_JSON STRING,
                    USUARIO STRING,
                    FECHA_CREACION TIMESTAMP,
                    FECHA_MODIFICACION TIMESTAMP,
                    MODELO_YAML_PATH STRING,       -- ruta del modelo semántico en el stage (ver js/semantic-model.js)
                    MODELO_YAML_FECHA TIMESTAMP    -- última vez que se generó/subió el YAML
                )`,
            JERARQUIAS: `
                CREATE TABLE IF NOT EXISTS ${t} (
                    JERARQUIA_ID STRING NOT NULL,
                    DIMENSION_ID STRING NOT NULL,
                    PROYECTO_ID STRING NOT NULL,
                    JERARQUIA STRING NOT NULL,
                    NIVELES_JSON STRING,
                    USUARIO STRING,
                    FECHA_CREACION TIMESTAMP,
                    FECHA_MODIFICACION TIMESTAMP
                )`,
            INTERFACES: `
                CREATE TABLE IF NOT EXISTS ${t} (
                    INTERFAZ_ID STRING NOT NULL,
                    PROYECTO_ID STRING NOT NULL,
                    INTERFAZ STRING NOT NULL,
                    TIPO STRING NOT NULL,
                    ORIGEN STRING,
                    CONECTOR STRING,
                    CUBO_ID STRING NOT NULL,
                    MAPPING_MODE STRING,
                    MAPPING_CODE STRING,
                    INPUT_TRANSFORM_CODE STRING,
                    OUTPUT_TRANSFORM_CODE STRING,
                    USUARIO STRING,
                    FECHA_CREACION TIMESTAMP,
                    FECHA_MODIFICACION TIMESTAMP
                )`,
            INTERFACES_VALUES: `
                CREATE TABLE IF NOT EXISTS ${t} (
                    PROYECTO_ID STRING NOT NULL,
                    INTERFAZ_ID STRING NOT NULL,
                    CARACTERISTICA STRING NOT NULL,
                    VALOR STRING
                )`,
            INTERFACES_INPUT: `
                CREATE TABLE IF NOT EXISTS ${t} (
                    PROYECTO_ID STRING NOT NULL,
                    INTERFAZ_ID STRING NOT NULL,
                    CAMPO STRING NOT NULL,
                    TIPO STRING,
                    ORDEN INTEGER
                )`,
            INTERFACES_INPUT_FILTERS: `
                CREATE TABLE IF NOT EXISTS ${t} (
                    PROYECTO_ID STRING NOT NULL,
                    INTERFAZ_ID STRING NOT NULL,
                    CAMPO STRING NOT NULL,
                    TIPO STRING NOT NULL,
                    VALOR STRING
                )`,
            INTERFACES_MAPPING: `
                CREATE TABLE IF NOT EXISTS ${t} (
                    PROYECTO_ID STRING NOT NULL,
                    INTERFAZ_ID STRING NOT NULL,
                    CAMPO_DESTINO STRING NOT NULL,
                    TIPO STRING NOT NULL,
                    VALOR STRING,
                    CODIGO STRING
                )`,
            FLUJOS: `
                CREATE TABLE IF NOT EXISTS ${t} (
                    FLUJO_ID STRING NOT NULL,
                    PROYECTO_ID STRING NOT NULL,
                    FLUJO STRING NOT NULL,
                    TIPO STRING NOT NULL,
                    SCHEDULE_JSON STRING,
                    SCREEN_TITLE STRING,
                    USUARIO STRING,
                    FECHA_CREACION TIMESTAMP,
                    FECHA_MODIFICACION TIMESTAMP
                )`,
            FLUJOS_INTERFACES: `
                CREATE TABLE IF NOT EXISTS ${t} (
                    PROYECTO_ID STRING NOT NULL,
                    FLUJO_ID STRING NOT NULL,
                    PASO_ID STRING NOT NULL,
                    INTERFAZ_ID STRING NOT NULL,
                    ORDEN INTEGER
                )`,
            FLUJOS_INTERFACES_TARGETS: `
                CREATE TABLE IF NOT EXISTS ${t} (
                    PROYECTO_ID STRING NOT NULL,
                    FLUJO_ID STRING NOT NULL,
                    PASO_ID STRING NOT NULL,
                    GRUPO STRING NOT NULL,
                    CLAVE STRING NOT NULL,
                    TIPO STRING NOT NULL,
                    VALOR STRING
                )`,
            FLUJOS_SCREEN_BLOCKS: `
                CREATE TABLE IF NOT EXISTS ${t} (
                    PROYECTO_ID STRING NOT NULL,
                    FLUJO_ID STRING NOT NULL,
                    BLOQUE_ID STRING NOT NULL,
                    TIPO STRING NOT NULL,
                    ORDEN INTEGER,
                    TITULO STRING,
                    CONTENIDO STRING
                )`,
            FLUJOS_SCREEN_VARIABLES: `
                CREATE TABLE IF NOT EXISTS ${t} (
                    PROYECTO_ID STRING NOT NULL,
                    FLUJO_ID STRING NOT NULL,
                    VARIABLE_ID STRING NOT NULL,
                    BLOQUE_ID STRING NOT NULL,
                    NOMBRE STRING NOT NULL,
                    ETIQUETA STRING,
                    TIPO STRING,
                    ORDEN INTEGER
                )`,
            FLUJOS_RUNS: `
                CREATE TABLE IF NOT EXISTS ${t} (
                    RUN_ID STRING NOT NULL,
                    PROYECTO_ID STRING NOT NULL,
                    FLUJO_ID STRING NOT NULL,
                    ESTADO STRING NOT NULL,
                    VARIABLES_JSON STRING,
                    MENSAJE STRING,
                    USUARIO STRING,
                    FECHA_INICIO TIMESTAMP,
                    FECHA_FIN TIMESTAMP
                )`,
            FLUJOS_RUN_STEPS: `
                CREATE TABLE IF NOT EXISTS ${t} (
                    RUN_ID STRING NOT NULL,
                    PASO_ID STRING NOT NULL,
                    ORDEN INTEGER,
                    INTERFAZ_ID STRING,
                    ESTADO STRING NOT NULL,
                    FILAS INTEGER,
                    MENSAJE STRING,
                    FECHA_INICIO TIMESTAMP,
                    FECHA_FIN TIMESTAMP
                )`,
            WORKFLOWS: `
                CREATE TABLE IF NOT EXISTS ${t} (
                    WORKFLOW_ID STRING NOT NULL,
                    PROYECTO_ID STRING NOT NULL,
                    WORKFLOW STRING NOT NULL,
                    DESCRIPCION STRING,
                    USUARIO STRING,
                    FECHA_CREACION TIMESTAMP,
                    FECHA_MODIFICACION TIMESTAMP
                )`,
            WORKFLOWS_PASOS: `
                CREATE TABLE IF NOT EXISTS ${t} (
                    PASO_ID STRING NOT NULL,
                    WORKFLOW_ID STRING NOT NULL,
                    PROYECTO_ID STRING NOT NULL,
                    PASO STRING NOT NULL,
                    ORDEN INTEGER,
                    ES_PASO0 BOOLEAN,
                    INICIO_TIPO STRING NOT NULL,
                    INICIO_FECHA STRING,
                    REVISION BOOLEAN,
                    FIN_TIPO STRING NOT NULL,
                    FIN_FECHA STRING,
                    DRIVER_DIMENSION_ID STRING,
                    DRIVER_MODO STRING
                )`,
            WORKFLOWS_PASOS_DRIVER_VALORES: `
                CREATE TABLE IF NOT EXISTS ${t} (
                    PROYECTO_ID STRING NOT NULL,
                    PASO_ID STRING NOT NULL,
                    VALOR STRING NOT NULL
                )`,
            WORKFLOWS_PASOS_VARIABLES: `
                CREATE TABLE IF NOT EXISTS ${t} (
                    PROYECTO_ID STRING NOT NULL,
                    PASO_ID STRING NOT NULL,
                    VARIABLE_ID STRING NOT NULL,
                    NOMBRE STRING NOT NULL,
                    ETIQUETA STRING,
                    TIPO STRING,
                    ORDEN INTEGER
                )`,
            WORKFLOWS_PASOS_BLOQUES: `
                CREATE TABLE IF NOT EXISTS ${t} (
                    PROYECTO_ID STRING NOT NULL,
                    PASO_ID STRING NOT NULL,
                    BLOQUE_ID STRING NOT NULL,
                    TITULO STRING,
                    ORDEN INTEGER
                )`,
            WORKFLOWS_PASOS_TAREAS: `
                CREATE TABLE IF NOT EXISTS ${t} (
                    PROYECTO_ID STRING NOT NULL,
                    PASO_ID STRING NOT NULL,
                    BLOQUE_ID STRING NOT NULL,
                    TAREA_ID STRING NOT NULL,
                    TIPO STRING NOT NULL,
                    NOMBRE STRING,
                    DESCRIPCION STRING,
                    REF_ID STRING,
                    REF_NOMBRE STRING,
                    ORDEN INTEGER
                )`,
            WORKFLOWS_PASOS_TAREAS_VALORES: `
                CREATE TABLE IF NOT EXISTS ${t} (
                    PROYECTO_ID STRING NOT NULL,
                    TAREA_ID STRING NOT NULL,
                    CLAVE STRING NOT NULL,
                    ETIQUETA STRING,
                    TIPO STRING NOT NULL,
                    VALOR STRING,
                    OCULTAR BOOLEAN
                )`,
            WORKFLOWS_RUNS: `
                CREATE TABLE IF NOT EXISTS ${t} (
                    RUN_ID STRING NOT NULL,
                    WORKFLOW_ID STRING NOT NULL,
                    PROYECTO_ID STRING NOT NULL,
                    NOMBRE STRING NOT NULL,
                    DESCRIPCION STRING,
                    ESTADO STRING NOT NULL,
                    USUARIO STRING,
                    FECHA_CREACION TIMESTAMP,
                    FECHA_FIN TIMESTAMP
                )`,
            WORKFLOWS_RUNS_INSTANCIAS: `
                CREATE TABLE IF NOT EXISTS ${t} (
                    RUN_ID STRING NOT NULL,
                    PASO_ID STRING NOT NULL,
                    INSTANCIA_ID STRING NOT NULL,
                    ORDEN INTEGER,
                    DRIVER_VALOR STRING,
                    ASIGNADO STRING,
                    ESTADO STRING NOT NULL,
                    FECHA_PROGRAMADA STRING,
                    FECHA_INICIO TIMESTAMP,
                    FECHA_FIN TIMESTAMP
                )`,
            WORKFLOWS_RUNS_VARIABLES: `
                CREATE TABLE IF NOT EXISTS ${t} (
                    RUN_ID STRING NOT NULL,
                    INSTANCIA_ID STRING NOT NULL,
                    NOMBRE STRING NOT NULL,
                    VALOR STRING
                )`,
            // Actualización de tablas: mantenimiento manual de una tabla física
            // (dimensión, cubo, o cualquier otra tabla del esquema del proyecto)
            // con pantalla de variables + validaciones por campo. Todo el diseño
            // (variables de pantalla y configuración de cada campo) se guarda
            // como JSON, igual que CUBOS.CAMPOS_JSON. Ver js/table-updates.js.
            ACTUALIZACIONES: `
                CREATE TABLE IF NOT EXISTS ${t} (
                    ACTUALIZACION_ID STRING NOT NULL,
                    PROYECTO_ID STRING NOT NULL,
                    NOMBRE STRING NOT NULL,
                    DESCRIPCION STRING,
                    TABLA STRING NOT NULL,       -- tabla física dentro del esquema del proyecto
                    VARIABLES_JSON STRING,       -- [{id,name,label,type}]  (pantalla de selección)
                    CAMPOS_JSON STRING,          -- [{name,description,order,filter,validation,allowEmpty,showText,searchHelp}]
                    USUARIO STRING,
                    FECHA_CREACION TIMESTAMP,
                    FECHA_MODIFICACION TIMESTAMP
                )`,
            PARAMETRIZACIONES: `
                CREATE TABLE IF NOT EXISTS ${t} (
                    PARAMETRIZACION_ID STRING NOT NULL,
                    PROYECTO_ID STRING NOT NULL,
                    NOMBRE STRING NOT NULL,
                    DESCRIPCION STRING,
                    TABLA STRING NOT NULL,       -- tabla física DRACO_PARAM_<nombre>
                    CAMPOS_JSON STRING,          -- definición de campos (nombre/descripción/tipo/clave) en JSON
                    USUARIO STRING,
                    FECHA_CREACION TIMESTAMP,
                    FECHA_MODIFICACION TIMESTAMP
                )`
        };
        return ddl[table];
    },

    TABLES: ["PROYECTOS", "DIMENSIONES", "CUBOS", "JERARQUIAS", "PARAMETRIZACIONES",
             "INTERFACES", "INTERFACES_VALUES", "INTERFACES_INPUT", "INTERFACES_INPUT_FILTERS", "INTERFACES_MAPPING",
             "FLUJOS", "FLUJOS_INTERFACES", "FLUJOS_INTERFACES_TARGETS", "FLUJOS_SCREEN_BLOCKS", "FLUJOS_SCREEN_VARIABLES",
             "FLUJOS_RUNS", "FLUJOS_RUN_STEPS",
             "WORKFLOWS", "WORKFLOWS_PASOS", "WORKFLOWS_PASOS_DRIVER_VALORES", "WORKFLOWS_PASOS_VARIABLES",
             "WORKFLOWS_PASOS_BLOQUES", "WORKFLOWS_PASOS_TAREAS", "WORKFLOWS_PASOS_TAREAS_VALORES",
             "WORKFLOWS_RUNS", "WORKFLOWS_RUNS_INSTANCIAS", "WORKFLOWS_RUNS_VARIABLES",
             "ACTUALIZACIONES"],

    /**
     * Comprobación ligera para la puerta de instalación (usada hoy solo
     * para BigQuery, ver js/auth.js::proceedBigQuery): solo mira si existe
     * el dataset/esquema DRACO_CONTROL y su PRIMERA tabla (PROYECTOS), en
     * vez de recorrer las ~25 tablas de control como hace isBootstrapped().
     */
    async controlSchemaExists() {
        const exists = await Provider.containerExists(DracoConfig.controlDataset);
        if (!exists) return false;
        try {
            await Provider.runQuery(`SELECT 1 FROM ${Provider.qualifyControl("PROYECTOS")} LIMIT 1`);
            return true;
        } catch (e) {
            return false;
        }
    },

    async isBootstrapped() {
        const exists = await Provider.containerExists(DracoConfig.controlDataset);
        if (!exists) return false;
        try {
            // Las 28 comprobaciones son independientes entre sí: lanzarlas
            // todas a la vez (en vez de una por una con await secuencial)
            // reduce el tiempo total al de la más lenta, no a la suma de
            // todas. Ver la misma optimización en bootstrap() más abajo.
            await Promise.all(this.TABLES.map(name =>
                Provider.runQuery(`SELECT 1 FROM ${Provider.qualifyControl(name)} LIMIT 1`)
            ));
            return true;
        } catch (e) {
            return false;
        }
    },

    /**
     * Cambios sobre tablas que ya existían antes de introducir estas
     * columnas (proyectos ya bootstrapeados). `CREATE TABLE IF NOT EXISTS`
     * no las añade retroactivamente, así que se hace aquí con
     * `ADD COLUMN IF NOT EXISTS`, soportado tal cual en BigQuery y en
     * Snowflake. Idempotente: se puede ejecutar en cada bootstrap.
     */
    async evolve(onProgress = () => {}) {
        onProgress("Verificando columnas del modelo semántico...");
        const cubos = Provider.qualifyControl("CUBOS");
        const workflowsRuns = Provider.qualifyControl("WORKFLOWS_RUNS");
        await Promise.all([
            Provider.runQuery(`ALTER TABLE ${cubos} ADD COLUMN IF NOT EXISTS MODELO_YAML_PATH STRING`),
            Provider.runQuery(`ALTER TABLE ${cubos} ADD COLUMN IF NOT EXISTS MODELO_YAML_FECHA TIMESTAMP`),
            Provider.runQuery(`ALTER TABLE ${workflowsRuns} ADD COLUMN IF NOT EXISTS DESCRIPCION STRING`)
        ]);
    },

    async bootstrap(onProgress = () => {}) {
        if (Provider.key() === "snowflake") {
            onProgress(`Verificando base de datos ${SF.getDatabase()}...`);
            await SF.ensureDatabase();
        }

        const exists = await Provider.containerExists(DracoConfig.controlDataset);
        if (!exists) {
            onProgress(`Creando ${DracoConfig.controlDataset}...`);
            await Provider.createContainer(DracoConfig.controlDataset, "Dataset/esquema de control de Draco Planning (proyectos, dimensiones, cubos, jerarquías, interfaces)");
        }

        // Las 28 tablas son independientes (ninguna DDL depende de otra),
        // así que se lanzan todas a la vez en vez de una por una: con
        // await secuencial, 28 tablas a ~200ms cada una son ~5-6s solo
        // aquí; en paralelo es ~el tiempo de la más lenta (~200-500ms).
        onProgress(`Verificando ${this.TABLES.length} tablas de control...`);
        await Promise.all(this.TABLES.map(name => Provider.runQuery(this.ddl(name))));

        await this.evolve(onProgress);
    }
};

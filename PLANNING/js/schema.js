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
                    FECHA_MODIFICACION TIMESTAMP
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
            FLUJOS_CHAIN: `
                CREATE TABLE IF NOT EXISTS ${t} (
                    PROYECTO_ID STRING NOT NULL,
                    FLUJO_ID STRING NOT NULL,
                    PASO_ID STRING NOT NULL,
                    INTERFAZ_ID STRING NOT NULL,
                    ORDEN INTEGER
                )`,
            FLUJOS_CHAIN_TARGETS: `
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
                )`
        };
        return ddl[table];
    },

    TABLES: ["PROYECTOS", "DIMENSIONES", "CUBOS", "JERARQUIAS",
             "INTERFACES", "INTERFACES_VALUES", "INTERFACES_INPUT", "INTERFACES_INPUT_FILTERS", "INTERFACES_MAPPING",
             "FLUJOS", "FLUJOS_CHAIN", "FLUJOS_CHAIN_TARGETS", "FLUJOS_SCREEN_BLOCKS", "FLUJOS_SCREEN_VARIABLES"],

    async isBootstrapped() {
        const exists = await Provider.containerExists(DracoConfig.controlDataset);
        if (!exists) return false;
        try {
            for (const name of this.TABLES) {
                await Provider.runQuery(`SELECT 1 FROM ${Provider.qualifyControl(name)} LIMIT 1`);
            }
            return true;
        } catch (e) {
            return false;
        }
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

        for (const name of this.TABLES) {
            onProgress(`Verificando tabla ${name}...`);
            await Provider.runQuery(this.ddl(name));
        }
    }
};

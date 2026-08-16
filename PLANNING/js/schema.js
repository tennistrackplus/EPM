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
                )`
        };
        return ddl[table];
    },

    async isBootstrapped() {
        const exists = await Provider.containerExists(DracoConfig.controlDataset);
        if (!exists) return false;
        try {
            for (const name of ["PROYECTOS", "DIMENSIONES", "CUBOS", "JERARQUIAS"]) {
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
            await Provider.createContainer(DracoConfig.controlDataset, "Dataset/esquema de control de Draco Planning (proyectos, dimensiones, cubos, jerarquías)");
        }

        for (const name of ["PROYECTOS", "DIMENSIONES", "CUBOS", "JERARQUIAS"]) {
            onProgress(`Verificando tabla ${name}...`);
            await Provider.runQuery(this.ddl(name));
        }
    }
};

/**
 * ============================================================
 * DRACO PLANNING — CAPA DE ABSTRACCIÓN DE PROVEEDOR
 * ============================================================
 * El resto de la app (schema.js, app.js, dimensions.js, cubes.js)
 * habla siempre con "Provider", nunca directamente con BQ o SF.
 * Así los mismos flujos (proyectos, dimensiones, cubos) funcionan
 * igual sobre BigQuery o sobre Snowflake.
 */
const Provider = {
    key() {
        return localStorage.getItem("draco_active_provider") || "bigquery";
    },
    setKey(k) {
        localStorage.setItem("draco_active_provider", k);
    },
    label() {
        return this.key() === "snowflake" ? "Snowflake" : "BigQuery";
    },

    isConnected() {
        return this.key() === "snowflake" ? SF.isConnected() : BQ.isConnected();
    },

    /** ¿Hay ya un "hogar" (proyecto GCP / cuenta+warehouse Snowflake) elegido? */
    isReady() {
        return this.key() === "snowflake" ? SF.isReady() : (BQ.isConnected() && !!BQ.getGcpProject());
    },

    /** Texto corto para mostrar en la topbar */
    homeLabel() {
        if (this.key() === "snowflake") {
            return `${SF.getAccount()} · ${SF.getDatabase()} · WH ${SF.getWarehouse()}`;
        }
        return BQ.getGcpProject();
    },

    logout() {
        if (this.key() === "snowflake") {
            SF.logout();
        } else {
            localStorage.removeItem("bigquery_access_token");
            localStorage.removeItem("bigquery_token_expires");
            localStorage.removeItem("draco_gcp_project");
        }
    },

    /** Ejecuta SQL y devuelve siempre un array de objetos { columna: valor } */
    async runQuery(sql) {
        if (this.key() === "snowflake") {
            return SF.runQuery(sql);
        }
        const result = await BQ.query(BQ.getGcpProject(), sql);
        return BQ.rowsToObjects(result);
    },

    esc(v) {
        return BQ.esc(v);
    },

    toIdentifier(v) {
        return BQ.toIdentifier(v);
    },

    /** SESSION_USER() en BigQuery, CURRENT_USER() en Snowflake */
    currentUserExpr() {
        return this.key() === "snowflake" ? "CURRENT_USER()" : "SESSION_USER()";
    },

    /** UUID generado en el propio motor (fallback si no hay crypto.randomUUID en el navegador) */
    newId() {
        return (crypto.randomUUID) ? crypto.randomUUID() : `${Date.now()}-${Math.random().toString(16).slice(2)}`;
    },

    /** Traduce un tipo de campo "canónico" al tipo físico del motor activo */
    mapFieldType(type) {
        if (this.key() === "bigquery" && type === "FLOAT") return "FLOAT64";
        return type; // STRING/INTEGER/NUMERIC/BOOLEAN/DATE/DATETIME/TIMESTAMP son válidos en ambos motores
    },

    /** dataset (BigQuery) o esquema (Snowflake) */
    async containerExists(name) {
        return this.key() === "snowflake"
            ? SF.schemaExists(name)
            : BQ.datasetExists(BQ.getGcpProject(), name);
    },

    async createContainer(name, description = "") {
        return this.key() === "snowflake"
            ? SF.createSchema(name, description)
            : BQ.createDataset(BQ.getGcpProject(), name, description);
    },

    async deleteContainer(name) {
        return this.key() === "snowflake"
            ? SF.dropSchema(name)
            : BQ.deleteDataset(BQ.getGcpProject(), name);
    },

    /** Referencia totalmente cualificada a una tabla dentro de un dataset/esquema */
    qualify(container, table) {
        if (this.key() === "snowflake") {
            return `${SF.getDatabase()}.${container}.${table}`;
        }
        return `\`${BQ.getGcpProject()}.${container}.${table}\``;
    },

    qualifyControl(table) {
        return this.qualify(DracoConfig.controlDataset, table);
    },

    /** Nombre de la "base de datos" (Snowflake) o proyecto GCP (BigQuery) activo */
    databaseLabel() {
        return this.key() === "snowflake" ? SF.getDatabase() : BQ.getGcpProject();
    },

    /** Lista los nombres de tabla dentro de un dataset/esquema (para selectores de tabla) */
    async listTablesInContainer(container) {
        const sql = this.key() === "snowflake"
            ? `SELECT TABLE_NAME FROM ${SF.getDatabase()}.INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = '${this.esc(container)}' ORDER BY TABLE_NAME`
            : `SELECT TABLE_NAME FROM \`${BQ.getGcpProject()}.${container}.INFORMATION_SCHEMA.TABLES\` ORDER BY TABLE_NAME`;
        const rows = await this.runQuery(sql);
        return rows.map(r => r.TABLE_NAME);
    },

    /** Lista { name, type } de las columnas físicas de una tabla, en orden */
    async listColumns(container, table) {
        const sql = this.key() === "snowflake"
            ? `SELECT COLUMN_NAME, DATA_TYPE, ORDINAL_POSITION FROM ${SF.getDatabase()}.INFORMATION_SCHEMA.COLUMNS
               WHERE TABLE_SCHEMA = '${this.esc(container)}' AND TABLE_NAME = '${this.esc(table)}' ORDER BY ORDINAL_POSITION`
            : `SELECT COLUMN_NAME, DATA_TYPE, ORDINAL_POSITION FROM \`${BQ.getGcpProject()}.${container}.INFORMATION_SCHEMA.COLUMNS\`
               WHERE TABLE_NAME = '${this.esc(table)}' ORDER BY ORDINAL_POSITION`;
        const rows = await this.runQuery(sql);
        return rows.map(r => ({ name: r.COLUMN_NAME, type: r.DATA_TYPE }));
    }
};

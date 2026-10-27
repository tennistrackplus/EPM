/**
 * ============================================================
 * DRACO PLANNING — CAPA DE ABSTRACCIÓN DE PROVEEDOR
 * ============================================================
 * El resto de la app (schema.js, app.js, dimensions.js, cubes.js)
 * habla siempre con "Provider", nunca directamente con BQ.
 * En esta versión el único motor es BigQuery; la interfaz se mantiene
 * para poder añadir otros sin tocar el resto de la app.
 */
const Provider = {
    key() {
        return "bigquery";
    },
    setKey(k) {
        // Solo hay un motor: se ignora y se limpia el valor heredado.
        try { localStorage.removeItem("draco_active_provider"); } catch (e) { /* sin acceso */ }
    },
    label() {
        return "BigQuery";
    },

    isConnected() {
        return BQ.isConnected();
    },

    /** ¿Hay ya un "hogar" (proyecto GCP) elegido? */
    isReady() {
        return BQ.isConnected() && !!BQ.getGcpProject();
    },

    /** Texto corto para mostrar en la topbar */
    homeLabel() {
        return BQ.getGcpProject();
    },

    logout() {
        BQ.logout();
        localStorage.removeItem("draco_gcp_project");
    },

    /** Ejecuta SQL y devuelve siempre un array de objetos { columna: valor } */
    async runQuery(sql) {
        const result = await BQ.query(BQ.getGcpProject(), sql);
        return BQ.rowsToObjects(result);
    },

    esc(v) {
        return BQ.esc(v);
    },

    toIdentifier(v) {
        return BQ.toIdentifier(v);
    },

    /** Usuario de la sesión en el motor */
    currentUserExpr() {
        return "SESSION_USER()";
    },

    /** UUID generado en el propio motor (fallback si no hay crypto.randomUUID en el navegador) */
    newId() {
        return (crypto.randomUUID) ? crypto.randomUUID() : `${Date.now()}-${Math.random().toString(16).slice(2)}`;
    },

    // ---------------------------------------------------------
    // Buscador de usuarios (asignación de ejecuciones de Workflows)
    // ---------------------------------------------------------
    /**
     * Buscador de usuarios a partir de la política IAM del proyecto. Necesita
     * un scope de Cloud Platform que Draco ya no pide (auditoría F-04), así
     * que solo se activa si la sesión lo tiene; si no, se escribe el email.
     */
    canSearchUsers() {
        const scopes = (BQ.getGrantedScopes && BQ.getGrantedScopes()) || [];
        return scopes.some(s => s.indexOf("https://www.googleapis.com/auth/cloud-platform") === 0);
    },

    /** Devuelve como mucho 20 emails que contengan `query` (o los primeros 20 si `query` está vacío) */
    async searchUsers(query) {
        if (!this.canSearchUsers()) return [];
        const all = await BQ.getProjectUsersCached();
        const q = String(query || "").trim().toLowerCase();
        const filtered = q ? all.filter(email => email.toLowerCase().includes(q)) : all;
        return filtered.slice(0, 20);
    },

    /** Traduce un tipo de campo "canónico" al tipo físico del motor activo */
    mapFieldType(type) {
        if (type === "FLOAT") return "FLOAT64";
        return type; // STRING/INTEGER/NUMERIC/BOOLEAN/DATE/DATETIME/TIMESTAMP son válidos en BigQuery
    },

    /** dataset de BigQuery */
    async containerExists(name) {
        return BQ.datasetExists(BQ.getGcpProject(), name);
    },

    async createContainer(name, description = "") {
        return BQ.createDataset(BQ.getGcpProject(), name, description);
    },

    async deleteContainer(name) {
        return BQ.deleteDataset(BQ.getGcpProject(), name);
    },

    /** Referencia totalmente cualificada a una tabla dentro de un dataset */
    qualify(container, table) {
        return `\`${this.projectId(BQ.getGcpProject())}.${this.ident(container, "dataset")}.${this.ident(table, "tabla")}\``;
    },

    // ---------------------------------------------------------
    // SEGURIDAD (auditoría F-07): los nombres de dataset, tabla y columna
    // que salen de las tablas de control los puede escribir otro usuario.
    // Antes de meterlos en una SQL se valida que sean identificadores
    // simples; si no, se rechaza la operación (no se intenta "arreglar").
    // ---------------------------------------------------------
    IDENT_RE: /^[A-Za-z_][A-Za-z0-9_]{0,1023}$/,
    PROJECT_RE: /^([a-z][a-z0-9.-]*[a-z0-9]:)?[a-z][a-z0-9-]{4,61}[a-z0-9]$/,

    /** Devuelve el identificador si es válido; si no, lanza un error. */
    ident(name, what = "identificador") {
        const v = String(name == null ? "" : name);
        if (!this.IDENT_RE.test(v)) {
            throw new Error(`Nombre de ${what} no válido: "${v.slice(0, 80)}". Solo se admiten letras, números y "_" (sin empezar por número).`);
        }
        return v;
    },

    /** Valida un ID de proyecto de Google Cloud. */
    projectId(id) {
        const v = String(id == null ? "" : id);
        if (!this.PROJECT_RE.test(v)) throw new Error(`ID de proyecto de Google Cloud no válido: "${v.slice(0, 80)}".`);
        return v;
    },

    qualifyControl(table) {
        return this.qualify(DracoConfig.controlDataset, table);
    },

    /** Proyecto GCP activo */
    databaseLabel() {
        return BQ.getGcpProject();
    },

    /** Lista los nombres de tabla dentro de un dataset (para selectores de tabla) */
    async listTablesInContainer(container) {
        const sql = `SELECT TABLE_NAME FROM \`${BQ.getGcpProject()}.${container}.INFORMATION_SCHEMA.TABLES\` ORDER BY TABLE_NAME`;
        const rows = await this.runQuery(sql);
        return rows.map(r => r.TABLE_NAME);
    },

    /** Lista { name, type } de las columnas físicas de una tabla, en orden */
    async listColumns(container, table) {
        const sql = `SELECT COLUMN_NAME, DATA_TYPE, ORDINAL_POSITION FROM \`${BQ.getGcpProject()}.${container}.INFORMATION_SCHEMA.COLUMNS\`
               WHERE TABLE_NAME = '${this.esc(table)}' ORDER BY ORDINAL_POSITION`;
        const rows = await this.runQuery(sql);
        return rows.map(r => ({ name: r.COLUMN_NAME, type: r.DATA_TYPE }));
    },

    /**
     * Sincroniza las columnas físicas de `container.table` con
     * `desiredFields` ([{name, type}], nombres ya como identificador
     * físico) SIN BORRAR los datos existentes:
     *   - Si la tabla no existe todavía, se crea entera de una vez
     *     (CREATE TABLE IF NOT EXISTS) con el esquema completo.
     *   - Si ya existe, se añaden (ALTER TABLE ADD COLUMN) las columnas
     *     que falten y se eliminan (ALTER TABLE DROP COLUMN) las que ya
     *     no estén en `desiredFields`. Las columnas que se mantienen (y
     *     sus datos) no se tocan.
     * Usado por Dimensions.save() y Cubes.save() al editar, en vez de
     * `CREATE OR REPLACE TABLE` (que recreaba la tabla vacía cada vez,
     * borrando todo lo que hubiera cargado).
     *
     * No migra cambios de TIPO en una columna que se mantiene (p.ej.
     * pasar un atributo de STRING a INTEGER): eso exigiría convertir los
     * valores ya existentes y el tipo físico normalizado no siempre
     * coincide con el canónico (BigQuery guarda INTEGER como INT64,
     * BOOLEAN como BOOL, etc.), así que se deja el tipo físico tal cual
     * estaba para no arriesgar un ALTER que falle o trunque datos.
     */
    async syncTableColumns(container, table, desiredFields) {
        const fullTable = this.qualify(container, table);
        const colDefs = desiredFields.map(f => `${f.name} ${this.mapFieldType(f.type)}`).join(", ");

        const existingCols = await this.listColumns(container, table);
        if (existingCols.length === 0) {
            // Alta nueva (o la tabla se borró manualmente fuera de Draco):
            // no hay nada que preservar, se crea entera de una vez.
            await this.runQuery(`CREATE TABLE IF NOT EXISTS ${fullTable} (${colDefs})`);
            return { created: true, added: desiredFields.map(f => f.name), dropped: [] };
        }

        const existingNames = new Set(existingCols.map(c => c.name.toUpperCase()));
        const desiredNames = new Set(desiredFields.map(f => f.name.toUpperCase()));

        const toAdd = desiredFields.filter(f => !existingNames.has(f.name.toUpperCase()));
        const toDrop = existingCols.filter(c => !desiredNames.has(c.name.toUpperCase()));

        for (const f of toAdd) {
            await this.runQuery(`ALTER TABLE ${fullTable} ADD COLUMN IF NOT EXISTS ${f.name} ${this.mapFieldType(f.type)}`);
        }
        for (const c of toDrop) {
            await this.runQuery(`ALTER TABLE ${fullTable} DROP COLUMN IF EXISTS ${c.name}`);
        }

        return { created: false, added: toAdd.map(f => f.name), dropped: toDrop.map(c => c.name) };
    }
};

window.Provider = Provider;

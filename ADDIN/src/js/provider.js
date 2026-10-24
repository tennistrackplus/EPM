/**
 * ============================================================
 * EPM ADD-IN — CAPA DE ABSTRACCIÓN DE PROVEEDOR
 * ============================================================
 * El resto del add-in habla con "Provider", nunca directamente con BQ.
 * En esta versión el único proveedor es BigQuery; la interfaz se mantiene
 * para poder añadir otros motores sin tocar el explorador del modelo
 * semántico ni el motor de informes.
 *
 * Jerarquía de metadatos: proyecto > dataset > tabla -> `p.d.t`
 */
const Provider = {
    key() {
        return "bigquery";
    },
    setKey(k) {
        // Solo hay un proveedor: se ignora. Se limpia el valor heredado
        // de versiones anteriores para no dejar restos en localStorage.
        try { localStorage.removeItem("draco_active_provider"); } catch (e) { /* sin acceso */ }
    },
    label() {
        return "BigQuery";
    },
    /** Etiqueta del primer nivel de la jerarquía de metadatos */
    level1Label() {
        return "proyecto";
    },
    /** Etiqueta del segundo nivel */
    level2Label() {
        return "dataset";
    },

    isConnected() {
        return BQ.isConnected();
    },

    logout() {
        BQ.logout();
    },

    esc(v) {
        return BQ.esc(v);
    },

    toIdentifier(v) {
        return BQ.toIdentifier(v);
    },

    /** Referencia totalmente cualificada a una tabla */
    qualify(c1, c2, table) {
        return `\`${c1}.${c2}.${table}\``;
    },

    /** Limita un SELECT a n filas */
    limit(selectSql, n) {
        const rows = parseInt(n, 10);
        return `${selectSql} LIMIT ${rows}`;
    },

    /** Tipo de texto para CAST(... AS <tipo>) */
    stringType() {
        return "STRING";
    },

    /** Ejecuta SQL y devuelve {fields:[{name}], rows:[{col: valor}]} */
    async runQuery(sql, level1Id, level2Id) {
        // El "proyecto" usado para lanzar el job de BigQuery es el proyecto
        // de facturación si se indicó uno al crear la conexión; las tablas
        // ya van totalmente cualificadas en el SQL
        // (qualify() incluye siempre el proyecto de los datos), así que esto
        // solo afecta a qué proyecto paga la consulta.
        const billingProject = BQ.getBillingProject() || level1Id;
        return BQ.runQuerySql(billingProject, sql);
    },

    // -----------------------------------------------------------
    // Explorador de metadatos (usado por semantic_model.js)
    // -----------------------------------------------------------
    /** Nivel 1: proyectos. Devuelve [{id, label}] */
    async listLevel1() {
        const projects = await BQ.listProjects();
        return projects.map(p => {
            const id = p.id || (p.projectReference && p.projectReference.projectId);
            return { id, label: id };
        });
    },

    /** Nivel 2: datasets dentro del proyecto */
    async listLevel2(level1Id) {
        const datasets = await BQ.listDatasets(level1Id);
        return datasets.map(d => {
            const id = d.datasetReference.datasetId;
            return { id, label: id };
        });
    },

    /** Tablas dentro de proyecto/dataset */
    async listTables(level1Id, level2Id) {
        const tables = await BQ.listTables(level1Id, level2Id);
        return tables.map(t => {
            const id = t.tableReference.tableId;
            return { id, label: id };
        });
    },

    /** Campos de una tabla: [{name, type}] con el tipo ya normalizado */
    async getTableFields(level1Id, level2Id, tableId) {
        return BQ.getTableFields(level1Id, level2Id, tableId);
    }
};

/**
 * ============================================================
 * EPM ADD-IN — CAPA DE ABSTRACCIÓN DE PROVEEDOR
 * ============================================================
 * El resto del add-in habla con "Provider", nunca directamente con
 * BQ, SF o FB, así el explorador del modelo semántico y el motor de
 * informes funcionan igual sobre BigQuery, Snowflake o Microsoft Fabric
 * (Lakehouse y Warehouse, que comparten SQL endpoint y dialecto T-SQL).
 *
 * Jerarquía de metadatos por proveedor:
 *   BigQuery : proyecto      > dataset > tabla   -> `p.d.t`
 *   Snowflake: base de datos > esquema > tabla   -> BD.ESQ.TABLA
 *   Fabric   : item (LH/WH)  > schema  > tabla   -> [item].[schema].[tabla]
 *              (el workspace va en la CONEXIÓN, no en el modelo)
 */
const Provider = {
    key() {
        return localStorage.getItem("draco_active_provider") || "bigquery";
    },
    setKey(k) {
        localStorage.setItem("draco_active_provider", k);
    },
    isFabric() {
        return this.key() === "fabric";
    },
    isSnowflake() {
        return this.key() === "snowflake";
    },
    label() {
        if (this.isSnowflake()) return "Snowflake";
        if (this.isFabric()) return "Microsoft Fabric";
        return "BigQuery";
    },
    /** Etiqueta del primer nivel de la jerarquía de metadatos */
    level1Label() {
        if (this.isSnowflake()) return "base de datos";
        if (this.isFabric()) return "lakehouse / warehouse";
        return "proyecto";
    },
    /** Etiqueta del segundo nivel */
    level2Label() {
        if (this.isSnowflake()) return "esquema";
        if (this.isFabric()) return "schema";
        return "dataset";
    },

    isConnected() {
        if (this.isSnowflake()) return SF.isConnected();
        if (this.isFabric()) return FB.isConnected();
        return BQ.isConnected();
    },

    logout() {
        if (this.isSnowflake()) SF.logout();
        else if (this.isFabric()) FB.logout();
        else BQ.logout();
    },

    esc(v) {
        return BQ.esc(v);
    },

    toIdentifier(v) {
        return BQ.toIdentifier(v);
    },

    /** Referencia totalmente cualificada a una tabla */
    qualify(c1, c2, table) {
        if (this.isSnowflake()) {
            return `${c1}.${c2}.${table}`;
        }
        if (this.isFabric()) {
            return `${FB.quoteIdent(c1)}.${FB.quoteIdent(c2)}.${FB.quoteIdent(table)}`;
        }
        return `\`${c1}.${c2}.${table}\``;
    },

    // -----------------------------------------------------------
    // Diferencias de dialecto SQL
    // -----------------------------------------------------------
    /**
     * Limita un SELECT a n filas: "... LIMIT n" en BigQuery/Snowflake,
     * "SELECT [DISTINCT] TOP (n) ..." en T-SQL (Fabric).
     */
    limit(selectSql, n) {
        const rows = parseInt(n, 10);
        if (this.isFabric()) {
            return String(selectSql).replace(/^\s*SELECT(\s+DISTINCT)?\s+/i, (m, distinct) =>
                "SELECT" + (distinct ? " DISTINCT" : "") + ` TOP (${rows}) `);
        }
        return `${selectSql} LIMIT ${rows}`;
    },

    /** Tipo de texto para CAST(... AS <tipo>) */
    stringType() {
        return this.isFabric() ? "VARCHAR(4000)" : "STRING";
    },

    /** Ejecuta SQL contra el proveedor activo y devuelve {fields:[{name}], rows:[{col: valor}]} */
    async runQuery(sql, level1Id, level2Id) {
        if (this.isSnowflake()) {
            return SF.runQuerySql(sql, { database: level1Id, schema: level2Id });
        }
        if (this.isFabric()) {
            return FB.runQuerySql(sql, { database: level1Id });
        }
        // El "proyecto" usado para lanzar el job de BigQuery es el proyecto
        // de facturación si se indicó uno al crear la conexión (igual que en
        // Power BI); las tablas ya van totalmente cualificadas en el SQL
        // (qualify() incluye siempre el proyecto de los datos), así que esto
        // solo afecta a qué proyecto paga la consulta.
        const billingProject = BQ.getBillingProject() || level1Id;
        return BQ.runQuerySql(billingProject, sql);
    },

    // -----------------------------------------------------------
    // Explorador de metadatos (usado por semantic_model.js)
    // -----------------------------------------------------------
    /** Nivel 1: proyectos (BQ), bases de datos (SF) o items del workspace (Fabric). Devuelve [{id, label}] */
    async listLevel1() {
        if (this.isSnowflake()) {
            const names = await SF.listDatabases();
            return names.map(n => ({ id: n, label: n }));
        }
        if (this.isFabric()) {
            const names = await FB.listItems();
            return names.map(n => ({ id: n, label: n }));
        }
        const projects = await BQ.listProjects();
        return projects.map(p => {
            const id = p.id || (p.projectReference && p.projectReference.projectId);
            return { id, label: id };
        });
    },

    /** Nivel 2: datasets (BQ) o esquemas (SF / Fabric) dentro del nivel 1 */
    async listLevel2(level1Id) {
        if (this.isSnowflake()) {
            const names = await SF.listSchemas(level1Id);
            return names.map(n => ({ id: n, label: n }));
        }
        if (this.isFabric()) {
            const names = await FB.listSchemas(level1Id);
            return names.map(n => ({ id: n, label: n }));
        }
        const datasets = await BQ.listDatasets(level1Id);
        return datasets.map(d => {
            const id = d.datasetReference.datasetId;
            return { id, label: id };
        });
    },

    /** Tablas dentro de nivel1/nivel2 */
    async listTables(level1Id, level2Id) {
        if (this.isSnowflake()) {
            const names = await SF.listTables(level1Id, level2Id);
            return names.map(n => ({ id: n, label: n }));
        }
        if (this.isFabric()) {
            const names = await FB.listTables(level1Id, level2Id);
            return names.map(n => ({ id: n, label: n }));
        }
        const tables = await BQ.listTables(level1Id, level2Id);
        return tables.map(t => {
            const id = t.tableReference.tableId;
            return { id, label: id };
        });
    },

    /** Campos de una tabla: [{name, type}] con el tipo ya normalizado */
    async getTableFields(level1Id, level2Id, tableId) {
        if (this.isSnowflake()) {
            return SF.getTableFields(level1Id, level2Id, tableId);
        }
        if (this.isFabric()) {
            return FB.getTableFields(level1Id, level2Id, tableId);
        }
        return BQ.getTableFields(level1Id, level2Id, tableId);
    }
};

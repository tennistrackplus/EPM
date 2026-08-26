/**
 * ============================================================
 * EPM ADD-IN — CLIENTE BIGQUERY (REST API)
 * ============================================================
 * Envuelve las llamadas a la API REST de BigQuery usando el token
 * OAuth guardado en localStorage por login.js. Mantiene las mismas
 * claves de localStorage ("bigquery_access_token", etc.) que ya
 * usaba el add-in, para no romper commands.js / semantic_model.js
 * mientras se migran (fase posterior).
 */
const BQ = {
    BASE: "https://bigquery.googleapis.com/bigquery/v2",

    // ---------------------------------------------------------
    // Configuración de la conexión activa (proyecto de facturación
    // opcional y repositorio de modelos semánticos, al estilo
    // Power BI / LookML). Se guardan junto con el resto de config
    // en Connections, pero se exponen aquí para el resto del código
    // que ya habla con BQ directamente.
    // ---------------------------------------------------------
    getBillingProject() {
        return localStorage.getItem("bq_billing_project") || "";
    },
    setBillingProject(v) {
        if (v) {
            localStorage.setItem("bq_billing_project", String(v).trim());
        } else {
            localStorage.removeItem("bq_billing_project");
        }
    },

    /** { type: "github"|"gitlab"|"azure_devops"|"local"|"", url: "" } */
    getSemanticRepo() {
        try {
            const raw = localStorage.getItem("bq_semantic_repo");
            return raw ? JSON.parse(raw) : { type: "", url: "" };
        } catch (e) {
            return { type: "", url: "" };
        }
    },
    setSemanticRepo(repo) {
        if (repo && (repo.type || repo.url)) {
            localStorage.setItem("bq_semantic_repo", JSON.stringify({
                type: repo.type || "",
                url: (repo.url || "").trim()
            }));
        } else {
            localStorage.removeItem("bq_semantic_repo");
        }
    },

    getToken() {
        const token = localStorage.getItem("bigquery_access_token");
        const expires = localStorage.getItem("bigquery_token_expires");
        if (!token || !expires || Date.now() >= parseInt(expires, 10)) {
            return null;
        }
        return token;
    },

    isConnected() {
        return !!this.getToken();
    },

    logout() {
        localStorage.removeItem("bigquery_access_token");
        localStorage.removeItem("bigquery_token_expires");
    },

    async request(path, options = {}) {
        const token = this.getToken();
        if (!token) {
            const err = new Error("Sesión de BigQuery no válida o expirada. Inicia sesión de nuevo.");
            err.code = "NO_AUTH";
            throw err;
        }
        const response = await fetch(`${this.BASE}${path}`, {
            ...options,
            headers: {
                "Authorization": `Bearer ${token}`,
                "Content-Type": "application/json",
                ...(options.headers || {})
            }
        });
        const data = await response.json().catch(() => ({}));
        if (!response.ok || data.error) {
            const msg = (data.error && data.error.message) || `Error HTTP ${response.status}`;
            const err = new Error(msg);
            err.details = data.error;
            throw err;
        }
        return data;
    },

    /** Lista los proyectos GCP visibles para el usuario autenticado */
    async listProjects() {
        const data = await this.request("/projects?maxResults=200");
        return data.projects || [];
    },

    async listDatasets(projectId) {
        const data = await this.request(`/projects/${encodeURIComponent(projectId)}/datasets?maxResults=500`);
        return data.datasets || [];
    },

    async listTables(projectId, datasetId) {
        const data = await this.request(`/projects/${encodeURIComponent(projectId)}/datasets/${encodeURIComponent(datasetId)}/tables?maxResults=500`);
        return data.tables || [];
    },

    /** Devuelve los campos de una tabla como [{name, type}] */
    async getTableFields(projectId, datasetId, tableId) {
        const data = await this.request(`/projects/${encodeURIComponent(projectId)}/datasets/${encodeURIComponent(datasetId)}/tables/${encodeURIComponent(tableId)}`);
        const fields = (data.schema && data.schema.fields) || [];
        return fields.map(f => ({ name: f.name, type: f.type }));
    },

    /** Ejecuta SQL de forma síncrona (jobs.query) */
    async query(projectId, sql) {
        const data = await this.request(`/projects/${encodeURIComponent(projectId)}/queries`, {
            method: "POST",
            body: JSON.stringify({ query: sql, useLegacySql: false, timeoutMs: 30000 })
        });
        return data;
    },

    /** Convierte el resultado de jobs.query en un array de objetos { columna: valor } */
    rowsToObjects(result) {
        const fields = (result.schema && result.schema.fields) || [];
        const rows = result.rows || [];
        return rows.map(r => {
            const obj = {};
            (r.f || []).forEach((cell, i) => {
                obj[fields[i].name] = cell.v;
            });
            return obj;
        });
    },

    /** Igual que query(), pero devuelve {fields:[{name}], rows:[{col: valor}]} ya normalizado (mismo formato que SF.runQuerySql) */
    async runQuerySql(projectId, sql) {
        const result = await this.query(projectId, sql);
        const fields = (result.schema && result.schema.fields) || [];
        const rows = this.rowsToObjects(result);
        return { fields: fields.map(f => ({ name: f.name })), rows };
    },

    /** Escapa comillas simples para literales SQL */
    esc(value) {
        if (value === null || value === undefined) return "";
        return String(value).replace(/\\/g, "\\\\").replace(/'/g, "\\'");
    },

    /** Convierte un texto libre en un identificador válido de BigQuery */
    toIdentifier(text) {
        return String(text || "")
            .trim()
            .toUpperCase()
            .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
            .replace(/[^A-Z0-9_]+/g, "_")
            .replace(/^_+|_+$/g, "")
            .replace(/^(\d)/, "N$1");
    }
};

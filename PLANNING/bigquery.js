/**
 * ============================================================
 * DRACO PLANNING — CLIENTE BIGQUERY (REST API)
 * ============================================================
 * Envuelve las llamadas a la API REST de BigQuery usando el token
 * OAuth guardado en localStorage por el módulo de login.
 * No usa ninguna librería externa: solo fetch().
 */
const BQ = {
    BASE: "https://bigquery.googleapis.com/bigquery/v2",

    getToken() {
        const token = localStorage.getItem("bigquery_access_token");
        const expires = localStorage.getItem("bigquery_token_expires");
        if (!token || !expires || Date.now() >= parseInt(expires, 10)) {
            return null;
        }
        return token;
    },

    getGcpProject() {
        return localStorage.getItem("draco_gcp_project") || "";
    },

    setGcpProject(projectId) {
        localStorage.setItem("draco_gcp_project", projectId);
    },

    // ---------------------------------------------------------
    // PAT de GitHub vía Google Secret Manager
    // ---------------------------------------------------------
    // GitHub bloquea el push si detecta el token de DracoConfig.
    // semanticModelGithub.token pegado tal cual en config.js, así que en
    // vez de leerlo de ahí se lee, justo antes de usarlo, desde Secret
    // Manager (secreto "github-pat-draco").
    GITHUB_PAT_SECRET_NAME: "github-pat-draco",

    /**
     * Lee la última versión del secreto "github-pat-draco" en Secret
     * Manager y devuelve su valor en texto (el PAT de GitHub). Usa el
     * mismo token OAuth de la sesión de BigQuery activa (hace falta el
     * scope "cloud-platform" en DracoConfig.googleScopes, además de que
     * el usuario tenga el rol "Secret Manager Secret Accessor" sobre ese
     * secreto en el proyecto indicado).
     *
     * projectId: proyecto GCP donde vive el secreto (por defecto, el
     * proyecto GCP activo en Planning, BQ.getGcpProject()).
     */
    async getGithubPatFromSecretManager(projectId) {
        const token = this.getToken();
        if (!token) {
            const err = new Error("Sesión de BigQuery no válida o expirada.");
            err.code = "NO_AUTH";
            throw err;
        }
        const proj = projectId || this.getGcpProject();
        if (!proj) {
            throw new Error(`Falta el proyecto GCP donde está el secreto "${this.GITHUB_PAT_SECRET_NAME}".`);
        }

        const url = `https://secretmanager.googleapis.com/v1/projects/${encodeURIComponent(proj)}/secrets/${this.GITHUB_PAT_SECRET_NAME}/versions/latest:access`;

        const res = await fetch(url, {
            headers: { "Authorization": `Bearer ${token}` }
        });

        if (!res.ok) {
            const errBody = await res.json().catch(() => null);
            const msg = (errBody && errBody.error && errBody.error.message) || `HTTP ${res.status}`;
            throw new Error(`No se pudo leer el secreto "${this.GITHUB_PAT_SECRET_NAME}" de Secret Manager: ${msg}`);
        }

        const data = await res.json();
        const b64 = data && data.payload && data.payload.data;
        if (!b64) {
            throw new Error(`El secreto "${this.GITHUB_PAT_SECRET_NAME}" no devolvió ningún valor en Secret Manager.`);
        }

        // El payload viene en base64 (posible UTF-8): decodificación segura.
        return decodeURIComponent(escape(atob(b64))).trim();
    },

    isConnected() {
        return !!this.getToken();
    },

    async request(path, options = {}) {
        const token = this.getToken();
        if (!token) {
            const err = new Error("Sesión de BigQuery no válida o expirada.");
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

    async datasetExists(projectId, datasetId) {
        try {
            await this.request(`/projects/${encodeURIComponent(projectId)}/datasets/${encodeURIComponent(datasetId)}`);
            return true;
        } catch (e) {
            return false;
        }
    },

    async createDataset(projectId, datasetId, description = "") {
        return this.request(`/projects/${encodeURIComponent(projectId)}/datasets`, {
            method: "POST",
            body: JSON.stringify({
                datasetReference: { projectId, datasetId },
                description,
                friendlyName: datasetId
            })
        });
    },

    async deleteDataset(projectId, datasetId) {
        return this.request(
            `/projects/${encodeURIComponent(projectId)}/datasets/${encodeURIComponent(datasetId)}?deleteContents=true`,
            { method: "DELETE" }
        );
    },

    async listTables(projectId, datasetId) {
        const data = await this.request(`/projects/${encodeURIComponent(projectId)}/datasets/${encodeURIComponent(datasetId)}/tables?maxResults=500`);
        return data.tables || [];
    },

    /** Ejecuta una consulta SQL (SELECT, DDL o DML) de forma síncrona */
    async query(projectId, sql, params = {}) {
        const data = await this.request(`/projects/${encodeURIComponent(projectId)}/queries`, {
            method: "POST",
            body: JSON.stringify({
                query: sql,
                useLegacySql: false,
                timeoutMs: 30000,
                ...params
            })
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

    /** Escapa comillas simples para literales SQL */
    esc(value) {
        if (value === null || value === undefined) return "";
        return String(value).replace(/\\/g, "\\\\").replace(/'/g, "\\'");
    },

    /** Convierte un texto libre en un identificador válido de BigQuery (dataset/tabla/columna) */
    toIdentifier(text) {
        return String(text || "")
            .trim()
            .toUpperCase()
            .normalize("NFD").replace(/[\u0300-\u036f]/g, "") // quita acentos
            .replace(/[^A-Z0-9_]+/g, "_")
            .replace(/^_+|_+$/g, "")
            .replace(/^(\d)/, "N$1"); // los identificadores no pueden empezar por número
    },

    // ---------------------------------------------------------
    // Buscador de usuarios (para asignar ejecuciones de Workflows)
    // ---------------------------------------------------------
    // No hay ningún directorio de usuarios propio: se reutiliza la
    // política IAM del proyecto GCP activo (Resource Manager) como
    // "directorio" — cualquier persona con algún rol en el proyecto es
    // candidata a que se le asigne una instancia. Requiere el scope
    // "cloud-platform" (ya solicitado en el login) y permiso de lectura
    // de IAM sobre el proyecto; si el usuario no tiene ese permiso, la
    // búsqueda simplemente no devuelve resultados y se puede seguir
    // escribiendo el nombre a mano.
    _iamMembersCache: { projectId: null, promise: null },

    async fetchProjectIamMembers(projectId) {
        const token = this.getToken();
        if (!token) {
            const err = new Error("Sesión de BigQuery no válida o expirada.");
            err.code = "NO_AUTH";
            throw err;
        }
        const res = await fetch(`https://cloudresourcemanager.googleapis.com/v1/projects/${encodeURIComponent(projectId)}:getIamPolicy`, {
            method: "POST",
            headers: { "Authorization": `Bearer ${token}`, "Content-Type": "application/json" },
            body: JSON.stringify({})
        });
        const data = await res.json().catch(() => ({}));
        if (!res.ok || data.error) {
            const msg = (data.error && data.error.message) || `HTTP ${res.status}`;
            throw new Error(`No se pudo leer los miembros del proyecto: ${msg}`);
        }
        const emails = new Set();
        (data.bindings || []).forEach(b => {
            (b.members || []).forEach(m => {
                const match = /^(?:user|group):(.+)$/.exec(m);
                if (match) emails.add(match[1]);
            });
        });
        return Array.from(emails).sort((a, b) => a.localeCompare(b));
    },

    /** Miembros IAM del proyecto activo, cacheados en memoria (una sola llamada por proyecto/sesión) */
    getProjectUsersCached() {
        const projectId = this.getGcpProject();
        const cache = this._iamMembersCache;
        if (cache.projectId !== projectId) {
            cache.projectId = projectId;
            cache.promise = null;
        }
        if (!cache.promise) {
            cache.promise = this.fetchProjectIamMembers(projectId).catch(err => {
                cache.promise = null; // reintentar en la siguiente búsqueda
                throw err;
            });
        }
        return cache.promise;
    }
};

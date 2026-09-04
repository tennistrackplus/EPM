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

    // ---------------------------------------------------------
    // Bucket de Google Cloud Storage donde se guarda una copia del
    // Excel activo al pulsar "Guardar en bucket" (ver js/gcsExport.js).
    // Se guarda junto al resto de config de la conexión BigQuery.
    // ---------------------------------------------------------
    getExportProject() {
        return localStorage.getItem("bq_export_project") || "";
    },
    setExportProject(v) {
        if (v) {
            localStorage.setItem("bq_export_project", String(v).trim());
        } else {
            localStorage.removeItem("bq_export_project");
        }
    },

    getExportBucket() {
        return localStorage.getItem("bq_export_bucket") || "";
    },
    setExportBucket(v) {
        if (v) {
            localStorage.setItem("bq_export_bucket", String(v).trim());
        } else {
            localStorage.removeItem("bq_export_bucket");
        }
    },

    /** { type: "github"|"gitlab"|"azure_devops"|"local"|"", url: "", branch: "", token: "" } */
    getSemanticRepo() {
        try {
            const raw = localStorage.getItem("bq_semantic_repo");
            return raw ? JSON.parse(raw) : { type: "", url: "", branch: "", token: "" };
        } catch (e) {
            return { type: "", url: "", branch: "", token: "" };
        }
    },
    setSemanticRepo(repo) {
        if (repo && (repo.type || repo.url)) {
            localStorage.setItem("bq_semantic_repo", JSON.stringify({
                type: repo.type || "",
                url: (repo.url || "").trim(),
                branch: (repo.branch || "").trim(),
                token: (repo.token || "").trim()
            }));
        } else {
            localStorage.removeItem("bq_semantic_repo");
        }
    },

    // ---------------------------------------------------------
    // El manifest de este add-in NO declara <Runtimes> (shared runtime),
    // así que commands.html (el runtime aislado que ejecuta los botones
    // del ribbon, p.ej. AbrirBucketButton/GuardarBucketButton -> commands.js)
    // tiene su PROPIO localStorage, distinto del de los paneles de tareas
    // (taskpane.html y el panel de login "ConexionPane"). El login solo
    // guarda el token en el localStorage del panel, así que ese runtime
    // del ribbon nunca lo ve y siempre parece "no conectado".
    //
    // Office.context.document.settings sí es visible desde CUALQUIER
    // runtime del mismo documento (paneles y ribbon) — es el mismo
    // mecanismo de "roaming settings" que ya usan reportStore.js,
    // filterRangeStore.js y SemanticModelStore en este proyecto — así que
    // lo usamos aquí también como respaldo del token.
    // ---------------------------------------------------------
    _syncTokenToDocumentSettings(token, expires) {
        try {
            if (typeof Office === "undefined" || !Office.context || !Office.context.document || !Office.context.document.settings) return;
            const settings = Office.context.document.settings;
            if (token && expires) {
                settings.set("epm_bq_token", token);
                settings.set("epm_bq_token_expires", String(expires));
            } else {
                settings.remove("epm_bq_token");
                settings.remove("epm_bq_token_expires");
            }
            settings.saveAsync();
        } catch (e) {
            // Fuera de Excel, o sin documento accesible desde este runtime: no pasa nada.
        }
    },

    _getTokenFromDocumentSettings() {
        try {
            if (typeof Office === "undefined" || !Office.context || !Office.context.document || !Office.context.document.settings) return null;
            const settings = Office.context.document.settings;
            const token = settings.get("epm_bq_token");
            const expires = settings.get("epm_bq_token_expires");
            if (!token || !expires || Date.now() >= parseInt(expires, 10)) return null;
            return { token, expires };
        } catch (e) {
            return null;
        }
    },

    /** Guarda el token tras un login correcto: en localStorage (rápido,
     *  para este mismo runtime) y en Office.context.document.settings
     *  (para que lo vean también el resto de runtimes, incluido el
     *  aislado de los botones del ribbon). Llamar desde login.js. */
    setToken(token, expiresAt) {
        localStorage.setItem("bigquery_access_token", token);
        localStorage.setItem("bigquery_token_expires", String(expiresAt));
        this._syncTokenToDocumentSettings(token, expiresAt);
    },

    getToken() {
        let token = localStorage.getItem("bigquery_access_token");
        let expires = localStorage.getItem("bigquery_token_expires");
        if (!token || !expires || Date.now() >= parseInt(expires, 10)) {
            // Este runtime puede no tener el token en su propio localStorage
            // (p.ej. el runtime aislado del ribbon): lo intentamos recuperar
            // de Office.context.document.settings antes de rendirnos.
            const fromSettings = this._getTokenFromDocumentSettings();
            if (!fromSettings) return null;
            token = fromSettings.token;
            expires = fromSettings.expires;
            // Lo copiamos a este localStorage para que las próximas lecturas
            // en este mismo runtime sean inmediatas.
            localStorage.setItem("bigquery_access_token", token);
            localStorage.setItem("bigquery_token_expires", expires);
        }
        return token;
    },

    isConnected() {
        return !!this.getToken();
    },

    // ---------------------------------------------------------
    // Los diálogos independientes (Office.context.ui.displayDialogAsync,
    // p.ej. bucketBrowser.html / saveBucket.html) se abren en su propia
    // ventana/proceso y, en varias plataformas de Office (sobre todo
    // Office de escritorio), NO comparten localStorage con el panel de
    // tareas que los abrió. Sin esto, el diálogo ve "no conectado" aunque
    // el usuario sí lo esté en el panel. Para evitarlo, quien abre el
    // diálogo añade el token vigente a la URL (ver getSessionQueryParams),
    // y el propio diálogo lo copia a SU localStorage nada más cargar
    // (ver hydrateSessionFromDialogParams), antes de comprobar isConnected().
    //
    // getSessionQueryParams() usa this.getToken() (no localStorage
    // directamente) precisamente para que esto funcione también cuando
    // quien abre el diálogo es el runtime aislado del ribbon: getToken()
    // ya sabe recuperar el token de Office.context.document.settings si
    // su propio localStorage está vacío.
    // ---------------------------------------------------------
    getSessionQueryParams() {
        const token = this.getToken();
        const expires = localStorage.getItem("bigquery_token_expires");
        if (!token || !expires) return "";
        return `epmTok=${encodeURIComponent(token)}&epmExp=${encodeURIComponent(expires)}`;
    },

    hydrateSessionFromDialogParams(searchParams) {
        const token = searchParams.get("epmTok");
        const expires = searchParams.get("epmExp");
        if (token && expires) {
            localStorage.setItem("bigquery_access_token", token);
            localStorage.setItem("bigquery_token_expires", expires);
        }
    },

    logout() {
        localStorage.removeItem("bigquery_access_token");
        localStorage.removeItem("bigquery_token_expires");
        this._syncTokenToDocumentSettings(null, null);
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

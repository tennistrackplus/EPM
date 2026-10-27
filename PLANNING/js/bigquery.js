/**
 * ============================================================
 * DRACO PLANNING — CLIENTE BIGQUERY (REST API)
 * ============================================================
 * Envuelve las llamadas a la API REST de BigQuery usando el token
 * OAuth obtenido por el módulo de login.
 *
 * SEGURIDAD: el token se guarda SOLO en sessionStorage (no en
 * localStorage): no se persiste en disco y desaparece al cerrar la
 * pestaña. Se comparte con los iframes del mismo origen y se copia a
 * las ventanas abiertas con window.open() desde la propia app.
 * No usa ninguna librería externa: solo fetch().
 */
const BQ = {
    BASE: "https://bigquery.googleapis.com/bigquery/v2",

    TOKEN_KEY: "bigquery_access_token",
    EXPIRES_KEY: "bigquery_token_expires",

    _store() {
        return window.sessionStorage;
    },

    /** Borra el token que versiones anteriores dejaban en localStorage. */
    purgeLegacyToken() {
        try {
            localStorage.removeItem(this.TOKEN_KEY);
            localStorage.removeItem(this.EXPIRES_KEY);
        } catch (e) { /* sin acceso a localStorage */ }
    },

    SCOPES_KEY: "bigquery_token_scopes",

    /** Scopes concedidos por Google en el último login (null si se desconocen). */
    getGrantedScopes() {
        const raw = this._store().getItem(this.SCOPES_KEY);
        return raw ? raw.split(/\s+/).filter(Boolean) : null;
    },

    /**
     * ¿El token permite ESCRIBIR en BigQuery? false si solo se concedió
     * bigquery.readonly (p. ej. el usuario desmarcó el permiso de escritura
     * en la pantalla de consentimiento, o el servicio de autenticación del
     * cliente solo admite lectura). Si no se conocen los scopes, true: la
     * seguridad real la pone IAM al escribir.
     */
    canWriteBigQuery() {
        const scopes = this.getGrantedScopes();
        if (!scopes) return true;
        return scopes.includes("https://www.googleapis.com/auth/bigquery")
            || scopes.includes("https://www.googleapis.com/auth/cloud-platform");
    },

    setToken(token, expiresAt, scope) {
        this._store().setItem(this.TOKEN_KEY, token);
        this._store().setItem(this.EXPIRES_KEY, String(expiresAt));
        if (scope) this._store().setItem(this.SCOPES_KEY, String(scope));
    },

    getToken() {
        const token = this._store().getItem(this.TOKEN_KEY);
        const expires = this._store().getItem(this.EXPIRES_KEY);
        if (!token || !expires || Date.now() >= parseInt(expires, 10)) {
            if (token) this.logout();
            return null;
        }
        return token;
    },

    logout() {
        this._store().removeItem(this.TOKEN_KEY);
        this._store().removeItem(this.EXPIRES_KEY);
        this._store().removeItem(this.SCOPES_KEY);
    },

    getGcpProject() {
        return localStorage.getItem("draco_gcp_project") || "";
    },

    setGcpProject(projectId) {
        localStorage.setItem("draco_gcp_project", projectId);
    },

    // ---------------------------------------------------------
    // Repositorio Git de modelos semánticos (.lkml)
    // ---------------------------------------------------------
    // SEGURIDAD (auditoría F-04): el PAT de GitHub ya NO se lee de Secret
    // Manager (eso obligaba a pedir el scope cloud-platform). Lo introduce
    // el propio usuario, como en el addin de Excel, y se guarda SOLO en la
    // sesión del navegador (sessionStorage): no se persiste en disco.
    // URL y rama del repositorio son datos de la conexión (no secretos).
    GIT_PAT_KEY: "draco_git_pat",

    getGitPat() {
        return sessionStorage.getItem(this.GIT_PAT_KEY) || "";
    },

    setGitPat(pat) {
        const v = String(pat || "").trim();
        if (v) sessionStorage.setItem(this.GIT_PAT_KEY, v);
        else sessionStorage.removeItem(this.GIT_PAT_KEY);
    },

    /** Pide el PAT al usuario (modal propio, sin dependencias). Devuelve "" si cancela. */
    _askGitPat(repoUrl) {
        return new Promise((resolve) => {
            const overlay = document.createElement("div");
            overlay.style.cssText = "position:fixed;inset:0;background:rgba(0,0,0,.45);display:flex;align-items:center;justify-content:center;z-index:100000;";
            const box = document.createElement("div");
            box.style.cssText = "background:#fff;color:#1f2937;border-radius:10px;padding:20px;width:min(440px,92vw);font:14px system-ui,sans-serif;box-shadow:0 10px 30px rgba(0,0,0,.25);";
            const title = document.createElement("h3");
            title.textContent = "Token de acceso a Git";
            title.style.cssText = "margin:0 0 8px;font-size:16px;";
            const p = document.createElement("p");
            p.textContent = `Introduce tu token de acceso (PAT) con permiso de escritura en ${repoUrl}. Solo se guarda mientras dure esta sesión del navegador.`;
            p.style.cssText = "margin:0 0 12px;line-height:1.4;";
            const input = document.createElement("input");
            input.type = "password";
            input.autocomplete = "off";
            input.style.cssText = "width:100%;box-sizing:border-box;padding:8px;border:1px solid #d1d5db;border-radius:6px;";
            const row = document.createElement("div");
            row.style.cssText = "display:flex;justify-content:flex-end;gap:8px;margin-top:14px;";
            const cancel = document.createElement("button");
            cancel.textContent = "Cancelar";
            const ok = document.createElement("button");
            ok.textContent = "Aceptar";
            [cancel, ok].forEach(b => { b.type = "button"; b.style.cssText = "padding:7px 14px;border-radius:6px;border:1px solid #d1d5db;cursor:pointer;background:#fff;"; });
            ok.style.background = "#2659d9"; ok.style.color = "#fff"; ok.style.borderColor = "#2659d9";
            row.append(cancel, ok);
            box.append(title, p, input, row);
            overlay.appendChild(box);
            document.body.appendChild(overlay);
            input.focus();
            const done = (val) => { overlay.remove(); resolve(val); };
            cancel.addEventListener("click", () => done(""));
            ok.addEventListener("click", () => done(input.value.trim()));
            input.addEventListener("keydown", (e) => {
                if (e.key === "Enter") done(input.value.trim());
                if (e.key === "Escape") done("");
            });
        });
    },

    /**
     * Configuración del repositorio para GithubRepo: { url, branch, token }.
     * URL y rama salen de la conexión activa (o, si no, de
     * DracoConfig.semanticModelGithub); el token, de la sesión o se pide.
     */
    async getGitRepoConfig() {
        const active = (window.Connections && typeof Connections.getActive === "function") ? Connections.getActive() : null;
        const repo = (active && active.config && active.config.semanticRepo) || {};
        const def = (typeof DracoConfig !== "undefined" && DracoConfig.semanticModelGithub) || {};
        const url = String(repo.url || def.url || "").trim();
        const branch = String(repo.branch || def.branch || "main").trim() || "main";
        if (!url) {
            throw new Error("No hay repositorio Git configurado: indícalo en la conexión (URL del repositorio de modelos).");
        }
        let token = this.getGitPat();
        if (!token) {
            token = await this._askGitPat(url);
            if (!token) throw new Error("No se ha indicado el token de acceso a Git.");
            this.setGitPat(token);
        }
        return { url, branch, token };
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
    // candidata a que se le asigne una instancia. Requiere un scope de
    // Cloud Platform que Draco ya no pide (ver Provider.canSearchUsers) y permiso de lectura
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

window.BQ = BQ;

// Limpieza de tokens que versiones anteriores guardaban en localStorage.
BQ.purgeLegacyToken();

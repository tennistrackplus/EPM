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
    // SEGURIDAD (corrige token-en-URL): los diálogos independientes
    // (Office.context.ui.displayDialogAsync, p.ej. bucketBrowser.html /
    // saveBucket.html) se abren en su propia ventana/proceso y, en varias
    // plataformas de Office, NO comparten localStorage con quien los
    // abrió. ANTES esto se resolvía añadiendo el token OAuth vigente como
    // parámetro en la URL del diálogo — visible en historiales de
    // navegación, en logs de depuración de WebView2 y en cualquier
    // monitorización intermedia. Se sustituye por un pequeño protocolo de
    // mensajería del propio Dialog API de Office: el diálogo pide el
    // token con Office.context.ui.messageParent(), y quien lo abrió
    // responde por dialog.messageChild() — el token nunca viaja por la
    // URL ni queda en ningún sitio persistente.
    // ---------------------------------------------------------
    TOKEN_REQUEST_MESSAGE_TYPE: "epm-bq-token-request",
    TOKEN_RESPONSE_MESSAGE_TYPE: "epm-bq-token-response",

    /** Llamar en el DialogMessageReceived de quien abre el diálogo, ANTES
     *  de tratar el mensaje como la respuesta "de negocio" habitual
     *  (selección de bucket, cierre, etc.): si es una petición de token,
     *  se responde aquí y se corta (return true); si no, el llamante
     *  sigue con su propio manejo del mensaje. */
    handleDialogMessage(dialog, rawMessage) {
        try {
            const msg = JSON.parse(rawMessage);
            if (msg && msg.type === this.TOKEN_REQUEST_MESSAGE_TYPE) {
                const token = this.getToken();
                const expires = localStorage.getItem("bigquery_token_expires");
                dialog.messageChild(JSON.stringify({
                    type: this.TOKEN_RESPONSE_MESSAGE_TYPE,
                    token: token || null,
                    expires: expires || null
                }));
                return true;
            }
        } catch (e) {
            // No era JSON o no era una petición de token: lo trata el llamante.
        }
        return false;
    },

    /** Llamar desde el propio diálogo (bucketBrowser.html/saveBucket.html)
     *  ANTES de comprobar isConnected(): pide el token a quien lo abrió y
     *  espera su respuesta un tiempo corto. Si no hay padre (p.ej. se abre
     *  la página suelta para depurar) o no responde a tiempo, continúa
     *  igualmente y isConnected() dará "no conectado" como siempre. */
    requestSessionFromOpener(timeoutMs = 1200) {
        return new Promise((resolve) => {
            if (typeof Office === "undefined" || !Office.context || !Office.context.ui
                || typeof Office.context.ui.messageParent !== "function"
                || typeof Office.context.ui.addHandlerAsync !== "function") {
                resolve(false);
                return;
            }

            let done = false;
            const finish = (ok) => {
                if (done) return;
                done = true;
                resolve(ok);
            };
            const timer = setTimeout(() => finish(false), timeoutMs);

            try {
                Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg) => {
                    try {
                        const msg = JSON.parse(arg.message);
                        if (msg && msg.type === this.TOKEN_RESPONSE_MESSAGE_TYPE) {
                            clearTimeout(timer);
                            if (msg.token && msg.expires) {
                                localStorage.setItem("bigquery_access_token", msg.token);
                                localStorage.setItem("bigquery_token_expires", msg.expires);
                            }
                            finish(true);
                        }
                    } catch (e) {
                        // Mensaje que no es la respuesta esperada: se ignora.
                    }
                });
                Office.context.ui.messageParent(JSON.stringify({ type: this.TOKEN_REQUEST_MESSAGE_TYPE }));
            } catch (e) {
                clearTimeout(timer);
                finish(false);
            }
        });
    },

    // ---------------------------------------------------------
    // C1 (auditoría de seguridad): ANTES este bloque sincronizaba el token
    // OAuth de BigQuery en Office.context.document.settings para que el
    // runtime aislado del ribbon (commands.html, sin Shared Runtime) lo
    // pudiera leer. Ese "roaming setting" se persiste dentro del propio
    // .xlsx al guardar el archivo: si el usuario compartía o subía ese
    // Excel mientras el token no había caducado, cualquiera que lo abriera
    // heredaba acceso de lectura/escritura a BigQuery y a Google Cloud
    // Storage con el scope concedido en el login.
    //
    // El manifest real de este add-in SÍ declara Shared Runtime
    // (<Runtimes><Runtime resid="TaskpaneUrl".../>) y los botones del
    // ribbon (AbrirBucketButton, GuardarBucketButton, etc.) son
    // ExecuteFunction sobre ese mismo runtime compartido con taskpane.html
    // — es decir, commands.js YA comparte localStorage con taskpane.js/
    // login.js en producción. El respaldo de document.settings era
    // innecesario (y es lo que embebía el token en el archivo) y se ha
    // eliminado: setToken/getToken/logout usan únicamente localStorage.
    // ---------------------------------------------------------
    setToken(token, expiresAt) {
        localStorage.setItem("bigquery_access_token", token);
        localStorage.setItem("bigquery_token_expires", String(expiresAt));
    },

    getToken() {
        const token = localStorage.getItem("bigquery_access_token");
        const expires = localStorage.getItem("bigquery_token_expires");
        if (!token || !expires || Date.now() >= parseInt(expires, 10)) return null;
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
    // getSessionQueryParams() / hydrateSessionFromDialogParams() ELIMINADOS:
    // pasaban el token OAuth en claro como parámetro de la URL del diálogo.
    // Sustituidos por el protocolo de mensajería de arriba
    // (handleDialogMessage / requestSessionFromOpener), que no expone el
    // token en ninguna URL. Ver también auth-callback.html (origen
    // concreto en vez de "*" en postMessage) y login.js (comprobación de
    // event.origin al recibir mensajes).

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

// IMPORTANTE: el resto de módulos (commands.js, login.js, gcsSaveBridge.js,
// bucketPickerUI.js) comprueban "window.BQ" antes de usarlo. Sin esta línea,
// "const BQ = {...}" de arriba NO cuelga de window (a diferencia de "var"),
// así que todos esos "window.BQ ? ... : ''" evaluaban siempre a falso, sin
// dar ningún error visible — el origen real de "Conéctate primero".
window.BQ = BQ;

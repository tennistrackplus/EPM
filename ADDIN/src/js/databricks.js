/**
 * ============================================================
 * EPM ADD-IN — CLIENTE DATABRICKS (vía DATABRICKS-GATEWAY)
 * ============================================================
 * Databricks no admite CORS, así que el add-in no habla con el workspace:
 * habla con la pasarela que el cliente despliega en SU nube
 * (DATABRICKS-GATEWAY/). Los datos van del Databricks del cliente a su
 * pasarela y de ahí al Excel; nunca pasan por infraestructura de Stratesys.
 *
 * Misma interfaz de sesión que BQ (getToken, isConnected, isAuthError,
 * onSessionExpired, validateSession, logout...) para que el resto del
 * add-in trate igual a los dos motores a través de Provider.session().
 *
 * Autenticación (config.authMode de la conexión):
 *   "oauth"  Producción. Login OAuth de Databricks a través de la pasarela
 *            (Authorization Code + PKCE, cliente confidencial). El refresh
 *            token nunca llega al navegador: solo un manejador cifrado.
 *   "pat"    SOLO PRUEBAS (DracoConfig.allowDatabricksPat). Token personal
 *            de Databricks; se guarda solo en sessionStorage, nunca en disco.
 *
 * SEGURIDAD: igual que BQ, el access token vive en sessionStorage (con
 * respaldo de 1 h en localStorage para el ribbon y los diálogos) y nunca
 * viaja por la URL ni se guarda en el .xlsx.
 */
const DBX = {
    TOKEN_KEY: "dbx_access_token",
    EXPIRES_KEY: "dbx_token_expires",
    MODE_KEY: "dbx_auth_mode",
    GATEWAY_KEY: "dbx_gateway_url",
    SESSION_KEY: "dbx_auth_session",          // manejador opaco de la pasarela (solo modo oauth)
    PAT_HOURS: 8,
    PAT_STORE_PREFIX: "dbx_pat_enc:",          // + URL de la pasarela: PAT cifrado (casilla «Recordar»)
    PAT_IDB: "draco-secure",                   // IndexedDB con la clave AES no exportable
    QUERY_MAX_WAIT_MS: 10 * 60 * 1000,
    _refreshTimer: null,
    _pendingState: null,
    _pendingGateway: null,
    _validatedAt: 0,

    // ---------------------------------------------------------
    // Configuración de la conexión activa
    // ---------------------------------------------------------
    normalizeGateway(url) {
        return String(url || "").trim().replace(/\/+$/, "");
    },

    patAllowed() {
        return !!(typeof DracoConfig !== "undefined" && DracoConfig.allowDatabricksPat === true);
    },

    normalizeConfig(cfg) {
        const c = cfg || {};
        return {
            authMode: c.authMode === "pat" && this.patAllowed() ? "pat" : "oauth",
            gatewayUrl: this.normalizeGateway(c.gatewayUrl)
        };
    },

    /** Devuelve un texto de error o null si la configuración es válida. */
    validateConfig(cfg) {
        if (cfg && cfg.authMode === "pat" && !this.patAllowed()) {
            return "En producción solo se admite el inicio de sesión de Databricks: edita la conexión.";
        }
        const c = this.normalizeConfig(cfg);
        // Error típico: poner la URL del workspace en vez de la de la pasarela
        if (/\.(azuredatabricks\.net|cloud\.databricks\.com|gcp\.databricks\.com|databricks\.azure\.cn)(\/|$)/i.test(c.gatewayUrl)) {
            return "Esa es la URL del workspace de Databricks, no la de la pasarela. Databricks no deja llamarle desde Excel: " +
                "arranca o despliega DATABRICKS-GATEWAY (con DATABRICKS_HOST = este workspace) y pon aquí SU dirección.";
        }
        if (!/^https:\/\/[^\s/?#]+(\/[^\s?#]*)?$/i.test(c.gatewayUrl)
            && !/^http:\/\/localhost(:\d+)?(\/[^\s?#]*)?$/i.test(c.gatewayUrl)) {
            return "Indica la URL (https://...) de la pasarela de Databricks del cliente.";
        }
        return null;
    },

    setGateway(url) {
        const u = this.normalizeGateway(url);
        // Otra pasarela = otro workspace: el token anterior no vale ahí.
        const prev = this.getGateway();
        if (prev && u && prev !== u) {
            this.logout();
            this._clearSession();
        }
        try { localStorage.setItem(this.GATEWAY_KEY, u); } catch (e) { /* sin acceso */ }
        try { sessionStorage.setItem(this.GATEWAY_KEY, u); } catch (e) { /* sin acceso */ }
    },

    getGateway() {
        try {
            return sessionStorage.getItem(this.GATEWAY_KEY) || localStorage.getItem(this.GATEWAY_KEY) || "";
        } catch (e) { return ""; }
    },

    // ---------------------------------------------------------
    // Token (mismo esquema que BQ)
    // ---------------------------------------------------------
    _store() { return window.sessionStorage; },
    _local() { try { return window.localStorage; } catch (e) { return null; } },

    setToken(token, expiresAt, mode) {
        const m = mode || "oauth";
        this._store().setItem(this.TOKEN_KEY, token);
        this._store().setItem(this.EXPIRES_KEY, String(expiresAt));
        this._store().setItem(this.MODE_KEY, m);
        // Un PAT no sale de sessionStorage (no se persiste en disco).
        const ls = this._local();
        if (ls && m !== "pat") {
            try {
                ls.setItem(this.TOKEN_KEY, token);
                ls.setItem(this.EXPIRES_KEY, String(expiresAt));
            } catch (e) { /* sin acceso */ }
        }
    },

    _rehydrate() {
        const ls = this._local();
        if (!ls || this._store().getItem(this.TOKEN_KEY)) return;
        try {
            const t = ls.getItem(this.TOKEN_KEY), e = ls.getItem(this.EXPIRES_KEY);
            if (t && e && Date.now() < parseInt(e, 10)) {
                this._store().setItem(this.TOKEN_KEY, t);
                this._store().setItem(this.EXPIRES_KEY, e);
            }
        } catch (err) { /* sin acceso */ }
    },

    getToken() {
        this._rehydrate();
        const token = this._store().getItem(this.TOKEN_KEY);
        const expires = this._store().getItem(this.EXPIRES_KEY);
        if (!token || !expires || Date.now() >= parseInt(expires, 10)) {
            if (token) this.logout();
            return null;
        }
        return token;
    },

    getExpiresAt() {
        this._rehydrate();
        const e = parseInt(this._store().getItem(this.EXPIRES_KEY) || "0", 10);
        return isNaN(e) ? 0 : e;
    },

    isConnected() {
        return !!this.getToken() && !!this.getGateway();
    },

    /** En Databricks la escritura la decide Unity Catalog: el add-in no la limita. */
    canWrite() { return true; },

    isAuthError(status) {
        // Solo 401: un 403 es "sin permiso sobre esa tabla", no sesión caducada.
        return status === 401;
    },

    onSessionExpired(source) {
        const hadToken = !!this._store().getItem(this.TOKEN_KEY);
        // Databricks ha rechazado el token (caducado o revocado): el recordado tampoco vale
        if (source === "query" && this._store().getItem(this.MODE_KEY) === "pat") this.forgetPat();
        this.logout();
        this._validatedAt = 0;
        const detail = { source: source || "", hadToken, provider: "databricks" };
        try { window.dispatchEvent(new CustomEvent("draco-session-expired", { detail })); } catch (e) { /* sin CustomEvent */ }
        try { if (typeof window.dracoSessionExpiredNotice === "function") window.dracoSessionExpiredNotice(detail); } catch (e) { /* sin aviso */ }
    },

    authExpiredError() {
        const err = new Error("La sesión de Databricks ha caducado. Vuelve a conectar en la pestaña Conexión.");
        err.code = "AUTH_EXPIRED";
        return err;
    },

    /**
     * A diferencia de BigQuery, aquí NO se lanza ninguna consulta para
     * comprobar la sesión: cualquier sentencia enciende el SQL Warehouse y
     * cuesta dinero (la lista de conexiones lo comprobaría cada minuto).
     * Solo se mira la caducidad local; un token rechazado se detecta en la
     * primera consulta real (401 -> onSessionExpired).
     */
    async validateSession() {
        return !!this.getToken();
    },

    logout() {
        [this.TOKEN_KEY, this.EXPIRES_KEY, this.MODE_KEY].forEach(k => {
            try { this._store().removeItem(k); } catch (e) { /* sin acceso */ }
        });
        const ls = this._local();
        if (ls) {
            try { ls.removeItem(this.TOKEN_KEY); ls.removeItem(this.EXPIRES_KEY); } catch (e) { /* sin acceso */ }
        }
    },

    // ---------------------------------------------------------
    // Diálogos: comparten la sesión por la mensajería del Dialog API
    // (ver BQ.handleDialogMessage / requestSessionFromOpener).
    // ---------------------------------------------------------
    exportForDialog() {
        const token = this.getToken();
        if (!token) return null;
        return { token, expires: this.getExpiresAt(), gateway: this.getGateway(), mode: this._store().getItem(this.MODE_KEY) || "oauth" };
    },

    importFromDialog(s) {
        if (!s || !s.token || !s.expires || !s.gateway) return;
        this.setGateway(s.gateway);
        this.setToken(s.token, s.expires, s.mode);
        try { sessionStorage.setItem("draco_session_provider", "databricks"); } catch (e) { /* sin acceso */ }
    },

    // ---------------------------------------------------------
    // Login
    // ---------------------------------------------------------
    _randomState() {
        const bytes = new Uint8Array(24);
        crypto.getRandomValues(bytes);
        return Array.from(bytes, b => b.toString(16).padStart(2, "0")).join("");
    },

    buildLoginUrl(cfg, redirectUri) {
        const c = this.normalizeConfig(cfg);
        const err = this.validateConfig(c);
        if (err) throw new Error(err);
        this._pendingState = this._randomState();
        this._pendingGateway = c.gatewayUrl;
        return `${c.gatewayUrl}/login?` + new URLSearchParams({ redirect_uri: redirectUri, state: this._pendingState }).toString();
    },

    /** Respuesta de auth-callback.html ({status, token, expiresIn, state, session}). */
    acceptResponse(response) {
        const expected = this._pendingState;
        this._pendingState = null;
        if (!expected || !response || response.state !== expected) {
            throw new Error("Respuesta de inicio de sesión no válida (state). Vuelve a intentarlo.");
        }
        if (response.status !== "success" || !response.token) return false;
        const expiresAt = Date.now() + (parseInt(response.expiresIn, 10) || 3600) * 1000;
        this.setGateway(this._pendingGateway);
        this.setToken(response.token, expiresAt, "oauth");
        if (response.session) {
            this._setSession({ gateway: this._pendingGateway, handle: response.session });
            this.scheduleRefresh(expiresAt);
        }
        this._pendingGateway = null;
        return true;
    },

    /** Solo pruebas: conecta con un token personal (PAT). */
    connectWithPat(cfg, pat) {
        if (!this.patAllowed()) throw new Error("El acceso con token personal está desactivado.");
        const c = this.normalizeConfig(cfg);
        const t = String(pat || "").trim();
        if (!/^dapi[0-9a-f]{20,}(-\d+)?$/i.test(t)) throw new Error("El token personal de Databricks empieza por «dapi».");
        this._clearSession();
        this.setGateway(c.gatewayUrl);
        this.setToken(t, Date.now() + this.PAT_HOURS * 3600 * 1000, "pat");
    },

    // ---------------------------------------------------------
    // «Recordar en este equipo» (solo PAT de pruebas)
    // El token se guarda CIFRADO (AES-GCM 256) en localStorage, por pasarela;
    // la clave es una CryptoKey NO exportable generada por el navegador y
    // guardada en IndexedDB: ningún script (ni Draco) puede leerla, solo
    // usarla para descifrar en este perfil de este equipo. Caduca a los
    // DracoConfig.rememberDatabricksPatDays días, al cerrar sesión, al borrar
    // la conexión o cuando Databricks rechaza el token (401).
    // ---------------------------------------------------------
    rememberPatDays() {
        const d = typeof DracoConfig !== "undefined" ? Number(DracoConfig.rememberDatabricksPatDays) : 0;
        return isFinite(d) && d > 0 ? d : 0;
    },

    canRememberPat() {
        return this.patAllowed() && this.rememberPatDays() > 0 && !!(window.crypto && crypto.subtle && window.indexedDB);
    },

    _patStoreKey(gateway) { return this.PAT_STORE_PREFIX + this.normalizeGateway(gateway); },

    _idb(mode, fn) {
        return new Promise((resolve, reject) => {
            const req = indexedDB.open(this.PAT_IDB, 1);
            req.onupgradeneeded = () => req.result.createObjectStore("keys");
            req.onerror = () => reject(req.error);
            req.onsuccess = () => {
                const db = req.result;
                const tx = db.transaction("keys", mode);
                const r = fn(tx.objectStore("keys"));
                tx.oncomplete = () => { db.close(); resolve(r && "result" in r ? r.result : undefined); };
                tx.onerror = () => { db.close(); reject(tx.error); };
            };
        });
    },

    async _patKey(create) {
        let key = await this._idb("readonly", st => st.get("dbx-pat"));
        if (!key && create) {
            key = await crypto.subtle.generateKey({ name: "AES-GCM", length: 256 }, false, ["encrypt", "decrypt"]);
            await this._idb("readwrite", st => st.put(key, "dbx-pat"));
        }
        return key || null;
    },

    hasRememberedPat(gateway) {
        try {
            const rec = JSON.parse(localStorage.getItem(this._patStoreKey(gateway || this.getGateway())) || "null");
            return !!(rec && rec.ct && Date.now() < Number(rec.exp));
        } catch (e) { return false; }
    },

    async rememberPat(gateway, pat) {
        if (!this.canRememberPat()) return false;
        try {
            const key = await this._patKey(true);
            const iv = crypto.getRandomValues(new Uint8Array(12));
            const ct = await crypto.subtle.encrypt({ name: "AES-GCM", iv }, key, new TextEncoder().encode(String(pat).trim()));
            const b64 = (buf) => btoa(String.fromCharCode.apply(null, new Uint8Array(buf)));
            const exp = Date.now() + this.rememberPatDays() * 86400 * 1000;
            localStorage.setItem(this._patStoreKey(gateway), JSON.stringify({ v: 1, iv: b64(iv), ct: b64(ct), exp }));
            return true;
        } catch (e) {
            console.warn("[Draco] No se pudo recordar el token de Databricks:", e);
            return false;
        }
    },

    forgetPat(gateway) {
        try { localStorage.removeItem(this._patStoreKey(gateway || this.getGateway())); } catch (e) { /* sin acceso */ }
    },

    /** Conecta con el PAT recordado para esa pasarela. true si lo ha conseguido. */
    async restoreRememberedPat(gateway) {
        const gw = this.normalizeGateway(gateway || this.getGateway());
        if (!gw || !this.canRememberPat()) return false;
        try {
            const rec = JSON.parse(localStorage.getItem(this._patStoreKey(gw)) || "null");
            if (!rec || !rec.ct) return false;
            if (Date.now() >= Number(rec.exp)) { this.forgetPat(gw); return false; }
            const key = await this._patKey(false);
            if (!key) { this.forgetPat(gw); return false; }
            const unb64 = (s) => Uint8Array.from(atob(s), c => c.charCodeAt(0));
            const pt = await crypto.subtle.decrypt({ name: "AES-GCM", iv: unb64(rec.iv) }, key, unb64(rec.ct));
            this.connectWithPat({ authMode: "pat", gatewayUrl: gw }, new TextDecoder().decode(pt));
            this._patFromStore = true;
            return true;
        } catch (e) {
            console.warn("[Draco] Token de Databricks recordado ilegible; se descarta:", e);
            this.forgetPat(gw);
            return false;
        }
    },

    _getSession() {
        try { return JSON.parse(sessionStorage.getItem(this.SESSION_KEY) || "null"); } catch (e) { return null; }
    },
    _setSession(s) { sessionStorage.setItem(this.SESSION_KEY, JSON.stringify(s)); },
    _clearSession() { sessionStorage.removeItem(this.SESSION_KEY); },

    async refresh() {
        const s = this._getSession();
        if (!s || !s.gateway || !s.handle) return false;
        try {
            const resp = await fetch(`${s.gateway}/token`, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ session: s.handle }),
                credentials: "omit",
                cache: "no-store"
            });
            if (!resp.ok) {
                if (resp.status === 401) { this._clearSession(); this.logout(); }
                return false;
            }
            const data = await resp.json();
            if (!data.access_token) return false;
            if (data.session) this._setSession({ gateway: s.gateway, handle: data.session });   // refresh token rotado
            const expiresAt = Date.now() + (parseInt(data.expires_in, 10) || 3600) * 1000;
            this.setToken(data.access_token, expiresAt, "oauth");
            this.scheduleRefresh(expiresAt);
            return true;
        } catch (e) {
            console.warn("[Draco] No se pudo renovar la sesión de Databricks:", e);
            return false;
        }
    },

    scheduleRefresh(expiresAt) {
        if (this._refreshTimer) clearTimeout(this._refreshTimer);
        if (!this._getSession()) return;
        const delay = Math.max(10000, Number(expiresAt) - Date.now() - 5 * 60 * 1000);
        this._refreshTimer = setTimeout(() => this.refresh(), delay);
    },

    async resume() {
        if (!this._getSession()) {
            if (!this.getToken() && this.hasRememberedPat()) await this.restoreRememberedPat();
            return;
        }
        if (!this.getToken() || this.getExpiresAt() - Date.now() < 5 * 60 * 1000) {
            await this.refresh();
        } else {
            this.scheduleRefresh(this.getExpiresAt());
        }
    },

    async fullLogout() {
        const s = this._getSession();
        if (this._refreshTimer) clearTimeout(this._refreshTimer);
        this._clearSession();
        this.logout();
        if (s && s.gateway && s.handle) {
            try {
                await fetch(`${s.gateway}/logout`, {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ session: s.handle }),
                    credentials: "omit"
                });
            } catch (e) { /* la sesión local ya está cerrada */ }
        }
    },

    // ---------------------------------------------------------
    // Consultas
    // ---------------------------------------------------------
    async _call(method, path, body) {
        if (!this.getToken() && this.hasRememberedPat()) await this.restoreRememberedPat();
        const token = this.getToken();
        const gw = this.getGateway();
        if (!token || !gw) {
            this.onSessionExpired("no-token");
            throw new Error("No hay una sesión activa de Databricks. Conéctate en la pestaña Conexión.");
        }
        let resp;
        try {
            resp = await fetch(gw + path, {
                method,
                headers: { "Authorization": "Bearer " + token, "Content-Type": "application/json" },
                body: body ? JSON.stringify(body) : undefined,
                credentials: "omit",
                cache: "no-store"
            });
        } catch (e) {
            // Un origen no permitido también llega aquí: la pasarela responde 403
            // sin cabecera CORS y el navegador lo ve como fallo de red.
            let origin = "";
            try { origin = window.location.origin; } catch (x) { /* sin origen */ }
            throw new Error("No se pudo contactar con la pasarela de Databricks (" + gw + "). Causas posibles: " +
                "la pasarela no está arrancada o la URL no es correcta; su ALLOWED_ORIGINS no incluye «" + origin + "»; " +
                "o Excel no deja llamar a esa dirección (p. ej. http://localhost desde una página https). Detalle: " + ((e && e.message) || e));
        }
        const data = await resp.json().catch(() => ({}));
        if (this.isAuthError(resp.status)) {
            this.onSessionExpired("query");
            throw this.authExpiredError();
        }
        if (resp.status === 413 || data.state === "TOO_LARGE") {
            const n = Number(data.row_count || 0), max = Number(data.max_rows || 0);
            const msg = `El informe devuelve ${n.toLocaleString("es-ES")} filas (máximo ${max.toLocaleString("es-ES")}). `
                + "Quita detalle (p. ej. documento) o añade filtros (sociedad, ejercicio...).";
            if (typeof dracoInfoNotice === "function") dracoInfoNotice("Informe demasiado grande", msg, "warn");
            throw new Error(msg);
        }
        if (!resp.ok || data.state === "FAILED" || data.state === "CANCELED" || data.state === "CLOSED") {
            const m = (data.error && data.error.message) || data.message || data.error || ("HTTP " + resp.status);
            throw new Error("Error de Databricks: " + m);
        }
        return data;
    },

    /** Ejecuta SQL y espera el resultado: {columns:[{name,type}], rows:[[...]]} */
    async execute(sql) {
        const t0 = Date.now();
        // La Statement Execution API ejecuta UNA sentencia: sin «;» final.
        const statement = String(sql || "").replace(/;\s*$/, "");
        let res = await this._call("POST", "/v1/statements", { statement });
        while (res.state === "RUNNING") {
            if (Date.now() - t0 > this.QUERY_MAX_WAIT_MS) {
                this._call("POST", "/v1/statements/" + encodeURIComponent(res.statement_id) + "/cancel").catch(() => {});
                throw new Error("La consulta lleva más de " + (this.QUERY_MAX_WAIT_MS / 60000) + " minutos en Databricks. Acota el informe (filtros, menos detalle) o inténtalo más tarde.");
            }
            res = await this._call("GET", "/v1/statements/" + encodeURIComponent(res.statement_id));
        }
        return { columns: res.columns || [], rows: res.rows || [] };
    },

    /**
     * Mismo texto JSON que devuelve jobs.query de BigQuery
     * ({schema:{fields}, rows:[{f:[{v}]}], totalRows, jobComplete}), que es lo
     * que leen los parsers del informe, del buscador de miembros y de los
     * filtros (buscan el literal `"v":` en el texto). Así el pintado no
     * distingue de qué motor viene el resultado.
     */
    async runQueryRaw(sql) {
        const { columns, rows } = await this.execute(sql);
        const out = {
            kind: "draco#databricksQueryResponse",
            schema: { fields: columns.map(c => ({ name: c.name, type: c.type })) },
            rows: rows.map(r => ({ f: r.map(v => ({ v: v === undefined ? null : v })) })),
            totalRows: String(rows.length),
            jobComplete: true
        };
        return JSON.stringify(out, null, 2);
    },

    /** {fields:[{name}], rows:[{col: valor}]} como BQ.runQuerySql */
    async runQuerySql(sql) {
        const { columns, rows } = await this.execute(sql);
        const names = columns.map(c => c.name);
        return {
            fields: names.map(name => ({ name })),
            rows: rows.map(r => {
                const o = {};
                names.forEach((n, i) => { o[n] = r[i]; });
                return o;
            })
        };
    },

    // ---------------------------------------------------------
    // Metadatos (catálogo > esquema > tabla), por SQL: la pasarela no
    // expone otras APIs de Databricks y Unity Catalog filtra lo visible.
    // ---------------------------------------------------------
    quoteIdent(name) {
        return "`" + String(name).replace(/`/g, "``") + "`";
    },

    _col(names, wanted) {
        const lower = names.map(n => String(n).toLowerCase());
        for (const w of wanted) {
            const i = lower.indexOf(w);
            if (i !== -1) return i;
        }
        return 0;
    },

    /** Consulta de metadatos anotada en el diario del Diagnóstico. */
    _meta(sql) {
        if (window.DracoQueryJournal) {
            return window.DracoQueryJournal.track(sql, async () => {
                const r = await this.execute(sql);
                return Object.assign({}, r, { rows: r.rows });
            }, { kind: "Metadatos", engine: "Databricks" });
        }
        return this.execute(sql);
    },

    async listCatalogs() {
        const { columns, rows } = await this._meta("SHOW CATALOGS");
        const i = this._col(columns.map(c => c.name), ["catalog", "catalog_name"]);
        return rows.map(r => r[i]).filter(Boolean).sort();
    },

    async listSchemas(catalog) {
        const { columns, rows } = await this._meta("SHOW SCHEMAS IN " + this.quoteIdent(catalog));
        const i = this._col(columns.map(c => c.name), ["databasename", "namespace", "schema_name"]);
        return rows.map(r => r[i]).filter(s => s && s !== "information_schema").sort();
    },

    /** Vistas del esquema (incluidas las Metric Views). */
    async listViews(catalog, schema) {
        const { columns, rows } = await this._meta("SHOW VIEWS IN " + this.quoteIdent(catalog) + "." + this.quoteIdent(schema));
        const names = columns.map(c => c.name);
        const i = this._col(names, ["viewname", "view_name", "tablename"]);
        const tmp = names.map(n => String(n).toLowerCase()).indexOf("istemporary");
        return rows.filter(r => tmp === -1 || String(r[tmp]) !== "true").map(r => r[i]).filter(Boolean).sort();
    },

    /** Metric Views del esquema (information_schema: table_type = 'METRIC_VIEW'). */
    async listMetricViews(catalog, schema) {
        const lit = "'" + String(schema).replace(/\\/g, "\\\\").replace(/'/g, "\\'") + "'";
        const { columns, rows } = await this._meta("SELECT table_name FROM " + this.quoteIdent(catalog)
            + ".information_schema.tables WHERE table_schema = " + lit + " AND table_type = 'METRIC_VIEW' ORDER BY table_name");
        const i = this._col(columns.map(c => c.name), ["table_name"]);
        return rows.map(r => r[i]).filter(Boolean);
    },

    /**
     * DESCRIBE TABLE EXTENDED ... AS JSON: tipo (METRIC_VIEW...), columnas
     * (con is_measure y display_name), view_text (el YAML, sin comentarios)
     * y table_properties. Devuelve el objeto ya interpretado.
     */
    async describeJson(catalog, schema, table) {
        const ref = this.quoteIdent(catalog) + "." + this.quoteIdent(schema) + "." + this.quoteIdent(table);
        const { rows } = await this._meta("DESCRIBE TABLE EXTENDED " + ref + " AS JSON");
        const raw = rows && rows[0] && rows[0][0];
        try { return JSON.parse(String(raw || "")); } catch (e) { throw new Error("Databricks no ha devuelto la descripción de «" + table + "»."); }
    },

    /** Propiedades (TBLPROPERTIES) de una tabla o vista: { clave: valor }. */
    async getTableProperties(catalog, schema, table) {
        const ref = this.quoteIdent(catalog) + "." + this.quoteIdent(schema) + "." + this.quoteIdent(table);
        const { columns, rows } = await this._meta("SHOW TBLPROPERTIES " + ref);
        const names = columns.map(c => c.name);
        const iKey = this._col(names, ["key"]), iVal = this._col(names, ["value"]);
        const out = {};
        rows.forEach(r => { if (r[iKey] != null) out[String(r[iKey])] = r[iVal] == null ? "" : String(r[iVal]); });
        return out;
    },

    async listTables(catalog, schema) {
        const { columns, rows } = await this._meta("SHOW TABLES IN " + this.quoteIdent(catalog) + "." + this.quoteIdent(schema));
        const names = columns.map(c => c.name);
        const i = this._col(names, ["tablename", "table_name"]);
        const tmp = names.map(n => String(n).toLowerCase()).indexOf("istemporary");
        return rows.filter(r => tmp === -1 || String(r[tmp]) !== "true").map(r => r[i]).filter(Boolean).sort();
    },

    /** Tipos de Databricks -> los nombres de BigQuery que ya entiende el modelo semántico. */
    normalizeType(t) {
        const s = String(t || "").toUpperCase();
        if (/^(TINYINT|SMALLINT|INT|INTEGER|BIGINT|LONG|SHORT|BYTE)\b/.test(s)) return "INTEGER";
        if (/^(FLOAT|DOUBLE|REAL)\b/.test(s)) return "FLOAT";
        if (/^DECIMAL|^NUMERIC|^DEC\b/.test(s)) return "NUMERIC";
        if (/^(STRING|VARCHAR|CHAR)\b/.test(s)) return "STRING";
        if (/^BOOLEAN\b/.test(s)) return "BOOLEAN";
        if (/^DATE\b/.test(s)) return "DATE";
        if (/^TIMESTAMP_NTZ\b/.test(s)) return "DATETIME";
        if (/^TIMESTAMP\b/.test(s)) return "TIMESTAMP";
        if (/^BINARY\b/.test(s)) return "BYTES";
        if (/^(ARRAY|MAP)\b/.test(s)) return "ARRAY";
        if (/^STRUCT\b/.test(s)) return "RECORD";
        return s || "STRING";
    },

    async getTableFields(catalog, schema, table) {
        const ref = this.quoteIdent(catalog) + "." + this.quoteIdent(schema) + "." + this.quoteIdent(table);
        const { columns, rows } = await this._meta("DESCRIBE TABLE " + ref);
        const names = columns.map(c => c.name);
        const iName = this._col(names, ["col_name"]), iType = this._col(names, ["data_type"]);
        const out = [];
        for (const r of rows) {
            const n = String(r[iName] || "").trim();
            if (!n || n.startsWith("#")) break;   // a partir de aquí: información de particiones
            out.push({ name: n, type: this.normalizeType(r[iType]) });
        }
        return out;
    }
};

window.DBX = DBX;

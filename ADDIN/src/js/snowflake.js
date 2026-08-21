/**
 * ============================================================
 * EPM ADD-IN — CLIENTE SNOWFLAKE (OAuth PKCE + SQL API v2)
 * ============================================================
 * Adaptado del cliente equivalente de Draco Planning. Diferencia
 * clave respecto a aquel: connect() NO abre la ventana de login él
 * mismo (Office Add-in necesita Office.context.ui.displayDialogAsync,
 * no window.open) — construye la URL y devuelve el control a quien
 * lo llama (login.js), igual que ya se hace con Google/BigQuery.
 *
 * ⚠️ El dominio *.snowflakecomputing.com no siempre da cabeceras CORS
 * a peticiones hechas desde el navegador/WebView del add-in. Si ves
 * errores de "Failed to fetch" al conectar o consultar, necesitas el
 * proxy CORS (ver PLANNING/proxy/cloudflare-worker.js) — apiBase ya
 * apunta al desplegado para Draco Planning; cámbialo si usas otro.
 *
 * ⚠️ El redirect_uri usado aquí (auth-callback-snowflake.html dentro
 * de ADDIN/src/) debe estar dado de alta en OAUTH_REDIRECT_URI de la
 * SECURITY INTEGRATION de Snowflake, además del que ya use Planning.
 */
const SF = {
    apiBase: "https://draco-snowflake-proxy.zipiwars.workers.dev",

    getAccount() {
        return localStorage.getItem("sf_account") || "";
    },
    setAccount(v) {
        localStorage.setItem("sf_account", v.trim());
    },

    getWarehouse() {
        return localStorage.getItem("sf_warehouse") || "";
    },
    setWarehouse(v) {
        localStorage.setItem("sf_warehouse", v.trim());
    },

    getRole() {
        return localStorage.getItem("sf_role") || "";
    },
    setRole(v) {
        localStorage.setItem("sf_role", v.trim());
    },

    getDatabase() {
        return localStorage.getItem("sf_database") || DracoConfig.snowflakeDatabase;
    },
    setDatabase(v) {
        localStorage.setItem("sf_database", v.trim());
    },

    getToken() {
        const token = localStorage.getItem("sf_access_token");
        const expires = localStorage.getItem("sf_token_expires");
        if (!token || !expires || Date.now() >= parseInt(expires, 10)) return null;
        return token;
    },

    isConnected() {
        return !!this.getToken();
    },

    logout() {
        ["sf_access_token", "sf_token_expires", "sf_refresh_token", "sf_pkce_verifier", "sf_oauth_state"]
            .forEach(k => localStorage.removeItem(k));
    },

    accountUrl() {
        return `https://${this.getAccount()}.snowflakecomputing.com`;
    },

    base() {
        return this.apiBase || this.accountUrl();
    },

    // ---------------------------------------------------------
    // PKCE
    // ---------------------------------------------------------
    randomString(len = 64) {
        const arr = new Uint8Array(len);
        crypto.getRandomValues(arr);
        return Array.from(arr, b => "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~"[b % 66]).join("");
    },

    base64url(buffer) {
        return btoa(String.fromCharCode(...new Uint8Array(buffer)))
            .replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
    },

    async pkceChallenge(verifier) {
        const data = new TextEncoder().encode(verifier);
        const digest = await crypto.subtle.digest("SHA-256", data);
        return this.base64url(digest);
    },

    // ---------------------------------------------------------
    // OAuth: construcción de la URL de autorización (PKCE)
    // ---------------------------------------------------------
    /** Genera y guarda el par PKCE + state, y devuelve la URL a abrir en el diálogo/popup de login. */
    async buildAuthUrl() {
        if (!this.getAccount()) {
            throw new Error("Indica primero el identificador de cuenta de Snowflake.");
        }
        if (DracoConfig.snowflakeClientId.startsWith("TU_SNOWFLAKE")) {
            throw new Error("Falta configurar snowflakeClientId en js/config.js.");
        }

        const verifier = this.randomString(64);
        const challenge = await this.pkceChallenge(verifier);
        const state = this.randomString(24);
        localStorage.setItem("sf_pkce_verifier", verifier);
        localStorage.setItem("sf_oauth_state", state);

        const redirectUri = new URL("auth-callback-snowflake.html", window.location.href).href;
        return `${this.accountUrl()}/oauth/authorize?` +
            `response_type=code` +
            `&client_id=${encodeURIComponent(DracoConfig.snowflakeClientId)}` +
            `&redirect_uri=${encodeURIComponent(redirectUri)}` +
            `&scope=${encodeURIComponent(DracoConfig.snowflakeScopes)}` +
            `&state=${encodeURIComponent(state)}` +
            `&code_challenge=${encodeURIComponent(challenge)}` +
            `&code_challenge_method=S256`;
    },

    /** Llamado tras recibir { code, state } desde auth-callback-snowflake.html */
    async handleAuthCode(code, state) {
        const savedState = localStorage.getItem("sf_oauth_state");
        const verifier = localStorage.getItem("sf_pkce_verifier");
        if (!verifier || state !== savedState) {
            throw new Error("Estado OAuth inválido (posible CSRF). Reintenta el login.");
        }

        const redirectUri = new URL("auth-callback-snowflake.html", window.location.href).href;
        const body = new URLSearchParams({
            grant_type: "authorization_code",
            code,
            redirect_uri: redirectUri,
            client_id: DracoConfig.snowflakeClientId,
            code_verifier: verifier
        });

        const response = await fetch(`${this.base()}/oauth/token-request`, {
            method: "POST",
            headers: { "Content-Type": "application/x-www-form-urlencoded" },
            body: body.toString()
        });

        const data = await response.json().catch(() => ({}));
        if (!response.ok || data.error) {
            throw new Error(data.error_description || data.error || `Error HTTP ${response.status} al canjear el código.`);
        }

        localStorage.setItem("sf_access_token", data.access_token);
        localStorage.setItem("sf_token_expires", Date.now() + (parseInt(data.expires_in || "3600", 10) * 1000));
        if (data.refresh_token) localStorage.setItem("sf_refresh_token", data.refresh_token);
        localStorage.removeItem("sf_pkce_verifier");
        localStorage.removeItem("sf_oauth_state");
    },

    // ---------------------------------------------------------
    // SQL API v2
    // ---------------------------------------------------------
    async execRaw(sql, { database, schema } = {}) {
        const token = this.getToken();
        if (!token) {
            const err = new Error("Sesión de Snowflake no válida o expirada.");
            err.code = "NO_AUTH";
            throw err;
        }
        if (!this.getWarehouse()) {
            throw new Error("No has indicado un warehouse de Snowflake.");
        }

        const payload = {
            statement: sql,
            timeout: 60,
            warehouse: this.getWarehouse(),
            database: database || this.getDatabase(),
            resultSetMetaData: { format: "json" }
        };
        if (schema) payload.schema = schema;
        if (this.getRole()) payload.role = this.getRole();

        const response = await fetch(`${this.base()}/api/v2/statements?requestId=${crypto.randomUUID()}`, {
            method: "POST",
            headers: {
                "Authorization": `Bearer ${token}`,
                "Content-Type": "application/json",
                "Accept": "application/json",
                "X-Snowflake-Authorization-Token-Type": "OAUTH"
            },
            body: JSON.stringify(payload)
        });

        let data = await response.json().catch(() => ({}));

        // 202 = todavía ejecutando; hacemos polling al handle hasta que resuelva.
        let attempts = 0;
        while (response.status === 202 && data.statementHandle && attempts < 30) {
            await new Promise(r => setTimeout(r, 1000));
            const poll = await fetch(`${this.base()}/api/v2/statements/${data.statementHandle}`, {
                headers: {
                    "Authorization": `Bearer ${token}`,
                    "Accept": "application/json",
                    "X-Snowflake-Authorization-Token-Type": "OAUTH"
                }
            });
            data = await poll.json().catch(() => ({}));
            if (poll.status !== 202) break;
            attempts++;
        }

        if (!response.ok && response.status !== 202) {
            throw new Error(data.message || `Error HTTP ${response.status} en Snowflake SQL API.`);
        }
        if (data.message && data.sqlState && data.sqlState !== "00000") {
            throw new Error(data.message);
        }
        return data;
    },

    /** Ejecuta SQL y devuelve directamente un array de objetos { columna: valor } */
    async runQuery(sql, opts) {
        const data = await this.execRaw(sql, opts);
        const cols = (data.resultSetMetaData && data.resultSetMetaData.rowType) || [];
        const rows = data.data || [];
        return rows.map(r => {
            const obj = {};
            cols.forEach((c, i) => { obj[c.name] = r[i]; });
            return obj;
        });
    },

    // ---------------------------------------------------------
    // Metadatos (para el explorador del modelo semántico)
    // ---------------------------------------------------------
    /** Lista las bases de datos visibles (equivalente a "proyectos" en BigQuery) */
    async listDatabases() {
        const rows = await this.runQuery("SHOW DATABASES");
        return rows.map(r => r.name);
    },

    /** Lista los esquemas de una base de datos (equivalente a "datasets") */
    async listSchemas(database) {
        const rows = await this.runQuery(`SHOW SCHEMAS IN DATABASE ${this.quoteIdent(database)}`);
        return rows.map(r => r.name);
    },

    /** Lista las tablas de un esquema */
    async listTables(database, schema) {
        const rows = await this.runQuery(`SHOW TABLES IN SCHEMA ${this.quoteIdent(database)}.${this.quoteIdent(schema)}`);
        return rows.map(r => r.name);
    },

    /** Devuelve los campos de una tabla como [{name, type}], con el tipo ya normalizado a los tipos canónicos de Draco */
    async getTableFields(database, schema, table) {
        const rows = await this.runQuery(
            `DESCRIBE TABLE ${this.quoteIdent(database)}.${this.quoteIdent(schema)}.${this.quoteIdent(table)}`
        );
        return rows.map(r => ({ name: r.name, type: this.mapType(r.type) }));
    },

    /** DESCRIBE TABLE devuelve tipos físicos (ej. "VARCHAR(16777216)", "NUMBER(38,0)"); los normalizamos */
    mapType(rawType) {
        const t = String(rawType || "").toUpperCase();
        if (t.startsWith("VARCHAR") || t.startsWith("CHAR") || t.startsWith("STRING") || t.startsWith("TEXT")) return "STRING";
        if (t.startsWith("NUMBER") || t.startsWith("DECIMAL") || t.startsWith("NUMERIC")) {
            // NUMBER(38,0) sin decimales -> entero; con decimales -> numérico
            const m = t.match(/\((\d+),(\d+)\)/);
            if (m && parseInt(m[2], 10) === 0) return "INTEGER";
            return "NUMERIC";
        }
        if (t.startsWith("FLOAT") || t.startsWith("DOUBLE") || t.startsWith("REAL")) return "FLOAT";
        if (t.startsWith("BOOLEAN")) return "BOOLEAN";
        if (t.startsWith("TIMESTAMP")) return "TIMESTAMP";
        if (t.startsWith("DATETIME")) return "DATETIME";
        if (t.startsWith("DATE")) return "DATE";
        return t;
    },

    quoteIdent(v) {
        return `"${String(v).replace(/"/g, '""')}"`;
    },

    esc(v) {
        return BQ.esc(v);
    }
};

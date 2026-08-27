/**
 * ============================================================
 * DRACO PLANNING — CLIENTE SNOWFLAKE (OAuth PKCE + SQL API v2)
 * ============================================================
 * Snowflake no soporta el flujo implícito de Google: usamos
 * "Authorization Code + PKCE" para cliente público (sin client
 * secret), tal como recomienda Snowflake para apps de navegador.
 *
 * ⚠️ Importante (léelo antes de desplegar): el dominio
 * *.snowflakecomputing.com no siempre devuelve cabeceras CORS para
 * peticiones fetch() hechas directamente desde el navegador. Si al
 * conectar ves errores de "CORS" / "Failed to fetch" en la consola,
 * necesitarás un pequeño proxy (ver README, sección Snowflake) que
 * reenvíe las peticiones a Snowflake añadiendo las cabeceras CORS.
 * El código de aquí ya está preparado para apuntar a ese proxy con
 * solo cambiar SF.apiBase.
 */
const SF = {
    // Si despliegas un proxy, cambia esto por su URL base (ej. "https://mi-proxy.workers.dev").
    // Si lo dejas vacío, se llama directamente a Snowflake.
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

    /** ¿Tenemos ya cuenta + warehouse listos para trabajar? */
    isReady() {
        return this.isConnected() && !!this.getAccount() && !!this.getWarehouse();
    },

    logout() {
        ["sf_access_token", "sf_token_expires", "sf_refresh_token", "sf_pkce_verifier", "sf_oauth_state"]
            .forEach(k => localStorage.removeItem(k));
    },

    accountUrl() {
        // El identificador de cuenta puede incluir puntos (org-account); se usa tal cual.
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
    // OAuth: inicio del flujo (popup)
    // ---------------------------------------------------------
    async connect() {
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
        const authUrl = `${this.accountUrl()}/oauth/authorize?` +
            `response_type=code` +
            `&client_id=${encodeURIComponent(DracoConfig.snowflakeClientId)}` +
            `&redirect_uri=${encodeURIComponent(redirectUri)}` +
            `&scope=${encodeURIComponent(DracoConfig.snowflakeScopes)}` +
            `&state=${encodeURIComponent(state)}` +
            `&code_challenge=${encodeURIComponent(challenge)}` +
            `&code_challenge_method=S256`;

        const width = 520, height = 700;
        const left = (window.screen.width / 2) - (width / 2);
        const top = (window.screen.height / 2) - (height / 2);
        window.open(authUrl, "SnowflakeAuthWindow", `width=${width},height=${height},top=${top},left=${left}`);
    },

    /** Llamado por auth-callback-snowflake.html vía postMessage con { code, state } o { error } */
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

    /**
     * Canjea el refresh token guardado por un access token nuevo, sin
     * abrir ningún popup. Se pedía el scope "refresh_token" al conectar
     * y se guardaba en sf_refresh_token, pero nada lo usaba: cada
     * caducidad del access token forzaba a reconectar desde cero
     * (y checkExistingTokens() en auth.js encima borraba el refresh
     * token al detectar el access token caducado, perdiendo la
     * posibilidad de renovar). Devuelve true si consiguió renovar.
     */
    async refreshAccessToken() {
        const refreshToken = localStorage.getItem("sf_refresh_token");
        if (!refreshToken || !this.getAccount()) return false;

        try {
            const body = new URLSearchParams({
                grant_type: "refresh_token",
                refresh_token: refreshToken,
                client_id: DracoConfig.snowflakeClientId
            });
            const response = await fetch(`${this.base()}/oauth/token-request`, {
                method: "POST",
                headers: { "Content-Type": "application/x-www-form-urlencoded" },
                body: body.toString()
            });
            const data = await response.json().catch(() => ({}));
            if (!response.ok || data.error) {
                console.warn("[Draco] No se pudo renovar el token de Snowflake:", data.error_description || data.error);
                return false;
            }
            localStorage.setItem("sf_access_token", data.access_token);
            localStorage.setItem("sf_token_expires", Date.now() + (parseInt(data.expires_in || "3600", 10) * 1000));
            // Snowflake puede rotar el refresh token en cada canje; si no manda uno nuevo, se conserva el actual.
            if (data.refresh_token) localStorage.setItem("sf_refresh_token", data.refresh_token);
            return true;
        } catch (e) {
            console.warn("[Draco] Error de red al renovar el token de Snowflake:", e.message);
            return false;
        }
    },

    /** Igual que isConnected(), pero intenta renovar en silencio con el refresh token antes de rendirse. */
    async ensureConnected() {
        if (this.isConnected()) return true;
        return this.refreshAccessToken();
    },

    // ---------------------------------------------------------
    // SQL API v2
    // ---------------------------------------------------------
    async execRaw(sql, { database, schema } = {}) {
        let token = this.getToken();
        if (!token) {
            // El token pudo caducar a media sesión (consulta larga, pestaña
            // abierta mucho rato...): intenta renovarlo antes de rendirte.
            if (await this.refreshAccessToken()) {
                token = this.getToken();
            }
        }
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

    async schemaExists(schemaName) {
        try {
            const rows = await this.runQuery(`SHOW SCHEMAS LIKE '${schemaName}' IN DATABASE ${this.getDatabase()}`);
            return rows.length > 0;
        } catch (e) {
            return false;
        }
    },

    async createSchema(schemaName, comment = "") {
        const commentSql = comment ? ` COMMENT = '${BQ.esc(comment)}'` : "";
        return this.runQuery(`CREATE SCHEMA IF NOT EXISTS ${this.getDatabase()}.${schemaName}${commentSql}`);
    },

    async dropSchema(schemaName) {
        return this.runQuery(`DROP SCHEMA IF EXISTS ${this.getDatabase()}.${schemaName} CASCADE`);
    },

    async ensureDatabase() {
        return this.runQuery(`CREATE DATABASE IF NOT EXISTS ${this.getDatabase()}`);
    }
};

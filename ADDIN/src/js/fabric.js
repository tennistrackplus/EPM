/**
 * ============================================================
 * EPM ADD-IN — CLIENTE MICROSOFT FABRIC (Lakehouse + Warehouse)
 * ============================================================
 * Mismo patrón que snowflake.js:
 *   - Login: OAuth 2.0 Authorization Code + PKCE contra Microsoft Entra ID
 *     (cliente público, app registrada como "Single-page application").
 *     connect() NO abre la ventana: buildAuthUrl() devuelve la URL y
 *     login.js la abre con Office.context.ui.displayDialogAsync.
 *   - Consultas: T-SQL contra el SQL endpoint del workspace
 *     (xxxx.datawarehouse.fabric.microsoft.com).
 *
 * ¿Por qué hay un "gateway"?
 *   El SQL endpoint de Fabric habla TDS (protocolo de SQL Server, TCP 1433),
 *   no HTTP. Un navegador / WebView de Office no puede abrir sockets TCP,
 *   así que las consultas pasan por un servicio mínimo (gateway/server.js)
 *   que recibe { server, database, sql } + el token del USUARIO y abre la
 *   conexión TDS con ESE token. El gateway no guarda credenciales ni
 *   tokens: cada consulta se ejecuta con la identidad y permisos de quien
 *   está conectado en Excel.
 *
 * Jerarquía (igual para Lakehouse y Warehouse):
 *   Conexión  = workspace  (su SQL endpoint)
 *   Nivel 1   = item       (Lakehouse o Warehouse; "base de datos" en T-SQL)
 *   Nivel 2   = schema     (dbo, finanzas...)
 *   Tabla
 *   SQL       = [item].[schema].[tabla]
 */
const FB = {
    // ---------------------------------------------------------
    // Configuración de la conexión activa (no sensible)
    // ---------------------------------------------------------
    getServer() {
        return localStorage.getItem("fb_server") || "";
    },
    setServer(v) {
        localStorage.setItem("fb_server", this.normalizeServer(v));
    },

    /** Warehouse/Lakehouse por defecto (opcional): contexto inicial de la conexión */
    getDefaultItem() {
        return localStorage.getItem("fb_default_item") || "";
    },
    setDefaultItem(v) {
        localStorage.setItem("fb_default_item", String(v || "").trim());
    },

    /** URL de un gateway propio del cliente (opcional). Vacío = el gateway de Draco (config.js) */
    getGatewayOverride() {
        return localStorage.getItem("fb_gateway_url") || "";
    },
    setGatewayOverride(v) {
        localStorage.setItem("fb_gateway_url", String(v || "").trim().replace(/\/+$/, ""));
    },

    /** Tenant de Entra ID (GUID o dominio). Vacío = "organizations" (cualquier cuenta de trabajo) */
    getTenant() {
        return localStorage.getItem("fb_tenant") || "";
    },
    setTenant(v) {
        localStorage.setItem("fb_tenant", String(v || "").trim());
    },

    tenantSegment() {
        return this.getTenant() || DracoConfig.fabricDefaultTenant || "organizations";
    },

    /**
     * Acepta tanto el host "xxxx.datawarehouse.fabric.microsoft.com" como la
     * cadena copiada de Fabric ("Data Source=xxxx...;..." o con ",1433").
     */
    normalizeServer(v) {
        let s = String(v || "").trim();
        const m = s.match(/([a-z0-9-]+\.(?:datawarehouse|database)\.fabric\.microsoft\.com)/i);
        if (m) s = m[1];
        return s.toLowerCase();
    },

    isValidServer(v) {
        return /^[a-z0-9-]+\.(datawarehouse|database)\.fabric\.microsoft\.com$/i.test(String(v || ""));
    },

    // ---------------------------------------------------------
    // Tokens
    // ---------------------------------------------------------
    getToken() {
        const token = localStorage.getItem("fb_access_token");
        const expires = localStorage.getItem("fb_token_expires");
        if (!token || !expires || Date.now() >= parseInt(expires, 10)) return null;
        return token;
    },

    /**
     * Conectado = hay token vigente O hay refresh token con el que renovarlo.
     * (Los refresh tokens de apps SPA en Entra duran 24 h.)
     */
    isConnected() {
        return !!this.getToken() || !!localStorage.getItem("fb_refresh_token");
    },

    logout() {
        ["fb_access_token", "fb_token_expires", "fb_refresh_token", "fb_pkce_verifier", "fb_oauth_state", "fb_account", "fb_api_token", "fb_api_token_expires"]
            .forEach(k => localStorage.removeItem(k));
    },

    getAccountName() {
        return localStorage.getItem("fb_account") || "";
    },

    _storeTokens(data) {
        localStorage.setItem("fb_access_token", data.access_token);
        // Se resta 1 minuto de margen para no usar un token a punto de caducar
        const ttl = (parseInt(data.expires_in || "3600", 10) - 60) * 1000;
        localStorage.setItem("fb_token_expires", String(Date.now() + ttl));
        if (data.refresh_token) localStorage.setItem("fb_refresh_token", data.refresh_token);
        if (data.id_token) {
            try {
                const payload = JSON.parse(atob(data.id_token.split(".")[1].replace(/-/g, "+").replace(/_/g, "/")));
                const name = payload.preferred_username || payload.upn || payload.email || "";
                if (name) localStorage.setItem("fb_account", name);
            } catch (e) { /* id_token opcional */ }
        }
    },

    /** Devuelve un token válido, renovándolo con el refresh token si hace falta */
    async ensureToken() {
        const current = this.getToken();
        if (current) return current;

        const refresh = localStorage.getItem("fb_refresh_token");
        if (!refresh) {
            const err = new Error("Sesión de Microsoft Fabric no válida o expirada. Vuelve a conectar.");
            err.code = "NO_AUTH";
            throw err;
        }

        const body = new URLSearchParams({
            grant_type: "refresh_token",
            client_id: DracoConfig.fabricClientId,
            refresh_token: refresh,
            scope: DracoConfig.fabricScopes
        });
        const response = await fetch(this.tokenUrl(), {
            method: "POST",
            headers: { "Content-Type": "application/x-www-form-urlencoded" },
            body: body.toString()
        });
        const data = await response.json().catch(() => ({}));
        if (!response.ok || data.error) {
            // Refresh token caducado o revocado: hay que volver a hacer login
            this.logout();
            const err = new Error("La sesión de Microsoft Fabric ha caducado. Vuelve a conectar. (" + (data.error_description || data.error || response.status) + ")");
            err.code = "NO_AUTH";
            throw err;
        }
        this._storeTokens(data);
        return data.access_token;
    },

    // ---------------------------------------------------------
    // PKCE (idéntico a snowflake.js)
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

    authorizeUrl() {
        return `https://login.microsoftonline.com/${encodeURIComponent(this.tenantSegment())}/oauth2/v2.0/authorize`;
    },

    tokenUrl() {
        return `https://login.microsoftonline.com/${encodeURIComponent(this.tenantSegment())}/oauth2/v2.0/token`;
    },

    redirectUri() {
        return new URL("auth-callback-fabric.html", window.location.href).href;
    },

    /**
     * En el login se pide consentimiento para el SQL endpoint Y para la API
     * REST de Fabric a la vez (un solo diálogo de consentimiento). El código
     * se canjea después solo por el token de SQL; el de la API se obtiene
     * con el refresh token cuando hace falta (getApiToken).
     */
    consentScopes() {
        return [DracoConfig.fabricScopes, DracoConfig.fabricApiScopes || ""].join(" ").trim();
    },

    // ---------------------------------------------------------
    // Token para la API REST de Fabric (api.fabric.microsoft.com)
    // ---------------------------------------------------------
    getApiTokenCached() {
        const token = localStorage.getItem("fb_api_token");
        const expires = localStorage.getItem("fb_api_token_expires");
        if (!token || !expires || Date.now() >= parseInt(expires, 10)) return null;
        return token;
    },

    async getApiToken() {
        const cached = this.getApiTokenCached();
        if (cached) return cached;

        const refresh = localStorage.getItem("fb_refresh_token");
        if (!refresh) {
            const err = new Error("No hay sesión de Microsoft Fabric. Conecta primero una conexión de Fabric desde el panel de Conexiones.");
            err.code = "NO_AUTH";
            throw err;
        }

        const body = new URLSearchParams({
            grant_type: "refresh_token",
            client_id: DracoConfig.fabricClientId,
            refresh_token: refresh,
            scope: (DracoConfig.fabricApiScopes || "") + " offline_access"
        });
        const response = await fetch(this.tokenUrl(), {
            method: "POST",
            headers: { "Content-Type": "application/x-www-form-urlencoded" },
            body: body.toString()
        });
        const data = await response.json().catch(() => ({}));
        if (!response.ok || data.error) {
            const desc = data.error_description || data.error || ("HTTP " + response.status);
            const err = new Error(
                /AADSTS65001|consent/i.test(desc)
                    ? "Faltan permisos para leer modelos semánticos. Vuelve a conectar tu conexión de Microsoft Fabric para aceptarlos."
                    : "No se pudo obtener acceso a la API de Fabric: " + desc
            );
            err.code = "NO_AUTH";
            throw err;
        }
        localStorage.setItem("fb_api_token", data.access_token);
        localStorage.setItem("fb_api_token_expires", String(Date.now() + (parseInt(data.expires_in || "3600", 10) - 60) * 1000));
        // Entra rota el refresh token: guardamos el nuevo para no invalidar la sesión
        if (data.refresh_token) localStorage.setItem("fb_refresh_token", data.refresh_token);
        return data.access_token;
    },

    // ---------------------------------------------------------
    // OAuth: URL de autorización
    // ---------------------------------------------------------
    async buildAuthUrl() {
        if (!DracoConfig.fabricClientId || DracoConfig.fabricClientId.startsWith("TU_")) {
            throw new Error("Falta configurar fabricClientId en js/config.js (Id. de aplicación de Entra ID).");
        }
        const verifier = this.randomString(64);
        const challenge = await this.pkceChallenge(verifier);
        const state = this.randomString(24);
        localStorage.setItem("fb_pkce_verifier", verifier);
        localStorage.setItem("fb_oauth_state", state);

        return this.authorizeUrl() + "?" +
            "response_type=code" +
            `&client_id=${encodeURIComponent(DracoConfig.fabricClientId)}` +
            `&redirect_uri=${encodeURIComponent(this.redirectUri())}` +
            "&response_mode=query" +
            `&scope=${encodeURIComponent(this.consentScopes())}` +
            `&state=${encodeURIComponent(state)}` +
            `&code_challenge=${encodeURIComponent(challenge)}` +
            "&code_challenge_method=S256" +
            "&prompt=select_account";
    },

    /** Llamado tras recibir { code, state } desde auth-callback-fabric.html */
    async handleAuthCode(code, state) {
        const savedState = localStorage.getItem("fb_oauth_state");
        const verifier = localStorage.getItem("fb_pkce_verifier");
        if (!verifier || state !== savedState) {
            throw new Error("Estado OAuth inválido (posible CSRF). Reintenta el login.");
        }

        const body = new URLSearchParams({
            grant_type: "authorization_code",
            client_id: DracoConfig.fabricClientId,
            code,
            redirect_uri: this.redirectUri(),
            code_verifier: verifier,
            scope: DracoConfig.fabricScopes
        });

        const response = await fetch(this.tokenUrl(), {
            method: "POST",
            headers: { "Content-Type": "application/x-www-form-urlencoded" },
            body: body.toString()
        });
        const data = await response.json().catch(() => ({}));
        if (!response.ok || data.error) {
            throw new Error(data.error_description || data.error || `Error HTTP ${response.status} al canjear el código.`);
        }

        this._storeTokens(data);
        localStorage.removeItem("fb_pkce_verifier");
        localStorage.removeItem("fb_oauth_state");
    },

    // ---------------------------------------------------------
    // Ejecución de T-SQL a través del gateway
    // ---------------------------------------------------------
    /** true si hay un gateway configurado (propio de la conexión o el de config.js) */
    isGatewayConfigured() {
        const base = String(this.getGatewayOverride() || DracoConfig.fabricGatewayUrl || "");
        return !!base && !base.startsWith("TU_");
    },

    gatewayUrl() {
        const base = String(this.getGatewayOverride() || DracoConfig.fabricGatewayUrl || "").replace(/\/+$/, "");
        if (!base || base.startsWith("TU_")) {
            throw new Error("Falta configurar fabricGatewayUrl en js/config.js (URL del gateway de Fabric).");
        }
        return base;
    },

    /**
     * Devuelve { columns:[{name,type}], rows:[[...]] , truncated } tal cual
     * lo manda el gateway (filas como arrays, en el orden del SELECT).
     */
    async execRaw(sql, { database } = {}) {
        const token = await this.ensureToken();
        const server = this.getServer();
        if (!this.isValidServer(server)) {
            throw new Error("SQL endpoint de Fabric no configurado en la conexión.");
        }

        const response = await fetch(this.gatewayUrl() + "/sql/query", {
            method: "POST",
            headers: {
                "Authorization": "Bearer " + token,
                "Content-Type": "application/json"
            },
            body: JSON.stringify({
                server,
                database: database || this.databaseFromSql(sql) || this.getDefaultItem() || "",
                sql
            })
        });

        const data = await response.json().catch(() => ({}));
        if (response.status === 401) {
            // Token rechazado por Fabric: forzamos a renovar la próxima vez
            localStorage.removeItem("fb_access_token");
            localStorage.removeItem("fb_token_expires");
            const err = new Error(data.error || "Fabric ha rechazado la autenticación. Vuelve a conectar.");
            err.code = "NO_AUTH";
            throw err;
        }
        if (!response.ok || data.error) {
            throw new Error(data.error || `Error HTTP ${response.status} en el gateway de Fabric.`);
        }
        if (data.truncated) {
            console.warn(`[Fabric] Resultado truncado a ${data.rows.length} filas por el gateway.`);
        }
        return data;
    },

    /** Array de objetos { columna: valor } (mismo contrato que SF.runQuery) */
    async runQuery(sql, opts) {
        const { columns, rows } = await this.execRaw(sql, opts);
        return rows.map(r => {
            const obj = {};
            columns.forEach((c, i) => { obj[c.name] = r[i]; });
            return obj;
        });
    },

    /** {fields:[{name}], rows:[{col: valor}]} (mismo contrato que SF.runQuerySql) */
    async runQuerySql(sql, opts) {
        const { columns, rows } = await this.execRaw(sql, opts);
        return {
            fields: columns.map(c => ({ name: c.name })),
            rows: rows.map(r => {
                const obj = {};
                columns.forEach((c, i) => { obj[c.name] = r[i]; });
                return obj;
            })
        };
    },

    /**
     * Mismo texto "pseudo JSON de BigQuery" que generan
     * snowflakeRowsToPseudoBqJson / svcRowsToPseudoBqJson, pero construido
     * a partir de las filas como ARRAYS: así se respeta el orden exacto de
     * columnas del SELECT aunque haya nombres repetidos.
     */
    async runQueryPseudoBqJson(sql, opts) {
        const { rows } = await this.execRaw(sql, opts);
        let out = '{"rows":[';
        rows.forEach((row, i) => {
            if (i > 0) out += ",";
            out += '{"f":[';
            row.forEach((val, j) => {
                if (j > 0) out += ",";
                if (val === null || val === undefined) {
                    out += '{"v": null}';
                } else {
                    const text = String(val).replace(/\\/g, "\\\\").replace(/"/g, '\\"');
                    out += '{"v": "' + text + '"}';
                }
            });
            out += "]}";
        });
        out += "]}";
        return out;
    },

    /**
     * Las consultas del add-in llevan nombres de 3 partes [item].[schema].[tabla];
     * el item de la primera tabla se usa como contexto de la conexión.
     */
    databaseFromSql(sql) {
        const m = String(sql || "").match(/\[((?:[^\]]|\]\])+)\]\s*\.\s*\[(?:[^\]]|\]\])+\]\s*\.\s*\[(?:[^\]]|\]\])+\]/);
        return m ? m[1].replace(/\]\]/g, "]") : "";
    },

    // ---------------------------------------------------------
    // Metadatos (explorador del modelo semántico)
    // ---------------------------------------------------------
    /** Nivel 1: items (Lakehouse / Warehouse) del workspace */
    async listItems() {
        const rows = await this.runQuery(
            "SELECT name FROM sys.databases WHERE name NOT IN ('master', 'tempdb', 'model', 'msdb') ORDER BY name",
            { database: this.getDefaultItem() || "" }
        );
        return rows.map(r => r.name);
    },

    /** Nivel 2: schemas con al menos una tabla o vista */
    async listSchemas(item) {
        const rows = await this.runQuery(
            "SELECT DISTINCT TABLE_SCHEMA FROM INFORMATION_SCHEMA.TABLES " +
            "WHERE TABLE_SCHEMA NOT IN ('sys', 'INFORMATION_SCHEMA', 'queryinsights') ORDER BY TABLE_SCHEMA",
            { database: item }
        );
        return rows.map(r => r.TABLE_SCHEMA);
    },

    async listTables(item, schema) {
        const rows = await this.runQuery(
            "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES " +
            `WHERE TABLE_SCHEMA = ${this.literal(schema)} ORDER BY TABLE_NAME`,
            { database: item }
        );
        return rows.map(r => r.TABLE_NAME);
    },

    /** [{name, type}] con el tipo ya normalizado a los tipos canónicos de Draco */
    async getTableFields(item, schema, table) {
        const rows = await this.runQuery(
            "SELECT COLUMN_NAME, DATA_TYPE, NUMERIC_SCALE FROM INFORMATION_SCHEMA.COLUMNS " +
            `WHERE TABLE_SCHEMA = ${this.literal(schema)} AND TABLE_NAME = ${this.literal(table)} ` +
            "ORDER BY ORDINAL_POSITION",
            { database: item }
        );
        return rows.map(r => ({ name: r.COLUMN_NAME, type: this.mapType(r.DATA_TYPE, r.NUMERIC_SCALE) }));
    },

    /** Tipos de INFORMATION_SCHEMA.COLUMNS (varchar, bigint, decimal...) -> tipos Draco */
    mapType(rawType, numericScale) {
        const t = String(rawType || "").toLowerCase();
        if (["varchar", "char", "nvarchar", "nchar", "text", "ntext", "uniqueidentifier"].includes(t)) return "STRING";
        if (["int", "bigint", "smallint", "tinyint"].includes(t)) return "INTEGER";
        if (t === "decimal" || t === "numeric" || t === "money" || t === "smallmoney") {
            return Number(numericScale) === 0 ? "INTEGER" : "NUMERIC";
        }
        if (t === "float" || t === "real") return "FLOAT";
        if (t === "bit") return "BOOLEAN";
        if (t === "date") return "DATE";
        if (t === "datetime2" || t === "datetime" || t === "smalldatetime") return "DATETIME";
        if (t === "datetimeoffset") return "TIMESTAMP";
        if (t === "time") return "TIME";
        return t.toUpperCase();
    },

    quoteIdent(v) {
        return "[" + String(v).replace(/\]/g, "]]") + "]";
    },

    literal(v) {
        return "N'" + String(v == null ? "" : v).replace(/'/g, "''") + "'";
    }
};

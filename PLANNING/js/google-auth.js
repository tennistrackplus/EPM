/**
 * ============================================================
 * DRACO — AUTENTICACIÓN CON GOOGLE (común al addin y a Planning)
 * ============================================================
 * Dos modos, elegidos en cada conexión (config.authMode):
 *
 *  "token"   Sencillo. OAuth 2.0 en el navegador con el OAuth Client ID
 *            del PROPIO CLIENTE (config.oauthClientId, tipo "Aplicación
 *            web"). Google devuelve un access token de 1 hora; al caducar
 *            hay que volver a conectar.
 *
 *  "backend" Recomendado por Google. Authorization Code + PKCE contra un
 *            servicio propio del cliente (config.authBrokerUrl, ver
 *            AUTH-BROKER/). El client secret y el refresh token NUNCA
 *            llegan al navegador: el servicio entrega access tokens de
 *            corta duración y los renueva solo mientras dura la sesión.
 *
 * PRODUCCIÓN (auditoría F-10): solo se admite el modo "backend". El modo
 * "token" (flujo implícito, desaconsejado por OAuth 2.1) queda disponible
 * únicamente si DracoConfig.allowSimpleOAuth = true, para pruebas.
 *
 * En ambos modos:
 *  - Se usa un parámetro `state` aleatorio que se comprueba a la vuelta
 *    (protección CSRF / inyección de tokens).
 *  - El access token se guarda solo en sessionStorage (BQ.setToken).
 *
 * Requiere: BQ (bigquery.js) y DracoConfig.googleScopes (config.js).
 * Este fichero es idéntico en ADDIN/src/js/google-auth.js y
 * PLANNING/js/google-auth.js.
 */
const GoogleAuth = {
    SESSION_KEY: "draco_auth_session",     // { brokerUrl, handle } (solo modo backend)
    CLIENT_ID_RE: /^[0-9]+-[a-z0-9]+\.apps\.googleusercontent\.com$/,
    _pendingState: null,
    _refreshTimer: null,

    // ---------------------------------------------------------
    // Validación de la configuración de la conexión
    // ---------------------------------------------------------
    /** ¿Se permite el modo sencillo (flujo implícito)? Solo para pruebas. */
    simpleModeAllowed() {
        return !!(typeof DracoConfig !== "undefined" && DracoConfig.allowSimpleOAuth === true);
    },

    /** Deja en el selector de modo solo las opciones permitidas. */
    applyModeOptions(selectEl) {
        if (!selectEl || this.simpleModeAllowed()) return;
        const opt = selectEl.querySelector('option[value="token"]');
        if (opt) opt.remove();
        selectEl.value = "backend";
    },

    normalizeConfig(cfg) {
        const c = cfg || {};
        const mode = c.authMode === "token" && this.simpleModeAllowed() ? "token" : "backend";
        return {
            authMode: mode,
            oauthClientId: String(c.oauthClientId || "").trim(),
            authBrokerUrl: String(c.authBrokerUrl || "").trim().replace(/\/+$/, "")
        };
    },

    /** Devuelve un texto de error o null si la configuración es válida. */
    validateConfig(cfg) {
        if (cfg && cfg.authMode === "token" && !this.simpleModeAllowed()) {
            return "En producción solo se admite el servicio de autenticación: edita la conexión e indica su URL.";
        }
        const c = this.normalizeConfig(cfg);
        if (c.authMode === "backend") {
            if (!/^https:\/\/[^\s/?#]+(\/[^\s?#]*)?$/i.test(c.authBrokerUrl)) {
                return "Indica la URL (https://...) del servicio de autenticación del cliente.";
            }
            return null;
        }
        if (!this.CLIENT_ID_RE.test(c.oauthClientId)) {
            return "Indica el OAuth Client ID de Google del cliente (termina en .apps.googleusercontent.com).";
        }
        return null;
    },

    // ---------------------------------------------------------
    // Inicio de sesión
    // ---------------------------------------------------------
    _randomState() {
        const bytes = new Uint8Array(24);
        crypto.getRandomValues(bytes);
        return Array.from(bytes, b => b.toString(16).padStart(2, "0")).join("");
    },

    /** URL a la que hay que mandar al usuario para iniciar sesión. */
    buildLoginUrl(cfg, redirectUri) {
        const c = this.normalizeConfig(cfg);
        const err = this.validateConfig(c);
        if (err) throw new Error(err);
        this._pendingState = this._randomState();

        if (c.authMode === "backend") {
            this._pendingBroker = c.authBrokerUrl;
            return `${c.authBrokerUrl}/login?` + new URLSearchParams({
                redirect_uri: redirectUri,
                scope: DracoConfig.googleScopes,
                state: this._pendingState
            }).toString();
        }

        this._pendingBroker = null;
        return "https://accounts.google.com/o/oauth2/v2/auth?" + new URLSearchParams({
            client_id: c.oauthClientId,
            redirect_uri: redirectUri,
            response_type: "token",
            scope: DracoConfig.googleScopes,
            include_granted_scopes: "true",
            prompt: "select_account",
            state: this._pendingState
        }).toString();
    },

    /**
     * Procesa la respuesta del callback ({status, token, expiresIn, state,
     * session}). Devuelve true si la sesión queda establecida. Lanza error
     * si el `state` no coincide con el de la petición en curso.
     */
    acceptResponse(response) {
        const expected = this._pendingState;
        this._pendingState = null;
        if (!expected || !response || response.state !== expected) {
            throw new Error("Respuesta de inicio de sesión no válida (state). Vuelve a intentarlo.");
        }
        if (response.status !== "success" || !response.token) return false;

        const expiresAt = Date.now() + (parseInt(response.expiresIn, 10) || 3600) * 1000;
        BQ.setToken(response.token, expiresAt, response.scope);

        if (this._pendingBroker && response.session) {
            this._setSession({ brokerUrl: this._pendingBroker, handle: response.session });
            this.scheduleRefresh(expiresAt);
        } else {
            this._clearSession();
        }
        this._pendingBroker = null;
        return true;
    },

    // ---------------------------------------------------------
    // Modo backend: renovación del access token
    // ---------------------------------------------------------
    _getSession() {
        try { return JSON.parse(sessionStorage.getItem(this.SESSION_KEY) || "null"); } catch (e) { return null; }
    },
    _setSession(s) { sessionStorage.setItem(this.SESSION_KEY, JSON.stringify(s)); },
    _clearSession() { sessionStorage.removeItem(this.SESSION_KEY); },

    hasBackendSession() {
        const s = this._getSession();
        return !!(s && s.brokerUrl && s.handle);
    },

    /** Pide un access token nuevo al servicio de autenticación. */
    async refresh() {
        const s = this._getSession();
        if (!s || !s.brokerUrl || !s.handle) return false;
        try {
            const resp = await fetch(`${s.brokerUrl}/token`, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ session: s.handle }),
                credentials: "omit",
                cache: "no-store"
            });
            if (!resp.ok) {
                if (resp.status === 401) { this._clearSession(); BQ.logout(); }
                return false;
            }
            const data = await resp.json();
            if (!data.access_token) return false;
            const expiresAt = Date.now() + (parseInt(data.expires_in, 10) || 3600) * 1000;
            BQ.setToken(data.access_token, expiresAt, data.scope);
            this.scheduleRefresh(expiresAt);
            return true;
        } catch (e) {
            console.warn("[Draco] No se pudo renovar la sesión de Google:", e);
            return false;
        }
    },

    /** Programa la renovación 5 minutos antes de que caduque el token. */
    scheduleRefresh(expiresAt) {
        if (this._refreshTimer) clearTimeout(this._refreshTimer);
        if (!this.hasBackendSession()) return;
        const delay = Math.max(10000, Number(expiresAt) - Date.now() - 5 * 60 * 1000);
        this._refreshTimer = setTimeout(() => this.refresh(), delay);
    },

    /** Al cargar una página: reanuda la renovación (o renueva ya si caducó). */
    async resume() {
        if (!this.hasBackendSession()) return;
        const expires = parseInt(sessionStorage.getItem(BQ.EXPIRES_KEY) || "0", 10);
        if (!BQ.getToken() || expires - Date.now() < 5 * 60 * 1000) {
            await this.refresh();
        } else {
            this.scheduleRefresh(expires);
        }
    },

    /** Cierra la sesión: borra el token y, en modo backend, revoca el refresh token. */
    async logout() {
        const s = this._getSession();
        if (this._refreshTimer) clearTimeout(this._refreshTimer);
        this._clearSession();
        BQ.logout();
        if (s && s.brokerUrl && s.handle) {
            try {
                await fetch(`${s.brokerUrl}/logout`, {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ session: s.handle }),
                    credentials: "omit"
                });
            } catch (e) { /* la sesión local ya está cerrada */ }
        }
    }
};

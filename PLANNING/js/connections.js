/**
 * ============================================================
 * DRACO PLANNING — ALMACÉN DE CONEXIONES GUARDADAS
 * ============================================================
 * Copia EXACTA (misma clave de localStorage `epm_connections` /
 * `epm_active_connection_id`, mismo formato de objeto) de
 * `ADDIN/src/js/connections.js`. Se mantiene idéntica a propósito:
 * si Draco Planning y el add-in de Excel se sirven desde el MISMO
 * origen (mismo dominio, aunque sea en rutas distintas, ej.
 * `https://tuapp.com/planning/` y `https://tuapp.com/addin/`), el
 * navegador comparte el mismo localStorage y las conexiones creadas
 * en uno aparecen automáticamente en el otro — no hay sincronización
 * de verdad, es el mismo almacén físico.
 *
 * Si en el futuro se sirven desde orígenes distintos, esto deja de
 * funcionar solo (localStorage es por origen); habría que sustituir
 * este almacén por un backend compartido (o guardarlo también como
 * fila en DRACO_CONTROL). De momento no hace falta.
 *
 * IMPORTANTE: aquí solo se guarda la CONFIGURACIÓN de cada conexión
 * (nombre visible, proveedor, proyecto GCP, cuenta de Snowflake,
 * etc.). Los tokens de acceso (OAuth) siguen gestionándose por
 * separado en BQ / SF: son credenciales de sesión, no datos de la
 * conexión en sí, y su expiración es independiente de si la conexión
 * sigue "guardada" en el menú.
 */
const Connections = {
    STORAGE_KEY: "epm_connections",
    ACTIVE_KEY: "epm_active_connection_id",

    PROVIDER_LABELS: {
        bigquery: "BigQuery",
        amazon: "Amazon",
        fabric: "Microsoft Fabric",
        snowflake: "Snowflake",
        datasphere: "SAP Datasphere",
        s4cds: "CDS de S4"
    },

    // Proveedores con flujo de conexión implementado hoy; el resto se
    // puede guardar como perfil pero mostrará "próximamente" al conectar.
    IMPLEMENTED_PROVIDERS: ["bigquery", "snowflake"],

    labelFor(providerKey) {
        return this.PROVIDER_LABELS[providerKey] || providerKey;
    },

    isImplemented(providerKey) {
        return this.IMPLEMENTED_PROVIDERS.includes(providerKey);
    },

    // ---------------------------------------------------------
    // Persistencia
    // ---------------------------------------------------------
    _load() {
        try {
            const raw = localStorage.getItem(this.STORAGE_KEY);
            const list = raw ? JSON.parse(raw) : [];
            return Array.isArray(list) ? list : [];
        } catch (e) {
            console.error("[EPM Connections] No se pudieron leer las conexiones guardadas:", e);
            return [];
        }
    },

    _save(list) {
        localStorage.setItem(this.STORAGE_KEY, JSON.stringify(list));
    },

    list() {
        return this._load().sort((a, b) => (a.createdAt || 0) - (b.createdAt || 0));
    },

    getById(id) {
        return this._load().find(c => c.id === id) || null;
    },

    generateId() {
        return "conn_" + Date.now().toString(36) + "_" + Math.random().toString(36).slice(2, 8);
    },

    defaultName(providerKey) {
        const base = this.labelFor(providerKey);
        const existing = this._load().filter(c => c.provider === providerKey).length;
        return existing ? `${base} (${existing + 1})` : base;
    },

    /** Crea y guarda una nueva conexión. config = objeto libre según el proveedor. */
    create({ name, provider, config }) {
        const list = this._load();
        const conn = {
            id: this.generateId(),
            name: (name || "").trim() || this.defaultName(provider),
            provider,
            config: config || {},
            createdAt: Date.now(),
            updatedAt: Date.now()
        };
        list.push(conn);
        this._save(list);
        return conn;
    },

    /** Actualiza una conexión existente (fusiona "config" en lugar de sobreescribirlo entero) */
    update(id, patch) {
        const list = this._load();
        const idx = list.findIndex(c => c.id === id);
        if (idx === -1) return null;
        const prev = list[idx];
        list[idx] = {
            ...prev,
            ...patch,
            config: { ...(prev.config || {}), ...((patch && patch.config) || {}) },
            updatedAt: Date.now()
        };
        this._save(list);
        return list[idx];
    },

    remove(id) {
        const list = this._load().filter(c => c.id !== id);
        this._save(list);
        if (this.getActiveId() === id) {
            this.setActiveId(null);
        }
    },

    // ---------------------------------------------------------
    // Conexión activa (la que está usando el resto del add-in ahora mismo)
    // ---------------------------------------------------------
    getActiveId() {
        return localStorage.getItem(this.ACTIVE_KEY) || null;
    },

    setActiveId(id) {
        if (id) {
            localStorage.setItem(this.ACTIVE_KEY, id);
        } else {
            localStorage.removeItem(this.ACTIVE_KEY);
        }
    },

    getActive() {
        const id = this.getActiveId();
        return id ? this.getById(id) : null;
    }
};

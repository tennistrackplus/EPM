/**
 * Lógica de la landing page: conexión OAuth a BigQuery o Snowflake,
 * configuración del "hogar" de trabajo (proyecto GCP / cuenta+warehouse)
 * y arranque del esquema de control sobre el motor elegido.
 *
 * Las conexiones se guardan con el mismo almacén que usa el add-in de
 * Excel (`js/connections.js`, clave de localStorage `epm_connections`):
 * si Planning y el add-in se sirven desde el mismo origen, lo que se
 * crea en uno aparece automáticamente en el otro. Ver connections.js
 * para más detalle.
 */
const LoginApp = {
    selectedProvider: null,
    connectedProviders: {},
    activeConnectionId: null, // conexión guardada que se está usando/creando ahora mismo
    pendingInstallProvider: null, // "bigquery" mientras se muestra el panel de instalación del esquema

    PROVIDER_ICONS: {
        bigquery: "https://www.vectorlogo.zone/logos/google_bigquery/google_bigquery-icon.svg",
        snowflake: "https://app.eu.peliqan.io/img/db/snowflake.svg",
        amazon: "https://www.vectorlogo.zone/logos/amazon_aws/amazon_aws-icon.svg",
        fabric: "https://community.fabric.microsoft.com/html/assets/fabric-expo-icon.svg",
        datasphere: "https://upload.wikimedia.org/wikipedia/commons/5/59/SAP_2011_logo.svg",
        s4cds: "https://upload.wikimedia.org/wikipedia/commons/5/59/SAP_2011_logo.svg"
    },

    init() {
        this.bindEvents();
        this.renderSavedConnections();
        this.checkExistingTokens();
        this.setupMessageListener();
    },

    showAlert(msg, isError = false) {
        const toast = document.getElementById("toastMessage");
        if (toast) {
            toast.innerText = msg;
            toast.className = `toast-message visible ${isError ? "error" : "info"}`;
            setTimeout(() => toast.classList.remove("visible"), 6000);
        }
    },

    escapeHtml(str) {
        const div = document.createElement("div");
        div.innerText = str == null ? "" : String(str);
        return div.innerHTML;
    },

    getProviderDisplayName(providerKey) {
        return (window.Connections && Connections.labelFor)
            ? Connections.labelFor(providerKey)
            : (providerKey === "bigquery" ? "BigQuery" : providerKey === "snowflake" ? "Snowflake" : providerKey);
    },

    bindEvents() {
        const cards = document.querySelectorAll(".connector-card");
        cards.forEach(card => {
            card.addEventListener("click", () => {
                cards.forEach(c => c.classList.remove("selected"));
                card.classList.add("selected");
                this.selectedProvider = card.getAttribute("data-provider");
                this.activeConnectionId = null; // vista "crear": no hay conexión guardada detrás todavía
                this.hideInstallPanel();
                this.togglePanels();
                this.updateActionButton();
            });
        });

        const gcpSelect = document.getElementById("gcpProjectSelect");
        if (gcpSelect) {
            gcpSelect.addEventListener("change", () => {
                BQ.setGcpProject(gcpSelect.value);
                this.persistActiveConnectionConfig();
                this.hideInstallPanel();
                this.updateActionButton();
            });
        }

        const btnToggleNew = document.getElementById("btnToggleNewConnection");
        if (btnToggleNew) {
            btnToggleNew.addEventListener("click", () => this.showNewConnectionForm());
        }

        const btnInstall = document.getElementById("btnInstallSchema");
        if (btnInstall) {
            btnInstall.addEventListener("click", () => this.installControlSchema());
        }

        const savedList = document.getElementById("savedConnectionsList");
        if (savedList) {
            savedList.addEventListener("click", (e) => {
                const card = e.target.closest(".conn-card");
                if (!card) return;
                this.connectExisting(card.getAttribute("data-id"));
            });
        }
    },

    // =====================================================================
    // Conexiones guardadas (compartidas con el add-in vía Connections)
    // =====================================================================
    renderSavedConnections() {
        const section = document.getElementById("savedConnectionsSection");
        const list = document.getElementById("savedConnectionsList");
        if (!section || !list || !window.Connections) return;

        const connections = Connections.list();
        if (connections.length === 0) {
            section.classList.remove("visible");
            this.showNewConnectionForm(); // sin conexiones guardadas, ir directo al formulario
            return;
        }

        section.classList.add("visible");
        const activeId = Connections.getActiveId();
        list.innerHTML = connections.map(conn => {
            const isActive = conn.id === activeId;
            const icon = this.PROVIDER_ICONS[conn.provider] || "";
            const subtitle = this.buildSubtitle(conn);
            return `
                <div class="conn-card ${isActive ? "active" : ""}" data-id="${conn.id}">
                    <div class="conn-info">
                        <div class="conn-icon"><img src="${icon}" alt="${this.getProviderDisplayName(conn.provider)}"></div>
                        <div class="conn-details">
                            <h3>${this.escapeHtml(conn.name)}</h3>
                            <span>${this.escapeHtml(subtitle)}</span>
                        </div>
                    </div>
                    <div class="conn-side">
                        <span class="badge-status">Usar</span>
                    </div>
                </div>`;
        }).join("");

        this.hideNewConnectionForm();
    },

    buildSubtitle(conn) {
        const label = this.getProviderDisplayName(conn.provider);
        const cfg = conn.config || {};
        if (conn.provider === "bigquery") {
            const project = cfg.homeProjectId || cfg.billingProjectId;
            return project ? `${label} · ${project}` : label;
        }
        if (conn.provider === "snowflake") {
            return cfg.account ? `${label} · ${cfg.account}` : label;
        }
        return label;
    },

    showNewConnectionForm() {
        document.getElementById("connectorsContainer").classList.remove("collapsed");
    },

    hideNewConnectionForm() {
        document.getElementById("connectorsContainer").classList.add("collapsed");
    },

    /** Carga en BQ/SF los parámetros guardados de una conexión (sin tocar los tokens de sesión) */
    applyConnectionConfig(conn) {
        const cfg = conn.config || {};
        if (conn.provider === "bigquery") {
            // homeProjectId: proyecto "hogar" donde vive DRACO_CONTROL (propio de Planning).
            // billingProjectId: campo que ya usaba el add-in; si la conexión se creó allí y
            // todavía no se ha elegido un proyecto hogar en Planning, se usa como punto de
            // partida razonable (normalmente coinciden).
            BQ.setGcpProject(cfg.homeProjectId || cfg.billingProjectId || "");
        } else if (conn.provider === "snowflake") {
            SF.setAccount(cfg.account || "");
            SF.setWarehouse(cfg.warehouse || "");
            SF.setDatabase(cfg.database || DracoConfig.snowflakeDatabase);
            SF.setRole(cfg.role || "");
        }
    },

    /** Crea o actualiza (merge) la conexión activa con el estado actual de BQ/SF */
    persistActiveConnectionConfig() {
        if (!window.Connections || !this.selectedProvider) return;

        let config = {};
        if (this.selectedProvider === "bigquery") {
            const project = BQ.getGcpProject() || "";
            // homeProjectId: lo usa Planning (dataset DRACO_CONTROL / DRACO_<proyecto>).
            // billingProjectId: mismo campo que ya usaba el add-in (proyecto que paga las
            // consultas). Se rellenan ambos con el mismo valor por defecto para que una
            // conexión creada en Planning ya sirva tal cual en el add-in, y viceversa
            // (ver applyConnectionConfig, que hace el fallback contrario).
            const existing = this.activeConnectionId && window.Connections ? Connections.getById(this.activeConnectionId) : null;
            const existingBilling = existing && existing.config ? existing.config.billingProjectId : "";
            config = { homeProjectId: project, billingProjectId: existingBilling || project };
        } else if (this.selectedProvider === "snowflake") {
            config = {
                account: SF.getAccount() || "",
                warehouse: SF.getWarehouse() || "",
                database: SF.getDatabase() || DracoConfig.snowflakeDatabase,
                role: SF.getRole() || ""
            };
        } else {
            return;
        }

        if (this.activeConnectionId) {
            Connections.update(this.activeConnectionId, { config });
        } else {
            const conn = Connections.create({ provider: this.selectedProvider, config });
            this.activeConnectionId = conn.id;
        }
        Connections.setActiveId(this.activeConnectionId);
    },

    /** Reutiliza los parámetros ya guardados de una conexión y conecta sin volver a preguntarlos */
    async connectExisting(id) {
        const conn = Connections.getById(id);
        if (!conn) return;

        this.activeConnectionId = id;
        this.selectedProvider = conn.provider;
        Connections.setActiveId(id);
        Provider.setKey(conn.provider);
        this.applyConnectionConfig(conn);
        this.hideInstallPanel();

        if (conn.provider === "bigquery") {
            if (BQ.isConnected()) {
                this.setProviderState("bigquery", true);
                await this.proceedBigQuery();
            } else {
                this.connectBigQuery();
            }
            return;
        }

        if (conn.provider === "snowflake") {
            if (SF.isConnected()) {
                this.setProviderState("snowflake", true);
                await this.proceedSnowflake();
            } else {
                try {
                    await SF.connect();
                } catch (err) {
                    this.showAlert("Error al conectar con Snowflake: " + err.message, true);
                }
            }
            return;
        }

        this.showAlert(`El conector para ${this.getProviderDisplayName(conn.provider)} estará disponible próximamente.`);
    },

    togglePanels() {
        document.getElementById("gcpProjectPicker").classList.toggle(
            "visible", this.selectedProvider === "bigquery" && BQ.isConnected());
        document.getElementById("snowflakeConfigPanel").classList.toggle(
            "visible", this.selectedProvider === "snowflake");

        if (this.selectedProvider === "snowflake") {
            document.getElementById("sfAccount").value = SF.getAccount();
            document.getElementById("sfWarehouse").value = SF.getWarehouse();
            document.getElementById("sfDatabase").value = SF.getDatabase();
            document.getElementById("sfRole").value = SF.getRole();
        }
    },

    setupMessageListener() {
        window.addEventListener("message", (event) => {
            if (typeof event.data !== "string") return;
            try {
                const data = JSON.parse(event.data);
                if (data.status) this.handleAuthResponse(event.data);
            } catch (e) { /* mensaje de otro origen, se ignora */ }
        });
    },

    checkExistingTokens() {
        if (BQ.isConnected()) {
            this.setProviderState("bigquery", true);
        } else {
            localStorage.removeItem("bigquery_access_token");
            localStorage.removeItem("bigquery_token_expires");
        }

        if (SF.isConnected()) {
            this.setProviderState("snowflake", true);
        } else {
            SF.logout();
        }
    },

    // ---------------------------------------------------------
    // BigQuery
    // ---------------------------------------------------------
    connectBigQuery() {
        if (DracoConfig.googleClientId.startsWith("TU_GOOGLE_CLIENT_ID")) {
            this.showAlert("Falta configurar el Client ID de Google en js/config.js antes de poder conectar.", true);
            return;
        }

        const redirectUri = new URL("auth-callback.html", window.location.href).href;
        const authUrl = "https://accounts.google.com/o/oauth2/v2/auth?" +
            `client_id=${encodeURIComponent(DracoConfig.googleClientId)}` +
            `&redirect_uri=${encodeURIComponent(redirectUri)}` +
            "&response_type=token" +
            `&scope=${encodeURIComponent(DracoConfig.googleScopes)}` +
            "&prompt=consent";

        const width = 520, height = 660;
        const left = (window.screen.width / 2) - (width / 2);
        const top = (window.screen.height / 2) - (height / 2);
        window.open(authUrl, "GoogleAuthWindow", `width=${width},height=${height},top=${top},left=${left}`);
    },

    async loadGcpProjects() {
        const picker = document.getElementById("gcpProjectPicker");
        const select = document.getElementById("gcpProjectSelect");
        if (!picker || !select) return;

        select.innerHTML = `<option value="">Cargando proyectos...</option>`;
        picker.classList.add("visible");

        try {
            const projects = await BQ.listProjects();
            if (!projects.length) {
                select.innerHTML = `<option value="">No se encontraron proyectos GCP</option>`;
                return;
            }
            const saved = BQ.getGcpProject();
            select.innerHTML = projects.map(p => {
                const id = p.id || (p.projectReference && p.projectReference.projectId);
                const name = p.friendlyName || id;
                return `<option value="${id}">${name} (${id})</option>`;
            }).join("");

            if (saved && projects.some(p => (p.id || p.projectReference.projectId) === saved)) {
                select.value = saved;
            } else {
                BQ.setGcpProject(select.value);
            }
            this.updateActionButton();
        } catch (err) {
            select.innerHTML = `<option value="">Error al listar proyectos</option>`;
            this.showAlert("No se pudieron listar los proyectos de GCP: " + err.message, true);
        }
    },

    // ---------------------------------------------------------
    // Snowflake
    // ---------------------------------------------------------
    async connectSnowflake() {
        const account = document.getElementById("sfAccount").value.trim();
        const warehouse = document.getElementById("sfWarehouse").value.trim();
        const database = document.getElementById("sfDatabase").value.trim() || DracoConfig.snowflakeDatabase;
        const role = document.getElementById("sfRole").value.trim();

        if (!account || !warehouse) {
            this.showAlert("Indica al menos la cuenta y el warehouse de Snowflake.", true);
            return;
        }

        SF.setAccount(account);
        SF.setWarehouse(warehouse);
        SF.setDatabase(database);
        SF.setRole(role);

        try {
            await SF.connect();
        } catch (err) {
            this.showAlert("Error al iniciar la conexión con Snowflake: " + err.message, true);
        }
    },

    // ---------------------------------------------------------
    // Respuesta común de los popups de autenticación
    // ---------------------------------------------------------
    async handleAuthResponse(messageString) {
        try {
            const response = JSON.parse(messageString);

            if (response.provider === "bigquery") {
                if (response.status === "success") {
                    localStorage.setItem("bigquery_access_token", response.token);
                    localStorage.setItem("bigquery_token_expires", Date.now() + (parseInt(response.expiresIn, 10) * 1000));
                    this.setProviderState("bigquery", true);
                    this.persistActiveConnectionConfig();
                    this.updateActionButton();
                    this.showAlert("¡Conexión con Google BigQuery establecida con éxito!");
                    await this.loadGcpProjects();
                    this.persistActiveConnectionConfig();
                } else {
                    this.showAlert("Error de autenticación con BigQuery: " + (response.error || "Desconocido"), true);
                }
                return;
            }

            if (response.provider === "snowflake") {
                if (response.status === "success" && response.code) {
                    await SF.handleAuthCode(response.code, response.state);
                    this.setProviderState("snowflake", true);
                    this.persistActiveConnectionConfig();
                    this.updateActionButton();
                    this.showAlert("¡Conexión con Snowflake establecida con éxito!");
                } else {
                    this.showAlert("Error de autenticación con Snowflake: " + (response.error || "Desconocido"), true);
                }
                return;
            }
        } catch (err) {
            console.error("Error leyendo respuesta de autenticación:", err);
            this.showAlert("Error al procesar la respuesta de login: " + err.message, true);
        }
    },

    // ---------------------------------------------------------
    // Puerta de instalación del esquema de control (BigQuery)
    // ---------------------------------------------------------
    showInstallPanel() {
        this.pendingInstallProvider = "bigquery";
        document.getElementById("installPanel").classList.add("visible");
    },

    hideInstallPanel() {
        this.pendingInstallProvider = null;
        const panel = document.getElementById("installPanel");
        if (panel) panel.classList.remove("visible");
    },

    async installControlSchema() {
        const btn = document.getElementById("btnInstallSchema");
        const btnAction = document.getElementById("btnAction");
        btn.disabled = true;
        try {
            await DracoSchema.bootstrap((msg) => { btn.innerText = msg; });
            window.location.href = "app.html";
        } catch (err) {
            btn.disabled = false;
            btn.innerText = "Instalar esquema de control";
            this.showAlert("Error al instalar el esquema de control: " + err.message, true);
        } finally {
            if (btnAction) btnAction.disabled = false;
        }
    },

    /**
     * BigQuery ya está conectado y con proyecto elegido: comprueba si el
     * esquema de control existe. Si existe, entra directamente (bootstrap
     * es idempotente y de paso aplica altas de columnas nuevas). Si no,
     * en vez de crearlo en silencio, muestra el panel de instalación.
     */
    async proceedBigQuery() {
        const btnAction = document.getElementById("btnAction");
        Provider.setKey("bigquery");
        this.persistActiveConnectionConfig();

        // Conexión reutilizada (p.ej. creada desde el add-in) sin proyecto
        // GCP "hogar" todavía: pedirlo aquí antes de comprobar el esquema.
        if (!BQ.getGcpProject()) {
            document.getElementById("gcpProjectPicker").classList.add("visible");
            await this.loadGcpProjects();
            this.updateActionButton();
            return;
        }

        if (btnAction) {
            btnAction.disabled = true;
            btnAction.innerText = "Comprobando esquema de control...";
        }
        try {
            const ready = await DracoSchema.controlSchemaExists();
            if (ready) {
                if (btnAction) btnAction.innerText = "Preparando tu espacio Draco...";
                await this.enterApp(btnAction);
            } else {
                this.showInstallPanel();
                if (btnAction) {
                    btnAction.disabled = false;
                    btnAction.innerText = "Entrar a Draco Planning →";
                }
            }
        } catch (err) {
            this.showAlert("Error al comprobar el esquema de control: " + err.message, true);
            if (btnAction) {
                btnAction.disabled = false;
                btnAction.innerText = "Reintentar";
            }
        }
    },

    async proceedSnowflake() {
        const btnAction = document.getElementById("btnAction");
        Provider.setKey("snowflake");
        this.persistActiveConnectionConfig();
        if (btnAction) {
            btnAction.disabled = true;
            btnAction.innerText = "Preparando tu espacio Draco...";
        }
        try {
            await this.enterApp(btnAction);
        } catch (e) {
            if (btnAction) {
                btnAction.disabled = false;
                btnAction.innerText = "Reintentar";
            }
        }
    },

    // ---------------------------------------------------------
    // Estado común
    // ---------------------------------------------------------
    logoutProvider(providerKey) {
        if (providerKey === "bigquery") {
            localStorage.removeItem("bigquery_access_token");
            localStorage.removeItem("bigquery_token_expires");
            localStorage.removeItem("draco_gcp_project");
            document.getElementById("gcpProjectPicker").classList.remove("visible");
        } else if (providerKey === "snowflake") {
            SF.logout();
        }
        this.setProviderState(providerKey, false);
        this.hideInstallPanel();
        this.updateActionButton();
        this.showAlert(`Sesión cerrada para ${this.getProviderDisplayName(providerKey)}`);
    },

    setProviderState(providerKey, isConnected) {
        this.connectedProviders[providerKey] = isConnected;
        const badge = document.getElementById(`badge-${providerKey}`);
        if (badge) {
            badge.innerText = isConnected ? "Conectado" : "Disponible";
            badge.classList.toggle("connected", isConnected);
        }
    },

    async enterApp(bootstrapMsgEl) {
        try {
            await DracoSchema.bootstrap((msg) => {
                if (bootstrapMsgEl) bootstrapMsgEl.innerText = msg;
            });
            window.location.href = "app.html";
        } catch (err) {
            this.showAlert("Error al inicializar Draco: " + err.message, true);
            throw err;
        }
    },

    updateActionButton() {
        const btnAction = document.getElementById("btnAction");
        if (!btnAction || !this.selectedProvider) return;

        const name = this.getProviderDisplayName(this.selectedProvider);
        btnAction.disabled = false;

        if (this.selectedProvider === "bigquery") {
            const isConnected = !!this.connectedProviders.bigquery;
            if (isConnected) {
                const hasProject = !!BQ.getGcpProject();
                btnAction.innerText = hasProject ? "Entrar a Draco Planning →" : "Selecciona un proyecto de GCP";
                btnAction.disabled = !hasProject;
                btnAction.classList.remove("disconnect-mode");
                btnAction.onclick = async () => {
                    if (!hasProject) return;
                    await this.proceedBigQuery();
                };
            } else {
                btnAction.innerText = `Conectar a ${name}`;
                btnAction.classList.remove("disconnect-mode");
                btnAction.onclick = () => this.connectBigQuery();
            }
            return;
        }

        if (this.selectedProvider === "snowflake") {
            const isConnected = !!this.connectedProviders.snowflake;
            if (isConnected) {
                btnAction.innerText = "Entrar a Draco Planning →";
                btnAction.classList.remove("disconnect-mode");
                btnAction.onclick = async () => { await this.proceedSnowflake(); };
            } else {
                btnAction.innerText = "Conectar a Snowflake";
                btnAction.classList.remove("disconnect-mode");
                btnAction.onclick = () => this.connectSnowflake();
            }
            return;
        }

        // Resto de conectores: todavía no implementados
        const isConnected = !!this.connectedProviders[this.selectedProvider];
        if (isConnected) {
            btnAction.innerText = `Cerrar sesión (${name})`;
            btnAction.classList.add("disconnect-mode");
            btnAction.onclick = () => this.logoutProvider(this.selectedProvider);
        } else {
            btnAction.innerText = `Conectar a ${name}`;
            btnAction.classList.remove("disconnect-mode");
            btnAction.onclick = () => {
                this.showAlert(`El conector para ${name} estará disponible próximamente.`);
            };
        }
    }
};

document.addEventListener("DOMContentLoaded", () => LoginApp.init());

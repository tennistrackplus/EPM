/**
 * Lógica de la landing page de Draco Planning (menú de conexiones,
 * igual en comportamiento a ADDIN/src/js/login.js):
 *
 *  - Vista "Lista": conexiones ya guardadas (Connections). Un clic
 *    reutiliza sus parámetros y continúa el flujo (conectar si hace
 *    falta, elegir proyecto GCP si falta, comprobar/instalar el
 *    esquema de control y entrar a la app).
 *  - Vista "Crear/Editar": nombre + selector de origen de datos +
 *    campos propios de cada proveedor, en un acordeón bajo la propia
 *    tarjeta del conector (BigQuery: proyecto de facturación opcional,
 *    seguido del selector de proyecto GCP y el panel de instalación
 *    del esquema una vez conectado; Snowflake: cuenta/warehouse/base
 *    de datos/rol).
 *
 * Las conexiones se guardan con el mismo almacén que usa el add-in de
 * Excel (js/connections.js, clave de localStorage "epm_connections"):
 * si Planning y el add-in se sirven desde el mismo origen (aunque sea
 * en rutas distintas, ej. https://tuapp.com/PLANNING/ y
 * https://tuapp.com/ADDIN/src/), el navegador comparte el mismo
 * localStorage y las conexiones creadas en uno aparecen
 * automáticamente en el otro.
 */
const LoginApp = {
    view: "list",               // "list" | "create"
    selectedProvider: null,     // proveedor elegido en la vista "crear/editar"
    editingConnectionId: null,  // != null si la vista "crear" está editando/reconectando en vez de creando
    connectMode: false,         // true si se abrió la vista para reconectar una conexión ya guardada
    connectedProviders: {},
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
        this.bindStaticEvents();
        this.checkExistingTokens();
        this.renderList();
        this.setupBrowserMessageListener();
    },

    // =====================================================================
    // Utilidades de UI
    // =====================================================================
    showAlert(msg, isError = false) {
        console.log("[Draco Planning]:", msg);
        const toastId = this.view === "create" ? "toastMessageCreate" : "toastMessageList";
        const toast = document.getElementById(toastId);
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

    switchView(view) {
        this.view = view;
        document.getElementById("viewList").classList.toggle("hidden", view !== "list");
        document.getElementById("viewCreate").classList.toggle("hidden", view !== "create");
        if (view === "list") {
            this.renderList();
        }
    },

    // =====================================================================
    // Eventos que no dependen de contenido generado dinámicamente
    // =====================================================================
    bindStaticEvents() {
        document.getElementById("btnNewConnection").addEventListener("click", () => {
            this.openConnectionView(null, "create");
        });

        document.getElementById("btnCancelCreate").addEventListener("click", () => {
            this.switchView("list");
        });

        // Selector de origen de datos dentro de la vista "crear/editar"
        document.querySelectorAll("#viewCreate .connector-card").forEach(card => {
            card.addEventListener("click", () => {
                const provider = card.getAttribute("data-provider");
                if (!Connections.isImplemented(provider)) {
                    this.showAlert(`El conector para ${this.getProviderDisplayName(provider)} estará disponible próximamente.`);
                    return;
                }
                document.querySelectorAll("#viewCreate .connector-card").forEach(c => c.classList.remove("selected"));
                card.classList.add("selected");
                this.selectedProvider = provider;
                this.connectMode = false;
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

        const btnInstall = document.getElementById("btnInstallSchema");
        if (btnInstall) {
            btnInstall.addEventListener("click", () => this.installControlSchema());
        }

        // Delegación de eventos sobre la lista de conexiones guardadas
        document.getElementById("savedConnectionsList").addEventListener("click", (e) => {
            const delBtn = e.target.closest(".btn-delete-conn");
            const editBtn = e.target.closest(".btn-edit-conn");
            const logoutBtn = e.target.closest(".btn-logout-conn");
            const card = e.target.closest(".conn-card");
            if (!card) return;
            const id = card.getAttribute("data-id");

            if (delBtn) {
                e.stopPropagation();
                this.handleDeleteClick(delBtn, id);
                return;
            }
            if (editBtn) {
                e.stopPropagation();
                this.openConnectionView(id, "edit");
                return;
            }
            if (logoutBtn) {
                e.stopPropagation();
                this.logoutConnectionCard(id);
                return;
            }
            this.openConnectionView(id, "connect");
        });
    },

    // =====================================================================
    // Vista LISTA
    // =====================================================================
    renderList() {
        const list = document.getElementById("savedConnectionsList");
        if (!list || !window.Connections) return;

        const connections = Connections.list();
        const activeId = Connections.getActiveId();

        if (connections.length === 0) {
            list.innerHTML = `
                <div class="empty-state">
                    <strong>Todavía no hay conexiones</strong>
                    <span>Crea la primera para empezar a modelar tus datos.</span>
                </div>`;
            return;
        }

        list.innerHTML = connections.map(conn => {
            const isActive = conn.id === activeId;
            const isConnectedNow = isActive && Provider.key() === conn.provider && Provider.isConnected();
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
                        ${isConnectedNow
                            ? `<button class="badge-status connected btn-logout-conn" title="Cerrar sesión" style="border:none;cursor:pointer;">Conectado</button>`
                            : `<span class="badge-status">Usar</span>`}
                        <button class="btn-icon btn-edit-conn" title="Editar conexión">✎</button>
                        <button class="btn-icon danger btn-delete-conn" title="Eliminar conexión">✕</button>
                    </div>
                </div>`;
        }).join("");
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

    /** Doble clic de seguridad para borrar: primer clic pide confirmación, segundo clic borra. */
    handleDeleteClick(btn, id) {
        if (btn.dataset.confirm === "1") {
            Connections.remove(id);
            this.showAlert("Conexión eliminada.");
            this.renderList();
            return;
        }
        btn.dataset.confirm = "1";
        btn.classList.add("confirm-danger");
        btn.title = "Pulsa de nuevo para confirmar";
        setTimeout(() => {
            btn.dataset.confirm = "";
            btn.classList.remove("confirm-danger");
            btn.title = "Eliminar conexión";
        }, 3000);
    },

    /** Cierra la sesión (token OAuth) de una conexión, sin borrar sus parámetros guardados */
    logoutConnectionCard(id) {
        const conn = Connections.getById(id);
        if (!conn) return;
        this.logoutProvider(conn.provider);
        this.renderList();
    },

    // =====================================================================
    // Vista CREAR / EDITAR / CONECTAR
    // =====================================================================
    /**
     * Abre la vista de conexión en uno de tres modos:
     *  - "create":  formulario en blanco para una conexión nueva.
     *  - "edit":    formulario relleno con los datos de una conexión guardada,
     *               para cambiarle el nombre o sus parámetros (✎ en la lista).
     *  - "connect": reutiliza una conexión guardada y continúa el flujo de
     *               conexión automáticamente (clic sobre la propia tarjeta).
     */
    openConnectionView(connectionId, mode) {
        this.editingConnectionId = connectionId;
        this.connectMode = mode === "connect";
        this.selectedProvider = null;

        document.getElementById("connName").value = "";
        document.getElementById("bqBillingProject").value = "";
        document.getElementById("sfAccount").value = "";
        document.getElementById("sfWarehouse").value = "";
        document.getElementById("sfDatabase").value = "";
        document.getElementById("sfRole").value = "";
        document.querySelectorAll("#viewCreate .connector-card").forEach(c => c.classList.remove("selected"));
        this.hideInstallPanel();

        const title = document.getElementById("createTitle");
        const subtitle = document.getElementById("createSubtitle");
        let conn = null;

        if (connectionId) {
            conn = Connections.getById(connectionId);
            if (!conn) return;
            this.selectedProvider = conn.provider;

            const cfg = conn.config || {};
            document.getElementById("connName").value = conn.name;
            if (conn.provider === "bigquery") {
                document.getElementById("bqBillingProject").value = cfg.billingProjectId || cfg.homeProjectId || "";
            } else if (conn.provider === "snowflake") {
                document.getElementById("sfAccount").value = cfg.account || "";
                document.getElementById("sfWarehouse").value = cfg.warehouse || "";
                document.getElementById("sfDatabase").value = cfg.database || "";
                document.getElementById("sfRole").value = cfg.role || "";
            }

            const card = document.querySelector(`#viewCreate .connector-card[data-provider="${conn.provider}"]`);
            if (card) card.classList.add("selected");

            if (mode === "edit") {
                title.innerText = "Editar conexión";
                subtitle.innerText = "Cambia el nombre o los parámetros de esta conexión.";
            } else {
                title.innerText = conn.name;
                subtitle.innerText = "Continuando con esta conexión guardada...";
            }
        } else {
            title.innerText = "Nueva conexión";
            subtitle.innerText = "Ponle un nombre y elige el origen de datos.";
        }

        this.togglePanels();
        this.updateActionButton();
        this.switchView("create");

        if (mode === "connect" && conn) {
            Connections.setActiveId(conn.id);
            Provider.setKey(conn.provider);
            this.applyConnectionConfig(conn);
            // Reutiliza el mismo botón de acción ya calculado por updateActionButton()
            // (que ya sabe si hay que conectar, elegir proyecto o entrar directamente).
            const btnAction = document.getElementById("btnAction");
            if (btnAction && typeof btnAction.onclick === "function" && !btnAction.disabled) {
                btnAction.onclick();
            }
        }
    },

    togglePanels() {
        document.getElementById("bigqueryConfigPanel").classList.toggle("visible", this.selectedProvider === "bigquery");
        document.getElementById("snowflakeConfigPanel").classList.toggle("visible", this.selectedProvider === "snowflake");

        if (this.selectedProvider === "bigquery") {
            const showPicker = BQ.isConnected();
            document.getElementById("gcpProjectPicker").classList.toggle("visible", showPicker);
            if (showPicker) this.maybeLoadGcpProjects();
        } else {
            document.getElementById("gcpProjectPicker").classList.remove("visible");
        }

        if (this.selectedProvider === "snowflake") {
            document.getElementById("sfAccount").value = document.getElementById("sfAccount").value || SF.getAccount();
            document.getElementById("sfWarehouse").value = document.getElementById("sfWarehouse").value || SF.getWarehouse();
            document.getElementById("sfDatabase").value = document.getElementById("sfDatabase").value || SF.getDatabase();
            document.getElementById("sfRole").value = document.getElementById("sfRole").value || SF.getRole();
        }
    },

    maybeLoadGcpProjects() {
        const select = document.getElementById("gcpProjectSelect");
        if (select && !select.options.length) {
            this.loadGcpProjects();
        }
    },

    /** Carga en BQ/SF los parámetros guardados de una conexión (sin tocar los tokens de sesión) */
    applyConnectionConfig(conn) {
        const cfg = conn.config || {};
        if (conn.provider === "bigquery") {
            // homeProjectId: lo usa Planning (dataset DRACO_CONTROL / DRACO_<proyecto>).
            // billingProjectId: mismo campo que usa el add-in de Excel; si la conexión
            // se creó allí y todavía no se ha elegido un proyecto hogar en Planning,
            // se usa como punto de partida razonable (normalmente coinciden).
            BQ.setGcpProject(cfg.homeProjectId || cfg.billingProjectId || "");
        } else if (conn.provider === "snowflake") {
            SF.setAccount(cfg.account || "");
            SF.setWarehouse(cfg.warehouse || "");
            SF.setDatabase(cfg.database || DracoConfig.snowflakeDatabase);
            SF.setRole(cfg.role || "");
        }
    },

    /** Construye el config a partir de los campos visibles, validando lo mínimo por proveedor */
    collectConfigFromForm() {
        if (this.selectedProvider === "bigquery") {
            const billingProjectId = document.getElementById("bqBillingProject").value.trim();
            return { billingProjectId, homeProjectId: BQ.getGcpProject() || billingProjectId || "" };
        }
        if (this.selectedProvider === "snowflake") {
            return {
                account: document.getElementById("sfAccount").value.trim(),
                warehouse: document.getElementById("sfWarehouse").value.trim(),
                database: document.getElementById("sfDatabase").value.trim() || DracoConfig.snowflakeDatabase,
                role: document.getElementById("sfRole").value.trim()
            };
        }
        return {};
    },

    validateConfig(provider, config) {
        if (provider === "snowflake") {
            if (!config.account || !config.warehouse) {
                return "Indica al menos la cuenta y el warehouse de Snowflake.";
            }
        }
        return null;
    },

    /** Crea o actualiza (según editingConnectionId) la conexión con lo que hay en el formulario */
    saveConnectionFromForm() {
        if (!this.selectedProvider) return null;

        const name = document.getElementById("connName").value.trim();
        const config = this.collectConfigFromForm();
        const validationError = this.validateConfig(this.selectedProvider, config);
        if (validationError) {
            this.showAlert(validationError, true);
            return null;
        }

        let conn;
        if (this.editingConnectionId) {
            conn = Connections.update(this.editingConnectionId, {
                name: name || undefined,
                provider: this.selectedProvider,
                config
            });
        } else {
            conn = Connections.create({ name, provider: this.selectedProvider, config });
        }
        this.editingConnectionId = conn.id;
        Connections.setActiveId(conn.id);
        Provider.setKey(conn.provider);
        return conn;
    },

    /** Guarda (merge) solo la config de la conexión activa, sin tocar su nombre */
    persistActiveConnectionConfig() {
        if (!window.Connections || !this.selectedProvider || !this.editingConnectionId) return;
        const config = this.collectConfigFromForm();
        Connections.update(this.editingConnectionId, { config });
    },

    // =====================================================================
    // Mensajería entre ventanas (popups de OAuth)
    // =====================================================================
    setupBrowserMessageListener() {
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
            this.connectedProviders.bigquery = true;
        } else {
            localStorage.removeItem("bigquery_access_token");
            localStorage.removeItem("bigquery_token_expires");
        }

        if (SF.isConnected()) {
            this.connectedProviders.snowflake = true;
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
                    this.connectedProviders.bigquery = true;
                    this.persistActiveConnectionConfig();
                    this.showAlert("¡Conexión con Google BigQuery establecida con éxito!");
                    this.togglePanels();
                    this.updateActionButton();
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
                    this.connectedProviders.snowflake = true;
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
            const picker = document.getElementById("gcpProjectPicker");
            if (picker) picker.classList.remove("visible");
        } else if (providerKey === "snowflake") {
            SF.logout();
        }
        this.connectedProviders[providerKey] = false;
        this.hideInstallPanel();
        if (this.view === "create") this.updateActionButton();
        this.showAlert(`Sesión cerrada para ${this.getProviderDisplayName(providerKey)}`);
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
        btnAction.classList.remove("disconnect-mode");

        if (this.selectedProvider === "bigquery") {
            const isConnected = !!this.connectedProviders.bigquery;
            if (isConnected) {
                const hasProject = !!BQ.getGcpProject();
                btnAction.innerText = hasProject ? "Entrar a Draco Planning →" : "Selecciona un proyecto de GCP";
                btnAction.disabled = !hasProject;
                btnAction.onclick = async () => {
                    if (!hasProject) return;
                    this.saveConnectionFromForm();
                    await this.proceedBigQuery();
                };
            } else {
                btnAction.innerText = `Conectar a ${name}`;
                btnAction.onclick = () => {
                    const conn = this.saveConnectionFromForm();
                    if (!conn) return;
                    this.applyConnectionConfig(conn);
                    this.connectBigQuery();
                };
            }
            return;
        }

        if (this.selectedProvider === "snowflake") {
            const isConnected = !!this.connectedProviders.snowflake;
            if (isConnected) {
                btnAction.innerText = "Entrar a Draco Planning →";
                btnAction.onclick = async () => {
                    this.saveConnectionFromForm();
                    await this.proceedSnowflake();
                };
            } else {
                btnAction.innerText = `Conectar a ${name}`;
                btnAction.onclick = () => {
                    const conn = this.saveConnectionFromForm();
                    if (!conn) return;
                    this.connectSnowflake();
                };
            }
            return;
        }

        // Resto de conectores: todavía no implementados
        btnAction.innerText = `El conector para ${name} estará disponible próximamente`;
        btnAction.disabled = true;
    }
};

document.addEventListener("DOMContentLoaded", () => LoginApp.init());

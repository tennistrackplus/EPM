/**
 * Lógica del Task Pane de Conexiones.
 *
 * Nuevo flujo (menú de conexiones, como Power BI / Looker):
 *  - Vista "Lista": conexiones ya guardadas (Connections). Un clic
 *    reutiliza sus parámetros y conecta sin volver a preguntarlos.
 *  - Vista "Crear": nombre + selector de origen de datos + campos
 *    propios de cada proveedor (BigQuery: proyecto de facturación y
 *    repositorio de modelos semánticos; Snowflake: cuenta/warehouse/
 *    base de datos/rol). Al guardar, se crea la conexión y se lanza
 *    el login del proveedor correspondiente.
 *
 * Soporta BigQuery (OAuth implícito) y Snowflake (OAuth PKCE) a
 * través de la capa Provider; el resto del add-in solo necesita
 * saber cuál está activo (Provider.key()).
 */
const LoginApp = {
    view: "list",              // "list" | "create"
    selectedProvider: null,    // proveedor elegido en la vista "crear"
    editingConnectionId: null, // != null si la vista "crear" está editando en vez de creando
    authDialog: null,
    pendingConnectionId: null, // conexión para la que se está esperando respuesta OAuth

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
        this.renderList();
        this.setupBrowserMessageListener();
    },

    // =====================================================================
    // Utilidades de UI
    // =====================================================================
    showAlert(msg, isError = false) {
        console.log("[EPM Add-in]:", msg);
        const toastId = this.view === "create" ? "toastMessageCreate" : "toastMessageList";
        const toast = document.getElementById(toastId);
        if (toast) {
            toast.innerText = msg;
            toast.className = `toast-message visible ${isError ? "error" : "info"}`;
            setTimeout(() => toast.classList.remove("visible"), 6000);
        }
    },

    switchView(view) {
        this.view = view;
        document.getElementById("viewList").classList.toggle("hidden", view !== "list");
        document.getElementById("viewCreate").classList.toggle("hidden", view !== "create");
        if (view === "list") {
            this.renderList();
        }
    },

    getProviderDisplayName(providerKey) {
        return Connections.labelFor(providerKey);
    },

    // =====================================================================
    // Eventos que no dependen de contenido generado dinámicamente
    // =====================================================================
    bindStaticEvents() {
        document.getElementById("btnNewConnection").addEventListener("click", () => {
            this.openCreateView(null);
        });

        document.getElementById("btnCancelCreate").addEventListener("click", () => {
            this.switchView("list");
        });

        document.getElementById("btnSaveConnection").addEventListener("click", () => {
            this.saveConnection();
        });

        // Selector de origen de datos dentro de la vista "crear"
        document.querySelectorAll("#viewCreate .connector-card").forEach(card => {
            card.addEventListener("click", () => {
                document.querySelectorAll("#viewCreate .connector-card").forEach(c => c.classList.remove("selected"));
                card.classList.add("selected");
                this.selectedProvider = card.getAttribute("data-provider");
                this.toggleProviderPanels();
                this.updateSaveButton();
            });
        });

        // Delegación de eventos sobre la lista de conexiones (se regenera dinámicamente)
        document.getElementById("connectionsContainer").addEventListener("click", (e) => {
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
                this.openCreateView(id);
                return;
            }
            if (logoutBtn) {
                e.stopPropagation();
                this.logoutConnection(id);
                return;
            }
            this.connectExisting(id);
        });
    },

    // =====================================================================
    // Vista LISTA
    // =====================================================================
    renderList() {
        const container = document.getElementById("connectionsContainer");
        const connections = Connections.list();
        const activeId = Connections.getActiveId();

        if (connections.length === 0) {
            container.innerHTML = `
                <div class="empty-state">
                    <strong>Todavía no hay conexiones</strong>
                    <span>Crea la primera para empezar a traer datos.</span>
                </div>`;
            return;
        }

        container.innerHTML = connections.map(conn => {
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
                            : `<span class="badge-status">Conectar</span>`}
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
            return cfg.billingProjectId ? `${label} · Facturación: ${cfg.billingProjectId}` : label;
        }
        if (conn.provider === "snowflake") {
            return cfg.account ? `${label} · ${cfg.account}` : label;
        }
        return label;
    },

    escapeHtml(str) {
        const div = document.createElement("div");
        div.innerText = str == null ? "" : String(str);
        return div.innerHTML;
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

    /** Reutiliza los parámetros ya guardados de una conexión y conecta sin volver a preguntarlos. */
    async connectExisting(id) {
        const conn = Connections.getById(id);
        if (!conn) return;

        Connections.setActiveId(conn.id);
        Provider.setKey(conn.provider);
        this.applyConnectionConfig(conn);

        if (!Connections.isImplemented(conn.provider)) {
            this.showAlert(`El conector para ${this.getProviderDisplayName(conn.provider)} estará disponible próximamente.`);
            this.renderList();
            return;
        }

        if (Provider.isConnected()) {
            this.showAlert(`Ya tienes una sesión activa de ${this.getProviderDisplayName(conn.provider)}.`);
            this.renderList();
            return;
        }

        this.pendingConnectionId = conn.id;
        try {
            if (conn.provider === "bigquery") {
                this.startBigQueryAuth();
            } else if (conn.provider === "snowflake") {
                await this.startSnowflakeAuth();
            }
        } catch (err) {
            console.error("Error al conectar:", err);
            this.showAlert("Error al conectar: " + err.message, true);
        }
    },

    /** Cierra la sesión (token OAuth) de una conexión, sin borrar sus parámetros guardados */
    logoutConnection(id) {
        const conn = Connections.getById(id);
        if (!conn) return;

        if (conn.provider === "snowflake") {
            SF.logout();
        } else {
            BQ.logout();
        }
        this.showAlert(`Sesión cerrada para ${this.getProviderDisplayName(conn.provider)}.`);
        this.renderList();
    },

    /** Carga en BQ/SF los parámetros guardados de una conexión (sin tocar los tokens de sesión) */
    applyConnectionConfig(conn) {
        const cfg = conn.config || {};
        if (conn.provider === "bigquery") {
            BQ.setBillingProject(cfg.billingProjectId || "");
            BQ.setSemanticRepo(cfg.semanticRepo || { type: "", url: "" });
        } else if (conn.provider === "snowflake") {
            SF.setAccount(cfg.account || "");
            SF.setWarehouse(cfg.warehouse || "");
            SF.setDatabase(cfg.database || DracoConfig.snowflakeDatabase);
            SF.setRole(cfg.role || "");
        }
    },

    // =====================================================================
    // Vista CREAR / EDITAR
    // =====================================================================
    openCreateView(connectionId) {
        this.editingConnectionId = connectionId;
        this.selectedProvider = null;

        document.getElementById("connName").value = "";
        document.getElementById("bqBillingProject").value = "";
        document.getElementById("bqRepoType").value = "";
        document.getElementById("bqRepoUrl").value = "";
        document.getElementById("bqRepoBranch").value = "";
        document.getElementById("bqRepoToken").value = "";
        document.getElementById("sfAccount").value = "";
        document.getElementById("sfWarehouse").value = "";
        document.getElementById("sfDatabase").value = "";
        document.getElementById("sfRole").value = "";
        document.querySelectorAll("#viewCreate .connector-card").forEach(c => c.classList.remove("selected"));

        const title = document.getElementById("createTitle");
        const saveBtn = document.getElementById("btnSaveConnection");

        if (connectionId) {
            const conn = Connections.getById(connectionId);
            if (!conn) return;
            title.innerText = "Editar conexión";
            document.getElementById("connName").value = conn.name;
            this.selectedProvider = conn.provider;

            const cfg = conn.config || {};
            if (conn.provider === "bigquery") {
                document.getElementById("bqBillingProject").value = cfg.billingProjectId || "";
                document.getElementById("bqRepoType").value = (cfg.semanticRepo && cfg.semanticRepo.type) || "";
                document.getElementById("bqRepoUrl").value = (cfg.semanticRepo && cfg.semanticRepo.url) || "";
                document.getElementById("bqRepoBranch").value = (cfg.semanticRepo && cfg.semanticRepo.branch) || "";
                document.getElementById("bqRepoToken").value = (cfg.semanticRepo && cfg.semanticRepo.token) || "";
            } else if (conn.provider === "snowflake") {
                document.getElementById("sfAccount").value = cfg.account || "";
                document.getElementById("sfWarehouse").value = cfg.warehouse || "";
                document.getElementById("sfDatabase").value = cfg.database || "";
                document.getElementById("sfRole").value = cfg.role || "";
            }

            const card = document.querySelector(`#viewCreate .connector-card[data-provider="${conn.provider}"]`);
            if (card) card.classList.add("selected");
        } else {
            title.innerText = "Nueva conexión";
        }

        this.toggleProviderPanels();
        this.updateSaveButton();
        this.switchView("create");
    },

    toggleProviderPanels() {
        document.getElementById("bigqueryConfigPanel").classList.toggle("visible", this.selectedProvider === "bigquery");
        document.getElementById("snowflakeConfigPanel").classList.toggle("visible", this.selectedProvider === "snowflake");
    },

    updateSaveButton() {
        const btn = document.getElementById("btnSaveConnection");
        if (!this.selectedProvider) {
            btn.disabled = true;
            btn.innerText = "Selecciona un origen de datos";
            return;
        }
        btn.disabled = false;
        const name = this.getProviderDisplayName(this.selectedProvider);
        btn.innerText = this.editingConnectionId ? `Guardar cambios (${name})` : `Crear y conectar (${name})`;
    },

    /** Construye el objeto config a partir de los campos visibles, validando lo mínimo por proveedor */
    collectConfigFromForm() {
        if (this.selectedProvider === "bigquery") {
            return {
                billingProjectId: document.getElementById("bqBillingProject").value.trim(),
                semanticRepo: {
                    type: document.getElementById("bqRepoType").value,
                    url: document.getElementById("bqRepoUrl").value.trim(),
                    branch: document.getElementById("bqRepoBranch").value.trim(),
                    token: document.getElementById("bqRepoToken").value.trim()
                }
            };
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

    async saveConnection() {
        if (!this.selectedProvider) return;

        const name = document.getElementById("connName").value.trim();
        const config = this.collectConfigFromForm();
        const validationError = this.validateConfig(this.selectedProvider, config);
        if (validationError) {
            this.showAlert(validationError, true);
            return;
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

        Connections.setActiveId(conn.id);
        Provider.setKey(conn.provider);
        this.applyConnectionConfig(conn);

        if (!Connections.isImplemented(conn.provider)) {
            this.showAlert(`Conexión guardada. El conector para ${this.getProviderDisplayName(conn.provider)} estará disponible próximamente.`);
            this.switchView("list");
            return;
        }

        this.pendingConnectionId = conn.id;
        try {
            if (conn.provider === "bigquery") {
                this.startBigQueryAuth();
            } else if (conn.provider === "snowflake") {
                await this.startSnowflakeAuth();
            }
        } catch (err) {
            console.error("Error al iniciar la conexión:", err);
            this.showAlert("Error al iniciar la conexión: " + err.message, true);
        }
    },

    // =====================================================================
    // OAuth: apertura de ventana de login (Office dialog o popup de navegador)
    // =====================================================================
    openAuthWindow(authUrl) {
        const isOfficeEnvironment = typeof Office !== "undefined" &&
            Office.context &&
            Office.context.ui &&
            typeof Office.context.ui.displayDialogAsync === "function";

        if (isOfficeEnvironment) {
            const dialogOptions = { height: 60, width: 40, displayInIframe: false };

            Office.context.ui.displayDialogAsync(authUrl, dialogOptions, (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.error("Error al abrir diálogo de autenticación:", asyncResult.error.message);
                    this.showAlert("No se pudo abrir la ventana de login. Código: " + asyncResult.error.code + " - " + asyncResult.error.message, true);
                    return;
                }

                this.authDialog = asyncResult.value;
                this.authDialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
                    this.handleAuthResponse(arg.message);
                });
            });
        } else {
            const width = 500, height = 650;
            const left = (window.screen.width / 2) - (width / 2);
            const top = (window.screen.height / 2) - (height / 2);
            window.open(authUrl, "AuthWindow", `width=${width},height=${height},top=${top},left=${left}`);
        }
    },

    /** Inicia el flujo OAuth 2.0 (implícito) para BigQuery */
    startBigQueryAuth() {
        if (window.location.protocol === "file:") {
            console.error("No se puede ejecutar desde file://. Debes usar un servidor web local (ej. Live Server o http://localhost).");
            this.showAlert("Error: no se puede ejecutar desde un archivo local (file://). Usa Live Server o un servidor HTTP.", true);
            return;
        }

        const redirectUri = new URL("auth-callback.html", window.location.href).href;
        const authUrl = "https://accounts.google.com/o/oauth2/v2/auth?" +
            `client_id=${encodeURIComponent(DracoConfig.googleClientId)}` +
            `&redirect_uri=${encodeURIComponent(redirectUri)}` +
            "&response_type=token" +
            `&scope=${encodeURIComponent(DracoConfig.googleScopes)}` +
            "&prompt=consent";

        this.openAuthWindow(authUrl);
    },

    /** Inicia el flujo OAuth 2.0 (Authorization Code + PKCE) para Snowflake, usando lo ya guardado en SF */
    async startSnowflakeAuth() {
        const authUrl = await SF.buildAuthUrl();
        this.openAuthWindow(authUrl);
    },

    /**
     * Listener para recibir mensajes cuando se ejecuta fuera del entorno de Office (navegador estándar)
     */
    setupBrowserMessageListener() {
        window.addEventListener("message", (event) => {
            if (typeof event.data === "string") {
                try {
                    const data = JSON.parse(event.data);
                    if (data.status) {
                        this.handleAuthResponse(event.data);
                    }
                } catch (e) {
                    // Ignorar mensajes de otros orígenes
                }
            }
        });
    },

    /**
     * Maneja la respuesta recibida tras el login, de cualquier proveedor
     */
    async handleAuthResponse(messageString) {
        try {
            const response = JSON.parse(messageString);

            if (response.provider === "bigquery") {
                if (response.status === "success") {
                    // BQ.setToken guarda el token en localStorage Y en
                    // Office.context.document.settings, para que también lo
                    // vea el runtime aislado de los botones del ribbon
                    // (commands.html) cuando se pulse "Abrir bucket" /
                    // "Guardar en bucket" (ver bigquery.js).
                    const expiresAt = Date.now() + (parseInt(response.expiresIn, 10) * 1000);
                    if (window.BQ && typeof BQ.setToken === "function") {
                        BQ.setToken(response.token, expiresAt);
                    } else {
                        localStorage.setItem("bigquery_access_token", response.token);
                        localStorage.setItem("bigquery_token_expires", expiresAt);
                    }
                    Provider.setKey("bigquery");
                    this.showAlert("¡Conexión con Google BigQuery establecida con éxito!");
                    this.switchView("list");
                } else {
                    console.error("Error en autenticación:", response.error);
                    this.showAlert("Error de autenticación: " + (response.error || "Desconocido"), true);
                }
            } else if (response.provider === "snowflake") {
                if (response.status === "success" && response.code) {
                    await SF.handleAuthCode(response.code, response.state);
                    Provider.setKey("snowflake");
                    this.showAlert("¡Conexión con Snowflake establecida con éxito!");
                    this.switchView("list");
                } else {
                    console.error("Error en autenticación:", response.error);
                    this.showAlert("Error de autenticación con Snowflake: " + (response.error || "Desconocido"), true);
                }
            }
        } catch (err) {
            console.error("Error leyendo respuesta de autenticación:", err);
            this.showAlert("Error al procesar la respuesta de login: " + err.message, true);
        } finally {
            this.pendingConnectionId = null;
            if (this.authDialog) {
                this.authDialog.close();
                this.authDialog = null;
            }
        }
    }
};

// Inicialización garantizada tras la definición del objeto LoginApp
if (typeof Office !== "undefined") {
    Office.onReady(() => {
        LoginApp.init();
    });
} else {
    document.addEventListener("DOMContentLoaded", () => {
        LoginApp.init();
    });
}

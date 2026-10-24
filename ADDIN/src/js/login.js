/**
 * Lógica del Task Pane de Conexiones.
 *
 * Nuevo flujo (menú de conexiones, como Looker):
 *  - Vista "Lista": conexiones ya guardadas (Connections). Un clic
 *    reutiliza sus parámetros y conecta sin volver a preguntarlos.
 *  - Vista "Crear": nombre + origen de datos + campos de BigQuery
 *    (proyecto de facturación y repositorio de modelos semánticos).
 *    Al guardar, se crea la conexión y se lanza el login.
 *
 * Proveedor soportado: BigQuery, a través de la capa Provider.
 */
const LoginApp = {
    view: "list",              // "list" | "create"
    selectedProvider: null,    // proveedor elegido en la vista "crear"
    editingConnectionId: null, // != null si la vista "crear" está editando en vez de creando
    authDialog: null,
    pendingConnectionId: null, // conexión para la que se está esperando respuesta OAuth

    PROVIDER_ICONS: {
        bigquery: "assets/datasource-db.svg"
    },

    init() {
        // Modo backend: reanuda la renovación automática del token si hay sesión.
        GoogleAuth.resume().finally(() => this.renderList());
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

        document.getElementById("bqAuthMode").addEventListener("change", () => this.toggleAuthModeFields());

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
                const cfgError = GoogleAuth.validateConfig(conn.config);
                if (cfgError) {
                    this.showAlert(cfgError + " Edita la conexión (✎).", true);
                    return;
                }
                this.startBigQueryAuth(conn);
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

        GoogleAuth.logout();
        this.showAlert(`Sesión cerrada para ${this.getProviderDisplayName(conn.provider)}.`);
        this.renderList();
    },

    /** Carga en BQ los parámetros guardados de una conexión (sin tocar los tokens de sesión) */
    applyConnectionConfig(conn) {
        const cfg = conn.config || {};
        if (conn.provider === "bigquery") {
            BQ.setBillingProject(cfg.billingProjectId || "");
            BQ.setSemanticRepo(cfg.semanticRepo || { type: "", url: "" });
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
        document.getElementById("bqAuthMode").value = "token";
        document.getElementById("bqOAuthClientId").value = "";
        document.getElementById("bqAuthBrokerUrl").value = "";
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
                const auth = GoogleAuth.normalizeConfig(cfg);
                document.getElementById("bqAuthMode").value = auth.authMode;
                document.getElementById("bqOAuthClientId").value = auth.oauthClientId;
                document.getElementById("bqAuthBrokerUrl").value = auth.authBrokerUrl;
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

    toggleAuthModeFields() {
        const backend = document.getElementById("bqAuthMode").value === "backend";
        document.getElementById("bqAuthTokenFields").style.display = backend ? "none" : "";
        document.getElementById("bqAuthBackendFields").style.display = backend ? "" : "none";
    },

    toggleProviderPanels() {
        document.getElementById("bigqueryConfigPanel").classList.toggle("visible", this.selectedProvider === "bigquery");
        this.toggleAuthModeFields();
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
                authMode: document.getElementById("bqAuthMode").value,
                oauthClientId: document.getElementById("bqOAuthClientId").value.trim(),
                authBrokerUrl: document.getElementById("bqAuthBrokerUrl").value.trim().replace(/\/+$/, ""),
                billingProjectId: document.getElementById("bqBillingProject").value.trim(),
                semanticRepo: {
                    type: document.getElementById("bqRepoType").value,
                    url: document.getElementById("bqRepoUrl").value.trim(),
                    branch: document.getElementById("bqRepoBranch").value.trim(),
                    token: document.getElementById("bqRepoToken").value.trim()
                }
            };
        }
        return {};
    },

    validateConfig(provider, config) {
        if (provider === "bigquery") return GoogleAuth.validateConfig(config);
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
                this.startBigQueryAuth(conn);
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

    /** Inicia sesión con Google según el modo de la conexión (ver js/google-auth.js) */
    startBigQueryAuth(conn) {
        if (window.location.protocol === "file:") {
            console.error("No se puede ejecutar desde file://. Debes usar un servidor web local (ej. Live Server o http://localhost).");
            this.showAlert("Error: no se puede ejecutar desde un archivo local (file://). Usa Live Server o un servidor HTTP.", true);
            return;
        }

        const redirectUri = new URL("auth-callback.html", window.location.href).href;
        const authUrl = GoogleAuth.buildLoginUrl((conn && conn.config) || {}, redirectUri);
        this.openAuthWindow(authUrl);
    },

    /**
     * Listener para recibir mensajes cuando se ejecuta fuera del entorno de Office (navegador estándar)
     */
    setupBrowserMessageListener() {
        window.addEventListener("message", (event) => {
            // SEGURIDAD: se comprueba el origen antes de procesar el
            // mensaje (antes se aceptaba cualquier origen). auth-callback.html
            // responde con window.location.origin en vez de "*", así que aquí debe
            // coincidir con el propio origen de este add-in.
            if (event.origin !== window.location.origin) return;
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
                // GoogleAuth comprueba el `state` y guarda el token (sessionStorage).
                // Con Shared Runtime (ver manifest), commands.js corre en el
                // mismo proceso que este panel y lo comparte.
                if (GoogleAuth.acceptResponse(response)) {
                    Provider.setKey("bigquery");
                    if (typeof TaskPaneApp !== "undefined" && typeof TaskPaneApp.applyPlanningGateUI === "function") {
                        TaskPaneApp.applyPlanningGateUI();
                    } else if (window.DracoPlanningGate) {
                        DracoPlanningGate.syncRibbon();
                    }
                    if (window.DracoPlanningGate && !DracoPlanningGate.canWrite()) {
                        this.showAlert("Conectado a BigQuery con permiso de solo lectura: la planificación no estará disponible.");
                    } else {
                        this.showAlert("¡Conexión con Google BigQuery establecida con éxito!");
                    }
                    this.switchView("list");
                } else {
                    console.error("Error en autenticación:", response.error);
                    this.showAlert("Error de autenticación: " + (response.error || "Desconocido"), true);
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

/**
 * Lógica del Task Pane de Login, Verificación de Sesión y Desconexión
 * Soporta BigQuery (OAuth implícito) y Snowflake (OAuth PKCE) a
 * través de la capa Provider; el resto del add-in solo necesita
 * saber cuál está activo (Provider.key()).
 */
const LoginApp = {
    selectedProvider: null,
    authDialog: null,
    connectedProviders: {},

    init() {
        this.bindEvents();
        this.checkExistingTokens();
        this.setupBrowserMessageListener();
    },

    /**
     * Muestra notificaciones visuales en el panel sin depender de window.alert (bloqueado por Office.js)
     */
    showAlert(msg, isError = false) {
        console.log("[EPM Add-in]:", msg);
        const toast = document.getElementById("toastMessage");
        if (toast) {
            toast.innerText = msg;
            toast.className = `toast-message visible ${isError ? "error" : "info"}`;

            setTimeout(() => {
                toast.classList.remove("visible");
            }, 6000);
        }
    },

    bindEvents() {
        const cards = document.querySelectorAll(".connector-card");
        const btnAction = document.getElementById("btnAction");

        cards.forEach(card => {
            card.addEventListener("click", () => {
                cards.forEach(c => c.classList.remove("selected"));
                card.classList.add("selected");

                this.selectedProvider = card.getAttribute("data-provider");
                this.togglePanels();
                this.updateActionButton();
            });
        });

        if (btnAction) {
            btnAction.addEventListener("click", () => {
                if (!this.selectedProvider) return;

                const isConnected = !!this.connectedProviders[this.selectedProvider];

                if (isConnected) {
                    this.logoutProvider(this.selectedProvider);
                } else if (this.selectedProvider === "bigquery") {
                    this.connectBigQuery();
                } else if (this.selectedProvider === "snowflake") {
                    this.connectSnowflake();
                } else {
                    this.showAlert(`El conector para ${this.getProviderDisplayName(this.selectedProvider)} estará disponible próximamente.`);
                }
            });
        }
    },

    /** Muestra/oculta el panel de configuración de Snowflake según el conector seleccionado */
    togglePanels() {
        const panel = document.getElementById("snowflakeConfigPanel");
        if (!panel) return;

        panel.classList.toggle("visible", this.selectedProvider === "snowflake");

        if (this.selectedProvider === "snowflake") {
            document.getElementById("sfAccount").value = SF.getAccount();
            document.getElementById("sfWarehouse").value = SF.getWarehouse();
            document.getElementById("sfDatabase").value = SF.getDatabase();
            document.getElementById("sfRole").value = SF.getRole();
        }
    },

    /**
     * Listener para recibir mensajes cuando se ejecuta fuera del entorno de Office (Navegador estándar)
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
     * Revisa al cargar el panel si existen tokens activos y actualiza el estado
     */
    checkExistingTokens() {
        if (BQ.isConnected()) {
            this.setProviderState("bigquery", true);
        } else {
            BQ.logout();
            this.setProviderState("bigquery", false);
        }

        if (SF.isConnected()) {
            this.setProviderState("snowflake", true);
        } else {
            SF.logout();
            this.setProviderState("snowflake", false);
        }
    },

    /**
     * Abre la URL de autenticación indicada en un diálogo de Office (o en una
     * ventana emergente si se ejecuta fuera de Excel, p.ej. para pruebas en navegador)
     */
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

    /**
     * Inicia el flujo OAuth 2.0 (implícito) para BigQuery
     */
    connectBigQuery() {
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

    /**
     * Inicia el flujo OAuth 2.0 (Authorization Code + PKCE) para Snowflake
     */
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
            const authUrl = await SF.buildAuthUrl();
            this.openAuthWindow(authUrl);
        } catch (err) {
            console.error("Error al iniciar la conexión con Snowflake:", err);
            this.showAlert("Error al iniciar la conexión con Snowflake: " + err.message, true);
        }
    },

    /**
     * Maneja la respuesta recibida tras el login, de cualquier proveedor
     */
    async handleAuthResponse(messageString) {
        try {
            const response = JSON.parse(messageString);

            if (response.provider === "bigquery") {
                if (response.status === "success") {
                    localStorage.setItem("bigquery_access_token", response.token);
                    localStorage.setItem("bigquery_token_expires", Date.now() + (parseInt(response.expiresIn, 10) * 1000));
                    Provider.setKey("bigquery");
                    this.setProviderState("bigquery", true);
                    this.updateActionButton();
                    this.showAlert("¡Conexión con Google BigQuery establecida con éxito!");
                } else {
                    console.error("Error en autenticación:", response.error);
                    this.showAlert("Error de autenticación: " + (response.error || "Desconocido"), true);
                }
            } else if (response.provider === "snowflake") {
                if (response.status === "success" && response.code) {
                    await SF.handleAuthCode(response.code, response.state);
                    Provider.setKey("snowflake");
                    this.setProviderState("snowflake", true);
                    this.updateActionButton();
                    this.showAlert("¡Conexión con Snowflake establecida con éxito!");
                } else {
                    console.error("Error en autenticación:", response.error);
                    this.showAlert("Error de autenticación con Snowflake: " + (response.error || "Desconocido"), true);
                }
            }
        } catch (err) {
            console.error("Error leyendo respuesta de autenticación:", err);
            this.showAlert("Error al procesar la respuesta de login: " + err.message, true);
        } finally {
            if (this.authDialog) {
                this.authDialog.close();
                this.authDialog = null;
            }
        }
    },

    /**
     * Cierra la sesión del proveedor indicado
     */
    logoutProvider(providerKey) {
        if (providerKey === "bigquery") {
            BQ.logout();
        } else if (providerKey === "snowflake") {
            SF.logout();
        }

        this.setProviderState(providerKey, false);
        this.updateActionButton();
        this.showAlert(`Sesión cerrada para ${this.getProviderDisplayName(providerKey)}`);
    },

    /**
     * Actualiza el estado interno y la insignia en la interfaz
     */
    setProviderState(providerKey, isConnected) {
        this.connectedProviders[providerKey] = isConnected;
        const badge = document.getElementById(`badge-${providerKey}`);

        if (badge) {
            if (isConnected) {
                badge.innerText = "Conectado";
                badge.classList.add("connected");
            } else {
                badge.innerText = "Disponible";
                badge.classList.remove("connected");
            }
        }
    },

    /**
     * Actualiza el botón principal de acción según el proveedor seleccionado y su estado de conexión
     */
    updateActionButton() {
        const btnAction = document.getElementById("btnAction");
        if (!btnAction || !this.selectedProvider) return;

        btnAction.disabled = false;
        const isConnected = !!this.connectedProviders[this.selectedProvider];
        const name = this.getProviderDisplayName(this.selectedProvider);

        if (isConnected) {
            btnAction.innerText = `Cerrar Sesión (${name})`;
            btnAction.classList.add("disconnect-mode");
        } else {
            btnAction.innerText = `Conectar a ${name}`;
            btnAction.classList.remove("disconnect-mode");
        }
    },

    getProviderDisplayName(providerKey) {
        const names = {
            bigquery: "BigQuery",
            amazon: "Amazon",
            fabric: "Microsoft Fabric",
            snowflake: "Snowflake",
            datasphere: "SAP Datasphere",
            s4cds: "CDS de S4"
        };
        return names[providerKey] || providerKey;
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

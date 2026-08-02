/**
 * Lógica del Task Pane de Login, Verificación de Sesión y Desconexión
 */
const LoginApp = {
    selectedProvider: null,
    authDialog: null,
    connectedProviders: {},

    // Configuración OAuth de Google BigQuery con tu Client ID real
    googleConfig: {
        clientId: "316357511817-lck6pdotv8mrb7n72pahuukt2e0fvsrt.apps.googleusercontent.com",
        scopes: "https://www.googleapis.com/auth/bigquery.readonly https://www.googleapis.com/auth/userinfo.email"
    },

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
                this.updateActionButton();
            });
        });

        if (btnAction) {
            btnAction.addEventListener("click", () => {
                if (!this.selectedProvider) return;

                const isConnected = !!this.connectedProviders[this.selectedProvider];

                if (isConnected) {
                    this.logoutProvider(this.selectedProvider);
                } else {
                    if (this.selectedProvider === "bigquery") {
                        this.connectBigQuery();
                    } else {
                        this.showAlert(`El conector para ${this.getProviderDisplayName(this.selectedProvider)} estará disponible próximamente.`);
                    }
                }
            });
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
        const token = localStorage.getItem("bigquery_access_token");
        const expires = localStorage.getItem("bigquery_token_expires");

        if (token && expires && Date.now() < parseInt(expires)) {
            this.setProviderState("bigquery", true);
        } else {
            localStorage.removeItem("bigquery_access_token");
            localStorage.removeItem("bigquery_token_expires");
            this.setProviderState("bigquery", false);
        }
    },

    /**
     * Inicia el flujo OAuth 2.0 en ventana emergente para BigQuery
     */
    connectBigQuery() {
        // Validación de protocolo: alertar si se abre directamente desde file://
        if (window.location.protocol === "file:") {
            console.error("No se puede ejecutar desde file://. Debes usar un servidor web local (ej. Live Server o http://localhost).");
            this.showAlert("Error: No se puede ejecutar desde un archivo local (file://). Usa Live Server o un servidor HTTP.", true);
            return;
        }

        // Construcción de la URL del callback en relación al origen actual
        const redirectUri = new URL("auth-callback.html", window.location.href).href;
        
        const authUrl = "https://accounts.google.com/o/oauth2/v2/auth?" +
            `client_id=${encodeURIComponent(this.googleConfig.clientId)}` +
            `&redirect_uri=${encodeURIComponent(redirectUri)}` +
            "&response_type=token" +
            `&scope=${encodeURIComponent(this.googleConfig.scopes)}` +
            "&prompt=consent";

        // Comprobación de si estamos ejecutando dentro de Microsoft Office
        const isOfficeEnvironment = typeof Office !== "undefined" && 
                                   Office.context && 
                                   Office.context.ui && 
                                   typeof Office.context.ui.displayDialogAsync === "function";

        if (isOfficeEnvironment) {
            console.log("Entorno detectado: Microsoft Office Add-in (Excel). Usando displayDialogAsync.");

            const dialogOptions = {
                height: 60,
                width: 40,
                displayInIframe: false
            };

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
            console.log("Entorno detectado: Navegador web estándar. Usando window.open.");

            const width = 500;
            const height = 650;
            const left = (window.screen.width / 2) - (width / 2);
            const top = (window.screen.height / 2) - (height / 2);
            alert("OPEN");
            window.open(
                authUrl,
                "GoogleAuthWindow",
                `width=${width},height=${height},top=${top},left=${left}`
            );
        }
    },

    /**
     * Maneja la respuesta recibida tras el login
     */
    handleAuthResponse(messageString) {
        try {
            const response = JSON.parse(messageString);

            if (response.status === "success" && response.provider === "bigquery") {
                localStorage.setItem("bigquery_access_token", response.token);
                localStorage.setItem("bigquery_token_expires", Date.now() + (parseInt(response.expiresIn) * 1000));

                this.setProviderState("bigquery", true);
                this.updateActionButton();
                this.showAlert("¡Conexión con Google BigQuery establecida con éxito!");
            } else {
                console.error("Error en autenticación:", response.error);
                this.showAlert("Error de autenticación: " + (response.error || "Desconocido"), true);
            }
        } catch (err) {
            console.error("Error leyendo respuesta de autenticación:", err);
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
            localStorage.removeItem("bigquery_access_token");
            localStorage.removeItem("bigquery_token_expires");
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
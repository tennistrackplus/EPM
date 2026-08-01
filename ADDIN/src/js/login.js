/**
 * Lógica del Task Pane de Login, Verificación de Sesión y Desconexión
 */
if (typeof Office !== "undefined") {
    Office.onReady(() => {
        LoginApp.init();
    });
} else {
    document.addEventListener("DOMContentLoaded", () => {
        LoginApp.init();
    });
}

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
                        alert(`El conector para ${this.getProviderDisplayName(this.selectedProvider)} estará disponible próximamente.`);
                    }
                }
            });
        }
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
        const redirectUri = location.origin + location.pathname.replace("login.html", "auth-callback.html");
        
        const authUrl = "https://accounts.google.com/o/oauth2/v2/auth?" +
            `client_id=${encodeURIComponent(this.googleConfig.clientId)}` +
            `&redirect_uri=${encodeURIComponent(redirectUri)}` +
            "&response_type=token" +
            `&scope=${encodeURIComponent(this.googleConfig.scopes)}` +
            "&prompt=consent";

        const dialogOptions = {
            height: 60,
            width: 40,
            displayInIframe: false
        };

        Office.context.ui.displayDialogAsync(authUrl, dialogOptions, (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Error al abrir diálogo de autenticación:", asyncResult.error.message);
                alert("No se pudo abrir la ventana de login: " + asyncResult.error.message);
                return;
            }

            this.authDialog = asyncResult.value;

            this.authDialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
                this.handleAuthResponse(arg.message);
            });
        });
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
                alert("¡Conexión con Google BigQuery establecida con éxito!");
            } else {
                console.error("Error en autenticación:", response.error);
                alert("Error de autenticación: " + (response.error || "Desconocido"));
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
        alert(`Sesión cerrada para ${this.getProviderDisplayName(providerKey)}`);
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
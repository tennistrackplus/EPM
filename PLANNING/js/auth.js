/**
 * Lógica de la landing page: conexión OAuth a BigQuery o Snowflake,
 * configuración del "hogar" de trabajo (proyecto GCP / cuenta+warehouse)
 * y arranque del esquema de control sobre el motor elegido.
 */
const LoginApp = {
    selectedProvider: null,
    connectedProviders: {},

    init() {
        this.bindEvents();
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

    bindEvents() {
        const cards = document.querySelectorAll(".connector-card");
        cards.forEach(card => {
            card.addEventListener("click", () => {
                cards.forEach(c => c.classList.remove("selected"));
                card.classList.add("selected");
                this.selectedProvider = card.getAttribute("data-provider");
                this.togglePanels();
                this.updateActionButton();
            });
        });

        const gcpSelect = document.getElementById("gcpProjectSelect");
        if (gcpSelect) {
            gcpSelect.addEventListener("change", () => {
                BQ.setGcpProject(gcpSelect.value);
                this.updateActionButton();
            });
        }
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
                    this.updateActionButton();
                    this.showAlert("¡Conexión con Google BigQuery establecida con éxito!");
                    await this.loadGcpProjects();
                } else {
                    this.showAlert("Error de autenticación con BigQuery: " + (response.error || "Desconocido"), true);
                }
                return;
            }

            if (response.provider === "snowflake") {
                if (response.status === "success" && response.code) {
                    await SF.handleAuthCode(response.code, response.state);
                    this.setProviderState("snowflake", true);
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
            await DracoSchema.bootstrap((msg) => { bootstrapMsgEl.innerText = msg; });
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
                    Provider.setKey("bigquery");
                    btnAction.disabled = true;
                    btnAction.innerText = "Preparando tu espacio Draco...";
                    try {
                        await this.enterApp(btnAction);
                    } catch (e) {
                        btnAction.disabled = false;
                        btnAction.innerText = "Reintentar";
                    }
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
                btnAction.onclick = async () => {
                    Provider.setKey("snowflake");
                    btnAction.disabled = true;
                    btnAction.innerText = "Preparando tu espacio Draco...";
                    try {
                        await this.enterApp(btnAction);
                    } catch (e) {
                        btnAction.disabled = false;
                        btnAction.innerText = "Reintentar";
                    }
                };
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
    },

    getProviderDisplayName(providerKey) {
        const names = {
            bigquery: "BigQuery", amazon: "Amazon", fabric: "Microsoft Fabric",
            snowflake: "Snowflake", datasphere: "SAP Datasphere", s4cds: "CDS de S4"
        };
        return names[providerKey] || providerKey;
    }
};

document.addEventListener("DOMContentLoaded", () => LoginApp.init());

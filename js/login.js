/**
 * Lógica del Task Pane de Login y Selección de Conectores
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

    init() {
        this.bindEvents();
    },

    bindEvents() {
        const cards = document.querySelectorAll(".connector-card");
        const btnConnect = document.getElementById("btnConnect");

        cards.forEach(card => {
            card.addEventListener("click", () => {
                // Desmarcar selecciones previas
                cards.forEach(c => c.classList.remove("selected"));
                
                // Marcar tarjeta actual
                card.classList.add("selected");
                this.selectedProvider = card.getAttribute("data-provider");

                // Habilitar botón de conexión
                if (btnConnect) {
                    btnConnect.disabled = false;
                    btnConnect.innerText = `Conectar a ${this.getProviderDisplayName(this.selectedProvider)}`;
                }
            });
        });

        if (btnConnect) {
            btnConnect.addEventListener("click", () => {
                if (this.selectedProvider) {
                    console.log("Iniciando flujo de conexión para:", this.selectedProvider);
                    alert(`Iniciando autenticación para ${this.getProviderDisplayName(this.selectedProvider)}...`);
                }
            });
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
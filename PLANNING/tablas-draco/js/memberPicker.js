/* ==========================================================================
 * memberPicker.js — buscador de miembros como diálogo de Office
 * independiente (Office.context.ui.displayDialogAsync), NO como overlay
 * dentro del taskpane. Se abre centrado sobre la ventana de Excel desde
 * commands.js (openMemberRecognitionPicker) cuando el "reconocimiento de
 * miembros" detecta texto escrito a mano en una celda de un eje Estático.
 *
 * Recibe dim/attr/search por querystring (?dim=...&attr=...&search=...) y
 * devuelve el resultado al que lo abrió con Office.context.ui.messageParent,
 * como { value, attribute } si se elige un valor, o null si se cancela.
 * ========================================================================== */

// Misma traducción de LoadJson que usa FilterModal (ver filterModal.js).
function loadJsonTree(json) {
    const fieldMatches = [...json.matchAll(/"name":\s*"([^"]+)"/g)];
    const campos = fieldMatches.map(m => m[1]);

    if (campos.length === 0) return [];

    const ultimos = campos.map(() => "");
    const valores = campos.map(() => "");

    const valueMatches = [...json.matchAll(/"v":\s*"([^"]*)"/g)];

    const items = [];
    let nivel = 0;

    for (const m of valueMatches) {
        valores[nivel] = m[1];
        nivel++;

        if (nivel > campos.length - 1) {
            for (let i = 0; i < campos.length; i++) {
                if (valores[i] !== ultimos[i]) {
                    items.push({
                        text: " ".repeat(i * 4) + valores[i],
                        attribute: campos[i],
                        value: valores[i]
                    });

                    ultimos[i] = valores[i];
                    for (let j = i + 1; j < campos.length; j++) ultimos[j] = "";
                }
            }
            nivel = 0;
        }
    }

    return items;
}

const MemberPicker = {
    allItems: [],
    selected: null,
    dim: "",
    attr: "",

    init() {
        const params = new URLSearchParams(window.location.search);
        this.dim = params.get("dim") || "";
        this.attr = params.get("attr") || "";
        const initialSearch = params.get("search") || "";

        const titleEl = document.getElementById("pickerTitle");
        if (titleEl) titleEl.innerText = `Buscar miembro: ${this.dim}.${this.attr}`;

        const searchInput = document.getElementById("pickerSearchInput");
        if (searchInput) {
            searchInput.value = initialSearch;
            searchInput.addEventListener("input", (e) => this.search(e.target.value));
            searchInput.focus();
        }

        const btnCancel = document.getElementById("btnPickerCancel");
        if (btnCancel) btnCancel.addEventListener("click", () => this.cancel());

        const btnApply = document.getElementById("btnPickerApply");
        if (btnApply) btnApply.addEventListener("click", () => this.apply());

        this.load(initialSearch);
    },

    async load(initialSearch) {
        const container = document.getElementById("pickerItemsContainer");
        try {
            const sql = await window.ExcelService.buildFilterValuesSQL(this.dim, this.attr);

            if (!sql) {
                container.innerHTML = "<div style='padding:10px;color:#a80000;'>No se ha encontrado el atributo o jerarquía.</div>";
                return;
            }

            const json = await window.ExcelService.executeSQL(sql);
            this.allItems = loadJsonTree(json);
            this.search(initialSearch || "");
        } catch (err) {
            console.error("Error cargando valores de miembros:", err);
            container.innerHTML = `<div style='padding:10px;color:#a80000;'>Error: ${err.message || err}</div>`;
        }
    },

    render(items) {
        const container = document.getElementById("pickerItemsContainer");
        container.innerHTML = "";

        if (items.length === 0) {
            container.innerHTML = "<div style='padding:10px;color:#605e5c;'>No hay valores para mostrar.</div>";
            return;
        }

        items.forEach((item) => {
            const row = document.createElement("div");
            row.className = "modal-item-row";
            row.style.whiteSpace = "pre";
            row.style.cursor = "pointer";
            row.textContent = item.text;

            if (this.selected && this.selected.attribute === item.attribute && this.selected.value === item.value) {
                row.classList.add("selected");
            }

            row.addEventListener("click", () => {
                container.querySelectorAll(".modal-item-row").forEach(r => r.classList.remove("selected"));
                row.classList.add("selected");
                this.selected = item;
            });

            row.addEventListener("dblclick", () => {
                this.selected = item;
                this.apply();
            });

            container.appendChild(row);
        });
    },

    search(query) {
        const q = String(query).toLowerCase().trim();
        const filtered = q === ""
            ? this.allItems
            : this.allItems.filter(it => it.text.toLowerCase().includes(q));
        this.render(filtered);
    },

    apply() {
        if (!this.selected) {
            this.cancel();
            return;
        }
        Office.context.ui.messageParent(JSON.stringify({ value: this.selected.value, attribute: this.selected.attribute }));
    },

    cancel() {
        Office.context.ui.messageParent(JSON.stringify(null));
    }
};

Office.onReady(() => {
    MemberPicker.init();
});

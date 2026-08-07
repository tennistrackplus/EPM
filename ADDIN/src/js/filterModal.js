/* ==========================================================================
 * FilterModal — traducción de frmFilterValues (VBA), en versión de
 * selección ÚNICA (sin checkboxes ni selección por rango, de momento).
 * ========================================================================== */

/**
 * Traducción de LoadJson: separa los nombres de campo ("name":"...") y los
 * valores ("v":"...") del JSON crudo de BigQuery, y construye una lista
 * "en árbol" (indentada) deduplicando cada nivel frente al valor anterior,
 * igual que hacía el bucle de Ultimos()/Valores() en VBA.
 */
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

const FilterModal = {
    allItems: [],
    selected: null,
    resolveFn: null,

    init() {
        const closeBtn = document.getElementById("closeModalBtn");
        const cancelBtn = document.getElementById("btnCancelModal");
        const applyBtn = document.getElementById("btnApplyModal");
        const searchInput = document.getElementById("modalSearchInput");

        if (closeBtn) closeBtn.addEventListener("click", () => this.cancel());
        if (cancelBtn) cancelBtn.addEventListener("click", () => this.cancel());
        if (applyBtn) applyBtn.addEventListener("click", () => this.apply());
        if (searchInput) searchInput.addEventListener("input", (e) => this.search(e.target.value));
    },

    /**
     * Abre el modal para el campo indicado y devuelve una Promise que
     * resuelve a { value, attribute } si el usuario elige un valor,
     * o null si cancela.
     */
    open(fieldData) {
        return new Promise(async (resolve) => {
            this.resolveFn = resolve;
            this.selected = null;
            this.allItems = [];

            const titleEl = document.getElementById("modalTitle");
            if (titleEl) titleEl.innerText = `Filtrar: ${fieldData.dim}.${fieldData.name}`;

            const searchInput = document.getElementById("modalSearchInput");
            if (searchInput) searchInput.value = "";

            const container = document.getElementById("modalItemsContainer");
            container.innerHTML = "<div style='padding:10px;color:#605e5c;'>Cargando valores…</div>";

            document.getElementById("filterModal").style.display = "flex";

            try {
                const sql = await window.ExcelService.buildFilterValuesSQL(fieldData.dim, fieldData.name);

                if (!sql) {
                    container.innerHTML = "<div style='padding:10px;color:#a80000;'>No se ha encontrado el atributo o jerarquía.</div>";
                    return;
                }

                const json = await window.ExcelService.executeSQL(sql);
                this.allItems = loadJsonTree(json);
                this.render(this.allItems);
            } catch (err) {
                console.error("Error cargando valores de filtro:", err);
                container.innerHTML = `<div style='padding:10px;color:#a80000;'>Error: ${err.message || err}</div>`;
            }
        });
    },

    render(items) {
        const container = document.getElementById("modalItemsContainer");
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
        const result = { value: this.selected.value, attribute: this.selected.attribute };
        const resolve = this.resolveFn;
        this.close();
        if (resolve) resolve(result);
    },

    cancel() {
        const resolve = this.resolveFn;
        this.close();
        if (resolve) resolve(null);
    },

    close() {
        document.getElementById("filterModal").style.display = "none";
        this.resolveFn = null;
    }
};

window.FilterModal = FilterModal;

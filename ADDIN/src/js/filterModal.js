/* ==========================================================================
 * FilterModal — selector de valores para la zona "Filtros" del taskpane.
 *
 * Soporta:
 *   - Selección MÚLTIPLE de valores (checkboxes), con buscador.
 *   - RANGO (Desde/Hasta) — solo tiene sentido en atributos planos, así que
 *     la pestaña "Rango" se oculta cuando el campo es una jerarquía.
 *   - INCLUIR / EXCLUIR — aplica tanto a "Valores" como a "Rango".
 *
 * El resultado que devuelve open() ya no es un único {value, attribute},
 * sino un objeto "filter" con esta forma:
 *
 *   Dimensión plana, varios valores:
 *     { mode: "values", include: true|false, attribute: "PAIS", values: ["ES","FR"] }
 *
 *   Dimensión plana, rango:
 *     { mode: "range", include: true|false, attribute: "IMPORTE", from: "0", to: "1000" }
 *
 *   Jerarquía, varios miembros (posiblemente de niveles/atributos distintos):
 *     { mode: "values", include: true|false, items: [{attribute:"PAIS", value:"España"}, ...] }
 *
 * taskpane.js guarda ese objeto tal cual (como JSON) en entry.value, y
 * commands.js (buildWhere) lo interpreta para construir IN/NOT IN/BETWEEN.
 * Un valor de filtro "antiguo" (una simple cadena, sin JSON) se sigue
 * interpretando como igualdad simple, por compatibilidad.
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

/** Clave única de un item del árbol (atributo + valor), usada para el Set de seleccionados. */
function itemKey(item) {
    return item.attribute + "||" + item.value;
}

const FilterModal = {
    allItems: [],
    fieldData: null,

    mode: "values",          // "values" | "range"
    include: true,           // true = incluir, false = excluir
    selectedKeys: new Set(), // claves (itemKey) de los valores marcados
    rangeFrom: "",
    rangeTo: "",

    resolveFn: null,

    init() {
        const closeBtn = document.getElementById("closeModalBtn");
        const cancelBtn = document.getElementById("btnCancelModal");
        const applyBtn = document.getElementById("btnApplyModal");
        const searchInput = document.getElementById("modalSearchInput");
        const selectAllBtn = document.getElementById("btnSelectAllVisible");
        const clearBtn = document.getElementById("btnClearSelection");
        const rangeFromInput = document.getElementById("modalRangeFrom");
        const rangeToInput = document.getElementById("modalRangeTo");

        if (closeBtn) closeBtn.addEventListener("click", () => this.cancel());
        if (cancelBtn) cancelBtn.addEventListener("click", () => this.cancel());
        if (applyBtn) applyBtn.addEventListener("click", () => this.apply());
        if (searchInput) searchInput.addEventListener("input", (e) => this.search(e.target.value));

        if (selectAllBtn) selectAllBtn.addEventListener("click", () => this.selectAllVisible());
        if (clearBtn) clearBtn.addEventListener("click", () => this.clearSelection());

        if (rangeFromInput) rangeFromInput.addEventListener("input", (e) => { this.rangeFrom = e.target.value; });
        if (rangeToInput) rangeToInput.addEventListener("input", (e) => { this.rangeTo = e.target.value; });

        document.querySelectorAll("#filterModeTabs .segmented-option").forEach(btn => {
            btn.addEventListener("click", () => this.setMode(btn.dataset.mode));
        });

        document.querySelectorAll("#filterIncludeToggle .segmented-option").forEach(btn => {
            btn.addEventListener("click", () => this.setInclude(btn.dataset.include === "true"));
        });
    },

    /**
     * Abre el modal para el campo indicado y devuelve una Promise que
     * resuelve al objeto "filter" descrito arriba, o null si el usuario
     * cancela.
     *
     * fieldData admite opcionalmente `currentFilter` (el filtro ya
     * aplicado anteriormente sobre este campo) para precargar la
     * selección al reabrir el modal.
     */
    open(fieldData) {
        return new Promise(async (resolve) => {
            this.resolveFn = resolve;
            this.fieldData = fieldData;
            this.allItems = [];
            this.selectedKeys = new Set();
            this.rangeFrom = "";
            this.rangeTo = "";
            this.include = true;
            this.mode = "values";

            this.preloadCurrentFilter(fieldData.currentFilter);

            const titleEl = document.getElementById("modalTitle");
            if (titleEl) titleEl.innerText = `Filtrar: ${fieldData.dim}.${fieldData.name}`;

            // El rango no tiene sentido en jerarquías: se oculta la pestaña.
            const rangeTab = document.querySelector('#filterModeTabs [data-mode="range"]');
            if (rangeTab) rangeTab.style.display = fieldData.isHierarchy ? "none" : "";

            const searchInput = document.getElementById("modalSearchInput");
            if (searchInput) searchInput.value = fieldData.initialSearch || "";

            document.getElementById("modalRangeFrom").value = this.rangeFrom;
            document.getElementById("modalRangeTo").value = this.rangeTo;

            this.renderModeUI();
            this.updateIncludeUI();
            this.updateSelectionCount();

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
                this.search(fieldData.initialSearch || "");
                this.renderChips();
            } catch (err) {
                console.error("Error cargando valores de filtro:", err);
                container.innerHTML = `<div style='padding:10px;color:#a80000;'>Error: ${err.message || err}</div>`;
            }
        });
    },

    /** Reconstruye mode/include/selectedKeys/range a partir de un filtro ya aplicado antes. */
    preloadCurrentFilter(currentFilter) {
        if (!currentFilter || !currentFilter.mode) return;

        this.mode = currentFilter.mode;
        this.include = currentFilter.include !== false;

        if (currentFilter.mode === "range") {
            this.rangeFrom = currentFilter.from || "";
            this.rangeTo = currentFilter.to || "";
        } else if (currentFilter.items) {
            currentFilter.items.forEach(it => this.selectedKeys.add(it.attribute + "||" + it.value));
        } else if (currentFilter.values) {
            // Formato antiguo (una cadena simple) no trae "attribute": se
            // asume el atributo del propio campo que se está filtrando.
            const attr = currentFilter.attribute || (this.fieldData && this.fieldData.name);
            currentFilter.values.forEach(v => this.selectedKeys.add(attr + "||" + v));
        }
    },

    /* -------------------------------------------------------------
     * Cambios de modo (Valores/Rango) e Incluir/Excluir
     * ----------------------------------------------------------- */

    setMode(mode) {
        if (mode === "range" && this.fieldData && this.fieldData.isHierarchy) return; // no aplica
        this.mode = mode;
        this.renderModeUI();
    },

    renderModeUI() {
        document.querySelectorAll("#filterModeTabs .segmented-option").forEach(btn => {
            btn.classList.toggle("active", btn.dataset.mode === this.mode);
        });
        document.getElementById("filterValuesPanel").style.display = this.mode === "values" ? "" : "none";
        document.getElementById("filterRangePanel").style.display = this.mode === "range" ? "" : "none";

        const chips = document.getElementById("modalSelectedChips");
        if (chips) chips.style.display = this.mode === "values" ? "" : "none";
    },

    setInclude(include) {
        this.include = include;
        this.updateIncludeUI();
    },

    updateIncludeUI() {
        document.querySelectorAll("#filterIncludeToggle .segmented-option").forEach(btn => {
            btn.classList.toggle("active", (btn.dataset.include === "true") === this.include);
        });
        const chips = document.getElementById("modalSelectedChips");
        if (chips) chips.classList.toggle("exclude-mode", !this.include);
    },

    /* -------------------------------------------------------------
     * Lista de valores (checkboxes) + búsqueda
     * ----------------------------------------------------------- */

    render(items) {
        const container = document.getElementById("modalItemsContainer");
        container.innerHTML = "";

        if (items.length === 0) {
            container.innerHTML = "<div style='padding:10px;color:#605e5c;'>No hay valores para mostrar.</div>";
            return;
        }

        items.forEach((item) => {
            const key = itemKey(item);

            const row = document.createElement("div");
            row.className = "modal-item-row";
            row.style.whiteSpace = "pre";

            const checkbox = document.createElement("input");
            checkbox.type = "checkbox";
            checkbox.checked = this.selectedKeys.has(key);

            const label = document.createElement("label");
            label.textContent = item.text;
            label.style.whiteSpace = "pre";

            row.appendChild(checkbox);
            row.appendChild(label);

            const toggle = () => {
                if (this.selectedKeys.has(key)) {
                    this.selectedKeys.delete(key);
                } else {
                    this.selectedKeys.add(key);
                }
                checkbox.checked = this.selectedKeys.has(key);
                row.classList.toggle("selected", checkbox.checked);
                this.renderChips();
                this.updateSelectionCount();
            };

            checkbox.addEventListener("click", (e) => { e.stopPropagation(); toggle(); });
            row.addEventListener("click", toggle);

            if (checkbox.checked) row.classList.add("selected");

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

    selectAllVisible() {
        const q = String(document.getElementById("modalSearchInput").value || "").toLowerCase().trim();
        const visible = q === "" ? this.allItems : this.allItems.filter(it => it.text.toLowerCase().includes(q));
        visible.forEach(it => this.selectedKeys.add(itemKey(it)));
        this.search(document.getElementById("modalSearchInput").value);
        this.renderChips();
        this.updateSelectionCount();
    },

    clearSelection() {
        this.selectedKeys.clear();
        this.search(document.getElementById("modalSearchInput").value);
        this.renderChips();
        this.updateSelectionCount();
    },

    /* -------------------------------------------------------------
     * Chips con la selección actual (feedback visual inmediato)
     * ----------------------------------------------------------- */

    renderChips() {
        const container = document.getElementById("modalSelectedChips");
        if (!container) return;

        if (this.selectedKeys.size === 0) {
            container.style.display = "none";
            container.innerHTML = "";
            return;
        }

        container.style.display = "flex";
        container.classList.toggle("exclude-mode", !this.include);
        container.innerHTML = "";

        const byKey = new Map(this.allItems.map(it => [itemKey(it), it]));

        this.selectedKeys.forEach(key => {
            const item = byKey.get(key);
            const label = item ? item.text.trim() : key.split("||")[1];

            const chip = document.createElement("div");
            chip.className = "filter-chip";
            chip.innerHTML = `<span class="chip-label"></span><button type="button" class="chip-remove">&times;</button>`;
            chip.querySelector(".chip-label").textContent = label;

            chip.querySelector(".chip-remove").addEventListener("click", () => {
                this.selectedKeys.delete(key);
                this.renderChips();
                this.search(document.getElementById("modalSearchInput").value);
                this.updateSelectionCount();
            });

            container.appendChild(chip);
        });
    },

    updateSelectionCount() {
        const el = document.getElementById("modalSelectionCount");
        if (!el) return;
        if (this.mode !== "values" || this.selectedKeys.size === 0) {
            el.textContent = "";
            return;
        }
        el.textContent = this.selectedKeys.size === 1
            ? "1 valor seleccionado"
            : `${this.selectedKeys.size} valores seleccionados`;
    },

    /* -------------------------------------------------------------
     * Aplicar / cancelar / cerrar
     * ----------------------------------------------------------- */

    apply() {
        let result = null;

        if (this.mode === "range") {
            const from = String(this.rangeFrom || "").trim();
            const to = String(this.rangeTo || "").trim();
            if (from === "" && to === "") {
                this.cancel();
                return;
            }
            result = {
                mode: "range",
                include: this.include,
                attribute: this.fieldData.name,
                from,
                to
            };
        } else {
            if (this.selectedKeys.size === 0) {
                this.cancel();
                return;
            }

            const byKey = new Map(this.allItems.map(it => [itemKey(it), it]));
            const chosen = [...this.selectedKeys].map(k => byKey.get(k)).filter(Boolean);

            if (this.fieldData.isHierarchy) {
                result = {
                    mode: "values",
                    include: this.include,
                    items: chosen.map(it => ({ attribute: it.attribute, value: it.value }))
                };
            } else {
                result = {
                    mode: "values",
                    include: this.include,
                    attribute: this.fieldData.name,
                    values: chosen.map(it => it.value)
                };
            }
        }

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

/**
 * Interpreta el valor bruto guardado para un filtro (columna "Valor").
 * Si es JSON con forma reconocible (mode: "values"|"range") lo usa tal
 * cual; si no, lo trata como el formato antiguo: una cadena simple que
 * equivale a una igualdad de un único valor.
 */
function parseFilterValue(raw) {
    const s = String(raw || "").trim();
    if (s === "") return null;

    if (s[0] === "{") {
        try {
            const parsed = JSON.parse(s);
            if (parsed && typeof parsed === "object" && parsed.mode) return parsed;
        } catch (e) {
            // No era JSON válido: se interpreta como valor simple (ver abajo).
        }
    }

    return { mode: "values", include: true, values: [s] };
}

/**
 * Construye un resumen legible de un filtro (para el "tag" del taskpane).
 * Acepta tanto el formato nuevo (objeto) como el antiguo (cadena simple),
 * por compatibilidad con informes guardados previamente.
 */
function describeFilter(filter) {
    if (!filter) return "";

    if (typeof filter === "string") return filter; // formato antiguo

    if (filter.mode === "range") {
        const desde = filter.from || "…";
        const hasta = filter.to || "…";
        return (filter.include ? "" : "≠ ") + `${desde} – ${hasta}`;
    }

    const values = filter.items ? filter.items.map(it => it.value) : (filter.values || []);
    if (values.length === 0) return "";

    const prefix = filter.include ? "" : "≠ ";
    if (values.length <= 2) return prefix + values.join(", ");
    return prefix + `${values.slice(0, 2).join(", ")} (+${values.length - 2} más)`;
}

window.FilterModal = FilterModal;
window.describeFilter = describeFilter;
window.parseFilterValue = parseFilterValue;

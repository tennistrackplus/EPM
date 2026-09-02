/**
 * filterDialog.js — lógica del diálogo de Office independiente "Filtrar
 * campo" (filterDialog.html).
 * ------------------------------------------------------------------------
 * Reemplaza al antiguo overlay embebido en taskpane.html: ahora el
 * selector de filtro vive en SU PROPIA ventana (igual que memberPicker.html
 * u openSemanticModel.html), lo que le da sitio de sobra para un panel
 * auxiliar a la derecha (modo, incluir/excluir, chips de selección) junto
 * a la lista de valores o el rango, sin pelearse con el ancho estrecho
 * del taskpane.
 *
 * Como cualquier diálogo de Office, esta ventana NO tiene acceso al modelo
 * de objetos de Excel ni a window.ExcelService: quien lo abre (ver
 * FilterModal.open en js/filterModal.js, ejecutado en el taskpane) ya ha
 * resuelto antes la lista de valores con ExcelService.buildFilterValuesSQL
 * + executeSQL, y nos la manda por mensaje en cuanto avisamos que estamos
 * listos. Cuando el usuario pulsa "Aceptar"/"Cancelar" devolvemos el
 * resultado de la misma forma (Office.context.ui.messageParent).
 *
 * El objeto "filter" que se envía de vuelta tiene la misma forma que
 * generaba el modal antiguo (ver cabecera de js/filterModal.js):
 *   - Valores (plano):   { mode:"values", include, attribute, values:[...] }
 *   - Valores (jerarquía): { mode:"values", include, items:[{attribute,value}] }
 *   - Rango (solo plano): { mode:"range", include, attribute, from, to }
 */

let allItems = [];
let fieldData = null;

let mode = "values";          // "values" | "range"
let include = true;           // true = incluir, false = excluir
let selectedKeys = new Set(); // claves (attribute||value) marcadas
let rangeFrom = "";
let rangeTo = "";

function itemKey(item) {
    return item.attribute + "||" + item.value;
}

function log(...args) {
    console.log("[FilterDialog]", ...args);
}

/* -------------------------------------------------------------
 * Comunicación con el host (taskpane)
 * ----------------------------------------------------------- */

function sendToParent(payload) {
    Office.context.ui.messageParent(JSON.stringify(payload));
}

function closeBySelf() {
    // No cerramos la ventana nosotros mismos: quien la abrió (FilterModal.open)
    // es quien llama a dialog.close() al recibir nuestro mensaje. Aquí solo
    // avisamos.
}

/* -------------------------------------------------------------
 * Modo (Valores/Rango) e Incluir/Excluir
 * ----------------------------------------------------------- */

function setMode(newMode) {
    if (newMode === "range" && fieldData && fieldData.isHierarchy) return; // no aplica
    mode = newMode;
    renderModeUI();
}

function renderModeUI() {
    document.querySelectorAll("#filterModeTabs .segmented-option").forEach(btn => {
        btn.classList.toggle("active", btn.dataset.mode === mode);
    });
    document.getElementById("filterValuesPanel").style.display = mode === "values" ? "" : "none";
    document.getElementById("filterRangePanel").style.display = mode === "range" ? "" : "none";
    document.getElementById("modalSelectedChips").parentElement.style.display = mode === "values" ? "" : "none";
    updateSelectionCount();
}

function setInclude(value) {
    include = value;
    updateIncludeUI();
}

function updateIncludeUI() {
    document.querySelectorAll("#filterIncludeToggle .segmented-option").forEach(btn => {
        btn.classList.toggle("active", (btn.dataset.include === "true") === include);
    });
    document.getElementById("modalSelectedChips").classList.toggle("exclude-mode", !include);
}

/* -------------------------------------------------------------
 * Lista de valores (checkboxes) + búsqueda
 * ----------------------------------------------------------- */

function render(items) {
    const container = document.getElementById("modalItemsContainer");
    container.innerHTML = "";

    if (items.length === 0) {
        container.innerHTML = "<div style='padding:10px;color:var(--text-muted);'>No hay valores para mostrar.</div>";
        return;
    }

    items.forEach((item) => {
        const key = itemKey(item);

        const row = document.createElement("div");
        row.className = "modal-item-row";
        row.style.whiteSpace = "pre";

        const checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.checked = selectedKeys.has(key);

        const label = document.createElement("label");
        label.textContent = item.text;
        label.style.whiteSpace = "pre";

        row.appendChild(checkbox);
        row.appendChild(label);

        const toggle = () => {
            if (selectedKeys.has(key)) {
                selectedKeys.delete(key);
            } else {
                selectedKeys.add(key);
            }
            checkbox.checked = selectedKeys.has(key);
            row.classList.toggle("selected", checkbox.checked);
            renderChips();
            updateSelectionCount();
        };

        checkbox.addEventListener("click", (e) => { e.stopPropagation(); toggle(); });
        row.addEventListener("click", toggle);

        if (checkbox.checked) row.classList.add("selected");

        container.appendChild(row);
    });
}

function applySearch(query) {
    const q = String(query || "").toLowerCase().trim();
    const filtered = q === "" ? allItems : allItems.filter(it => it.text.toLowerCase().includes(q));
    render(filtered);
}

function selectAllVisible() {
    const q = document.getElementById("modalSearchInput").value;
    const visible = String(q || "").trim() === ""
        ? allItems
        : allItems.filter(it => it.text.toLowerCase().includes(q.toLowerCase()));
    visible.forEach(it => selectedKeys.add(itemKey(it)));
    applySearch(q);
    renderChips();
    updateSelectionCount();
}

function clearSelection() {
    selectedKeys.clear();
    applySearch(document.getElementById("modalSearchInput").value);
    renderChips();
    updateSelectionCount();
}

/* -------------------------------------------------------------
 * Chips con la selección actual (panel auxiliar de la derecha)
 * ----------------------------------------------------------- */

function renderChips() {
    const container = document.getElementById("modalSelectedChips");

    if (selectedKeys.size === 0) {
        container.style.display = "none";
        container.innerHTML = "";
        return;
    }

    container.style.display = "flex";
    container.classList.toggle("exclude-mode", !include);
    container.innerHTML = "";

    const byKey = new Map(allItems.map(it => [itemKey(it), it]));

    selectedKeys.forEach(key => {
        const item = byKey.get(key);
        const label = item ? item.text.trim() : key.split("||")[1];

        const chip = document.createElement("div");
        chip.className = "filter-chip";
        chip.innerHTML = `<span class="chip-label"></span><button type="button" class="chip-remove">&times;</button>`;
        chip.querySelector(".chip-label").textContent = label;

        chip.querySelector(".chip-remove").addEventListener("click", () => {
            selectedKeys.delete(key);
            renderChips();
            applySearch(document.getElementById("modalSearchInput").value);
            updateSelectionCount();
        });

        container.appendChild(chip);
    });
}

function updateSelectionCount() {
    const el = document.getElementById("modalSelectionCount");
    if (mode !== "values" || selectedKeys.size === 0) {
        el.textContent = "";
        return;
    }
    el.textContent = selectedKeys.size === 1
        ? "1 valor seleccionado"
        : `${selectedKeys.size} valores seleccionados`;
}

/* -------------------------------------------------------------
 * Precarga de un filtro ya aplicado antes (al reabrir sobre el mismo campo)
 * ----------------------------------------------------------- */

function preloadCurrentFilter(currentFilter) {
    if (!currentFilter || !currentFilter.mode) return;

    mode = currentFilter.mode;
    include = currentFilter.include !== false;

    if (currentFilter.mode === "range") {
        rangeFrom = currentFilter.from || "";
        rangeTo = currentFilter.to || "";
    } else if (currentFilter.items) {
        currentFilter.items.forEach(it => selectedKeys.add(it.attribute + "||" + it.value));
    } else if (currentFilter.values) {
        const attr = currentFilter.attribute || (fieldData && fieldData.name);
        currentFilter.values.forEach(v => selectedKeys.add(attr + "||" + v));
    }
}

/* -------------------------------------------------------------
 * Aplicar / cancelar
 * ----------------------------------------------------------- */

function apply() {
    let result = null;

    if (mode === "range") {
        const from = String(rangeFrom || "").trim();
        const to = String(rangeTo || "").trim();
        if (from === "" && to === "") {
            sendToParent({ type: "cancel" });
            return;
        }
        result = { mode: "range", include, attribute: fieldData.name, from, to };
    } else {
        if (selectedKeys.size === 0) {
            sendToParent({ type: "cancel" });
            return;
        }

        const byKey = new Map(allItems.map(it => [itemKey(it), it]));
        const chosen = [...selectedKeys].map(k => byKey.get(k)).filter(Boolean);

        if (fieldData.isHierarchy) {
            result = { mode: "values", include, items: chosen.map(it => ({ attribute: it.attribute, value: it.value })) };
        } else {
            result = { mode: "values", include, attribute: fieldData.name, values: chosen.map(it => it.value) };
        }
    }

    sendToParent({ type: "apply", filter: result });
}

function cancel() {
    sendToParent({ type: "cancel" });
}

/* -------------------------------------------------------------
 * Arranque: pide los datos al host y engancha los eventos
 * ----------------------------------------------------------- */

function initEvents() {
    document.getElementById("btnCloseDialog").addEventListener("click", cancel);
    document.getElementById("btnCancelModal").addEventListener("click", cancel);
    document.getElementById("btnApplyModal").addEventListener("click", apply);

    document.getElementById("modalSearchInput").addEventListener("input", (e) => applySearch(e.target.value));
    document.getElementById("btnSelectAllVisible").addEventListener("click", selectAllVisible);
    document.getElementById("btnClearSelection").addEventListener("click", clearSelection);

    document.getElementById("modalRangeFrom").addEventListener("input", (e) => { rangeFrom = e.target.value; });
    document.getElementById("modalRangeTo").addEventListener("input", (e) => { rangeTo = e.target.value; });

    document.querySelectorAll("#filterModeTabs .segmented-option").forEach(btn => {
        btn.addEventListener("click", () => setMode(btn.dataset.mode));
    });
    document.querySelectorAll("#filterIncludeToggle .segmented-option").forEach(btn => {
        btn.addEventListener("click", () => setInclude(btn.dataset.include === "true"));
    });
}

Office.onReady(() => {
    log("Office.onReady. Enviando 'ready' al host…");
    initEvents();

    try {
        Office.context.ui.messageParent(JSON.stringify({ type: "ready" }));
    } catch (err) {
        log("Error enviando 'ready':", err);
    }

    Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg) => {
        log("Datos recibidos del host.");
        try {
            const payload = JSON.parse(arg.message);

            allItems = payload.items || [];
            fieldData = payload.fieldData || {};

            document.getElementById("dialogTitle").textContent = `Filtrar: ${fieldData.dim}.${fieldData.name}`;

            const rangeTab = document.querySelector('#filterModeTabs [data-mode="range"]');
            if (rangeTab) rangeTab.style.display = fieldData.isHierarchy ? "none" : "";

            preloadCurrentFilter(payload.currentFilter);

            document.getElementById("modalSearchInput").value = payload.initialSearch || "";
            document.getElementById("modalRangeFrom").value = rangeFrom;
            document.getElementById("modalRangeTo").value = rangeTo;

            renderModeUI();
            updateIncludeUI();
            applySearch(payload.initialSearch || "");
            renderChips();
            updateSelectionCount();
        } catch (err) {
            log("Error interpretando los datos recibidos:", err);
            document.getElementById("modalItemsContainer").innerHTML =
                "<div class='error'>Error cargando los valores.</div>";
        }
    });
});

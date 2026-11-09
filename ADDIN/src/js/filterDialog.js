/**
 * filterDialog.js — lógica del diálogo de Office independiente "Filtrar
 * campo" (filterDialog.html).
 * ------------------------------------------------------------------------
 * Vive en SU PROPIA ventana (igual que memberPicker.html u
 * openSemanticModel.html), lo que le da sitio de sobra para un panel
 * auxiliar a la derecha (modo, selección actual) junto a la lista de
 * valores o el rango, sin pelearse con el ancho estrecho del taskpane.
 *
 * Como cualquier diálogo de Office, esta ventana NO tiene acceso al modelo
 * de objetos de Excel ni a window.ExcelService: quien lo abre (ver
 * FilterModal.open en js/filterModal.js, ejecutado en el taskpane) ya ha
 * resuelto antes la lista de valores con ExcelService.buildFilterValuesSQL
 * + executeSQL, y nos la manda por mensaje en cuanto avisamos que estamos
 * listos. Cuando el usuario pulsa "Aceptar"/"Cancelar" devolvemos el
 * resultado de la misma forma (Office.context.ui.messageParent).
 *
 * A diferencia de versiones anteriores (un filtro = un modo Valores O
 * Rango, con un único incluir/excluir para todo el campo), aquí el
 * usuario puede combinar libremente:
 *   - Varios valores sueltos incluidos
 *   - Varios rangos incluidos
 *   - Varios valores sueltos excluidos
 *   - Varios rangos excluidos
 * Cada elemento cae en una lista u otra (incluidos/excluidos) según el
 * Incluir/Excluir activo EN EL MOMENTO de marcarlo/añadirlo — cambiar el
 * interruptor después no mueve lo ya añadido: son selecciones
 * independientes, no una sola que se invierte para todos a la vez. El
 * "Modo" (Valores/Rango) solo decide qué panel se edita ahora mismo; no
 * es excluyente con lo que ya haya en el otro panel.
 *
 * El objeto "filter" que se envía de vuelta (mode:"list"):
 *   Dimensión plana:
 *     { mode:"list", attribute,
 *       values:[...], ranges:[{from,to}, ...],
 *       excludeValues:[...], excludeRanges:[{from,to}, ...] }
 *   Jerarquía (el rango no aplica a jerarquías):
 *     { mode:"list",
 *       items:[{attribute,value}, ...], excludeItems:[{attribute,value}, ...] }
 *
 * Se conserva la LECTURA (no la escritura) de los formatos de versiones
 * anteriores del diálogo ("values", "range", "mixed") al reabrir un
 * filtro ya guardado — ver preloadCurrentFilter. commands.js debe
 * interpretar igualmente todos estos formatos (ver parseStoredFilterValue
 * / buildSimpleFilterCondition / buildHierarchyFilterCondition).
 */

let allItems = [];
let fieldData = null;
let isHierarchy = false;

let currentMode = "values";     // "values" | "range" — qué panel se edita
let currentInclude = true;      // incluir/excluir de lo que se añada AHORA

let includedValues = new Set(); // claves (attribute||value) incluidas
let excludedValues = new Set(); // claves (attribute||value) excluidas
let includedRanges = [];        // [{from,to}, ...]
let excludedRanges = [];        // [{from,to}, ...]

// Árbol (jerarquías): cada item lleva depth; aquí se le añaden _idx,
// _parent (índice del padre o -1) y _hasChildren al recibir los datos.
let expandedNodes = new Set();  // índices de nodos expandidos
let lastVisibleItems = [];      // filas pintadas ahora mismo (para "Todo")

// Qué se muestra de cada miembro: "ID" | "TEXT" | "ID_TEXT". El filtro se
// guarda SIEMPRE por ID; esto solo cambia lo que se ve (y se recuerda).
const DISPLAY_MODE_KEY = "draco_filter_display_mode";
let displayMode = "ID_TEXT";
let hasTexts = false;

function itemKey(item) {
    return item.attribute + "||" + item.value;
}

/*
 * Clave de SELECCIÓN. En un árbol, un mismo miembro puede colgar de varias
 * ramas (p. ej. la cuenta 100 bajo dos grupos distintos): con la clave por
 * valor (atributo||valor), marcar un padre marcaba también ese miembro en
 * las OTRAS ramas, y como allí su padre no estaba marcado, aparecía además
 * como selección propia en la lista de la derecha y en el filtro. En el
 * árbol se selecciona por NODO (#posición); en una lista plana, por valor.
 */
function selKey(item) {
    return hasTree() ? "#" + item._idx : itemKey(item);
}

// Valores (atributo||valor) que aparecen en más de una rama del árbol.
let duplicatedKeys = new Set();
function indexDuplicates() {
    const seen = new Set();
    duplicatedKeys = new Set();
    allItems.forEach(it => {
        const k = itemKey(it);
        if (seen.has(k)) duplicatedKeys.add(k); else seen.add(k);
    });
}

// Ruta de antepasados de un nodo, de la raíz al padre: [{attribute, value}]
function ancestorPath(item) {
    const path = [];
    for (let p = item._parent; p >= 0; p = allItems[p]._parent) {
        path.unshift({ attribute: allItems[p].attribute, value: allItems[p].value });
    }
    return path;
}

function displayLabel(item) {
    if (!item) return "";
    const id = String(item.value);
    const txt = item.label ? String(item.label) : "";
    if (!txt || displayMode === "ID") return id;
    if (displayMode === "TEXT") return txt;
    return `${id} - ${txt}`;
}

function itemMatches(item, q) {
    if (!q) return true;
    if (String(item.value).toLowerCase().includes(q)) return true;
    return !!item.label && String(item.label).toLowerCase().includes(q);
}

// Calcula padre/hijos a partir de depth (los items llegan en orden de
// recorrido en profundidad, ver buildMemberTree en filterModal.js).
function indexTree(items) {
    const stack = [];
    items.forEach((it, i) => {
        it._idx = i;
        it.depth = Number(it.depth) || 0;
        while (stack.length && items[stack[stack.length - 1]].depth >= it.depth) stack.pop();
        it._parent = stack.length ? stack[stack.length - 1] : -1;
        it._hasChildren = false;
        if (it._parent >= 0) items[it._parent]._hasChildren = true;
        stack.push(i);
    });
}

function hasTree() {
    return allItems.some(it => it._hasChildren);
}

/* -------------------------------------------------------------
 * Selección en cascada (jerarquías)
 *   - marcar un nodo marca también todos sus descendientes;
 *   - desmarcar un nodo desmarca sus descendientes y sus ANTEPASADOS
 *     (el padre deja de estar «entero»);
 *   - marcar todos los hijos NO marca el padre (decisión explícita).
 * Al aplicar se envían solo los nodos «de arriba» (los que no tienen el
 * padre marcado): el padre ya cubre a sus hijos en el WHERE, así la
 * consulta no lleva cientos de valores redundantes. Al reabrir, la
 * precarga vuelve a marcar los descendientes (cascadeLoadedSelection).
 * ----------------------------------------------------------- */
function descendantIdx(idx) {
    const out = [];
    const d = allItems[idx].depth;
    for (let i = idx + 1; i < allItems.length && allItems[i].depth > d; i++) out.push(i);
    return out;
}

function ancestorIdx(idx) {
    const out = [];
    for (let p = allItems[idx]._parent; p >= 0; p = allItems[p]._parent) out.push(p);
    return out;
}

function selectNode(idx, set, other) {
    [idx, ...descendantIdx(idx)].forEach(i => {
        const k = selKey(allItems[i]);
        set.add(k);
        other.delete(k);   // un valor no puede estar incluido Y excluido a la vez
    });
}

function deselectNode(idx, set) {
    [idx, ...descendantIdx(idx), ...ancestorIdx(idx)].forEach(i => set.delete(selKey(allItems[i])));
}

// Precarga: los nodos guardados arrastran a sus descendientes.
function cascadeLoadedSelection() {
    if (!hasTree()) return;
    [includedValues, excludedValues].forEach(set => {
        const other = set === includedValues ? excludedValues : includedValues;
        allItems.forEach((it, idx) => { if (set.has(selKey(it))) selectNode(idx, set, other); });
    });
}

// Solo los nodos marcados cuyo padre NO está marcado (en ese mismo conjunto).
function topSelectedKeys(set) {
    if (!hasTree()) return [...set];
    const out = [];
    const seen = new Set();
    allItems.forEach(it => {
        const k = selKey(it);
        if (!set.has(k) || seen.has(k)) return;
        const parentSelected = it._parent >= 0 && set.has(selKey(allItems[it._parent]));
        if (!parentSelected) { out.push(k); seen.add(k); }
    });
    // claves que no están en la lista (p. ej. de un filtro antiguo): se mantienen
    set.forEach(k => { if (!seen.has(k) && !k.startsWith("#")) out.push(k); });
    return out;
}

// Expande los antepasados de los valores ya seleccionados (al reabrir un
// filtro guardado), para que se vean sin tener que buscarlos.
function expandAncestorsOfSelected() {
    // Solo hasta los nodos «de arriba» seleccionados (sus hijos, que van en
    // cascada, no abren todo el subárbol)
    const top = new Set([...topSelectedKeys(includedValues), ...topSelectedKeys(excludedValues)]);
    allItems.forEach(it => {
        if (!top.has(selKey(it))) return;
        let p = it._parent;
        while (p >= 0) { expandedNodes.add(p); p = allItems[p]._parent; }
    });
}

function setDisplayMode(mode) {
    displayMode = mode;
    try { localStorage.setItem(DISPLAY_MODE_KEY, mode); } catch (e) { /* sin storage */ }
    updateDisplayModeUI();
    applySearch(document.getElementById("modalSearchInput").value);
    renderChips();
}

function updateDisplayModeUI() {
    const box = document.getElementById("displayModeBox");
    if (box) box.style.display = hasTexts ? "" : "none";
    document.querySelectorAll("#displayModeTabs .segmented-option").forEach(btn => {
        btn.classList.toggle("active", btn.dataset.display === displayMode);
    });
    const treeBtns = document.getElementById("treeButtons");
    if (treeBtns) treeBtns.style.display = hasTree() ? "" : "none";
}

function expandAll(expand) {
    expandedNodes = new Set();
    if (expand) allItems.forEach(it => { if (it._hasChildren) expandedNodes.add(it._idx); });
    applySearch(document.getElementById("modalSearchInput").value);
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
 * Modo (Valores/Rango): solo decide qué panel se ve y se edita; ambos
 * "acumulan" en la misma selección compartida (no se pierde lo del otro
 * panel al cambiar de modo).
 * ----------------------------------------------------------- */

function setMode(newMode) {
    if (newMode === "range" && isHierarchy) return; // el rango no aplica a jerarquías
    currentMode = newMode;
    renderModeUI();
}

function renderModeUI() {
    document.querySelectorAll("#filterModeTabs .segmented-option").forEach(btn => {
        btn.classList.toggle("active", btn.dataset.mode === currentMode);
    });
    document.getElementById("filterValuesPanel").style.display = currentMode === "values" ? "" : "none";
    document.getElementById("filterRangePanel").style.display = currentMode === "range" ? "" : "none";
}

/* -------------------------------------------------------------
 * Incluir/Excluir: un único concepto compartido por los dos paneles (los
 * dos interruptores de la página se mantienen sincronizados vía la clase
 * .js-include-toggle), que decide bajo qué lista cae lo próximo que se
 * marque o añada.
 * ----------------------------------------------------------- */

function setInclude(value) {
    currentInclude = value;
    updateIncludeUI();
    // El estado "marcado" de los checkboxes depende de qué incluir/excluir
    // está activo (ver render): al cambiar, hay que repintar la lista.
    applySearch(document.getElementById("modalSearchInput").value);
}

function updateIncludeUI() {
    document.querySelectorAll(".js-include-toggle .segmented-option").forEach(btn => {
        btn.classList.toggle("active", (btn.dataset.include === "true") === currentInclude);
    });
}

/* -------------------------------------------------------------
 * Lista de valores (checkboxes) + búsqueda
 * ----------------------------------------------------------- */

function activeValuesSet() {
    return currentInclude ? includedValues : excludedValues;
}

function otherValuesSet() {
    return currentInclude ? excludedValues : includedValues;
}

function render(items, searching) {
    const container = document.getElementById("modalItemsContainer");
    container.innerHTML = "";
    lastVisibleItems = items;

    if (items.length === 0) {
        container.innerHTML = "<div style='padding:10px;color:var(--text-muted);'>No hay valores para mostrar.</div>";
        return;
    }

    const activeSet = activeValuesSet();
    const tree = hasTree();
    const frag = document.createDocumentFragment();

    items.forEach((item) => {
        const key = selKey(item);

        const row = document.createElement("div");
        row.className = "modal-item-row";
        if (tree) row.style.paddingLeft = (6 + item.depth * 16) + "px";

        // Flecha de contraer/expandir (solo nodos con hijos). Durante una
        // búsqueda se muestran abiertos los antepasados de lo encontrado.
        if (tree) {
            const caret = document.createElement("span");
            caret.className = "tree-caret";
            if (item._hasChildren) {
                const open = searching ? true : expandedNodes.has(item._idx);
                caret.textContent = open ? "▾" : "▸";
                caret.title = open ? "Contraer" : "Expandir";
                caret.addEventListener("click", (e) => {
                    e.stopPropagation();
                    if (searching) return;
                    if (expandedNodes.has(item._idx)) expandedNodes.delete(item._idx);
                    else expandedNodes.add(item._idx);
                    applySearch(document.getElementById("modalSearchInput").value);
                });
            } else {
                caret.classList.add("leaf");
            }
            row.appendChild(caret);
        }

        const checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        // El checkbox refleja la pertenencia al conjunto ACTIVO (incluidos
        // si "Incluir" está activo, excluidos si "Excluir" está activo): un
        // valor ya marcado en el OTRO conjunto aparece sin marcar aquí, sin
        // dejar de estar seleccionado — se ve en los chips de la derecha.
        checkbox.checked = activeSet.has(key);

        const label = document.createElement("label");
        label.textContent = displayLabel(item);
        if (item.label) label.title = `${item.value} - ${item.label}`;

        row.appendChild(checkbox);
        row.appendChild(label);

        const toggle = () => {
            if (tree) {
                // Jerarquía: en cascada (ver selectNode / deselectNode) y se
                // repinta, porque cambian también otras filas.
                if (activeSet.has(key)) deselectNode(item._idx, activeSet);
                else selectNode(item._idx, activeSet, otherValuesSet());
                applySearch(document.getElementById("modalSearchInput").value);
                renderChips();
                updateSelectionCount();
                return;
            }
            if (activeSet.has(key)) {
                activeSet.delete(key);
            } else {
                activeSet.add(key);
                otherValuesSet().delete(key); // un valor no puede estar incluido Y excluido a la vez
            }
            checkbox.checked = activeSet.has(key);
            row.classList.toggle("selected", checkbox.checked);
            renderChips();
            updateSelectionCount();
        };

        checkbox.addEventListener("click", (e) => { e.stopPropagation(); toggle(); });
        row.addEventListener("click", toggle);

        if (checkbox.checked) row.classList.add("selected");
        // Marcado solo porque su padre lo está: es visual, el filtro usa el
        // padre (al aplicar solo se envían los nodos de arriba, topSelectedKeys)
        if (tree && checkbox.checked && item._parent >= 0 && activeSet.has(selKey(allItems[item._parent]))) {
            row.classList.add("inherited");
            row.title = "Incluido por su nodo padre (el filtro usa el padre)";
        }

        frag.appendChild(row);
    });

    container.appendChild(frag);
}

// Filas visibles: sin búsqueda, las que tienen todos sus antepasados
// expandidos; con búsqueda, lo que coincide (por ID o por texto) y sus
// antepasados, para no perder el contexto del árbol.
function visibleItems(query) {
    const q = String(query || "").toLowerCase().trim();
    if (!hasTree()) return allItems.filter(it => itemMatches(it, q));

    if (q === "") {
        return allItems.filter(it => {
            let p = it._parent;
            while (p >= 0) {
                if (!expandedNodes.has(p)) return false;
                p = allItems[p]._parent;
            }
            return true;
        });
    }

    const show = new Set();
    allItems.forEach(it => {
        if (!itemMatches(it, q)) return;
        show.add(it._idx);
        let p = it._parent;
        while (p >= 0 && !show.has(p)) { show.add(p); p = allItems[p]._parent; }
    });
    return allItems.filter(it => show.has(it._idx));
}

function applySearch(query) {
    const searching = String(query || "").trim() !== "";
    render(visibleItems(query), searching);
}

function selectAllVisible() {
    const q = String(document.getElementById("modalSearchInput").value || "").toLowerCase().trim();
    // Con búsqueda, solo lo que coincide (no los antepasados que se pintan
    // como contexto); sin búsqueda, lo que se ve en pantalla.
    const visible = q === "" ? lastVisibleItems : lastVisibleItems.filter(it => itemMatches(it, q));
    const activeSet = activeValuesSet();
    const other = otherValuesSet();
    visible.forEach(it => {
        if (hasTree() && typeof it._idx === "number") { selectNode(it._idx, activeSet, other); return; }
        const key = selKey(it);
        activeSet.add(key);
        other.delete(key);
    });
    applySearch(document.getElementById("modalSearchInput").value);
    renderChips();
    updateSelectionCount();
}

function clearSelection() {
    // Solo vacía la lista ACTUALMENTE activa (incluidos o excluidos): son
    // selecciones independientes, "Ninguno" no debe borrar la otra.
    activeValuesSet().clear();
    applySearch(document.getElementById("modalSearchInput").value);
    renderChips();
    updateSelectionCount();
}

/* -------------------------------------------------------------
 * Rango: "Añadir" mete el rango actual en la lista de incluidos o
 * excluidos (según el incluir/excluir activo) y limpia los campos para
 * poder cargar otro rango a continuación.
 * ----------------------------------------------------------- */

function addRange() {
    const fromInput = document.getElementById("modalRangeFrom");
    const toInput = document.getElementById("modalRangeTo");
    const from = String(fromInput.value || "").trim();
    const to = String(toInput.value || "").trim();

    if (from === "" && to === "") return;

    (currentInclude ? includedRanges : excludedRanges).push({ from, to });

    fromInput.value = "";
    toInput.value = "";

    renderChips();
    updateSelectionCount();
}

/* -------------------------------------------------------------
 * Chips con la selección actual (panel auxiliar de la derecha). Orden
 * fijo: valores incluidos, rangos incluidos, valores excluidos, rangos
 * excluidos — el mismo orden en que commands.js compone la condición SQL
 * (ver buildListFilterCondition).
 * ----------------------------------------------------------- */

function rangeLabel(r) {
    return `${r.from || "…"} – ${r.to || "…"}`;
}

function renderChips() {
    const container = document.getElementById("modalSelectedChips");
    const byKey = new Map(allItems.map(it => [selKey(it), it]));

    const chips = [];

    const removeKey = (set, key) => {
        const item = byKey.get(key);
        if (hasTree() && item && typeof item._idx === "number") deselectNode(item._idx, set);
        else set.delete(key);
        afterChipRemoved();
    };
    topSelectedKeys(includedValues).forEach(key => {
        const item = byKey.get(key);
        chips.push({
            label: item ? displayLabel(item) : key.split("||")[1],
            include: true,
            onRemove: () => removeKey(includedValues, key)
        });
    });
    includedRanges.forEach((r, idx) => {
        chips.push({
            label: rangeLabel(r),
            include: true,
            onRemove: () => { includedRanges.splice(idx, 1); afterChipRemoved(); }
        });
    });
    topSelectedKeys(excludedValues).forEach(key => {
        const item = byKey.get(key);
        chips.push({
            label: item ? displayLabel(item) : key.split("||")[1],
            include: false,
            onRemove: () => removeKey(excludedValues, key)
        });
    });
    excludedRanges.forEach((r, idx) => {
        chips.push({
            label: rangeLabel(r),
            include: false,
            onRemove: () => { excludedRanges.splice(idx, 1); afterChipRemoved(); }
        });
    });

    if (chips.length === 0) {
        container.style.display = "none";
        container.innerHTML = "";
        return;
    }

    container.style.display = "flex";
    container.innerHTML = "";

    chips.forEach(spec => {
        const chip = document.createElement("div");
        chip.className = "filter-chip" + (spec.include ? "" : " exclude-mode");
        chip.innerHTML = `<span class="chip-label"></span><button type="button" class="chip-remove">&times;</button>`;
        chip.querySelector(".chip-label").textContent = spec.label;
        chip.querySelector(".chip-remove").addEventListener("click", spec.onRemove);
        container.appendChild(chip);
    });
}

function afterChipRemoved() {
    renderChips();
    applySearch(document.getElementById("modalSearchInput").value);
    updateSelectionCount();
}

function updateSelectionCount() {
    const el = document.getElementById("modalSelectionCount");
    const total = topSelectedKeys(includedValues).length + includedRanges.length
        + topSelectedKeys(excludedValues).length + excludedRanges.length;
    if (total === 0) {
        el.textContent = "";
        return;
    }
    el.textContent = total === 1 ? "1 seleccionado" : `${total} seleccionados`;
}

/* -------------------------------------------------------------
 * Precarga de un filtro ya aplicado antes (al reabrir sobre el mismo
 * campo). Acepta el formato nuevo ("list") y, por compatibilidad, los
 * formatos de versiones anteriores del diálogo ("values", "range",
 * "mixed").
 * ----------------------------------------------------------- */

function preloadCurrentFilter(currentFilter) {
    if (!currentFilter || !currentFilter.mode) return;

    const addValues = (values, attr, include) => {
        const set = include ? includedValues : excludedValues;
        (values || []).forEach(v => set.add((attr || (fieldData && fieldData.name)) + "||" + v));
    };
    const addItems = (items, include) => {
        const set = include ? includedValues : excludedValues;
        (items || []).forEach(it => {
            const key = it.attribute + "||" + it.value;
            if (!hasTree()) { set.add(key); return; }
            const samePath = (node) => {
                const path = ancestorPath(node);
                return path.length === it.path.length && path.every((p, i) =>
                    String(p.attribute) === String(it.path[i].attribute) && String(p.value) === String(it.path[i].value));
            };
            const nodes = allItems.filter(n => itemKey(n) === key && (!Array.isArray(it.path) || samePath(n)));
            if (nodes.length) nodes.forEach(n => set.add(selKey(n)));
            else set.add(key);   // ya no está en la lista: se conserva tal cual
        });
    };
    const addRangeEntry = (from, to, include) => {
        if ((from === undefined || from === "") && (to === undefined || to === "")) return;
        (include ? includedRanges : excludedRanges).push({ from: from || "", to: to || "" });
    };

    if (currentFilter.mode === "list") {
        addValues(currentFilter.values, currentFilter.attribute, true);
        addValues(currentFilter.excludeValues, currentFilter.attribute, false);
        addItems(currentFilter.items, true);
        addItems(currentFilter.excludeItems, false);
        (currentFilter.ranges || []).forEach(r => addRangeEntry(r.from, r.to, true));
        (currentFilter.excludeRanges || []).forEach(r => addRangeEntry(r.from, r.to, false));
    } else if (currentFilter.mode === "range") {
        addRangeEntry(currentFilter.from, currentFilter.to, currentFilter.include !== false);
    } else if (currentFilter.mode === "mixed") {
        addValues(currentFilter.values, currentFilter.attribute, currentFilter.valuesInclude !== false);
        addRangeEntry(currentFilter.from, currentFilter.to, currentFilter.rangeInclude !== false);
    } else if (currentFilter.items) {
        addItems(currentFilter.items, currentFilter.include !== false);
    } else if (currentFilter.values) {
        addValues(currentFilter.values, currentFilter.attribute, currentFilter.include !== false);
    }
}

/* -------------------------------------------------------------
 * Aplicar / cancelar
 * ----------------------------------------------------------- */

function apply() {
    const hasAny = includedValues.size > 0 || excludedValues.size > 0
        || includedRanges.length > 0 || excludedRanges.length > 0;

    if (!hasAny) {
        sendToParent({ type: "cancel" });
        return;
    }

    const byKey = new Map(allItems.map(it => [selKey(it), it]));
    let result;

    if (isHierarchy) {
        // Solo los nodos «de arriba»: el padre ya incluye a sus hijos
        const toItems = (set) => topSelectedKeys(set).map(k => {
            const it = byKey.get(k);
            if (it) {
                const out = { attribute: it.attribute, value: it.value };
                // Miembro repetido en otra rama: se envía su ruta para que el
                // filtro afecte SOLO a la rama elegida (ver commands.js,
                // buildHierarchyListCondition).
                if (duplicatedKeys.has(itemKey(it)) && it._parent >= 0) out.path = ancestorPath(it);
                return out;
            }
            // Clave de un filtro antiguo que ya no está en la lista: se mantiene
            const i = k.indexOf("||");
            return i > 0 ? { attribute: k.slice(0, i), value: k.slice(i + 2) } : null;
        }).filter(Boolean);
        result = {
            mode: "list",
            items: toItems(includedValues),
            excludeItems: toItems(excludedValues)
        };
    } else {
        const toValues = (set) => [...set].map(k => byKey.get(k)).filter(Boolean).map(it => it.value);
        result = {
            mode: "list",
            attribute: fieldData.name,
            values: toValues(includedValues),
            ranges: includedRanges.slice(),
            excludeValues: toValues(excludedValues),
            excludeRanges: excludedRanges.slice()
        };
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

    document.getElementById("btnAddRange").addEventListener("click", addRange);
    const onRangeEnter = (e) => { if (e.key === "Enter") addRange(); };
    document.getElementById("modalRangeFrom").addEventListener("keydown", onRangeEnter);
    document.getElementById("modalRangeTo").addEventListener("keydown", onRangeEnter);

    document.querySelectorAll("#filterModeTabs .segmented-option").forEach(btn => {
        btn.addEventListener("click", () => setMode(btn.dataset.mode));
    });
    document.querySelectorAll(".js-include-toggle .segmented-option").forEach(btn => {
        btn.addEventListener("click", () => setInclude(btn.dataset.include === "true"));
    });
    document.querySelectorAll("#displayModeTabs .segmented-option").forEach(btn => {
        btn.addEventListener("click", () => setDisplayMode(btn.dataset.display));
    });
    const btnExpandAll = document.getElementById("btnExpandAll");
    if (btnExpandAll) btnExpandAll.addEventListener("click", () => expandAll(true));
    const btnCollapseAll = document.getElementById("btnCollapseAll");
    if (btnCollapseAll) btnCollapseAll.addEventListener("click", () => expandAll(false));
}

Office.onReady(() => {
    initEvents();

    // Primero se registra el receptor y DESPUÉS se avisa «ready» (callback de
    // addHandlerAsync). Al revés, con los valores ya en la caché de la sesión
    // la respuesta podía llegar antes de estar escuchando y perderse.
    Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg) => {
        log("Datos recibidos del host.");
        try {
            const payload = JSON.parse(arg.message);
            if (payload && payload.error) {
                // El panel no pudo cargar los valores: se explica aquí (antes
                // la ventana se cerraba sin más)
                const box = document.getElementById("modalItemsContainer");
                box.innerHTML = "";
                const div = document.createElement("div");
                div.className = "error";
                div.textContent = "No se pudieron cargar los valores: " + payload.error;
                box.appendChild(div);
                return;
            }

            allItems = payload.items || [];
            fieldData = payload.fieldData || {};
            isHierarchy = !!fieldData.isHierarchy;

            indexTree(allItems);
            indexDuplicates();
            hasTexts = allItems.some(it => it.label);
            try {
                const saved = localStorage.getItem(DISPLAY_MODE_KEY);
                if (saved === "ID" || saved === "TEXT" || saved === "ID_TEXT") displayMode = saved;
            } catch (e) { /* sin storage */ }
            expandedNodes = new Set();

            document.getElementById("dialogTitle").textContent = `Filtrar: ${fieldData.dim}.${fieldData.name}`;

            const rangeTab = document.querySelector('#filterModeTabs [data-mode="range"]');
            if (rangeTab) rangeTab.style.display = isHierarchy ? "none" : "";

            preloadCurrentFilter(payload.currentFilter);
            cascadeLoadedSelection();
            expandAncestorsOfSelected();
            updateDisplayModeUI();

            document.getElementById("modalSearchInput").value = payload.initialSearch || "";

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
    }, () => {
        log("Receptor listo. Enviando 'ready' al host…");
        try {
            Office.context.ui.messageParent(JSON.stringify({ type: "ready" }));
        } catch (err) {
            log("Error enviando 'ready':", err);
        }
    });
});

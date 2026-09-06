/**
 * WidgetPivot — panel de "Informe" del editor de widgets tipo Tabla.
 *
 * Réplica del taskpane de ADDIN/src (mismas cajas: selector de modelo
 * semántico + buscador + árbol de campos por dimensión a la izquierda;
 * zonas Filtros / Filas / Columnas como dropzones a la derecha). La única
 * diferencia deliberada es de dónde sale el modelo semántico: el add-in
 * lee modelos ya guardados (semanticModelStore); aquí se lee DIRECTAMENTE
 * el .lkml que SemanticModel.generateAndSave sube a GitHub al guardar
 * cada cubo (uno por cubo, ver js/lkml-parse.js) — así el propio LookML
 * es la fuente de verdad de dimensiones/jerarquías/medidas, ni siquiera
 * se consulta DIMENSIONES/JERARQUIAS/CAMPOS_JSON.
 *
 * Los tres botones de "ribbon" (Reconocimiento de miembros, Actualizar,
 * Añadir filtro) viven en la barra de herramientas PRINCIPAL del editor
 * (ver widget-table-editor.js), igual que en el add-in viven en la cinta
 * de Excel — no dentro de este panel, que es solo el diseño del informe.
 *
 * Motor de datos (una sola consulta, pivotado en el cliente) y pintado
 * sobre la rejilla: sin cambios respecto a la versión anterior — ver
 * buildAxisTree/buildValueMap/paint más abajo.
 */
const WidgetPivot = {
    lkmlCache: {},   // cuboId -> { factTable, measures, dimensions, joins }
    cubes: [],

    editor() { return WidgetTableEditor; },
    editorProject() { return WidgetTableEditor.project; },

    ensureReport() {
        const st = this.editor().state;
        if (!st.report) st.report = this.editor().defaultReport();
        return st.report;
    },

    // Se llama desde WidgetTableEditor.open() cada vez que se abre CUALQUIER
    // widget: limpia el resultado cacheado del widget anterior (el catálogo
    // de cubos/modelos del proyecto sí se conserva).
    onWidgetOpened() {
        this.panelEl = null;
        this.lastRowTree = null;
        this.lastColTree = null;
        this.lastValueMap = null;
        this.lastRowLevels = null;
        this.lastColLevels = null;
        this.toggleMap = {};
        this._lastPaintedW = 0;
        this._lastPaintedH = 0;
    },

    async render(panelEl) {
        this.panelEl = panelEl;
        if (!this.cubes.length) await this.loadCubes();
        this.renderPanel();
        this.syncTopButtons();

        const report = this.ensureReport();
        if (report.cuboId && (report.rowField || report.colField || report.values.length) && !this.lastRowTree) {
            this.refresh().catch(() => {});
        }
    },

    syncTopButtons() {
        const report = this.ensureReport();
        const btn = document.getElementById("wteRecognition");
        if (btn) btn.classList.toggle("active", !!report.memberRecognition);
    },

    async loadCubes() {
        const project = this.editorProject();
        const sql = `SELECT CUBO_ID, CUBOS, MODELO_YAML_PATH
                     FROM ${Provider.qualifyControl("CUBOS")}
                     WHERE PROYECTO_ID = '${Provider.esc(project.PROYECTO_ID)}'
                     ORDER BY CUBOS`;
        try {
            this.cubes = await Provider.runQuery(sql);
        } catch (err) {
            UI.toast("Error al cargar los cubos: " + err.message, "error");
            this.cubes = [];
        }
    },

    lkmlPathFor(cube) {
        if (!cube || !cube.MODELO_YAML_PATH) return null;
        return cube.MODELO_YAML_PATH.replace(/\.yaml$/i, ".lkml");
    },

    // Descarga (si hace falta) y parsea el .lkml del cubo. Lanza si el
    // cubo no tiene modelo semántico generado (solo se genera para
    // proyectos con BigQuery activo, ver SemanticModel.generateAndSave).
    async getLkmlModel(cuboId) {
        if (this.lkmlCache[cuboId]) return this.lkmlCache[cuboId];
        const cube = this.cubes.find(c => c.CUBO_ID === cuboId);
        const path = this.lkmlPathFor(cube);
        if (!path) {
            throw new Error(`El cubo "${cube ? cube.CUBOS : cuboId}" todavía no tiene un modelo semántico (.lkml) generado. Guárdalo de nuevo con BigQuery como proveedor activo.`);
        }
        if (typeof GithubRepo === "undefined") throw new Error("js/github-repo.js no está cargado.");
        if (typeof LkmlParse === "undefined") throw new Error("js/lkml-parse.js no está cargado.");

        const text = await GithubRepo.getFile(path);
        if (text === null) throw new Error(`No se encontró ${path} en el repositorio configurado.`);

        const model = LkmlParse.parse(text);
        this.lkmlCache[cuboId] = model;
        return model;
    },

    // ------------------------------------------------------------
    // Niveles de un eje: la jerarquía elegida, un único atributo plano,
    // o la clave de la dimensión si no se ha elegido nada más específico.
    // ------------------------------------------------------------
    levelsFor(dim, field) {
        if (!dim) return [];
        if (!field || !field.ref) return [{ colId: dim.keyColumn, label: dim.id }];
        if (field.kind === "hierarchy") {
            const hier = dim.hierarchies.find(h => h.name === field.ref);
            if (!hier || !hier.levels.length) return [{ colId: dim.keyColumn, label: dim.id }];
            return hier.levels.map(l => ({ colId: l.colId, label: l.label }));
        }
        // atributo plano: un único nivel
        return [{ colId: field.ref, label: field.label || field.ref }];
    },

    // ------------------------------------------------------------
    // Panel: fields-section (izquierda) + zones-section (derecha),
    // igual que taskpane.html
    // ------------------------------------------------------------
    renderPanel() {
        const report = this.ensureReport();
        const model = report.cuboId ? this.lkmlCache[report.cuboId] : null;

        this.panelEl.innerHTML = `
            <div class="taskpane-container wtp-embedded">
                <div class="taskpane-body">
                <section class="fields-section">
                    <select id="wtpCubeSelect" title="Modelo semántico" class="wtp-select">
                        <option value="">— Sin modelos semánticos —</option>
                        ${this.cubes.map(c => `<option value="${c.CUBO_ID}" ${c.CUBO_ID === report.cuboId ? "selected" : ""}>${UI.escapeHtml(c.CUBOS)}${c.MODELO_YAML_PATH ? "" : " (sin .lkml)"}</option>`).join("")}
                    </select>
                    <div class="search-input-wrapper">
                        <svg class="search-icon" width="14" height="14" viewBox="0 0 16 16" fill="currentColor">
                            <path d="M11.742 10.344a6.5 6.5 0 10-1.397 1.398h-.001c.03.04.062.078.098.115l3.85 3.85a1 1 0 001.415-1.414l-3.85-3.85a1.007 1.007 0 00-.115-.1zM12 6.5a5.5 5.5 0 11-11 0 5.5 5.5 0 0111 0z"/>
                        </svg>
                        <input type="text" id="wtpFieldSearch" placeholder="Buscar campo..." autocomplete="off">
                    </div>
                    <div id="wtpFieldsTree" class="fields-tree">
                        ${model ? this.fieldsTreeHtml(model) : `<div class="field-options-empty">Elige un modelo semántico.</div>`}
                    </div>
                </section>

                <section class="zones-section">
                    <div class="zone-card" data-zone="filters">
                        <div class="zone-header">
                            <div class="zone-title-group">
                                <svg class="zone-icon" width="14" height="14" viewBox="0 0 16 16" fill="currentColor"><path d="M1.5 1.5A.5.5 0 012 1h12a.5.5 0 01.354.854L9.5 6.707V13.5a.5.5 0 01-.757.429l-2-1.25A.5.5 0 016.5 12.25V6.707L1.146 1.854A.5.5 0 011.5 1.5z"/></svg>
                                <span>Filtros</span>
                            </div>
                        </div>
                        <div class="dropzone-content" id="wtpZoneFilters">${this.filtersChipsHtml(report)}</div>
                    </div>

                    <div class="zone-card" data-zone="rows">
                        <div class="zone-header">
                            <div class="zone-title-group">
                                <svg class="zone-icon" width="14" height="14" viewBox="0 0 16 16" fill="currentColor"><path d="M2 3h12a1 1 0 011 1v2a1 1 0 01-1 1H2a1 1 0 01-1-1V4a1 1 0 011-1zm0 6h12a1 1 0 011 1v2a1 1 0 01-1 1H2a1 1 0 01-1-1v-2a1 1 0 011-1z"/></svg>
                                <span>Filas</span>
                            </div>
                        </div>
                        <div class="dropzone-content" id="wtpZoneRows">${this.axisFieldChipHtml("rows", report)}</div>
                    </div>

                    <div class="zone-card" data-zone="cols">
                        <div class="zone-header">
                            <div class="zone-title-group">
                                <svg class="zone-icon" width="14" height="14" viewBox="0 0 16 16" fill="currentColor"><path d="M3 2h2a1 1 0 011 1v10a1 1 0 01-1 1H3a1 1 0 01-1-1V3a1 1 0 011-1zm6 0h2a1 1 0 011 1v10a1 1 0 01-1 1H9a1 1 0 01-1-1V3a1 1 0 011-1z"/></svg>
                                <span>Columnas</span>
                            </div>
                        </div>
                        <div class="dropzone-content" id="wtpZoneCols">${this.axisFieldChipHtml("cols", report)}</div>
                    </div>

                    <div class="zone-card" data-zone="values">
                        <div class="zone-header">
                            <div class="zone-title-group">
                                <span>Σ</span>
                                <span>Valores</span>
                            </div>
                        </div>
                        <div class="dropzone-content" id="wtpZoneValues">${this.valueChipsHtml(report)}</div>
                    </div>
                </section>
                </div>
            </div>`;

        document.getElementById("wtpCubeSelect").addEventListener("change", (e) => this.onSelectCube(e.target.value || null));
        document.getElementById("wtpFieldSearch").addEventListener("input", (e) => this.filterFields(e.target.value));

        this.panelEl.querySelectorAll("[data-del-filter]").forEach(btn => {
            btn.addEventListener("click", () => {
                report.filters.splice(parseInt(btn.dataset.delFilter, 10), 1);
                this.renderPanel();
                this.refresh();
            });
        });
        this.panelEl.querySelectorAll("[data-pick-filter]").forEach(btn => {
            btn.addEventListener("click", () => this.pickFilterValues(parseInt(btn.dataset.pickFilter, 10)));
        });
        this.wireFieldChipRemovals(report, model);
        this.wireDragAndDrop(model);
    },

    filterFields(query) {
        const q = (query || "").toLowerCase();
        this.panelEl.querySelectorAll(".field-item").forEach(item => {
            const label = item.querySelector(".field-label").textContent.toLowerCase();
            item.style.display = label.includes(q) ? "flex" : "none";
        });
    },

    // Árbol de campos: un grupo colapsable por dimensión (jerarquías +
    // atributos sueltos, ambos arrastrables) + un grupo "Medidas" al
    // final — igual estructura que ADDIN/src/js/taskpane.js::loadFields.
    fieldsTreeHtml(model) {
        const dimGroups = model.dimensions.map(dim => `
            <div class="dimension-group">
                <div class="dimension-header" data-group-toggle>
                    <span class="dimension-caret">▾</span><span>${UI.escapeHtml(dim.id)}</span>
                </div>
                ${dim.hierarchies.map(h => this.fieldItemHtml({ kind: "hierarchy", dimId: dim.id, ref: h.name, label: h.name }, "🗂️")).join("")}
                ${dim.attributes.map(a => this.fieldItemHtml({ kind: "attribute", dimId: dim.id, ref: a.colId, label: a.colId }, "📄")).join("")}
            </div>`).join("");

        const measureGroup = model.measures.length ? `
            <div class="dimension-group">
                <div class="dimension-header" data-group-toggle>
                    <span class="dimension-caret">▾</span><span>medidas</span>
                </div>
                ${model.measures.map(m => this.fieldItemHtml({ kind: "measure", ref: m.column, label: m.name }, "📄")).join("")}
            </div>` : "";

        return dimGroups + measureGroup;
    },

    fieldItemHtml(fieldData, icon) {
        return `<div class="field-item" draggable="true" data-field="${UI.escapeHtml(JSON.stringify(fieldData))}">
            <span class="field-icon">${icon}</span><span class="field-label">${UI.escapeHtml(fieldData.label)}</span>
        </div>`;
    },

    axisFieldChipHtml(zone, report) {
        const field = zone === "rows" ? report.rowField : report.colField;
        if (!field) return "";
        return `<span class="dropped-tag">
            <span class="dropped-tag-title">${field.kind === "hierarchy" ? "🗂️" : "📄"} ${UI.escapeHtml(field.label || field.ref)}</span>
            <span class="dropped-tag-remove" data-remove-field="${zone}">&times;</span>
        </span>`;
    },

    valueChipsHtml(report) {
        return report.values.map((v, i) => `
            <span class="dropped-tag measure-tag">
                <span class="dropped-tag-title">Σ ${UI.escapeHtml(v.name)}</span>
                <span class="dropped-tag-remove" data-remove-value="${i}">&times;</span>
            </span>`).join("");
    },

    filtersChipsHtml(report) {
        return report.filters.map((f, i) => `
            <span class="dropped-tag">
                <span class="dropped-tag-title" data-pick-filter="${i}" style="cursor:pointer;">${UI.escapeHtml(f.dimLabel)}: ${f.values.length ? f.values.length + " miembro(s)" : "todos"}</span>
                <span class="dropped-tag-remove" data-del-filter="${i}">&times;</span>
            </span>`).join("");
    },

    wireFieldChipRemovals(report) {
        this.panelEl.querySelectorAll("[data-remove-field]").forEach(btn => {
            btn.addEventListener("click", () => {
                const zone = btn.dataset.removeField;
                if (zone === "rows") { report.rowField = null; report.expandedRows = []; }
                else { report.colField = null; report.expandedCols = []; }
                this.renderPanel();
            });
        });
        this.panelEl.querySelectorAll("[data-remove-value]").forEach(btn => {
            btn.addEventListener("click", () => {
                report.values.splice(parseInt(btn.dataset.removeValue, 10), 1);
                this.renderPanel();
            });
        });
        this.panelEl.querySelectorAll("[data-pick-filter]").forEach(el => {
            el.addEventListener("click", () => this.pickFilterValues(parseInt(el.dataset.pickFilter, 10)));
        });
    },

    wireDragAndDrop(model) {
        this.panelEl.querySelectorAll(".dimension-header[data-group-toggle]").forEach(header => {
            header.addEventListener("click", () => {
                const group = header.parentElement;
                const collapsed = group.classList.toggle("collapsed");
                header.querySelector(".dimension-caret").textContent = collapsed ? "▸" : "▾";
            });
        });
        this.panelEl.querySelectorAll(".field-item").forEach(item => {
            item.addEventListener("dragstart", (e) => {
                item.classList.add("dragging");
                e.dataTransfer.effectAllowed = "move";
                e.dataTransfer.setData("text/plain", item.dataset.field);
            });
            item.addEventListener("dragend", () => item.classList.remove("dragging"));
        });
        this.panelEl.querySelectorAll(".zone-card").forEach(zoneEl => {
            zoneEl.addEventListener("dragover", (e) => { e.preventDefault(); zoneEl.classList.add("drag-over"); });
            zoneEl.addEventListener("dragleave", () => zoneEl.classList.remove("drag-over"));
            zoneEl.addEventListener("drop", async (e) => {
                e.preventDefault();
                zoneEl.classList.remove("drag-over");
                let data;
                try { data = JSON.parse(e.dataTransfer.getData("text/plain")); } catch (err) { return; }
                await this.onDropField(zoneEl.dataset.zone, data, model);
            });
        });
    },

    async onDropField(zone, data, model) {
        const report = this.ensureReport();

        if (zone === "values") {
            if (data.kind !== "measure") { UI.toast("En Valores solo se pueden soltar medidas.", "error"); return; }
            if (report.values.some(v => v.column === data.ref)) return;
            report.values.push({ name: data.label, column: data.ref });
            this.renderPanel();
            return;
        }

        if (data.kind === "measure") { UI.toast("Las medidas solo se pueden soltar en Valores.", "error"); return; }

        const dim = model.dimensions.find(d => d.id === data.dimId);
        if (!dim) return;
        const field = { dimId: data.dimId, kind: data.kind, ref: data.ref, label: data.label };

        if (zone === "filters") {
            if (report.filters.some(f => f.dimId === data.dimId && f.colId === (data.kind === "attribute" ? data.ref : dim.keyColumn))) return;
            report.filters.push({ dimId: data.dimId, dimLabel: dim.id, colId: data.kind === "attribute" ? data.ref : dim.keyColumn, values: [] });
            this.renderPanel();
            return;
        }
        if (zone === "rows") { report.rowField = field; report.expandedRows = []; }
        else if (zone === "cols") { report.colField = field; report.expandedCols = []; }
        this.renderPanel();
    },

    onSelectCube(cuboId) {
        const report = this.ensureReport();
        if (report.cuboId === cuboId) return;
        report.cuboId = cuboId;
        report.rowField = null;
        report.colField = null;
        report.values = [];
        report.filters = [];
        report.expandedRows = [];
        report.expandedCols = [];
        this.lastRowTree = null;
        this.lastColTree = null;
        this.lastValueMap = null;
        this.renderPanel();
        if (cuboId) {
            this.getLkmlModel(cuboId).then(() => this.renderPanel()).catch(err => UI.toast(err.message, "error"));
        }
    },

    toggleRecognition() {
        const report = this.ensureReport();
        report.memberRecognition = !report.memberRecognition;
        this.syncTopButtons();
        UI.toast(`Reconocimiento de miembros ${report.memberRecognition ? "activado" : "desactivado"}.`, "info");
    },

    // ------------------------------------------------------------
    // Filtros: elegir miembros reales de la dimensión (SELECT DISTINCT
    // directamente sobre su tabla física, con la ruta ya resuelta por
    // el .lkml — no hace falta Provider.qualify).
    // ------------------------------------------------------------
    async pickFilterValues(idx) {
        const report = this.ensureReport();
        const f = report.filters[idx];
        if (!f) return;
        let model;
        try { model = await this.getLkmlModel(report.cuboId); } catch (err) { UI.toast(err.message, "error"); return; }
        const dim = model.dimensions.find(d => d.id === f.dimId);
        if (!dim) return;

        let rows;
        try {
            rows = await Provider.runQuery(`SELECT DISTINCT ${f.colId} AS V FROM ${dim.table} ORDER BY V LIMIT 500`);
        } catch (err) {
            UI.toast("Error al leer los miembros: " + err.message, "error");
            return;
        }

        const chosen = await this._choiceModal({
            title: `Miembros de ${dim.id}`,
            options: rows.map(r => ({ value: String(r.V), label: String(r.V) })),
            multi: true,
            preselected: f.values
        });
        if (chosen === null) return;
        f.values = chosen;
        this.renderPanel();
        this.refresh();
    },

    // Botón "▽ Filtro" de la barra principal: atajo para añadir un filtro
    // sin arrastrar — elige la dimensión y a continuación sus miembros.
    async addFilter() {
        const report = this.ensureReport();
        if (!report.cuboId) { UI.toast("Elige un modelo semántico para el informe.", "error"); return; }
        let model;
        try { model = await this.getLkmlModel(report.cuboId); } catch (err) { UI.toast(err.message, "error"); return; }
        if (!model.dimensions.length) { UI.toast("Este modelo no tiene dimensiones.", "error"); return; }

        const dimId = await this._choiceModal({
            title: "Filtrar por dimensión",
            options: model.dimensions.map(d => ({ value: d.id, label: d.id })),
            multi: false
        });
        if (!dimId) return;
        const dim = model.dimensions.find(d => d.id === dimId);
        if (report.filters.some(f => f.dimId === dimId && f.colId === dim.keyColumn)) {
            UI.toast("Esa dimensión ya tiene un filtro.", "info");
            return;
        }
        report.filters.push({ dimId, dimLabel: dim.id, colId: dim.keyColumn, values: [] });
        this.renderPanel();
        await this.pickFilterValues(report.filters.length - 1);
    },

    // ------------------------------------------------------------
    // Ejecutar el informe
    // ------------------------------------------------------------
    async refresh() {
        const report = this.ensureReport();
        if (!report.cuboId) { UI.toast("Elige un modelo semántico para el informe.", "error"); return; }
        if (!report.values.length) { UI.toast("Añade al menos una medida en Valores.", "error"); return; }

        let model;
        try { model = await this.getLkmlModel(report.cuboId); } catch (err) { UI.toast(err.message, "error"); return; }
        if (!model.factTable) { UI.toast("El modelo semántico no tiene tabla de hechos.", "error"); return; }

        const rowDim = report.rowField ? model.dimensions.find(d => d.id === report.rowField.dimId) : null;
        const colDim = report.colField ? model.dimensions.find(d => d.id === report.colField.dimId) : null;
        const rowLevels = rowDim ? this.levelsFor(rowDim, report.rowField) : [];
        const colLevels = colDim ? this.levelsFor(colDim, report.colField) : [];

        const selectParts = [];
        const groupParts = [];
        const joinParts = [];

        if (rowDim) {
            const join = model.joins.find(j => j.dimId === rowDim.id);
            if (!join) { UI.toast(`No se encontró el join de la dimensión "${rowDim.id}" en el explore del .lkml.`, "error"); return; }
            joinParts.push(`INNER JOIN ${rowDim.table} rdim ON c.${join.fkColumn} = rdim.${join.refColumn}`);
            rowLevels.forEach((lvl, i) => { selectParts.push(`rdim.${lvl.colId} AS RL${i}`); groupParts.push(`rdim.${lvl.colId}`); });
        }
        if (colDim) {
            const join = model.joins.find(j => j.dimId === colDim.id);
            if (!join) { UI.toast(`No se encontró el join de la dimensión "${colDim.id}" en el explore del .lkml.`, "error"); return; }
            joinParts.push(`INNER JOIN ${colDim.table} cdim ON c.${join.fkColumn} = cdim.${join.refColumn}`);
            colLevels.forEach((lvl, i) => { selectParts.push(`cdim.${lvl.colId} AS CL${i}`); groupParts.push(`cdim.${lvl.colId}`); });
        }
        report.values.forEach((v, i) => selectParts.push(`SUM(c.${v.column}) AS V${i}`));

        const filterSql = this.buildFilterJoinsAndWhere(report, model, rowDim, colDim);

        const sql = `SELECT ${selectParts.join(", ")}
FROM ${model.factTable} c
${joinParts.concat(filterSql.joins).join("\n")}
${filterSql.wheres.length ? "WHERE " + filterSql.wheres.join(" AND ") : ""}
${groupParts.length ? "GROUP BY " + groupParts.join(", ") : ""}`;

        let data;
        try {
            data = await Provider.runQuery(sql);
        } catch (err) {
            UI.toast("Error al ejecutar el informe: " + err.message, "error");
            return;
        }

        this.lastRowLevels = rowLevels;
        this.lastColLevels = colLevels;

        const rowTree = this.buildAxisTree(data, rowLevels, "RL");
        const colTree = this.buildAxisTree(data, colLevels, "CL");
        const valueMap = this.buildValueMap(data, rowTree, colTree, rowLevels, colLevels, report);

        const editor = this.editor();
        const originR = editor.selection.r1, originC = editor.selection.c1;
        this.paint(rowTree, colTree, valueMap, report, originR, originC);
        UI.toast("Informe actualizado.", "success");
    },

    buildFilterJoinsAndWhere(report, model, rowDim, colDim) {
        const joins = [];
        const wheres = [];
        report.filters.forEach((f, i) => {
            if (!f.values || !f.values.length) return; // sin miembros elegidos = sin restringir
            let alias;
            if (rowDim && f.dimId === rowDim.id) alias = "rdim";
            else if (colDim && f.dimId === colDim.id) alias = "cdim";
            else {
                const dim = model.dimensions.find(d => d.id === f.dimId);
                if (!dim) return;
                const join = model.joins.find(j => j.dimId === f.dimId);
                if (!join) return;
                alias = `fdim${i}`;
                joins.push(`INNER JOIN ${dim.table} ${alias} ON c.${join.fkColumn} = ${alias}.${join.refColumn}`);
            }
            const list = f.values.map(v => `'${Provider.esc(v)}'`).join(", ");
            wheres.push(`${alias}.${f.colId} IN (${list})`);
        });
        return { joins, wheres };
    },

    // ------------------------------------------------------------
    // Árbol por eje + mapa de sumas (con roll-up a cualquier ancestro)
    // ------------------------------------------------------------
    buildAxisTree(data, levels, prefix) {
        const ROOT = "";
        const nodes = { [ROOT]: { path: ROOT, parentPath: null, depth: -1, label: "Total", children: [] } };
        data.forEach(row => {
            let parentPath = ROOT;
            for (let i = 0; i < levels.length; i++) {
                const raw = row[`${prefix}${i}`];
                const label = (raw === null || raw === undefined || raw === "") ? "(en blanco)" : String(raw);
                const path = parentPath === ROOT ? `${i}\u241F${label}` : `${parentPath}\u241E${i}\u241F${label}`;
                if (!nodes[path]) {
                    nodes[path] = { path, parentPath, depth: i, label, children: [] };
                }
                nodes[parentPath].children.push(path);
                parentPath = path;
            }
        });
        Object.keys(nodes).forEach(p => {
            const uniq = Array.from(new Set(nodes[p].children));
            uniq.sort((a, b) => nodes[a].label.localeCompare(nodes[b].label, "es"));
            nodes[p].children = uniq;
        });
        return { ROOT, nodes };
    },

    ancestorsOf(tree, path) {
        const out = [];
        let p = path;
        while (true) {
            out.push(p);
            if (p === tree.ROOT) break;
            p = tree.nodes[p].parentPath;
        }
        return out;
    },

    buildValueMap(data, rowTree, colTree, rowLevels, colLevels, report) {
        const map = {};
        const addTo = (rp, cp, vals) => {
            if (!map[rp]) map[rp] = {};
            if (!map[rp][cp]) map[rp][cp] = new Array(report.values.length).fill(0);
            for (let k = 0; k < vals.length; k++) map[rp][cp][k] += vals[k];
        };
        const pathFor = (row, levels, prefix, ROOT) => {
            let p = ROOT;
            for (let i = 0; i < levels.length; i++) {
                const raw = row[`${prefix}${i}`];
                const label = (raw === null || raw === undefined || raw === "") ? "(en blanco)" : String(raw);
                p = p === ROOT ? `${i}\u241F${label}` : `${p}\u241E${i}\u241F${label}`;
            }
            return p;
        };
        data.forEach(row => {
            const rp = pathFor(row, rowLevels, "RL", rowTree.ROOT);
            const cp = pathFor(row, colLevels, "CL", colTree.ROOT);
            const vals = report.values.map((v, i) => Number(row[`V${i}`]) || 0);
            const rowAnc = this.ancestorsOf(rowTree, rp);
            const colAnc = this.ancestorsOf(colTree, cp);
            rowAnc.forEach(ra => colAnc.forEach(ca => addTo(ra, ca, vals)));
        });
        return map;
    },

    collectVisibleRowNodes(tree, expandedSet) {
        const out = [];
        const root = tree.nodes[tree.ROOT];
        if (root.children.length === 0) { out.push(tree.ROOT); return out; }
        const walk = (path) => {
            if (path !== tree.ROOT) out.push(path);
            const node = tree.nodes[path];
            if (path === tree.ROOT || (node.children.length > 0 && expandedSet.has(path))) {
                node.children.forEach(walk);
            }
        };
        walk(tree.ROOT);
        return out;
    },

    collectVisibleLeaves(tree, expandedSet) {
        const leaves = [];
        const root = tree.nodes[tree.ROOT];
        if (root.children.length === 0) { leaves.push(tree.ROOT); return leaves; }
        const walk = (path) => {
            const node = tree.nodes[path];
            if (node.children.length === 0 || !expandedSet.has(path)) leaves.push(path);
            else node.children.forEach(walk);
        };
        root.children.forEach(walk);
        return leaves;
    },

    ancestorAtDepth(tree, leafPath, depth) {
        let p = leafPath;
        while (p !== tree.ROOT && tree.nodes[p].depth > depth) p = tree.nodes[p].parentPath;
        return p;
    },

    buildColumnHeaderRows(tree, visibleLeaves, colTreeRows) {
        const rowsOut = [];
        for (let d = 0; d < colTreeRows; d++) {
            const cells = [];
            let i = 0;
            while (i < visibleLeaves.length) {
                const anc = this.ancestorAtDepth(tree, visibleLeaves[i], d);
                let span = 1;
                while (i + span < visibleLeaves.length && this.ancestorAtDepth(tree, visibleLeaves[i + span], d) === anc) span++;
                cells.push({ path: anc, colStart: i, colSpan: span });
                i += span;
            }
            rowsOut.push(cells);
        }
        return rowsOut;
    },

    formatNumber(val) {
        if (val === null || val === undefined || isNaN(val)) return "";
        return Number(val).toLocaleString("es-ES", { maximumFractionDigits: 2 });
    },

    // ------------------------------------------------------------
    // Pintar el árbol actual (con el expand/collapse vigente) en la rejilla
    // ------------------------------------------------------------
    paint(rowTree, colTree, valueMap, report, originR, originC) {
        this.lastRowTree = rowTree;
        this.lastColTree = colTree;
        this.lastValueMap = valueMap;
        this.lastOriginR = originR;
        this.lastOriginC = originC;
        this.toggleMap = {};

        const editor = this.editor();
        const rowLevels = this.lastRowLevels || [];
        const colLevels = this.lastColLevels || [];
        const expandedRows = new Set(report.expandedRows);
        const expandedCols = new Set(report.expandedCols);

        const visibleRowPaths = this.collectVisibleRowNodes(rowTree, expandedRows);
        const visibleColLeaves = this.collectVisibleLeaves(colTree, expandedCols);

        const colTreeRows = colLevels.length;
        const needsMeasureRow = report.values.length > 1 || colTreeRows === 0;
        const colHeaderRows = colTreeRows + (needsMeasureRow ? 1 : 0);
        const perLeafCols = needsMeasureRow ? report.values.length : 1;

        const totalDataCols = visibleColLeaves.length * perLeafCols;
        const totalDataRows = visibleRowPaths.length;
        const newW = 1 + totalDataCols;
        const newH = colHeaderRows + totalDataRows + 1; // +1 = fila de Total general

        const prevW = this._lastPaintedW || 0, prevH = this._lastPaintedH || 0;
        const clearW = Math.max(prevW, newW), clearH = Math.max(prevH, newH);

        editor.clearRegion(originR, originC, originR + clearH - 1, originC + clearW - 1);
        editor.state.merges = editor.state.merges.filter(m =>
            !(m.r >= originR && m.r < originR + clearH && m.c >= originC && m.c < originC + clearW));
        this._lastPaintedW = newW;
        this._lastPaintedH = newH;

        if (colTreeRows > 0) {
            const headerRowsByDepth = this.buildColumnHeaderRows(colTree, visibleColLeaves, colTreeRows);
            headerRowsByDepth.forEach((cells, d) => {
                const sheetRow = originR + d;
                cells.forEach(cellDef => {
                    const node = colTree.nodes[cellDef.path];
                    const sheetColStart = originC + 1 + cellDef.colStart * perLeafCols;
                    const spanCols = cellDef.colSpan * perLeafCols;
                    if (spanCols > 1) {
                        editor.state.merges.push({ r: sheetRow, c: sheetColStart, rowSpan: 1, colSpan: spanCols });
                    }
                    const extra = { b: 1, bg: "#F1F5F9", al: "center" };
                    if (node.children.length > 0) {
                        extra.tg = { expanded: expandedCols.has(node.path) };
                        this.toggleMap[`${sheetRow}_${sheetColStart}`] = { axis: "col", path: node.path };
                    }
                    editor.writeCell(sheetRow, sheetColStart, node.label, extra);
                });
            });
        }
        if (needsMeasureRow) {
            const sheetRow = originR + colTreeRows;
            visibleColLeaves.forEach((leafPath, li) => {
                report.values.forEach((v, vi) => {
                    const sheetCol = originC + 1 + li * perLeafCols + vi;
                    editor.writeCell(sheetRow, sheetCol, v.name, { b: 1, bg: "#F1F5F9", al: "center", bt: 1 });
                });
            });
        }

        editor.writeCell(originR, originC, report.rowField ? report.rowField.dimId : "", { b: 1, bg: "#E4E7EC" });
        for (let d = 1; d < colHeaderRows; d++) editor.writeCell(originR + d, originC, "", { bg: "#E4E7EC" });

        const dataStartRow = originR + colHeaderRows;
        visibleRowPaths.forEach((path, ri) => {
            const node = rowTree.nodes[path];
            const sheetRow = dataStartRow + ri;
            const indent = "\u00A0\u00A0".repeat(Math.max(0, node.depth));
            const extra = { al: "left" };
            if (node.depth === 0) extra.b = 1;
            if (node.children.length > 0) {
                extra.tg = { expanded: expandedRows.has(node.path) };
                this.toggleMap[`${sheetRow}_${originC}`] = { axis: "row", path: node.path };
            }
            editor.writeCell(sheetRow, originC, indent + node.label, extra);

            visibleColLeaves.forEach((leafPath, li) => {
                for (let vi = 0; vi < perLeafCols; vi++) {
                    const measureIdx = needsMeasureRow ? vi : 0;
                    const vals = (valueMap[node.path] && valueMap[node.path][leafPath]) || null;
                    const val = vals ? vals[measureIdx] : 0;
                    editor.writeCell(sheetRow, originC + 1 + li * perLeafCols + vi, this.formatNumber(val), { al: "right" });
                }
            });
        });

        {
            const sheetRow = dataStartRow + visibleRowPaths.length;
            editor.writeCell(sheetRow, originC, "Total general", { b: 1, bt: 1, bg: "#F8F9FB" });
            visibleColLeaves.forEach((leafPath, li) => {
                for (let vi = 0; vi < perLeafCols; vi++) {
                    const measureIdx = needsMeasureRow ? vi : 0;
                    const vals = (valueMap[rowTree.ROOT] && valueMap[rowTree.ROOT][leafPath]) || null;
                    const val = vals ? vals[measureIdx] : 0;
                    editor.writeCell(sheetRow, originC + 1 + li * perLeafCols + vi, this.formatNumber(val), { al: "right", b: 1, bt: 1, bg: "#F8F9FB" });
                }
            });
        }

        editor.markDirty();
        editor.renderGrid();
    },

    toggleMember(r, c) {
        const info = this.toggleMap[`${r}_${c}`];
        if (!info || !this.lastRowTree) return;
        const report = this.ensureReport();
        const set = info.axis === "row" ? report.expandedRows : report.expandedCols;
        const idx = set.indexOf(info.path);
        if (idx >= 0) set.splice(idx, 1); else set.push(info.path);
        this.paint(this.lastRowTree, this.lastColTree, this.lastValueMap, report, this.lastOriginR, this.lastOriginC);
    },

    // ------------------------------------------------------------
    // Reconocimiento de miembros (acción explícita, nunca automática
    // mientras se teclea): al confirmar texto en la columna de cabecera
    // de filas de un informe pintado, se busca entre los miembros reales
    // de la dimensión de Filas.
    // ------------------------------------------------------------
    onCellCommitted(r, c) {
        const report = this.ensureReport();
        if (!report.memberRecognition) return;
        if (!this.lastRowTree || this.lastOriginR === undefined) return;
        if (c !== this.lastOriginC || r < this.lastOriginR) return;
        if (!report.rowField) return;

        const editor = this.editor();
        const cell = editor.getCell(r, c);
        const text = (cell.v || "").replace(/^[\s\u00A0]+/, "").trim();
        if (!text) return;
        this.recognizeMember(r, c, text);
    },

    async recognizeMember(r, c, text) {
        const report = this.ensureReport();
        let model;
        try { model = await this.getLkmlModel(report.cuboId); } catch (err) { return; }
        const dim = model.dimensions.find(d => d.id === report.rowField.dimId);
        if (!dim) return;
        const levels = this.levelsFor(dim, report.rowField);
        const lvl = levels[0];

        const sql = `SELECT DISTINCT ${lvl.colId} AS LBL FROM ${dim.table}
                     WHERE UPPER(${lvl.colId}) LIKE UPPER('%${Provider.esc(text)}%') ORDER BY LBL LIMIT 20`;
        let rows;
        try {
            rows = await Provider.runQuery(sql);
        } catch (err) {
            UI.toast("Error al reconocer el miembro: " + err.message, "error");
            return;
        }
        if (!rows.length) {
            UI.toast(`No se ha reconocido ningún miembro para "${text}".`, "error");
            return;
        }
        const editor = this.editor();
        if (rows.length === 1) {
            editor.writeCell(r, c, rows[0].LBL);
            editor.renderGrid();
            UI.toast(`Miembro reconocido: ${rows[0].LBL}`, "success");
            return;
        }
        const chosen = await this._choiceModal({
            title: `Varios miembros coinciden con "${text}"`,
            options: rows.map(x => ({ value: x.LBL, label: x.LBL })),
            multi: false
        });
        if (chosen) {
            editor.writeCell(r, c, chosen);
            editor.renderGrid();
        }
    },

    // ------------------------------------------------------------
    // Modal genérico de selección única/múltiple
    // ------------------------------------------------------------
    _choiceModal({ title, options, multi, preselected }) {
        return new Promise((resolve) => {
            let overlay = document.getElementById("wtpChoiceModal");
            if (!overlay) {
                overlay = document.createElement("div");
                overlay.className = "modal-overlay";
                overlay.id = "wtpChoiceModal";
                document.body.appendChild(overlay);
            }
            const selected = new Set(preselected || []);

            overlay.innerHTML = `
                <div class="modal-box" style="max-width:420px;">
                    <div class="modal-header">
                        <h3>${UI.escapeHtml(title)}</h3>
                        <button class="modal-close" id="wtpChoiceClose">&times;</button>
                    </div>
                    <div class="modal-body">
                        <input type="text" id="wtpChoiceSearch" placeholder="Buscar..." style="width:100%;margin-bottom:8px;">
                        <div class="wtp-choice-list" id="wtpChoiceList"></div>
                    </div>
                    <div class="modal-footer">
                        <button class="btn btn-secondary" id="wtpChoiceCancel">Cancelar</button>
                        <button class="btn btn-primary" id="wtpChoiceOk">${multi ? "Aplicar" : "Elegir"}</button>
                    </div>
                </div>`;

            const listEl = overlay.querySelector("#wtpChoiceList");
            const renderList = (filter) => {
                const f = (filter || "").toLowerCase();
                const filtered = options.filter(o => o.label.toLowerCase().includes(f));
                listEl.innerHTML = filtered.length
                    ? filtered.map(o => `
                        <label class="wtp-choice-item">
                            <input type="${multi ? "checkbox" : "radio"}" name="wtpChoice" value="${UI.escapeHtml(o.value)}" ${selected.has(o.value) ? "checked" : ""}>
                            ${UI.escapeHtml(o.label)}
                        </label>`).join("")
                    : `<div class="wtp-empty-hint">Sin resultados</div>`;
                listEl.querySelectorAll("input").forEach(inp => {
                    inp.addEventListener("change", () => {
                        if (multi) { if (inp.checked) selected.add(inp.value); else selected.delete(inp.value); }
                    });
                });
            };
            renderList("");
            overlay.querySelector("#wtpChoiceSearch").addEventListener("input", (e) => renderList(e.target.value));

            overlay.classList.add("visible");
            const cleanup = (result) => { overlay.classList.remove("visible"); resolve(result); };
            overlay.querySelector("#wtpChoiceClose").onclick = () => cleanup(null);
            overlay.querySelector("#wtpChoiceCancel").onclick = () => cleanup(null);
            overlay.querySelector("#wtpChoiceOk").onclick = () => {
                if (multi) cleanup(Array.from(selected));
                else {
                    const checked = overlay.querySelector('input[name="wtpChoice"]:checked');
                    cleanup(checked ? checked.value : null);
                }
            };
        });
    }
};

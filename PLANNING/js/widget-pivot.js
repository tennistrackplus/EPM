/**
 * WidgetPivot — panel de "Informe" del editor de widgets tipo Tabla.
 *
 * Permite atar el widget a un cubo del proyecto y diseñar un pivot simple:
 * una dimensión (opcional, con jerarquía opcional) en Filas, una en
 * Columnas, y una o varias medidas en Valores. El resultado se pinta
 * directamente sobre la rejilla del widget (WidgetTableEditor), a partir
 * de la celda seleccionada en el momento de pulsar "Actualizar".
 *
 * Diseño de datos (una sola consulta, pivotado en el cliente):
 * ------------------------------------------------------------------
 * En vez de relanzar una consulta cada vez que se expande o contrae un
 * miembro, se trae TODA la jerarquía de filas y de columnas en una única
 * consulta agrupada por todos sus niveles a la vez, y se construye en
 * JavaScript un árbol por eje (buildAxisTree) más un mapa de sumas para
 * cualquier combinación fila×columna a cualquier profundidad
 * (buildValueMap, con roll-up hacia todos los ancestros). Expandir o
 * contraer un miembro (toggleMember) solo cambia qué nodos del árbol se
 * pintan — no vuelve a consultar el motor.
 *
 * El icono de expandir/contraer (▸/▾) se pinta como un <span> propio
 * dentro de la celda (ver cellHtml en widget-table-editor.js) con su
 * propio manejador de clic — tal y como se pidió, en vez de un listener
 * de hover.
 *
 * Persistencia: todo el diseño vive en widget.CONFIG_JSON.report (ver
 * defaultReport() en widget-table-editor.js), así que viaja con el resto
 * del widget al guardar.
 */
const WidgetPivot = {
    dimMetaCache: {},
    _hierCache: {},
    cubes: [],

    editor() { return WidgetTableEditor; },
    editorProject() { return WidgetTableEditor.project; },

    ensureReport() {
        const st = this.editor().state;
        if (!st.report) st.report = this.editor().defaultReport();
        return st.report;
    },

    // Se llama desde WidgetTableEditor.open() cada vez que se abre CUALQUIER
    // widget: limpia el resultado cacheado del widget anterior (los catálogos
    // de cubos/dimensiones/jerarquías del proyecto sí se conservan).
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

        const report = this.ensureReport();
        if (report.cuboId && (report.rowField || report.colField || report.values.length) && !this.lastRowTree) {
            this.refresh().catch(() => {});
        }
    },

    async loadCubes() {
        const project = this.editorProject();
        const sql = `SELECT CUBO_ID, CUBOS, CAMPOS_JSON, TABLA
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

    parseCubeSpec(cube) {
        let spec;
        try { spec = JSON.parse(cube.CAMPOS_JSON || "{}"); } catch (e) { spec = {}; }
        const dimensions = spec.dimensions || [];
        const measures = (spec.measures || []).map(m => ({ name: m.name, type: m.type, colId: Provider.toIdentifier(m.name) }));
        return { dimensions, measures };
    },

    async getDimMeta(dimId) {
        if (this.dimMetaCache[dimId]) return this.dimMetaCache[dimId];
        const rows = await Provider.runQuery(
            `SELECT DIMENSION_ID, DIMENSION, TABLA, CAMPOS_JSON FROM ${Provider.qualifyControl("DIMENSIONES")} WHERE DIMENSION_ID = '${Provider.esc(dimId)}'`
        );
        const row = rows[0];
        if (!row) throw new Error("Dimensión no encontrada.");
        const hierRows = await Provider.runQuery(
            `SELECT JERARQUIA_ID, JERARQUIA, NIVELES_JSON FROM ${Provider.qualifyControl("JERARQUIAS")} WHERE DIMENSION_ID = '${Provider.esc(dimId)}' ORDER BY JERARQUIA`
        );
        let fields;
        try { fields = JSON.parse(row.CAMPOS_JSON || "[]"); } catch (e) { fields = []; }
        const meta = { row, keyColId: Provider.toIdentifier(row.DIMENSION), fields };
        this.dimMetaCache[dimId] = meta;
        this._hierCache[dimId] = hierRows;
        return meta;
    },

    levelsFor(meta, hierarchyId) {
        const flatLevel = [{ colId: meta.keyColId, label: meta.row.DIMENSION }];
        if (!hierarchyId) return flatLevel;
        const hiers = this._hierCache[meta.row.DIMENSION_ID] || [];
        const hier = hiers.find(h => h.JERARQUIA_ID === hierarchyId);
        if (!hier) return flatLevel;
        let levelColIds;
        try { levelColIds = JSON.parse(hier.NIVELES_JSON || "[]"); } catch (e) { levelColIds = []; }
        if (!levelColIds.length) return flatLevel;
        return levelColIds.map(colId => {
            const f = meta.fields.find(ff => Provider.toIdentifier(ff.name) === colId);
            return { colId, label: f ? f.name : colId };
        });
    },

    // ------------------------------------------------------------
    // Panel lateral: cubo, ribbon, filtros, zonas de arrastre, campos
    // ------------------------------------------------------------
    renderPanel() {
        const report = this.ensureReport();
        const cube = this.cubes.find(c => c.CUBO_ID === report.cuboId) || null;
        const spec = cube ? this.parseCubeSpec(cube) : { dimensions: [], measures: [] };

        this.panelEl.innerHTML = `
            <div class="wtp-panel">
                <div class="wtp-section">
                    <label class="wtp-label">Cubo</label>
                    <select id="wtpCubeSelect" class="wtp-select">
                        <option value="">— Elige un cubo —</option>
                        ${this.cubes.map(c => `<option value="${c.CUBO_ID}" ${c.CUBO_ID === report.cuboId ? "selected" : ""}>${UI.escapeHtml(c.CUBOS)}</option>`).join("")}
                    </select>
                </div>
                <div class="wtp-ribbon">
                    <button class="wtp-ribbon-btn${report.memberRecognition ? " active" : ""}" id="wtpRecognition" title="Al confirmar un valor tecleado en la cabecera de filas, lo busca entre los miembros reales de la dimensión">🔎 Reconocimiento</button>
                    <button class="wtp-ribbon-btn" id="wtpRefresh" title="Ejecuta la consulta y repinta el informe en la celda seleccionada">⟳ Actualizar</button>
                    <button class="wtp-ribbon-btn" id="wtpAddFilter" title="Añade un filtro por miembros de una dimensión del cubo">▽ Filtro</button>
                </div>
                ${report.filters.length ? `
                    <div class="wtp-section">
                        <label class="wtp-label">Filtros</label>
                        <div class="wtp-chip-list">
                            ${report.filters.map((f, i) => `
                                <span class="wtp-chip">${UI.escapeHtml(f.dimName)}: ${f.values.length} miembro(s)
                                    <button class="wtp-chip-x" data-del-filter="${i}" title="Quitar filtro">&times;</button>
                                </span>`).join("")}
                        </div>
                    </div>` : ""}
                <div class="wtp-dropzone" data-zone="rows">
                    <label class="wtp-label">Filas</label>
                    <div class="wtp-zone-body" id="wtpRowsZone">${this.fieldChipHtml("rows", report, spec)}</div>
                </div>
                <div class="wtp-dropzone" data-zone="cols">
                    <label class="wtp-label">Columnas</label>
                    <div class="wtp-zone-body" id="wtpColsZone">${this.fieldChipHtml("cols", report, spec)}</div>
                </div>
                <div class="wtp-dropzone" data-zone="values">
                    <label class="wtp-label">Valores</label>
                    <div class="wtp-zone-body" id="wtpValuesZone">${this.valueChipsHtml(report)}</div>
                </div>
                <div class="wtp-fieldlist">
                    <label class="wtp-label">Campos del cubo</label>
                    <div class="wtp-field-group-title">Dimensiones</div>
                    ${spec.dimensions.map(d => `<div class="wtp-field-pill" draggable="true" data-field-kind="dim" data-dim-id="${d.id}">⊞ ${UI.escapeHtml(d.name)}</div>`).join("") || `<div class="wtp-empty-hint">Sin dimensiones</div>`}
                    <div class="wtp-field-group-title">Medidas</div>
                    ${spec.measures.map(m => `<div class="wtp-field-pill" draggable="true" data-field-kind="measure" data-measure-name="${UI.escapeHtml(m.name)}">Σ ${UI.escapeHtml(m.name)}</div>`).join("") || `<div class="wtp-empty-hint">Sin medidas</div>`}
                </div>
                ${!cube ? `<div class="wtp-empty-hint">Elige un cubo para ver sus campos.</div>` : ""}
            </div>`;

        document.getElementById("wtpCubeSelect").addEventListener("change", (e) => this.onSelectCube(e.target.value || null));
        document.getElementById("wtpRecognition").addEventListener("click", () => this.toggleRecognition());
        document.getElementById("wtpRefresh").addEventListener("click", () => this.refresh());
        document.getElementById("wtpAddFilter").addEventListener("click", () => this.addFilter());

        this.panelEl.querySelectorAll("[data-del-filter]").forEach(btn => {
            btn.addEventListener("click", () => {
                report.filters.splice(parseInt(btn.dataset.delFilter, 10), 1);
                this.renderPanel();
                this.refresh();
            });
        });

        this.wireFieldChipRemovals(report);
        this.wireDragAndDrop(spec);
    },

    fieldChipHtml(zone, report, spec) {
        const field = zone === "rows" ? report.rowField : report.colField;
        if (!field) return `<div class="wtp-drop-hint">Arrastra aquí una dimensión</div>`;
        const dim = spec.dimensions.find(d => d.id === field.dimId);
        const dimName = dim ? dim.name : "(dimensión eliminada)";
        const hierOptions = this._hierCache[field.dimId] || [];
        return `
            <div class="wtp-chip wtp-chip-field">
                ${UI.escapeHtml(dimName)}
                <button class="wtp-chip-x" data-remove-field="${zone}" title="Quitar campo">&times;</button>
                ${hierOptions.length ? `
                    <select class="wtp-hier-select" data-hier-for="${zone}" title="Jerarquía a usar para expandir">
                        <option value="">(sin jerarquía)</option>
                        ${hierOptions.map(h => `<option value="${h.JERARQUIA_ID}" ${field.hierarchyId === h.JERARQUIA_ID ? "selected" : ""}>${UI.escapeHtml(h.JERARQUIA)}</option>`).join("")}
                    </select>` : ""}
            </div>`;
    },

    valueChipsHtml(report) {
        if (!report.values.length) return `<div class="wtp-drop-hint">Arrastra aquí una o varias medidas</div>`;
        return report.values.map((v, i) => `
            <span class="wtp-chip">Σ ${UI.escapeHtml(v.name)}
                <button class="wtp-chip-x" data-remove-value="${i}" title="Quitar medida">&times;</button>
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
        this.panelEl.querySelectorAll("[data-hier-for]").forEach(sel => {
            sel.addEventListener("change", () => {
                const zone = sel.dataset.hierFor;
                const field = zone === "rows" ? report.rowField : report.colField;
                if (!field) return;
                field.hierarchyId = sel.value || null;
                if (zone === "rows") report.expandedRows = []; else report.expandedCols = [];
            });
        });
    },

    wireDragAndDrop(spec) {
        this.panelEl.querySelectorAll(".wtp-field-pill").forEach(pill => {
            pill.addEventListener("dragstart", (e) => {
                e.dataTransfer.setData("text/plain", JSON.stringify({
                    kind: pill.dataset.fieldKind,
                    dimId: pill.dataset.dimId || null,
                    measureName: pill.dataset.measureName || null
                }));
            });
        });
        this.panelEl.querySelectorAll(".wtp-dropzone").forEach(zoneEl => {
            zoneEl.addEventListener("dragover", (e) => { e.preventDefault(); zoneEl.classList.add("wtp-drop-over"); });
            zoneEl.addEventListener("dragleave", () => zoneEl.classList.remove("wtp-drop-over"));
            zoneEl.addEventListener("drop", async (e) => {
                e.preventDefault();
                zoneEl.classList.remove("wtp-drop-over");
                let data;
                try { data = JSON.parse(e.dataTransfer.getData("text/plain")); } catch (err) { return; }
                await this.onDropField(zoneEl.dataset.zone, data, spec);
            });
        });
    },

    async onDropField(zone, data, spec) {
        const report = this.ensureReport();
        if (zone === "values") {
            if (data.kind !== "measure") { UI.toast("En Valores solo se pueden soltar medidas.", "error"); return; }
            if (report.values.some(v => v.name === data.measureName)) return;
            const m = spec.measures.find(x => x.name === data.measureName);
            if (!m) return;
            report.values.push({ name: m.name, colId: m.colId, agg: "SUM" });
            this.renderPanel();
            return;
        }
        if (data.kind !== "dim") { UI.toast("En Filas/Columnas solo se pueden soltar dimensiones.", "error"); return; }
        try {
            await this.getDimMeta(data.dimId);
        } catch (err) {
            UI.toast(err.message, "error");
            return;
        }
        const field = { dimId: data.dimId, hierarchyId: null };
        if (zone === "rows") { report.rowField = field; report.expandedRows = []; }
        else { report.colField = field; report.expandedCols = []; }
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
    },

    toggleRecognition() {
        const report = this.ensureReport();
        report.memberRecognition = !report.memberRecognition;
        this.renderPanel();
        UI.toast(`Reconocimiento de miembros ${report.memberRecognition ? "activado" : "desactivado"}.`, "info");
    },

    // ------------------------------------------------------------
    // Filtros
    // ------------------------------------------------------------
    async addFilter() {
        const report = this.ensureReport();
        if (!report.cuboId) { UI.toast("Elige primero un cubo.", "error"); return; }
        const cube = this.cubes.find(c => c.CUBO_ID === report.cuboId);
        const spec = this.parseCubeSpec(cube);
        if (!spec.dimensions.length) { UI.toast("Este cubo no tiene dimensiones.", "error"); return; }

        const dimId = await this._choiceModal({
            title: "Filtrar por dimensión",
            options: spec.dimensions.map(d => ({ value: d.id, label: d.name })),
            multi: false
        });
        if (!dimId) return;
        const dim = spec.dimensions.find(d => d.id === dimId);

        let meta;
        try { meta = await this.getDimMeta(dimId); } catch (err) { UI.toast(err.message, "error"); return; }

        const project = this.editorProject();
        const dimTable = Provider.qualify(project.DATASET, meta.row.TABLA);
        let rows;
        try {
            rows = await Provider.runQuery(`SELECT DISTINCT ${meta.keyColId} AS V FROM ${dimTable} ORDER BY V LIMIT 500`);
        } catch (err) {
            UI.toast("Error al leer los miembros: " + err.message, "error");
            return;
        }

        const chosen = await this._choiceModal({
            title: `Miembros de ${dim.name}`,
            options: rows.map(r => ({ value: String(r.V), label: String(r.V) })),
            multi: true
        });
        if (!chosen || !chosen.length) return;

        report.filters.push({ dimId, dimName: dim.name, colId: meta.keyColId, values: chosen });
        this.renderPanel();
        this.refresh();
    },

    async buildFilterJoinsAndWhere(report, rowMeta, colMeta, project) {
        const joins = [];
        const wheres = [];
        for (let i = 0; i < report.filters.length; i++) {
            const f = report.filters[i];
            let alias;
            if (rowMeta && f.dimId === rowMeta.row.DIMENSION_ID) {
                alias = "rdim";
            } else if (colMeta && f.dimId === colMeta.row.DIMENSION_ID) {
                alias = "cdim";
            } else {
                let meta;
                try { meta = await this.getDimMeta(f.dimId); } catch (err) { continue; }
                alias = `fdim${i}`;
                const dimTable = Provider.qualify(project.DATASET, meta.row.TABLA);
                joins.push(`INNER JOIN ${dimTable} ${alias} ON c.${meta.keyColId} = ${alias}.${meta.keyColId}`);
            }
            if (f.values && f.values.length) {
                const list = f.values.map(v => `'${Provider.esc(v)}'`).join(", ");
                wheres.push(`${alias}.${f.colId} IN (${list})`);
            }
        }
        return { joins, wheres };
    },

    // ------------------------------------------------------------
    // Ejecutar el informe
    // ------------------------------------------------------------
    async refresh() {
        const report = this.ensureReport();
        if (!report.cuboId) { UI.toast("Elige un cubo para el informe.", "error"); return; }
        if (!report.values.length) { UI.toast("Añade al menos una medida en Valores.", "error"); return; }
        const cube = this.cubes.find(c => c.CUBO_ID === report.cuboId);
        if (!cube) { UI.toast("No se encontró el cubo seleccionado.", "error"); return; }

        const project = this.editorProject();
        const cubeTable = Provider.qualify(project.DATASET, cube.TABLA);

        let rowMeta = null, rowLevels = [];
        if (report.rowField) {
            try { rowMeta = await this.getDimMeta(report.rowField.dimId); } catch (err) { UI.toast(err.message, "error"); return; }
            rowLevels = this.levelsFor(rowMeta, report.rowField.hierarchyId);
        }
        let colMeta = null, colLevels = [];
        if (report.colField) {
            try { colMeta = await this.getDimMeta(report.colField.dimId); } catch (err) { UI.toast(err.message, "error"); return; }
            colLevels = this.levelsFor(colMeta, report.colField.hierarchyId);
        }

        const selectParts = [];
        const groupParts = [];
        const joinParts = [];

        if (rowMeta) {
            const dimTable = Provider.qualify(project.DATASET, rowMeta.row.TABLA);
            joinParts.push(`INNER JOIN ${dimTable} rdim ON c.${rowMeta.keyColId} = rdim.${rowMeta.keyColId}`);
            rowLevels.forEach((lvl, i) => { selectParts.push(`rdim.${lvl.colId} AS RL${i}`); groupParts.push(`rdim.${lvl.colId}`); });
        }
        if (colMeta) {
            const dimTable = Provider.qualify(project.DATASET, colMeta.row.TABLA);
            joinParts.push(`INNER JOIN ${dimTable} cdim ON c.${colMeta.keyColId} = cdim.${colMeta.keyColId}`);
            colLevels.forEach((lvl, i) => { selectParts.push(`cdim.${lvl.colId} AS CL${i}`); groupParts.push(`cdim.${lvl.colId}`); });
        }
        report.values.forEach((v, i) => selectParts.push(`SUM(c.${v.colId}) AS V${i}`));

        const filterSql = await this.buildFilterJoinsAndWhere(report, rowMeta, colMeta, project);

        const sql = `SELECT ${selectParts.join(", ")}
FROM ${cubeTable} c
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

        // --- Cabecera de columnas (una fila por nivel de la jerarquía) ---
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

        // --- Esquina superior izquierda ---
        editor.writeCell(originR, originC, report.rowField ? (this.dimMetaCache[report.rowField.dimId] || {}).row.DIMENSION || "" : "",
            { b: 1, bg: "#E4E7EC" });
        for (let d = 1; d < colHeaderRows; d++) editor.writeCell(originR + d, originC, "", { bg: "#E4E7EC" });

        // --- Filas de datos ---
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

        // --- Total general ---
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

    // Clic en el icono ▸/▾ de una celda: alterna expandido/contraído y
    // repinta con los datos YA cacheados (no vuelve a consultar el motor).
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
    // Reconocimiento de miembros: al confirmar texto tecleado a mano en la
    // columna de cabecera de filas de un informe pintado, se busca entre
    // los miembros reales de la dimensión (aproximación web al buscador
    // de miembros del add-in — aquí siempre disparado por una acción
    // explícita, nunca de forma automática mientras se teclea).
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
        let meta;
        try { meta = await this.getDimMeta(report.rowField.dimId); } catch (err) { return; }
        const levels = this.levelsFor(meta, report.rowField.hierarchyId);
        const lvl = levels[0];
        const project = this.editorProject();
        const dimTable = Provider.qualify(project.DATASET, meta.row.TABLA);
        const sql = `SELECT DISTINCT ${lvl.colId} AS LBL FROM ${dimTable}
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
    // Modal genérico de selección única/múltiple (elegir dimensión,
    // elegir miembros de un filtro, resolver el reconocimiento)
    // ------------------------------------------------------------
    _choiceModal({ title, options, multi }) {
        return new Promise((resolve) => {
            let overlay = document.getElementById("wtpChoiceModal");
            if (!overlay) {
                overlay = document.createElement("div");
                overlay.className = "modal-overlay";
                overlay.id = "wtpChoiceModal";
                document.body.appendChild(overlay);
            }
            const selected = new Set();

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
                if (multi) {
                    cleanup(Array.from(selected));
                } else {
                    const checked = overlay.querySelector('input[name="wtpChoice"]:checked');
                    cleanup(checked ? checked.value : null);
                }
            };
        });
    }
};

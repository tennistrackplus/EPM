/**
 * ============================================================
 * DRACO PLANNING — FLUJOS DE CARGA
 * ============================================================
 * Listado de flujos (procesos) agrupados en "Automáticos" y
 * "Manuales", con alta/edición en dos pasos:
 *   1) Modal pequeño: nombre del proceso y tipo (automático/manual).
 *   2) Modal a pantalla completa con 4 bloques, en este orden:
 *      A) Cabecera: nombre/tipo, y si es automático, planificación
 *         y ejecución.
 *      B) Pantalla de entrada de variables (solo flujos manuales,
 *         contraíble): título, variables sueltas, agrupadas en
 *         "frames" y textos explicativos con formato mínimo. Cada
 *         elemento añadido se previsualiza como se vería de verdad
 *         (sin recuadro alrededor, con un input real para la
 *         variable) y se puede reordenar arrastrando, también
 *         dentro de un frame.
 *      C) Cadena de interfaces (cargas de datos) a ejecutar, en
 *         orden, sin bifurcaciones — se puede reordenar arrastrando.
 *      D) Mapeo de variables del paso seleccionado en la cadena
 *         (clic en una tarjeta de la cadena para seleccionarla):
 *         fichero (ruta/local-servidor), filtro y mapeo, asignadas
 *         por constante o arrastrando una variable de pantalla (si
 *         el flujo es manual).
 * Persistencia real en DRACO_CONTROL, repartida en 5 tablas:
 *   - FLUJOS                     cabecera (nombre, tipo, planificación, pantalla)
 *   - FLUJOS_INTERFACES          cadena de interfaces, ordenada
 *   - FLUJOS_INTERFACES_TARGETS  variables asignadas por paso (fichero/filtro/mapeo)
 *   - FLUJOS_SCREEN_BLOCKS       bloques de la pantalla de variables (var/frame/texto)
 *   - FLUJOS_SCREEN_VARIABLES    variables sueltas o dentro de un frame
 * La ejecución real corre en `python/flow_runner.py` (orquestador que
 * llama, paso a paso, a `python/interface_loader.py`), lanzada y
 * monitorizada desde `flow_run.html` (botones "Ejecutar" / "Monitor").
 * El botón de este editor abre esa pantalla en una pestaña nueva, ya
 * con el FLUJO_ID como parámetro.
 */
const Flows = {
    list: [],
    interfaces: [],
    cubes: [],
    dimensionsCache: [],
    editing: null,
    editingIsNew: true,
    collapsed: { automatico: false, manual: false },
    screenCollapsed: false,
    selectedStepId: null,
    dragChainIdx: null,
    dragBlockIdx: null,
    dragFrameVar: null,
    dragScreenVar: null,

    async render(container, project) {
        this.container = container;
        this.project = project;

        container.innerHTML = `
            <div class="module-header">
                <div>
                    <h3>Flujos de carga</h3>
                    <p>Proyecto: ${UI.escapeHtml(project.PROYECTO)} · dataset ${project.DATASET}</p>
                </div>
                <button class="btn btn-primary btn-sm" id="btnNewFlow">+ Nuevo flujo</button>
            </div>
            <div id="flowsListWrap"><span class="spinner"></span></div>
        `;

        document.getElementById("btnNewFlow").addEventListener("click", () => this.openForm());

        await this.loadInterfacesAndCubes();
        await this.loadDimensions();
        await this.loadList();
        this.renderList();
    },

    /** Dimensiones del proyecto, para el selector de validación de las variables de pantalla. */
    async loadDimensions() {
        try {
            this.dimensionsCache = await Provider.runQuery(`
                SELECT DIMENSION_ID, DIMENSION, TABLA, CAMPOS_JSON
                FROM ${Provider.qualifyControl("DIMENSIONES")}
                WHERE PROYECTO_ID = '${Provider.esc(this.project.PROYECTO_ID)}'
                ORDER BY DIMENSION`);
        } catch (err) {
            this.dimensionsCache = [];
        }
    },

    // ------------------------------------------------------------
    // Carga del listado (resumen) desde las tablas de control
    // ------------------------------------------------------------
    async loadList() {
        try {
            const rows = await Provider.runQuery(`
                SELECT FLUJO_ID, FLUJO, TIPO, SCHEDULE_JSON
                FROM ${Provider.qualifyControl("FLUJOS")}
                WHERE PROYECTO_ID = '${Provider.esc(this.project.PROYECTO_ID)}'
                ORDER BY FLUJO`);

            let chainCounts = {};
            if (rows.length) {
                const chains = await Provider.runQuery(`
                    SELECT FLUJO_ID, COUNT(*) AS N
                    FROM ${Provider.qualifyControl("FLUJOS_INTERFACES")}
                    WHERE PROYECTO_ID = '${Provider.esc(this.project.PROYECTO_ID)}'
                    GROUP BY FLUJO_ID`);
                chains.forEach(c => { chainCounts[c.FLUJO_ID] = parseInt(c.N || "0", 10); });
            }

            this.list = rows.map(r => {
                let schedule = null;
                if (r.SCHEDULE_JSON) { try { schedule = JSON.parse(r.SCHEDULE_JSON); } catch (e) { schedule = null; } }
                return {
                    id: r.FLUJO_ID,
                    name: r.FLUJO,
                    type: r.TIPO === "MANUAL" ? "manual" : "automatico",
                    schedule,
                    chainCount: chainCounts[r.FLUJO_ID] || 0
                };
            });
        } catch (err) {
            this.list = [];
            UI.toast("Error al cargar los flujos: " + err.message, "error");
        }
    },

    /** Carga el detalle completo de un flujo (para editarlo) en la forma que espera el editor. */
    async loadDetail(id) {
        const rows = await Provider.runQuery(`SELECT * FROM ${Provider.qualifyControl("FLUJOS")} WHERE FLUJO_ID = '${Provider.esc(id)}'`);
        const row = rows[0];
        if (!row) return null;

        const chainRows = await Provider.runQuery(`
            SELECT PASO_ID, INTERFAZ_ID, ORDEN FROM ${Provider.qualifyControl("FLUJOS_INTERFACES")}
            WHERE FLUJO_ID = '${Provider.esc(id)}' ORDER BY ORDEN`);
        const targetRows = await Provider.runQuery(`
            SELECT PASO_ID, GRUPO, CLAVE, TIPO, VALOR FROM ${Provider.qualifyControl("FLUJOS_INTERFACES_TARGETS")}
            WHERE FLUJO_ID = '${Provider.esc(id)}'`);
        const targetsByPaso = {};
        targetRows.forEach(t => {
            const grupo = (t.GRUPO || "").toLowerCase();
            targetsByPaso[t.PASO_ID] = targetsByPaso[t.PASO_ID] || { file: {}, filter: {}, mapping: {} };
            targetsByPaso[t.PASO_ID][grupo][t.CLAVE] = { type: t.TIPO === "VARIABLE" ? "variable" : "constante", value: t.VALOR || "" };
        });
        const chain = chainRows.map(c => ({
            id: c.PASO_ID,
            interfaceId: c.INTERFAZ_ID,
            targets: targetsByPaso[c.PASO_ID] || { file: {}, filter: {}, mapping: {} }
        }));

        const blockRows = await Provider.runQuery(`
            SELECT BLOQUE_ID, TIPO, ORDEN, TITULO, CONTENIDO FROM ${Provider.qualifyControl("FLUJOS_SCREEN_BLOCKS")}
            WHERE FLUJO_ID = '${Provider.esc(id)}' ORDER BY ORDEN`);
        const varRows = await Provider.runQuery(`
            SELECT VARIABLE_ID, BLOQUE_ID, NOMBRE, ETIQUETA, TIPO, SELECT_MODE, ORDEN FROM ${Provider.qualifyControl("FLUJOS_SCREEN_VARIABLES")}
            WHERE FLUJO_ID = '${Provider.esc(id)}' ORDER BY ORDEN`);
        const varsByBloque = {};
        varRows.forEach(v => {
            (varsByBloque[v.BLOQUE_ID] = varsByBloque[v.BLOQUE_ID] || [])
                .push({ id: v.VARIABLE_ID, name: v.NOMBRE, label: v.ETIQUETA || v.NOMBRE, type: v.TIPO || "STRING", selectMode: v.SELECT_MODE || "unico" });
        });

        const blocks = blockRows.map(b => {
            if (b.TIPO === "VARIABLE") {
                const v = (varsByBloque[b.BLOQUE_ID] || [])[0] || { id: Provider.newId(), name: "", label: "", type: "STRING", selectMode: "unico" };
                return { id: b.BLOQUE_ID, kind: "variable", variable: v };
            }
            if (b.TIPO === "TEXTO") {
                return { id: b.BLOQUE_ID, kind: "text", text: b.CONTENIDO || "" };
            }
            if (b.TIPO === "SKIP") {
                return { id: b.BLOQUE_ID, kind: "skip" };
            }
            if (b.TIPO === "ULINE") {
                return { id: b.BLOQUE_ID, kind: "line" };
            }
            return { id: b.BLOQUE_ID, kind: "frame", title: b.TITULO || "Frame", variables: varsByBloque[b.BLOQUE_ID] || [] };
        });

        let schedule = null;
        if (row.SCHEDULE_JSON) { try { schedule = JSON.parse(row.SCHEDULE_JSON); } catch (e) { schedule = null; } }

        return {
            id: row.FLUJO_ID,
            name: row.FLUJO,
            type: row.TIPO === "MANUAL" ? "manual" : "automatico",
            schedule,
            chain,
            screen: { title: row.SCREEN_TITLE || "", blocks }
        };
    },

    /** Interfaces (cargas de datos) y cubos del proyecto, para poder listarlas, leer sus campos y mapear variables. */
    async loadInterfacesAndCubes() {
        try {
            const rows = await Provider.runQuery(`
                SELECT INTERFAZ_ID, INTERFAZ, TIPO, CUBO_ID
                FROM ${Provider.qualifyControl("INTERFACES")}
                WHERE PROYECTO_ID = '${Provider.esc(this.project.PROYECTO_ID)}'
                ORDER BY INTERFAZ`);

            let fieldsByIface = {}, mappingByIface = {};
            if (rows.length) {
                const inputs = await Provider.runQuery(`
                    SELECT INTERFAZ_ID, CAMPO, TIPO, ORDEN FROM ${Provider.qualifyControl("INTERFACES_INPUT")}
                    WHERE PROYECTO_ID = '${Provider.esc(this.project.PROYECTO_ID)}' ORDER BY ORDEN`);
                inputs.forEach(i => {
                    (fieldsByIface[i.INTERFAZ_ID] = fieldsByIface[i.INTERFAZ_ID] || []).push({ name: i.CAMPO, type: i.TIPO || "STRING" });
                });

                const filters = await Provider.runQuery(`
                    SELECT INTERFAZ_ID, CAMPO, TIPO, VALOR FROM ${Provider.qualifyControl("INTERFACES_INPUT_FILTERS")}
                    WHERE PROYECTO_ID = '${Provider.esc(this.project.PROYECTO_ID)}'`);
                const filtersByIface = {};
                filters.forEach(f => {
                    filtersByIface[f.INTERFAZ_ID] = filtersByIface[f.INTERFAZ_ID] || {};
                    filtersByIface[f.INTERFAZ_ID][f.CAMPO] = { type: f.TIPO === "VARIABLE" ? "variable" : "constante", value: f.VALOR || "" };
                });
                Object.keys(fieldsByIface).forEach(ifaceId => {
                    fieldsByIface[ifaceId] = fieldsByIface[ifaceId].map(fl =>
                        ({ ...fl, filter: (filtersByIface[ifaceId] || {})[fl.name] || null }));
                });

                const mapping = await Provider.runQuery(`
                    SELECT INTERFAZ_ID, CAMPO_DESTINO, TIPO FROM ${Provider.qualifyControl("INTERFACES_MAPPING")}
                    WHERE PROYECTO_ID = '${Provider.esc(this.project.PROYECTO_ID)}'`);
                mapping.forEach(m => {
                    mappingByIface[m.INTERFAZ_ID] = mappingByIface[m.INTERFAZ_ID] || {};
                    mappingByIface[m.INTERFAZ_ID][m.CAMPO_DESTINO] = { type: (m.TIPO || "").toLowerCase() };
                });
            }

            this.interfaces = rows.map(r => ({
                id: r.INTERFAZ_ID,
                name: r.INTERFAZ,
                cuboId: r.CUBO_ID,
                originType: r.TIPO === "FICHERO" ? "fichero" : "tabla",
                origin: { fields: fieldsByIface[r.INTERFAZ_ID] || [] },
                outputMappings: mappingByIface[r.INTERFAZ_ID] || {}
            }));
        } catch (e) {
            this.interfaces = [];
        }

        try {
            const sql = `SELECT CUBO_ID, CUBOS, CAMPOS_JSON
                         FROM ${Provider.qualifyControl("CUBOS")}
                         WHERE PROYECTO_ID = '${Provider.esc(this.project.PROYECTO_ID)}'
                         ORDER BY CUBOS`;
            this.cubes = await Provider.runQuery(sql);
        } catch (err) {
            this.cubes = [];
            UI.toast("Error al cargar los cubos: " + err.message, "error");
        }
    },

    cubeNameById() {
        const map = {};
        this.cubes.forEach(c => { map[c.CUBO_ID] = c.CUBOS; });
        return map;
    },

    cuboFields(cuboId) {
        const cubo = this.cubes.find(c => c.CUBO_ID === cuboId);
        if (!cubo) return [];
        try {
            const spec = JSON.parse(cubo.CAMPOS_JSON || "{}");
            const dims = (spec.dimensions || []).map(d => ({ id: d.colId, name: d.name }));
            const meas = (spec.measures || []).map(m => ({ id: Provider.toIdentifier(m.name), name: m.name }));
            return [...dims, ...meas];
        } catch (e) {
            return [];
        }
    },

    interfaceById(id) {
        return this.interfaces.find(i => i.id === id);
    },

    // ------------------------------------------------------------
    // Listado agrupado (automáticos / manuales), contraíble
    // ------------------------------------------------------------
    renderList() {
        const wrap = document.getElementById("flowsListWrap");
        if (!wrap) return;

        const autos = this.list.filter(f => f.type === "automatico");
        const manuals = this.list.filter(f => f.type === "manual");

        wrap.innerHTML =
            this.groupHtml("automatico", "Flujos automáticos", "⟲", autos) +
            this.groupHtml("manual", "Flujos manuales", "☺", manuals);

        wrap.querySelectorAll("[data-toggle-group]").forEach(btn => btn.addEventListener("click", () => {
            const key = btn.dataset.toggleGroup;
            this.collapsed[key] = !this.collapsed[key];
            this.renderList();
        }));
        wrap.querySelectorAll("[data-edit-flow]").forEach(btn =>
            btn.addEventListener("click", () => this.openForm(btn.dataset.editFlow)));
        wrap.querySelectorAll("[data-del-flow]").forEach(btn =>
            btn.addEventListener("click", () => this.remove(btn.dataset.delFlow)));
    },

    groupHtml(key, label, icon, items) {
        const collapsed = this.collapsed[key];
        return `
            <div class="flow-group">
                <button type="button" class="flow-group-header" data-toggle-group="${key}">
                    <span class="flow-group-caret ${collapsed ? "is-collapsed" : ""}">▾</span>
                    <span class="admin-menu-icon">${icon}</span>
                    <strong>${label}</strong>
                    <span class="col-type">${items.length} flujo${items.length === 1 ? "" : "s"}</span>
                </button>
                ${!collapsed ? (items.length ? `
                    <div class="data-list">
                        <table>
                            <thead><tr><th>Flujo</th><th>Estado</th><th>Cadena</th><th></th></tr></thead>
                            <tbody>
                                ${items.map(f => `
                                    <tr>
                                        <td><strong>${UI.escapeHtml(f.name)}</strong></td>
                                        <td>${this.statusHtml(f)}</td>
                                        <td>${f.chainCount} paso${f.chainCount === 1 ? "" : "s"}</td>
                                        <td>
                                            <div class="row-actions">
                                                <button data-edit-flow="${f.id}" title="Editar">✎</button>
                                                <button data-del-flow="${f.id}" class="danger" title="Eliminar">🗑</button>
                                            </div>
                                        </td>
                                    </tr>`).join("")}
                            </tbody>
                        </table>
                    </div>` : `<div class="module-empty module-empty--inline">Todavía no hay flujos ${key === "automatico" ? "automáticos" : "manuales"}.</div>`) : ""}
            </div>`;
    },

    statusHtml(f) {
        if (f.type !== "automatico") return `<span class="table-tag">Manual</span>`;
        if (f.schedule && f.schedule.active) return `<span class="table-tag flow-status-scheduled">⏱ Planificado</span>`;
        return `<span class="table-tag">Sin planificar</span>`;
    },

    async remove(id) {
        const flow = this.list.find(f => f.id === id);
        if (!flow) return;
        const ok = await UI.confirm("Eliminar flujo", `Se eliminará el flujo <strong>${UI.escapeHtml(flow.name)}</strong> y toda su configuración (cadena, pantalla y mapeo de variables).`);
        if (!ok) return;
        try {
            for (const table of ["FLUJOS_SCREEN_VARIABLES", "FLUJOS_SCREEN_BLOCKS", "FLUJOS_INTERFACES_TARGETS", "FLUJOS_INTERFACES", "FLUJOS"]) {
                await Provider.runQuery(`DELETE FROM ${Provider.qualifyControl(table)} WHERE FLUJO_ID = '${Provider.esc(id)}'`);
            }
            await this.loadList();
            this.renderList();
            UI.toast(`Flujo "${flow.name}" eliminado.`, "success");
        } catch (err) {
            UI.toast("Error al eliminar el flujo: " + err.message, "error");
        }
    },

    // ------------------------------------------------------------
    // Datos "en blanco" de un flujo nuevo
    // ------------------------------------------------------------
    blankFlow() {
        return {
            id: Provider.newId(),
            name: "",
            type: "automatico",
            schedule: null,
            chain: [],
            screen: { title: "", blocks: [] }
        };
    },

    // ------------------------------------------------------------
    // Paso 1 (solo para flujos NUEVOS): nombre y tipo. Al editar un flujo
    // existente ya no se pasa por aquí — se entra directo al editor completo
    // (ver openForm); el nombre se renombra desde el propio título.
    // ------------------------------------------------------------
    openBasicsModal(initial, isNew) {
        return new Promise((resolve) => {
            let overlay = document.getElementById("flowBasicsModal");
            if (!overlay) {
                overlay = document.createElement("div");
                overlay.className = "modal-overlay";
                overlay.id = "flowBasicsModal";
                document.body.appendChild(overlay);
            }

            overlay.innerHTML = `
                <div class="modal-box">
                    <div class="modal-header">
                        <h3>${isNew ? "Nuevo flujo" : "Datos básicos del flujo"}</h3>
                        <button class="modal-close" id="flowBasicsClose">&times;</button>
                    </div>
                    <div class="modal-body">
                        <div class="flow-type-toggle-wrap">
                            <div class="flow-type-toggle" id="flowBasicsType">
                                <button type="button" class="flow-type-toggle-btn ${initial.type !== "manual" ? "active" : ""}" data-ftype="automatico">⟲ Automático</button>
                                <button type="button" class="flow-type-toggle-btn ${initial.type === "manual" ? "active" : ""}" data-ftype="manual">☺ Manual</button>
                            </div>
                        </div>
                        <div class="form-group">
                            <label>Nombre del proceso</label>
                            <input type="text" id="flowBasicsName" placeholder="Ej. Carga diaria ventas" value="${UI.escapeHtml(initial.name || "")}">
                        </div>
                    </div>
                    <div class="modal-footer">
                        <button class="btn btn-secondary" id="flowBasicsCancel">Cancelar</button>
                        <button class="btn btn-primary" id="flowBasicsNext">Continuar</button>
                    </div>
                </div>`;

            let type = initial.type === "manual" ? "manual" : "automatico";
            overlay.querySelectorAll("#flowBasicsType [data-ftype]").forEach(btn => {
                btn.addEventListener("click", () => {
                    type = btn.dataset.ftype;
                    overlay.querySelectorAll("#flowBasicsType [data-ftype]").forEach(b => b.classList.toggle("active", b === btn));
                });
            });

            overlay.classList.add("visible");
            const nameInput = overlay.querySelector("#flowBasicsName");
            setTimeout(() => { nameInput.focus(); nameInput.select(); }, 50);

            const cleanup = (result) => { overlay.classList.remove("visible"); resolve(result); };
            overlay.querySelector("#flowBasicsClose").onclick = () => cleanup(null);
            overlay.querySelector("#flowBasicsCancel").onclick = () => cleanup(null);
            overlay.querySelector("#flowBasicsNext").onclick = () => {
                const name = nameInput.value.trim();
                if (!name) { UI.toast("Indica un nombre para el flujo.", "error"); return; }
                cleanup({ name, type });
            };
            nameInput.onkeydown = (e) => { if (e.key === "Enter") overlay.querySelector("#flowBasicsNext").click(); };
        });
    },

    // ------------------------------------------------------------
    // Orquesta los 2 pasos de alta/edición
    // ------------------------------------------------------------
    async openForm(editId = null) {
        this.editingIsNew = !editId;
        this.screenCollapsed = false;
        this.selectedStepId = null;

        // Editar entra directo al editor completo — nombre y tipo ya no se
        // piden en un popup previo (el nombre se edita en el propio título).
        if (editId) {
            const draft = await this.loadDetail(editId);
            if (!draft) { UI.toast("No se ha podido cargar el flujo.", "error"); return; }
            this.editing = draft;
            this.openMainModal();
            return;
        }

        const draft = this.blankFlow();
        const basics = await this.openBasicsModal(draft, true);
        if (!basics) return;
        Object.assign(draft, basics);

        this.editing = draft;
        this.openMainModal();
    },

    // ------------------------------------------------------------
    // Paso 2: modal grande (4 bloques)
    // ------------------------------------------------------------
    openMainModal() {
        let overlay = document.getElementById("flowFormModal");
        if (!overlay) {
            overlay = document.createElement("div");
            overlay.className = "modal-overlay";
            overlay.id = "flowFormModal";
            document.body.appendChild(overlay);
        }
        this.overlay = overlay;

        overlay.innerHTML = `
            <div class="modal-box modal-full">
                <div class="modal-header">
                    <div class="modal-title-edit-wrap">
                        <h3 id="flowModalTitle" class="modal-title-editable" contenteditable="true" spellcheck="false" title="Clic para renombrar el flujo"></h3>
                        <span class="modal-subtitle" id="flowModalSubtitle"></span>
                    </div>
                    <div class="modal-header-right">
                        <button class="modal-close" id="flowFormClose">&times;</button>
                    </div>
                </div>
                <div class="modal-body modal-body-flush">
                    <div id="flowHeaderPart" class="flow-part flow-part--header"></div>
                    <div id="flowScreenPart" class="flow-part flow-part--screen"></div>
                    <div id="flowChainPart" class="flow-part flow-part--chain"></div>
                    <div id="flowMappingPart" class="flow-part flow-part--mapping"></div>
                </div>
                <div class="modal-footer">
                    <button class="btn btn-secondary" id="flowFormRun" style="display:none;">▶ Ejecutar / Monitor</button>
                    <span class="load-fn-toolbar-spacer"></span>
                    <button class="btn btn-secondary" id="flowFormCancel">Cancelar</button>
                    <button class="btn btn-primary" id="flowFormSave">Guardar flujo</button>
                </div>
            </div>`;

        document.getElementById("flowFormClose").addEventListener("click", () => this.closeForm());
        document.getElementById("flowFormCancel").addEventListener("click", () => this.closeForm());
        document.getElementById("flowFormSave").addEventListener("click", () => this.save());
        document.getElementById("flowFormRun").addEventListener("click", () => {
            window.open(`flow_run.html?flujo_id=${encodeURIComponent(this.editing.id)}`, "_blank");
        });
        this.updateRunButtonVisibility();

        const titleEl = document.getElementById("flowModalTitle");
        titleEl.addEventListener("keydown", (e) => {
            if (e.key === "Enter") { e.preventDefault(); titleEl.blur(); }
        });
        titleEl.addEventListener("blur", () => {
            const newName = titleEl.textContent.trim();
            if (!newName) {
                UI.toast("El flujo necesita un nombre.", "error");
                titleEl.textContent = this.editing.name;
                return;
            }
            this.editing.name = newName;
            titleEl.textContent = newName;
        });

        overlay.classList.add("visible");

        this.updateModalHeader();
        this.renderHeaderPart();
        this.renderScreenBlock();
        this.renderChainBlock();
        this.renderMappingBlock();
    },

    updateModalHeader() {
        document.getElementById("flowModalTitle").textContent = this.editing.name;
        document.getElementById("flowModalSubtitle").innerHTML =
            `Tipo: <strong>${this.editing.type === "automatico" ? "Automático" : "Manual"}</strong>`;
    },

    closeForm() {
        if (this.overlay) this.overlay.classList.remove("visible");
        this.editing = null;
    },

    // ------------------------------------------------------------
    // A) Cabecera: nombre/tipo ya está en el título del modal;
    //    aquí van planificación y ejecución (solo automáticos)
    // ------------------------------------------------------------
    renderHeaderPart() {
        const part = document.getElementById("flowHeaderPart");
        const f = this.editing;
        this.updateRunButtonVisibility();

        if (f.type !== "automatico") {
            part.innerHTML = this.editingIsNew ? `
                <div class="flow-header-row">
                    <span class="form-hint">Guarda el flujo para poder ejecutarlo.</span>
                </div>` : "";
            return;
        }

        const scheduled = !!(f.schedule && f.schedule.active);
        const statusText = scheduled ? this.scheduleSummary(f.schedule) : "Sin planificar";

        part.innerHTML = `
            <div class="flow-header-row">
                <span class="table-tag">Flujo automático</span>
                <span class="flow-schedule-status ${scheduled ? "is-active" : ""}">⏱ ${UI.escapeHtml(statusText)}</span>
                <span class="load-fn-toolbar-spacer"></span>
                <button class="btn btn-secondary btn-sm" id="btnPlanFlow">📅 Planificar</button>
            </div>`;

        document.getElementById("btnPlanFlow").addEventListener("click", async () => {
            const result = await UI.openScheduleModal({ current: f.schedule });
            if (result === null) return;
            f.schedule = result === "remove" ? null : result;
            this.renderHeaderPart();
        });
    },

    /** Muestra/oculta el botón "Ejecutar / Monitor" del footer según si el flujo ya está guardado */
    updateRunButtonVisibility() {
        const btnRun = document.getElementById("flowFormRun");
        if (btnRun) btnRun.style.display = this.editingIsNew ? "none" : "";
    },

    scheduleSummary(s) {
        if (s.repeat === "ninguna") return `una vez el ${s.startDate || "—"} a las ${s.startTime}`;
        if (s.repeat === "diaria") return `cada día a las ${s.startTime}`;
        if (s.repeat === "semanal") {
            const names = ["L", "M", "X", "J", "V", "S", "D"];
            const days = (s.weekDays || []).slice().sort().map(d => names[d]).join(", ");
            return `semanal (${days || "sin días"}) a las ${s.startTime}`;
        }
        if (s.repeat === "mensual") return `día ${s.dayOfMonth} de cada mes a las ${s.startTime}`;
        if (s.repeat === "personalizada") return `cada ${s.intervalValue} ${s.intervalUnit}`;
        return "";
    },

    // ------------------------------------------------------------
    // B) Cadena de interfaces — sin bifurcaciones, drag&drop
    // ------------------------------------------------------------
    renderChainBlock() {
        const part = document.getElementById("flowChainPart");
        const f = this.editing;
        const cubeNames = this.cubeNameById();

        const cards = f.chain.map((step, idx) => {
            const iface = this.interfaceById(step.interfaceId);
            const label = this.chainInstanceLabel(idx);
            const selected = step.id === this.selectedStepId;
            const card = `
                <div class="flow-chain-card ${selected ? "is-selected" : ""}" draggable="true" data-chain-idx="${idx}" title="Clic para ver su mapeo de variables">
                    <div class="flow-chain-card-name">${iface ? UI.escapeHtml(iface.name) : "<em>Interfaz eliminada</em>"}</div>
                    <div class="flow-chain-card-meta">${iface ? UI.escapeHtml(cubeNames[iface.cuboId] || "—") : ""} · ${UI.escapeHtml(label)}</div>
                    <button type="button" class="flow-chain-card-remove" data-remove-chain="${idx}" title="Quitar de la cadena">✕</button>
                </div>`;
            const arrow = idx < f.chain.length - 1 ? `<div class="flow-chain-arrow">→</div>` : "";
            return card + arrow;
        }).join("");

        part.innerHTML = `
            <div class="flow-part-header"><span>Cadena de interfaces</span></div>
            <div class="flow-chain-wrap" id="flowChainWrap">
                ${cards}
                <button type="button" class="flow-chain-add" id="btnAddChainStep">+ Añadir</button>
            </div>`;

        document.getElementById("btnAddChainStep").addEventListener("click", async () => {
            const picked = await UI.openInterfacePickerModal({ interfaces: this.interfaces, cubeNameById: cubeNames });
            if (!picked) return;
            const step = { id: Provider.newId(), interfaceId: picked.id, targets: { file: {}, filter: {}, mapping: {} } };
            f.chain.push(step);
            this.selectedStepId = step.id;
            this.renderChainBlock();
            this.renderMappingBlock();
        });

        part.querySelectorAll("[data-remove-chain]").forEach(btn => btn.addEventListener("click", (e) => {
            e.stopPropagation();
            const idx = parseInt(btn.dataset.removeChain, 10);
            const removedId = f.chain[idx].id;
            f.chain.splice(idx, 1);
            if (this.selectedStepId === removedId) this.selectedStepId = null;
            this.renderChainBlock();
            this.renderMappingBlock();
        }));

        part.querySelectorAll(".flow-chain-card").forEach(card => {
            card.addEventListener("click", () => {
                const idx = parseInt(card.dataset.chainIdx, 10);
                this.selectedStepId = f.chain[idx].id;
                this.renderChainBlock();
                this.renderMappingBlock();
            });
            card.addEventListener("dragstart", () => {
                this.dragChainIdx = parseInt(card.dataset.chainIdx, 10);
                card.classList.add("dragging");
            });
            card.addEventListener("dragend", () => card.classList.remove("dragging"));
            card.addEventListener("dragover", (e) => e.preventDefault());
            card.addEventListener("drop", (e) => {
                e.preventDefault();
                const targetIdx = parseInt(card.dataset.chainIdx, 10);
                if (this.dragChainIdx === null || this.dragChainIdx === targetIdx) return;
                const [moved] = f.chain.splice(this.dragChainIdx, 1);
                f.chain.splice(targetIdx, 0, moved);
                this.dragChainIdx = null;
                this.renderChainBlock();
                this.renderMappingBlock();
            });
        });
    },

    /** Etiqueta de instancia "(n)" — cuenta cuántas veces aparece esa misma interfaz hasta esta posición. */
    chainInstanceLabel(idx) {
        const f = this.editing;
        const step = f.chain[idx];
        let n = 0;
        for (let i = 0; i <= idx; i++) if (f.chain[i].interfaceId === step.interfaceId) n++;
        return `(${n})`;
    },

    // ------------------------------------------------------------
    // C) Pantalla de entrada de variables (solo flujos manuales)
    // ------------------------------------------------------------
    renderScreenBlock() {
        const part = document.getElementById("flowScreenPart");
        const f = this.editing;

        if (f.type !== "manual") {
            part.innerHTML = "";
            part.style.display = "none";
            return;
        }
        part.style.display = "";

        part.innerHTML = `
            <div class="flow-part-header flow-part-header--screen">
                <button type="button" class="flow-part-toggle" id="btnToggleScreen">
                    <span class="flow-group-caret ${this.screenCollapsed ? "is-collapsed" : ""}">▾</span>
                    <span>Pantalla de entrada de variables</span>
                </button>
                <div class="flow-screen-toolbar-mini">
                    <button type="button" class="flow-mini-btn" id="btnAddScreenVar" title="Añadir variable">+ Var</button>
                    <button type="button" class="flow-mini-btn" id="btnAddScreenFrame" title="Añadir frame">+ Frame</button>
                    <button type="button" class="flow-mini-btn" id="btnAddScreenText" title="Añadir texto">+ Texto</button>
                    <button type="button" class="flow-mini-btn" id="btnAddScreenSkip" title="Añadir espacio en blanco (tipo SKIP de ABAP)">+ Espacio</button>
                    <button type="button" class="flow-mini-btn" id="btnAddScreenLine" title="Añadir línea separadora (tipo ULINE de ABAP)">+ Línea</button>
                </div>
            </div>
            <div class="flow-screen-box ${this.screenCollapsed ? "is-collapsed" : ""}" id="flowScreenBox">
                <div class="form-group">
                    <label>Título de la pantalla</label>
                    <input type="text" id="screenTitle" placeholder="Ej. Parámetros de carga mensual" value="${UI.escapeHtml(f.screen.title || "")}">
                </div>
                <div class="flow-screen-blocks" id="flowScreenBlocks"></div>
            </div>`;

        document.getElementById("btnToggleScreen").addEventListener("click", () => {
            this.screenCollapsed = !this.screenCollapsed;
            const box = document.getElementById("flowScreenBox");
            box.classList.toggle("is-collapsed", this.screenCollapsed);
            part.querySelector(".flow-group-caret").classList.toggle("is-collapsed", this.screenCollapsed);
        });

        document.getElementById("screenTitle").addEventListener("input", (e) => { f.screen.title = e.target.value; });

        document.getElementById("btnAddScreenVar").addEventListener("click", async () => {
            const v = await UI.openScreenVariableModal({ dimensions: this.dimensionsCache });
            if (!v) return;
            f.screen.blocks.push({ id: Provider.newId(), kind: "variable", variable: { id: Provider.newId(), ...v } });
            this.renderScreenBlocksList();
            this.renderMappingBlock();
        });
        document.getElementById("btnAddScreenFrame").addEventListener("click", () => {
            f.screen.blocks.push({ id: Provider.newId(), kind: "frame", title: "Nuevo frame", variables: [] });
            this.renderScreenBlocksList();
        });
        document.getElementById("btnAddScreenText").addEventListener("click", async () => {
            const text = await UI.openScreenTextModal({ current: "" });
            if (text === null) return;
            f.screen.blocks.push({ id: Provider.newId(), kind: "text", text });
            this.renderScreenBlocksList();
        });
        document.getElementById("btnAddScreenSkip").addEventListener("click", () => {
            f.screen.blocks.push({ id: Provider.newId(), kind: "skip" });
            this.renderScreenBlocksList();
        });
        document.getElementById("btnAddScreenLine").addEventListener("click", () => {
            f.screen.blocks.push({ id: Provider.newId(), kind: "line" });
            this.renderScreenBlocksList();
        });

        this.renderScreenBlocksList();
    },

    renderScreenBlocksList() {
        const wrap = document.getElementById("flowScreenBlocks");
        const f = this.editing;
        if (!wrap) return;

        if (!f.screen.blocks.length) {
            wrap.innerHTML = `<div class="module-empty module-empty--inline">Añade variables, frames o textos para construir la pantalla.</div>`;
            return;
        }

        wrap.innerHTML = f.screen.blocks.map((b, idx) => this.screenBlockHtml(b, idx)).join("");
        this.bindScreenBlocksEvents();
    },

    screenBlockHtml(b, idx) {
        if (b.kind === "variable") {
            return `
                <div class="flow-screen-block flow-screen-block--var" draggable="true" data-block-idx="${idx}">
                    <span class="load-drag-handle">⠿</span>
                    <div class="flow-field-preview-click" data-edit-var="${idx}" title="Clic para configurar la variable">
                        ${this.fieldPreviewHtml(b.variable)}
                    </div>
                    <button type="button" class="field-remove" data-remove-block="${idx}" title="Eliminar">✕</button>
                </div>`;
        }
        if (b.kind === "text") {
            return `
                <div class="flow-screen-block flow-screen-block--text" draggable="true" data-block-idx="${idx}">
                    <span class="load-drag-handle">⠿</span>
                    <div class="flow-screen-text-content" data-edit-text="${idx}" title="Clic para editar">${UI.renderFormattedText(b.text)}</div>
                    <button type="button" class="field-remove" data-remove-block="${idx}" title="Eliminar">✕</button>
                </div>`;
        }
        if (b.kind === "skip") {
            return `
                <div class="flow-screen-block flow-screen-block--skip" draggable="true" data-block-idx="${idx}">
                    <span class="load-drag-handle">⠿</span>
                    <div class="flow-screen-skip-marker">· · · espacio en blanco · · ·</div>
                    <button type="button" class="field-remove" data-remove-block="${idx}" title="Eliminar">✕</button>
                </div>`;
        }
        if (b.kind === "line") {
            return `
                <div class="flow-screen-block flow-screen-block--line" draggable="true" data-block-idx="${idx}">
                    <span class="load-drag-handle">⠿</span>
                    <div class="flow-screen-line-marker"><hr></div>
                    <button type="button" class="field-remove" data-remove-block="${idx}" title="Eliminar">✕</button>
                </div>`;
        }
        // frame
        return `
            <div class="flow-screen-block flow-screen-block--frame" draggable="true" data-block-idx="${idx}">
                <div class="flow-frame-header">
                    <span class="load-drag-handle">⠿</span>
                    <strong data-edit-frame-title="${idx}" title="Clic para renombrar">${UI.escapeHtml(b.title || "Frame")}</strong>
                    <span class="load-fn-toolbar-spacer"></span>
                    <button type="button" class="flow-mini-btn" data-add-frame-var="${idx}">+ Var</button>
                    <button type="button" class="field-remove" data-remove-block="${idx}" title="Eliminar frame">✕</button>
                </div>
                <div class="flow-frame-vars" data-frame-idx="${idx}">
                    ${b.variables.length ? b.variables.map((v, vi) => `
                        <div class="flow-frame-var-row" draggable="true" data-frame-idx="${idx}" data-var-idx="${vi}">
                            <span class="load-drag-handle">⠿</span>
                            <div class="flow-field-preview-click" data-edit-frame-var="${idx}:${vi}" title="Clic para configurar la variable">
                                ${this.fieldPreviewHtml(v)}
                            </div>
                            <button type="button" class="field-remove" data-remove-frame-var="${idx}:${vi}" title="Eliminar">✕</button>
                        </div>`).join("") : `<div class="hierarchy-pool-empty">Sin variables en este frame.</div>`}
                </div>
            </div>`;
    },

    /** Previsualización de una variable de pantalla tal y como se vería de verdad: etiqueta + input real, sin recuadro alrededor. */
    fieldPreviewHtml(v) {
        const modeLabels = { rango: "Rango", multiple: "Varios valores", cualquiera: "Select-options" };
        const modeBadge = v.selectMode && modeLabels[v.selectMode] ? `<span class="flow-var-mode-badge">${modeLabels[v.selectMode]}</span>` : "";
        return `
            <div class="flow-field-preview">
                <label>${UI.escapeHtml(v.label || v.name)}${modeBadge}</label>
                <input type="text" disabled placeholder="${UI.escapeHtml(v.name)}">
            </div>`;
    },

    /** Chip compacto (con tipo) para la paleta de variables arrastrables del bloque de mapeo. */
    screenVarChipHtml(v) {
        return `<span class="flow-screen-var-chip"><strong>${UI.escapeHtml(v.label || v.name)}</strong><span class="load-output-type-abbr">${UI.escapeHtml(v.type)}</span></span>`;
    },

    bindScreenBlocksEvents() {
        const wrap = document.getElementById("flowScreenBlocks");
        const f = this.editing;

        wrap.querySelectorAll("[data-remove-block]").forEach(btn => btn.addEventListener("click", () => {
            const idx = parseInt(btn.dataset.removeBlock, 10);
            f.screen.blocks.splice(idx, 1);
            this.renderScreenBlocksList();
            this.renderMappingBlock();
        }));

        wrap.querySelectorAll("[data-add-frame-var]").forEach(btn => btn.addEventListener("click", async () => {
            const idx = parseInt(btn.dataset.addFrameVar, 10);
            const v = await UI.openScreenVariableModal({ dimensions: this.dimensionsCache });
            if (!v) return;
            f.screen.blocks[idx].variables.push({ id: Provider.newId(), ...v });
            this.renderScreenBlocksList();
            this.renderMappingBlock();
        }));

        wrap.querySelectorAll("[data-remove-frame-var]").forEach(btn => btn.addEventListener("click", () => {
            const [bIdx, vIdx] = btn.dataset.removeFrameVar.split(":").map(Number);
            f.screen.blocks[bIdx].variables.splice(vIdx, 1);
            this.renderScreenBlocksList();
            this.renderMappingBlock();
        }));

        wrap.querySelectorAll("[data-edit-frame-title]").forEach(el => el.addEventListener("click", async () => {
            const idx = parseInt(el.dataset.editFrameTitle, 10);
            const val = await UI.openTextPromptModal({ title: "Nombre del frame", label: "Título", value: f.screen.blocks[idx].title || "" });
            if (val === null) return;
            f.screen.blocks[idx].title = val || "Frame";
            this.renderScreenBlocksList();
        }));

        wrap.querySelectorAll("[data-edit-text]").forEach(el => el.addEventListener("click", async () => {
            const idx = parseInt(el.dataset.editText, 10);
            const val = await UI.openScreenTextModal({ current: f.screen.blocks[idx].text || "" });
            if (val === null) return;
            f.screen.blocks[idx].text = val;
            this.renderScreenBlocksList();
        }));

        wrap.querySelectorAll("[data-edit-var]").forEach(el => el.addEventListener("click", async (e) => {
            e.stopPropagation();
            const idx = parseInt(el.dataset.editVar, 10);
            const current = f.screen.blocks[idx].variable;
            const v = await UI.openScreenVariableModal({ current, dimensions: this.dimensionsCache });
            if (!v) return;
            f.screen.blocks[idx].variable = { ...current, ...v };
            this.renderScreenBlocksList();
            this.renderMappingBlock();
        }));

        wrap.querySelectorAll("[data-edit-frame-var]").forEach(el => el.addEventListener("click", async (e) => {
            e.stopPropagation();
            const [bIdx, vIdx] = el.dataset.editFrameVar.split(":").map(Number);
            const current = f.screen.blocks[bIdx].variables[vIdx];
            const v = await UI.openScreenVariableModal({ current, dimensions: this.dimensionsCache });
            if (!v) return;
            f.screen.blocks[bIdx].variables[vIdx] = { ...current, ...v };
            this.renderScreenBlocksList();
            this.renderMappingBlock();
        }));

        // Reordenar bloques de primer nivel (variable suelta / frame / texto) arrastrando.
        wrap.querySelectorAll(":scope > .flow-screen-block").forEach(block => {
            block.addEventListener("dragstart", (e) => {
                e.stopPropagation();
                this.dragBlockIdx = parseInt(block.dataset.blockIdx, 10);
                block.classList.add("dragging");
            });
            block.addEventListener("dragend", () => block.classList.remove("dragging"));
            block.addEventListener("dragover", (e) => e.preventDefault());
            block.addEventListener("drop", (e) => {
                e.preventDefault();
                e.stopPropagation();
                const targetIdx = parseInt(block.dataset.blockIdx, 10);
                if (this.dragBlockIdx === null || this.dragBlockIdx === targetIdx) return;
                const [moved] = f.screen.blocks.splice(this.dragBlockIdx, 1);
                f.screen.blocks.splice(targetIdx, 0, moved);
                this.dragBlockIdx = null;
                this.renderScreenBlocksList();
            });
        });

        // Reordenar variables dentro de un mismo frame arrastrando.
        wrap.querySelectorAll(".flow-frame-var-row").forEach(row => {
            row.addEventListener("dragstart", (e) => {
                e.stopPropagation();
                this.dragFrameVar = { frameIdx: parseInt(row.dataset.frameIdx, 10), varIdx: parseInt(row.dataset.varIdx, 10) };
                row.classList.add("dragging");
            });
            row.addEventListener("dragend", () => row.classList.remove("dragging"));
            row.addEventListener("dragover", (e) => { e.preventDefault(); e.stopPropagation(); });
            row.addEventListener("drop", (e) => {
                e.preventDefault();
                e.stopPropagation();
                if (!this.dragFrameVar) return;
                const frameIdx = parseInt(row.dataset.frameIdx, 10);
                const targetVarIdx = parseInt(row.dataset.varIdx, 10);
                if (this.dragFrameVar.frameIdx !== frameIdx) { this.dragFrameVar = null; return; }
                const vars = f.screen.blocks[frameIdx].variables;
                const [moved] = vars.splice(this.dragFrameVar.varIdx, 1);
                vars.splice(targetVarIdx, 0, moved);
                this.dragFrameVar = null;
                this.renderScreenBlocksList();
            });
        });
    },

    /** Lista plana de variables de pantalla (sueltas + las de todos los frames), para arrastrar/mapear. */
    flatScreenVariables() {
        const f = this.editing;
        if (f.type !== "manual") return [];
        const out = [];
        f.screen.blocks.forEach(b => {
            if (b.kind === "variable") out.push(b.variable);
            else if (b.kind === "frame") out.push(...b.variables);
        });
        return out;
    },

    // ------------------------------------------------------------
    // D) Mapeo: variables de cada paso de la cadena, por constante
    //    o arrastrando una variable de pantalla (si hay pantalla)
    // ------------------------------------------------------------
    renderMappingBlock() {
        const part = document.getElementById("flowMappingPart");
        const f = this.editing;
        const screenVars = this.flatScreenVariables();
        const isManual = f.type === "manual";

        // El paso seleccionado sobrevive a reordenar/renombrar; si se eliminó o no hay ninguno, se elige el primero.
        let selectedIdx = f.chain.findIndex(s => s.id === this.selectedStepId);
        if (selectedIdx < 0 && f.chain.length) { selectedIdx = 0; this.selectedStepId = f.chain[0].id; }
        if (!f.chain.length) this.selectedStepId = null;
        const selectedStep = selectedIdx >= 0 ? f.chain[selectedIdx] : null;
        const selectedIface = selectedStep ? this.interfaceById(selectedStep.interfaceId) : null;

        const leftHtml = isManual ? `
            <div class="flow-mapping-col flow-mapping-col--vars">
                <div class="load-mapping-col-header"><span>Variables de pantalla</span></div>
                <div class="flow-mapping-vars-list">
                    ${screenVars.length ? screenVars.map(v => `
                        <div class="load-input-field-row flow-screen-var-drag" draggable="true" data-var-name="${UI.escapeHtml(v.name)}">
                            <span class="load-drag-handle">⠿</span>
                            ${this.screenVarChipHtml(v)}
                        </div>`).join("") : `<div class="hierarchy-pool-empty">Añade variables en el bloque de pantalla.</div>`}
                </div>
            </div>` : "";

        const rightHtml = `
            <div class="flow-mapping-col flow-mapping-col--targets">
                <div class="load-mapping-col-header">
                    <span>${selectedStep ? `Variables de: ${UI.escapeHtml(selectedIface ? selectedIface.name : "Interfaz eliminada")} ${UI.escapeHtml(this.chainInstanceLabel(selectedIdx))}` : "Variables por paso"}</span>
                </div>
                <div class="flow-mapping-steps" id="flowMappingSteps">
                    ${selectedStep
                        ? this.stepMappingHtml(selectedStep, selectedIdx)
                        : `<div class="hierarchy-pool-empty">${f.chain.length ? "Selecciona un paso de la cadena para ver sus variables." : "Añade interfaces a la cadena para poder mapear sus variables."}</div>`}
                </div>
            </div>`;

        part.innerHTML = `
            <div class="flow-part-header">
                <span>Mapeo de variables</span>
                ${!isManual ? `<span class="form-hint">Flujo automático: solo se puede asignar por constante.</span>` : ""}
            </div>
            <div class="flow-mapping-cols ${!isManual ? "is-background-only" : ""}">
                ${leftHtml}
                ${rightHtml}
            </div>`;

        if (isManual) {
            part.querySelectorAll(".flow-screen-var-drag").forEach(row => {
                row.addEventListener("dragstart", () => {
                    this.dragScreenVar = row.dataset.varName;
                    row.classList.add("dragging");
                });
                row.addEventListener("dragend", () => row.classList.remove("dragging"));
            });
        }

        this.bindMappingTargets();
    },

    stepMappingHtml(step, idx) {
        const iface = this.interfaceById(step.interfaceId);

        if (!iface) {
            return `<p class="form-hint">Esta interfaz ya no existe; quítala de la cadena.</p>`;
        }

        step.targets = step.targets || { file: {}, filter: {}, mapping: {} };

        const isFile = iface.originType === "fichero";
        const filterFields = (iface.origin.fields || []).filter(fl => fl.filter && fl.filter.type === "variable");
        const cubeFields = this.cuboFields(iface.cuboId);
        const mappingFields = cubeFields.filter(cf => iface.outputMappings && iface.outputMappings[cf.id] && iface.outputMappings[cf.id].type === "variable");

        const groupRows = (groupKey, entries) => entries.map(entry => {
            const m = step.targets[groupKey][entry.key];
            const valueHtml = m
                ? (m.type === "constante"
                    ? `Constante: <strong>${UI.escapeHtml(m.value || "—")}</strong>`
                    : `Variable: <strong>${UI.escapeHtml(m.value || "—")}</strong>`)
                : `<span class="load-map-empty">Sin asignar</span>`;
            return `
                <div class="flow-target-row" data-step="${step.id}" data-group="${groupKey}" data-key="${UI.escapeHtml(entry.key)}">
                    <span class="flow-target-label">${UI.escapeHtml(entry.label)}</span>
                    <div class="load-map-target flow-target-drop">${valueHtml}</div>
                </div>`;
        }).join("");

        let fileHtml = "";
        if (isFile) {
            fileHtml = `
                <div class="flow-step-group">
                    <div class="flow-step-group-title">Variables de fichero</div>
                    ${groupRows("file", [{ key: "ruta_local", label: "Ruta local" }, { key: "ruta_storage", label: "Ruta storage" }])}
                </div>`;
        }

        const filterHtml = filterFields.length ? `
            <div class="flow-step-group">
                <div class="flow-step-group-title">Variables de filtro</div>
                ${groupRows("filter", filterFields.map(fl => ({ key: fl.name, label: fl.name })))}
            </div>` : "";

        const mappingHtml = mappingFields.length ? `
            <div class="flow-step-group">
                <div class="flow-step-group-title">Variables de mapeo</div>
                ${groupRows("mapping", mappingFields.map(cf => ({ key: cf.id, label: cf.name })))}
            </div>` : "";

        const nothingToMap = !isFile && !filterFields.length && !mappingFields.length;

        return `
            <div class="flow-step-detail">
                ${fileHtml}${filterHtml}${mappingHtml}
                ${nothingToMap ? `<p class="form-hint">Esta interfaz no tiene variables que asignar.</p>` : ""}
            </div>`;
    },

    bindMappingTargets() {
        const f = this.editing;
        const screenVars = this.flatScreenVariables();

        document.querySelectorAll(".flow-target-drop").forEach(zone => {
            const row = zone.closest(".flow-target-row");
            const stepId = row.dataset.step, groupKey = row.dataset.group, key = row.dataset.key;

            zone.addEventListener("dragover", (e) => e.preventDefault());
            zone.addEventListener("dragenter", () => zone.classList.add("is-drop-hover"));
            zone.addEventListener("dragleave", () => zone.classList.remove("is-drop-hover"));
            zone.addEventListener("drop", (e) => {
                e.preventDefault();
                zone.classList.remove("is-drop-hover");
                if (!this.dragScreenVar) return;
                const step = f.chain.find(s => s.id === stepId);
                step.targets[groupKey][key] = { type: "variable", value: this.dragScreenVar };
                this.dragScreenVar = null;
                this.renderMappingBlock();
            });

            zone.addEventListener("click", async () => {
                const step = f.chain.find(s => s.id === stepId);
                const current = step.targets[groupKey][key] || null;
                const result = await UI.openFlowTargetModal({
                    title: "Asignar valor",
                    targetLabel: row.querySelector(".flow-target-label").textContent,
                    screenVariables: screenVars,
                    current
                });
                if (result === null) return;
                if (result === "remove") delete step.targets[groupKey][key];
                else step.targets[groupKey][key] = result;
                this.renderMappingBlock();
            });
        });
    },

    // ------------------------------------------------------------
    // Guardado — reparte los datos en las 5 tablas de control
    // ------------------------------------------------------------
    async save() {
        const f = this.editing;
        if (!f.name) { UI.toast("El flujo necesita un nombre.", "error"); return; }

        const btn = document.getElementById("flowFormSave");
        if (btn) { btn.disabled = true; btn.textContent = "Guardando…"; }
        try {
            await this.persist();
            const name = f.name;
            this.editingIsNew = false;
            this.closeForm();
            await this.loadList();
            this.renderList();
            if (window.Draco && Draco.renderProgress) Draco.renderProgress();
            UI.toast(`Flujo "${name}" guardado.`, "success");
        } catch (err) {
            UI.toast("Error al guardar el flujo: " + err.message, "error");
        } finally {
            if (btn) { btn.disabled = false; btn.textContent = "Guardar flujo"; }
        }
    },

    async persist() {
        const f = this.editing;
        const id = f.id;
        const pid = this.project.PROYECTO_ID;
        const tipo = f.type === "manual" ? "MANUAL" : "AUTOMATICO";
        const scheduleJson = (tipo === "AUTOMATICO" && f.schedule) ? Provider.esc(JSON.stringify(f.schedule)) : "";
        const screenTitle = tipo === "MANUAL" ? (f.screen.title || "") : "";

        // 1) Cabecera --------------------------------------------------
        if (this.editingIsNew) {
            const sql = `INSERT INTO ${Provider.qualifyControl("FLUJOS")}
                (FLUJO_ID, PROYECTO_ID, FLUJO, TIPO, SCHEDULE_JSON, SCREEN_TITLE, USUARIO, FECHA_CREACION, FECHA_MODIFICACION)
                VALUES ('${Provider.esc(id)}', '${Provider.esc(pid)}', '${Provider.esc(f.name)}', '${Provider.esc(tipo)}',
                        '${scheduleJson}', '${Provider.esc(screenTitle)}', ${Provider.currentUserExpr()}, CURRENT_TIMESTAMP(), CURRENT_TIMESTAMP())`;
            await Provider.runQuery(sql);
        } else {
            const sql = `UPDATE ${Provider.qualifyControl("FLUJOS")}
                SET FLUJO = '${Provider.esc(f.name)}',
                    TIPO = '${Provider.esc(tipo)}',
                    SCHEDULE_JSON = '${scheduleJson}',
                    SCREEN_TITLE = '${Provider.esc(screenTitle)}',
                    FECHA_MODIFICACION = CURRENT_TIMESTAMP()
                WHERE FLUJO_ID = '${Provider.esc(id)}'`;
            await Provider.runQuery(sql);
        }

        // 2) Cadena de interfaces + variables asignadas por paso --------
        await Provider.runQuery(`DELETE FROM ${Provider.qualifyControl("FLUJOS_INTERFACES")} WHERE FLUJO_ID = '${Provider.esc(id)}'`);
        await Provider.runQuery(`DELETE FROM ${Provider.qualifyControl("FLUJOS_INTERFACES_TARGETS")} WHERE FLUJO_ID = '${Provider.esc(id)}'`);
        if (f.chain.length) {
            const chainVals = f.chain.map((s, idx) =>
                `('${Provider.esc(pid)}', '${Provider.esc(id)}', '${Provider.esc(s.id)}', '${Provider.esc(s.interfaceId)}', ${idx})`).join(",\n");
            await Provider.runQuery(`INSERT INTO ${Provider.qualifyControl("FLUJOS_INTERFACES")} (PROYECTO_ID, FLUJO_ID, PASO_ID, INTERFAZ_ID, ORDEN) VALUES ${chainVals}`);

            const targetRows = [];
            f.chain.forEach(s => {
                const targets = s.targets || { file: {}, filter: {}, mapping: {} };
                ["file", "filter", "mapping"].forEach(grupo => {
                    Object.entries(targets[grupo] || {}).forEach(([clave, t]) => {
                        if (!t || !t.type) return;
                        targetRows.push([s.id, grupo.toUpperCase(), clave, t.type === "variable" ? "VARIABLE" : "CONSTANTE", t.value || ""]);
                    });
                });
            });
            if (targetRows.length) {
                const vals = targetRows.map(([pasoId, grupo, clave, tipoT, valor]) =>
                    `('${Provider.esc(pid)}', '${Provider.esc(id)}', '${Provider.esc(pasoId)}', '${Provider.esc(grupo)}', '${Provider.esc(clave)}', '${Provider.esc(tipoT)}', '${Provider.esc(valor)}')`).join(",\n");
                await Provider.runQuery(`INSERT INTO ${Provider.qualifyControl("FLUJOS_INTERFACES_TARGETS")} (PROYECTO_ID, FLUJO_ID, PASO_ID, GRUPO, CLAVE, TIPO, VALOR) VALUES ${vals}`);
            }
        }

        // 3) Pantalla de variables (solo manual; se limpia siempre por si cambió de tipo)
        await Provider.runQuery(`DELETE FROM ${Provider.qualifyControl("FLUJOS_SCREEN_VARIABLES")} WHERE FLUJO_ID = '${Provider.esc(id)}'`);
        await Provider.runQuery(`DELETE FROM ${Provider.qualifyControl("FLUJOS_SCREEN_BLOCKS")} WHERE FLUJO_ID = '${Provider.esc(id)}'`);
        if (tipo === "MANUAL" && f.screen.blocks.length) {
            const tipoBMap = { variable: "VARIABLE", frame: "FRAME", text: "TEXTO", skip: "SKIP", line: "ULINE" };
            const blockVals = f.screen.blocks.map((b, idx) => {
                const tipoB = tipoBMap[b.kind] || "FRAME";
                const titulo = b.kind === "frame" ? (b.title || "") : "";
                const contenido = b.kind === "text" ? (b.text || "") : "";
                return `('${Provider.esc(pid)}', '${Provider.esc(id)}', '${Provider.esc(b.id)}', '${Provider.esc(tipoB)}', ${idx}, '${Provider.esc(titulo)}', '${Provider.esc(contenido)}')`;
            }).join(",\n");
            await Provider.runQuery(`INSERT INTO ${Provider.qualifyControl("FLUJOS_SCREEN_BLOCKS")} (PROYECTO_ID, FLUJO_ID, BLOQUE_ID, TIPO, ORDEN, TITULO, CONTENIDO) VALUES ${blockVals}`);

            const varRows = [];
            f.screen.blocks.forEach(b => {
                if (b.kind === "variable" && b.variable) {
                    varRows.push([b.variable.id || Provider.newId(), b.id, b.variable.name, b.variable.label, b.variable.type, b.variable.selectMode || "unico", 0]);
                } else if (b.kind === "frame") {
                    (b.variables || []).forEach((v, vi) => {
                        varRows.push([v.id || Provider.newId(), b.id, v.name, v.label, v.type, v.selectMode || "unico", vi]);
                    });
                }
            });
            if (varRows.length) {
                const vals = varRows.map(([varId, bloqueId, nombre, etiqueta, tipoV, selectMode, orden]) =>
                    `('${Provider.esc(pid)}', '${Provider.esc(id)}', '${Provider.esc(varId)}', '${Provider.esc(bloqueId)}', '${Provider.esc(nombre)}', '${Provider.esc(etiqueta)}', '${Provider.esc(tipoV)}', '${Provider.esc(selectMode)}', ${orden})`).join(",\n");
                await Provider.runQuery(`INSERT INTO ${Provider.qualifyControl("FLUJOS_SCREEN_VARIABLES")} (PROYECTO_ID, FLUJO_ID, VARIABLE_ID, BLOQUE_ID, NOMBRE, ETIQUETA, TIPO, SELECT_MODE, ORDEN) VALUES ${vals}`);
            }
        }
    }
};

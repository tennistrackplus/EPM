/**
 * ============================================================
 * DRACO PLANNING — WORKFLOWS (definición)
 * ============================================================
 * Un Workflow es una secuencia de Pasos. Cada paso define:
 *   - Nombre.
 *   - Inicio: al iniciar el workflow / al completar el paso anterior / fecha.
 *   - Revisión (sí/no).
 *   - Finalización: N/A, al enviarse a revisión (si revisión=true),
 *     al completarse (si revisión=false), o fecha — se recalcula sola
 *     al cambiar Revisión salvo que el usuario haya elegido N/A o Fecha.
 *   - Driver: dimensión (opcional) por la que se reparte la ejecución del
 *     paso (ej. nodo superior de la jerarquía de CECOs) + valores
 *     concretos opcionales (si no se eligen, se reparte entre TODOS los
 *     valores de la dimensión). Sin driver, el paso se asigna en bloque.
 *   - Variables del paso: variables de valor único para parametrizar
 *     distintas ejecuciones.
 *   - Tareas del paso, agrupadas en bloques: flujo manual, plantilla,
 *     función, actualización de tabla de parametrización o página HTML.
 *     Cada tarea puede completar sus propias variables por constante o
 *     por una variable del paso, y ocultarlas de la pantalla de ejecución.
 *
 * Persistencia en DRACO_CONTROL:
 *   - WORKFLOWS                        cabecera (nombre, descripción)
 *   - WORKFLOWS_PASOS                  pasos (propiedades + driver simple)
 *   - WORKFLOWS_PASOS_DRIVER_VALORES   valores concretos del driver
 *   - WORKFLOWS_PASOS_VARIABLES        variables del paso
 *   - WORKFLOWS_PASOS_BLOQUES          bloques de tareas
 *   - WORKFLOWS_PASOS_TAREAS           tareas dentro de cada bloque
 *   - WORKFLOWS_PASOS_TAREAS_VALORES   valores/variables de cada tarea
 *
 * La ejecución (crear/gestionar instancias de un workflow ya definido)
 * se aborda en una fase posterior.
 */
const Workflows = {
    TABLE: "WORKFLOWS",
    ID_COL: "WORKFLOW_ID",
    NAME_COL: "WORKFLOW",

    TASK_TYPES: {
        FLUJO_MANUAL: { label: "Flujo manual", icon: "☺" },
        PLANTILLA: { label: "Plantilla", icon: "▤" },
        FUNCION: { label: "Función", icon: "ƒ" },
        PARAMETRIZACION: { label: "Actualizar tabla de parametrización", icon: "◆" },
        HTML: { label: "Página HTML", icon: "⌗" }
    },

    list: [],
    dimensions: [],
    manualFlows: [],
    editing: null,
    editingIsNew: true,
    selectedStepId: null,
    activeTab: "propiedades",
    dragStepIdx: null,

    async render(container, project) {
        this.container = container;
        this.project = project;

        container.innerHTML = `
            <div class="module-header">
                <div>
                    <h3>Workflows</h3>
                    <p>Proyecto: ${UI.escapeHtml(project.PROYECTO)} · dataset ${project.DATASET}</p>
                </div>
                <button class="btn btn-primary btn-sm" id="btnNewWorkflow">+ Nuevo workflow</button>
            </div>
            <div id="workflowsListWrap"><span class="spinner"></span></div>
        `;

        document.getElementById("btnNewWorkflow").addEventListener("click", () => this.openForm());

        await this.loadList();
        this.renderList();
    },

    // ------------------------------------------------------------
    // Listado
    // ------------------------------------------------------------
    async loadList() {
        try {
            const rows = await Provider.runQuery(`
                SELECT ${this.ID_COL}, ${this.NAME_COL}, DESCRIPCION
                FROM ${Provider.qualifyControl(this.TABLE)}
                WHERE PROYECTO_ID = '${Provider.esc(this.project.PROYECTO_ID)}'
                ORDER BY ${this.NAME_COL}`);

            let stepCounts = {};
            if (rows.length) {
                const steps = await Provider.runQuery(`
                    SELECT WORKFLOW_ID, COUNT(*) AS N
                    FROM ${Provider.qualifyControl("WORKFLOWS_PASOS")}
                    WHERE PROYECTO_ID = '${Provider.esc(this.project.PROYECTO_ID)}'
                    GROUP BY WORKFLOW_ID`);
                steps.forEach(s => { stepCounts[s.WORKFLOW_ID] = parseInt(s.N || "0", 10); });
            }

            this.list = rows.map(r => ({
                id: r[this.ID_COL],
                name: r[this.NAME_COL],
                description: r.DESCRIPCION || "",
                stepCount: stepCounts[r[this.ID_COL]] || 0
            }));
        } catch (err) {
            this.list = [];
            UI.toast("Error al cargar los workflows: " + err.message, "error");
        }
    },

    renderList() {
        const wrap = document.getElementById("workflowsListWrap");
        if (!wrap) return;

        if (!this.list.length) {
            wrap.innerHTML = `<div class="module-empty">Todavía no hay workflows en este proyecto. Crea el primero con "Nuevo workflow".</div>`;
            return;
        }

        wrap.innerHTML = `
            <div class="data-list">
                <table>
                    <thead><tr><th>Workflow</th><th>Descripción</th><th>Pasos</th><th></th></tr></thead>
                    <tbody>
                        ${this.list.map(w => `
                            <tr>
                                <td><strong>${UI.escapeHtml(w.name)}</strong></td>
                                <td>${UI.escapeHtml(w.description || "—")}</td>
                                <td>${w.stepCount} paso${w.stepCount === 1 ? "" : "s"}</td>
                                <td>
                                    <div class="row-actions">
                                        <button data-runs-wf="${w.id}" title="Ejecuciones">▶</button>
                                        <button data-edit-wf="${w.id}" title="Editar">✎</button>
                                        <button data-del-wf="${w.id}" class="danger" title="Eliminar">🗑</button>
                                    </div>
                                </td>
                            </tr>`).join("")}
                    </tbody>
                </table>
            </div>`;

        wrap.querySelectorAll("[data-edit-wf]").forEach(btn =>
            btn.addEventListener("click", () => this.openForm(btn.dataset.editWf)));
        wrap.querySelectorAll("[data-del-wf]").forEach(btn =>
            btn.addEventListener("click", () => this.remove(btn.dataset.delWf)));
        wrap.querySelectorAll("[data-runs-wf]").forEach(btn =>
            btn.addEventListener("click", () => WorkflowRuns.open(this.container, this.project, btn.dataset.runsWf)));
    },

    async remove(id) {
        const wf = this.list.find(w => w.id === id);
        if (!wf) return;
        const ok = await UI.confirm("Eliminar workflow", `Se eliminará el workflow <strong>${UI.escapeHtml(wf.name)}</strong> y todos sus pasos, variables y tareas.`);
        if (!ok) return;
        try {
            const pasoIds = (await Provider.runQuery(
                `SELECT PASO_ID FROM ${Provider.qualifyControl("WORKFLOWS_PASOS")} WHERE WORKFLOW_ID = '${Provider.esc(id)}'`
            )).map(r => r.PASO_ID);

            for (const table of ["WORKFLOWS_PASOS_DRIVER_VALORES", "WORKFLOWS_PASOS_VARIABLES", "WORKFLOWS_PASOS_BLOQUES"]) {
                await Provider.runQuery(`DELETE FROM ${Provider.qualifyControl(table)} WHERE PASO_ID IN (SELECT PASO_ID FROM ${Provider.qualifyControl("WORKFLOWS_PASOS")} WHERE WORKFLOW_ID = '${Provider.esc(id)}')`);
            }
            if (pasoIds.length) {
                const tareaIds = (await Provider.runQuery(
                    `SELECT TAREA_ID FROM ${Provider.qualifyControl("WORKFLOWS_PASOS_TAREAS")} WHERE PASO_ID IN (${pasoIds.map(p => `'${Provider.esc(p)}'`).join(",")})`
                )).map(r => r.TAREA_ID);
                if (tareaIds.length) {
                    await Provider.runQuery(`DELETE FROM ${Provider.qualifyControl("WORKFLOWS_PASOS_TAREAS_VALORES")} WHERE TAREA_ID IN (${tareaIds.map(t => `'${Provider.esc(t)}'`).join(",")})`);
                }
                await Provider.runQuery(`DELETE FROM ${Provider.qualifyControl("WORKFLOWS_PASOS_TAREAS")} WHERE PASO_ID IN (${pasoIds.map(p => `'${Provider.esc(p)}'`).join(",")})`);
            }
            await Provider.runQuery(`DELETE FROM ${Provider.qualifyControl("WORKFLOWS_PASOS")} WHERE WORKFLOW_ID = '${Provider.esc(id)}'`);
            await Provider.runQuery(`DELETE FROM ${Provider.qualifyControl(this.TABLE)} WHERE ${this.ID_COL} = '${Provider.esc(id)}'`);

            await this.loadList();
            this.renderList();
            UI.toast(`Workflow "${wf.name}" eliminado.`, "success");
        } catch (err) {
            UI.toast("Error al eliminar el workflow: " + err.message, "error");
        }
    },

    // ------------------------------------------------------------
    // Datos auxiliares (dimensiones del proyecto, flujos manuales)
    // ------------------------------------------------------------
    async loadDimensions() {
        try {
            this.dimensions = await Provider.runQuery(`
                SELECT DIMENSION_ID, DIMENSION, DESCRIPCION, TABLA
                FROM ${Provider.qualifyControl("DIMENSIONES")}
                WHERE PROYECTO_ID = '${Provider.esc(this.project.PROYECTO_ID)}'
                ORDER BY DIMENSION`);
        } catch (e) { this.dimensions = []; }
    },

    async loadManualFlows() {
        try {
            const rows = await Provider.runQuery(`
                SELECT FLUJO_ID, FLUJO
                FROM ${Provider.qualifyControl("FLUJOS")}
                WHERE PROYECTO_ID = '${Provider.esc(this.project.PROYECTO_ID)}' AND TIPO = 'MANUAL'
                ORDER BY FLUJO`);
            this.manualFlows = rows.map(r => ({ id: r.FLUJO_ID, name: r.FLUJO }));
        } catch (e) { this.manualFlows = []; }
    },

    async flowScreenVariables(flujoId) {
        try {
            const rows = await Provider.runQuery(`
                SELECT NOMBRE, ETIQUETA, TIPO
                FROM ${Provider.qualifyControl("FLUJOS_SCREEN_VARIABLES")}
                WHERE FLUJO_ID = '${Provider.esc(flujoId)}' ORDER BY ORDEN`);
            return rows;
        } catch (e) { return []; }
    },

    dimensionById(id) {
        return this.dimensions.find(d => d.DIMENSION_ID === id) || null;
    },

    // ------------------------------------------------------------
    // Blancos
    // ------------------------------------------------------------
    blankWorkflow() {
        return { id: Provider.newId(), name: "", description: "", steps: [] };
    },

    blankStep(name) {
        return {
            id: Provider.newId(),
            name: name || "",
            inicio: { tipo: "INICIO_WORKFLOW", fecha: "" },
            revision: false,
            fin: { tipo: "COMPLETAR", fecha: "" },
            driver: { dimensionId: null, valores: [] },
            variables: [],
            bloques: []
        };
    },

    // ------------------------------------------------------------
    // Carga completa de un workflow existente
    // ------------------------------------------------------------
    async loadDetail(id) {
        const rows = await Provider.runQuery(`SELECT * FROM ${Provider.qualifyControl(this.TABLE)} WHERE ${this.ID_COL} = '${Provider.esc(id)}'`);
        const row = rows[0];
        if (!row) return null;

        const pasoRows = await Provider.runQuery(`
            SELECT * FROM ${Provider.qualifyControl("WORKFLOWS_PASOS")}
            WHERE WORKFLOW_ID = '${Provider.esc(id)}' ORDER BY ORDEN`);

        const pasoIds = pasoRows.map(p => p.PASO_ID);
        const inClause = pasoIds.length ? pasoIds.map(p => `'${Provider.esc(p)}'`).join(",") : "''";

        const driverValRows = pasoIds.length ? await Provider.runQuery(`
            SELECT PASO_ID, VALOR FROM ${Provider.qualifyControl("WORKFLOWS_PASOS_DRIVER_VALORES")}
            WHERE PASO_ID IN (${inClause})`) : [];
        const varRows = pasoIds.length ? await Provider.runQuery(`
            SELECT PASO_ID, VARIABLE_ID, NOMBRE, ETIQUETA, TIPO, ORDEN FROM ${Provider.qualifyControl("WORKFLOWS_PASOS_VARIABLES")}
            WHERE PASO_ID IN (${inClause}) ORDER BY ORDEN`) : [];
        const bloqueRows = pasoIds.length ? await Provider.runQuery(`
            SELECT PASO_ID, BLOQUE_ID, TITULO, ORDEN FROM ${Provider.qualifyControl("WORKFLOWS_PASOS_BLOQUES")}
            WHERE PASO_ID IN (${inClause}) ORDER BY ORDEN`) : [];
        const tareaRows = pasoIds.length ? await Provider.runQuery(`
            SELECT PASO_ID, BLOQUE_ID, TAREA_ID, TIPO, NOMBRE, REF_ID, REF_NOMBRE, ORDEN FROM ${Provider.qualifyControl("WORKFLOWS_PASOS_TAREAS")}
            WHERE PASO_ID IN (${inClause}) ORDER BY ORDEN`) : [];
        const tareaIds = tareaRows.map(t => t.TAREA_ID);
        const tareaValRows = tareaIds.length ? await Provider.runQuery(`
            SELECT TAREA_ID, CLAVE, ETIQUETA, TIPO, VALOR, OCULTAR FROM ${Provider.qualifyControl("WORKFLOWS_PASOS_TAREAS_VALORES")}
            WHERE TAREA_ID IN (${tareaIds.map(t => `'${Provider.esc(t)}'`).join(",")})`) : [];

        const driverValsByPaso = {};
        driverValRows.forEach(d => { (driverValsByPaso[d.PASO_ID] = driverValsByPaso[d.PASO_ID] || []).push(d.VALOR); });

        const varsByPaso = {};
        varRows.forEach(v => { (varsByPaso[v.PASO_ID] = varsByPaso[v.PASO_ID] || []).push({ id: v.VARIABLE_ID, name: v.NOMBRE, label: v.ETIQUETA || v.NOMBRE, type: v.TIPO || "STRING" }); });

        const valsByTarea = {};
        tareaValRows.forEach(v => {
            (valsByTarea[v.TAREA_ID] = valsByTarea[v.TAREA_ID] || []).push({
                clave: v.CLAVE, etiqueta: v.ETIQUETA || v.CLAVE,
                tipo: (v.TIPO || "CONSTANTE").toLowerCase() === "variable" ? "variable" : "constante",
                valor: v.VALOR || "", ocultar: !!(v.OCULTAR === true || v.OCULTAR === "true" || v.OCULTAR === 1)
            });
        });

        const tareasByBloque = {};
        tareaRows.forEach(t => {
            (tareasByBloque[t.BLOQUE_ID] = tareasByBloque[t.BLOQUE_ID] || []).push({
                id: t.TAREA_ID, tipo: t.TIPO, nombre: t.NOMBRE || "",
                refId: t.REF_ID || null, refNombre: t.REF_NOMBRE || "",
                valores: valsByTarea[t.TAREA_ID] || []
            });
        });

        const bloquesByPaso = {};
        bloqueRows.forEach(b => {
            (bloquesByPaso[b.PASO_ID] = bloquesByPaso[b.PASO_ID] || []).push({
                id: b.BLOQUE_ID, titulo: b.TITULO || "Bloque",
                tareas: tareasByBloque[b.BLOQUE_ID] || []
            });
        });

        const steps = pasoRows.map(p => ({
            id: p.PASO_ID,
            name: p.PASO,
            inicio: { tipo: p.INICIO_TIPO || "INICIO_WORKFLOW", fecha: p.INICIO_FECHA || "" },
            revision: !!(p.REVISION === true || p.REVISION === "true" || p.REVISION === 1),
            fin: { tipo: p.FIN_TIPO || "COMPLETAR", fecha: p.FIN_FECHA || "" },
            driver: { dimensionId: p.DRIVER_DIMENSION_ID || null, valores: driverValsByPaso[p.PASO_ID] || [] },
            variables: varsByPaso[p.PASO_ID] || [],
            bloques: bloquesByPaso[p.PASO_ID] || []
        }));

        return { id: row[this.ID_COL], name: row[this.NAME_COL], description: row.DESCRIPCION || "", steps };
    },

    // ------------------------------------------------------------
    // Alta: nombre + descripción, luego editor completo
    // ------------------------------------------------------------
    async openForm(editId = null) {
        this.editingIsNew = !editId;
        this.selectedStepId = null;
        this.activeTab = "propiedades";

        await this.loadDimensions();
        await this.loadManualFlows();

        if (editId) {
            const draft = await this.loadDetail(editId);
            if (!draft) { UI.toast("No se ha podido cargar el workflow.", "error"); return; }
            this.editing = draft;
            this.selectedStepId = draft.steps.length ? draft.steps[0].id : null;
            this.openMainModal();
            return;
        }

        const draft = this.blankWorkflow();
        const basics = await this.openBasicsModal(draft, true);
        if (!basics) return;
        Object.assign(draft, basics);
        this.editing = draft;
        this.openMainModal();
    },

    openBasicsModal(initial, isNew) {
        return new Promise((resolve) => {
            let overlay = document.getElementById("wfBasicsModal");
            if (!overlay) {
                overlay = document.createElement("div");
                overlay.className = "modal-overlay";
                overlay.id = "wfBasicsModal";
                document.body.appendChild(overlay);
            }
            overlay.innerHTML = `
                <div class="modal-box">
                    <div class="modal-header">
                        <h3>${isNew ? "Nuevo workflow" : "Datos básicos"}</h3>
                        <button class="modal-close" id="wfBasicsClose">&times;</button>
                    </div>
                    <div class="modal-body">
                        <div class="form-group">
                            <label>Nombre del workflow</label>
                            <input type="text" id="wfBasicsName" placeholder="Ej. Cierre mensual" value="${UI.escapeHtml(initial.name || "")}">
                        </div>
                        <div class="form-group">
                            <label>Descripción</label>
                            <textarea id="wfBasicsDesc" rows="3" placeholder="Describe brevemente el propósito de este workflow...">${UI.escapeHtml(initial.description || "")}</textarea>
                        </div>
                    </div>
                    <div class="modal-footer">
                        <button class="btn btn-secondary" id="wfBasicsCancel">Cancelar</button>
                        <button class="btn btn-primary" id="wfBasicsNext">Continuar</button>
                    </div>
                </div>`;

            overlay.classList.add("visible");
            const nameInput = overlay.querySelector("#wfBasicsName");
            setTimeout(() => { nameInput.focus(); }, 50);

            const cleanup = (result) => { overlay.classList.remove("visible"); resolve(result); };
            overlay.querySelector("#wfBasicsClose").onclick = () => cleanup(null);
            overlay.querySelector("#wfBasicsCancel").onclick = () => cleanup(null);
            overlay.querySelector("#wfBasicsNext").onclick = () => {
                const name = nameInput.value.trim();
                if (!name) { UI.toast("Indica un nombre para el workflow.", "error"); return; }
                cleanup({ name, description: overlay.querySelector("#wfBasicsDesc").value.trim() });
            };
            nameInput.onkeydown = (e) => { if (e.key === "Enter") overlay.querySelector("#wfBasicsNext").click(); };
        });
    },

    // ------------------------------------------------------------
    // Editor completo
    // ------------------------------------------------------------
    openMainModal() {
        let overlay = document.getElementById("wfFormModal");
        if (!overlay) {
            overlay = document.createElement("div");
            overlay.className = "modal-overlay";
            overlay.id = "wfFormModal";
            document.body.appendChild(overlay);
        }
        this.overlay = overlay;

        overlay.innerHTML = `
            <div class="modal-box modal-full">
                <div class="modal-header">
                    <div class="modal-title-edit-wrap">
                        <h3 id="wfModalTitle" class="modal-title-editable" contenteditable="true" spellcheck="false" title="Clic para renombrar el workflow"></h3>
                        <span class="modal-subtitle" id="wfModalSubtitle" contenteditable="true" spellcheck="false" title="Clic para editar la descripción"></span>
                    </div>
                    <div class="modal-header-right">
                        <button class="modal-close" id="wfFormClose">&times;</button>
                    </div>
                </div>
                <div class="modal-body modal-body-flush">
                    <div id="wfStepsPart" class="flow-part flow-part--chain"></div>
                    <div id="wfStepDetailPart" class="flow-part flow-part--mapping"></div>
                </div>
                <div class="modal-footer">
                    <button class="btn btn-secondary" id="wfFormCancel">Cancelar</button>
                    <button class="btn btn-primary" id="wfFormSave">Guardar workflow</button>
                </div>
            </div>`;

        document.getElementById("wfFormClose").addEventListener("click", () => this.closeForm());
        document.getElementById("wfFormCancel").addEventListener("click", () => this.closeForm());
        document.getElementById("wfFormSave").addEventListener("click", () => this.save());

        const titleEl = document.getElementById("wfModalTitle");
        titleEl.addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); titleEl.blur(); } });
        titleEl.addEventListener("blur", () => {
            const newName = titleEl.textContent.trim();
            if (!newName) { UI.toast("El workflow necesita un nombre.", "error"); titleEl.textContent = this.editing.name; return; }
            this.editing.name = newName;
            titleEl.textContent = newName;
        });

        const subEl = document.getElementById("wfModalSubtitle");
        subEl.addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); subEl.blur(); } });
        subEl.addEventListener("blur", () => { this.editing.description = subEl.textContent.trim(); });

        overlay.classList.add("visible");

        document.getElementById("wfModalTitle").textContent = this.editing.name;
        document.getElementById("wfModalSubtitle").textContent = this.editing.description || "Añade una descripción...";

        this.renderStepsPart();
        this.renderStepDetail();
    },

    closeForm() {
        if (this.overlay) this.overlay.classList.remove("visible");
        this.editing = null;
    },

    // ------------------------------------------------------------
    // Pasos (cadena, arrastrable)
    // ------------------------------------------------------------
    renderStepsPart() {
        const part = document.getElementById("wfStepsPart");
        const wf = this.editing;

        const cards = wf.steps.map((s, idx) => {
            const selected = s.id === this.selectedStepId;
            const card = `
                <div class="flow-chain-card ${selected ? "is-selected" : ""}" draggable="true" data-step-idx="${idx}" title="Clic para editar el paso">
                    <div class="flow-chain-card-name">${idx + 1}. ${UI.escapeHtml(s.name || "(sin nombre)")}</div>
                    <div class="flow-chain-card-meta">${this.stepMetaLabel(s)}</div>
                    <button type="button" class="flow-chain-card-remove" data-remove-step="${idx}" title="Eliminar paso">✕</button>
                </div>`;
            const arrow = idx < wf.steps.length - 1 ? `<div class="flow-chain-arrow">→</div>` : "";
            return card + arrow;
        }).join("");

        part.innerHTML = `
            <div class="flow-part-header"><strong>Pasos del workflow</strong></div>
            <div class="flow-chain-wrap" id="wfStepsWrap">
                ${cards}
                <button type="button" class="flow-chain-add" id="btnAddStep">+ Añadir paso</button>
            </div>`;

        document.getElementById("btnAddStep").addEventListener("click", () => this.addStep());

        part.querySelectorAll("[data-step-idx]").forEach(card => {
            card.addEventListener("click", (e) => {
                if (e.target.closest("[data-remove-step]")) return;
                this.selectedStepId = wf.steps[parseInt(card.dataset.stepIdx, 10)].id;
                this.activeTab = "propiedades";
                this.renderStepsPart();
                this.renderStepDetail();
            });
            card.addEventListener("dragstart", () => { this.dragStepIdx = parseInt(card.dataset.stepIdx, 10); });
            card.addEventListener("dragover", (e) => e.preventDefault());
            card.addEventListener("drop", (e) => {
                e.preventDefault();
                const toIdx = parseInt(card.dataset.stepIdx, 10);
                if (this.dragStepIdx === null || this.dragStepIdx === toIdx) return;
                const [moved] = wf.steps.splice(this.dragStepIdx, 1);
                wf.steps.splice(toIdx, 0, moved);
                this.dragStepIdx = null;
                this.renderStepsPart();
            });
        });
        part.querySelectorAll("[data-remove-step]").forEach(btn => {
            btn.addEventListener("click", async (e) => {
                e.stopPropagation();
                const idx = parseInt(btn.dataset.removeStep, 10);
                const step = wf.steps[idx];
                const ok = await UI.confirm("Eliminar paso", `Se eliminará el paso <strong>${UI.escapeHtml(step.name || "(sin nombre)")}</strong> con toda su configuración.`);
                if (!ok) return;
                wf.steps.splice(idx, 1);
                if (this.selectedStepId === step.id) this.selectedStepId = wf.steps.length ? wf.steps[0].id : null;
                this.renderStepsPart();
                this.renderStepDetail();
            });
        });
    },

    stepMetaLabel(s) {
        const bits = [];
        bits.push(s.revision ? "Con revisión" : "Sin revisión");
        if (s.driver.dimensionId) {
            const dim = this.dimensionById(s.driver.dimensionId);
            bits.push(`Driver: ${dim ? UI.escapeHtml(dim.DIMENSION) : "—"}${s.driver.valores.length ? ` (${s.driver.valores.length})` : " (todos)"}`);
        }
        const taskCount = s.bloques.reduce((n, b) => n + b.tareas.length, 0);
        bits.push(`${taskCount} tarea${taskCount === 1 ? "" : "s"}`);
        return bits.join(" · ");
    },

    async addStep() {
        const name = await UI.openTextPromptModal({ title: "Nuevo paso", label: "Nombre del paso", placeholder: "Ej. Carga de datos reales" });
        if (name === null) return;
        if (!name.trim()) { UI.toast("Indica un nombre para el paso.", "error"); return; }
        const step = this.blankStep(name.trim());
        this.editing.steps.push(step);
        this.selectedStepId = step.id;
        this.activeTab = "propiedades";
        this.renderStepsPart();
        this.renderStepDetail();
    },

    currentStep() {
        return this.editing.steps.find(s => s.id === this.selectedStepId) || null;
    },

    // ------------------------------------------------------------
    // Detalle del paso seleccionado (pestañas)
    // ------------------------------------------------------------
    renderStepDetail() {
        const part = document.getElementById("wfStepDetailPart");
        const step = this.currentStep();

        if (!step) {
            part.innerHTML = `<div class="module-empty module-empty--inline">Selecciona o añade un paso para configurarlo.</div>`;
            return;
        }

        const tabs = [
            ["propiedades", "Propiedades"],
            ["driver", "Driver"],
            ["variables", "Variables"],
            ["tareas", "Tareas"]
        ];

        part.innerHTML = `
            <div class="flow-part-header"><strong>Paso: ${UI.escapeHtml(step.name || "(sin nombre)")}</strong></div>
            <div class="segmented" id="wfStepTabs">
                ${tabs.map(([key, label]) => `<button type="button" class="segmented-btn ${this.activeTab === key ? "active" : ""}" data-tab="${key}">${label}</button>`).join("")}
            </div>
            <div class="flow-step-detail" id="wfStepTabBody"></div>`;

        part.querySelectorAll("[data-tab]").forEach(btn => {
            btn.addEventListener("click", () => {
                this.activeTab = btn.dataset.tab;
                this.renderStepDetail();
            });
        });

        const body = document.getElementById("wfStepTabBody");
        if (this.activeTab === "propiedades") this.renderPropiedadesTab(body, step);
        else if (this.activeTab === "driver") this.renderDriverTab(body, step);
        else if (this.activeTab === "variables") this.renderVariablesTab(body, step);
        else if (this.activeTab === "tareas") this.renderTareasTab(body, step);
    },

    // -------------------- Propiedades --------------------
    finOptions(step) {
        const opts = [["NA", "N/A"]];
        opts.push(step.revision ? ["REVISION", "Al enviarse a revisión"] : ["COMPLETAR", "Al completarse el paso"]);
        opts.push(["FECHA", "Fecha concreta"]);
        return opts;
    },

    renderPropiedadesTab(body, step) {
        const finOpts = this.finOptions(step);
        if (!finOpts.some(([k]) => k === step.fin.tipo)) step.fin.tipo = finOpts[1][0];

        body.innerHTML = `
            <div class="flow-step-group">
                <div class="form-group">
                    <label>Nombre del paso</label>
                    <input type="text" id="stepName" value="${UI.escapeHtml(step.name)}">
                </div>
                <div class="form-group">
                    <label>Inicio</label>
                    <select id="stepInicioTipo">
                        <option value="INICIO_WORKFLOW" ${step.inicio.tipo === "INICIO_WORKFLOW" ? "selected" : ""}>Al iniciar el workflow</option>
                        <option value="PASO_ANTERIOR" ${step.inicio.tipo === "PASO_ANTERIOR" ? "selected" : ""}>Al completar el paso anterior</option>
                        <option value="FECHA" ${step.inicio.tipo === "FECHA" ? "selected" : ""}>Fecha concreta</option>
                    </select>
                    ${step.inicio.tipo === "FECHA" ? `<input type="date" id="stepInicioFecha" value="${UI.escapeHtml(step.inicio.fecha)}" style="margin-top:8px;">` : ""}
                </div>
                <div class="form-group">
                    <label><input type="checkbox" id="stepRevision" ${step.revision ? "checked" : ""}> Requiere revisión</label>
                    <p class="form-hint">Si el paso requiere revisión, se considera terminado al enviarse a revisión (no al completarse).</p>
                </div>
                <div class="form-group">
                    <label>Finalización</label>
                    <select id="stepFinTipo">
                        ${finOpts.map(([k, l]) => `<option value="${k}" ${step.fin.tipo === k ? "selected" : ""}>${l}</option>`).join("")}
                    </select>
                    ${step.fin.tipo === "FECHA" ? `<input type="date" id="stepFinFecha" value="${UI.escapeHtml(step.fin.fecha)}" style="margin-top:8px;">` : ""}
                    <p class="form-hint">Este momento es el que dispara el "Al completar el paso anterior" del siguiente paso.</p>
                </div>
            </div>`;

        document.getElementById("stepName").addEventListener("input", (e) => {
            step.name = e.target.value;
            this.renderStepsPart();
            document.querySelector("#wfStepDetailPart .flow-part-header strong").textContent = `Paso: ${step.name || "(sin nombre)"}`;
        });
        document.getElementById("stepInicioTipo").addEventListener("change", (e) => {
            step.inicio.tipo = e.target.value;
            this.renderPropiedadesTab(body, step);
        });
        const inicioFecha = document.getElementById("stepInicioFecha");
        if (inicioFecha) inicioFecha.addEventListener("input", (e) => { step.inicio.fecha = e.target.value; });

        document.getElementById("stepRevision").addEventListener("change", (e) => {
            step.revision = e.target.checked;
            if (step.fin.tipo === "REVISION" || step.fin.tipo === "COMPLETAR") {
                step.fin.tipo = step.revision ? "REVISION" : "COMPLETAR";
            }
            this.renderPropiedadesTab(body, step);
            this.renderStepsPart();
        });
        document.getElementById("stepFinTipo").addEventListener("change", (e) => {
            step.fin.tipo = e.target.value;
            this.renderPropiedadesTab(body, step);
        });
        const finFecha = document.getElementById("stepFinFecha");
        if (finFecha) finFecha.addEventListener("input", (e) => { step.fin.fecha = e.target.value; });
    },

    // -------------------- Driver --------------------
    renderDriverTab(body, step) {
        const dim = step.driver.dimensionId ? this.dimensionById(step.driver.dimensionId) : null;

        body.innerHTML = `
            <div class="flow-step-group">
                <p class="form-hint">Reparte la ejecución de este paso por los valores de una dimensión (ej. nodo superior de la jerarquía de CECOs → una ejecución por CECO, cada una asignable a una persona distinta). Sin driver, el paso se asigna en bloque a una persona o grupo.</p>
                ${dim ? `
                    <div class="hier-chip-card" style="margin-bottom:12px;">
                        <div>
                            <strong>${UI.escapeHtml(dim.DIMENSION)}</strong>
                            <span class="hier-chip-levels">${step.driver.valores.length ? `${step.driver.valores.length} valor(es) seleccionado(s)` : "Todos los valores de la dimensión"}</span>
                        </div>
                        <div class="row-actions">
                            <button id="btnDriverValues" title="Elegir valores concretos">▤</button>
                            <button id="btnDriverClear" class="danger" title="Quitar driver">🗑</button>
                        </div>
                    </div>
                    ${step.driver.valores.length ? `
                        <div class="chip-row">
                            ${step.driver.valores.map(v => `<span class="hier-chip">${UI.escapeHtml(v)} <a href="#" data-remove-driver-val="${UI.escapeHtml(v)}" style="margin-left:4px;">✕</a></span>`).join("")}
                        </div>` : ""}
                ` : `<button class="btn btn-secondary btn-sm" id="btnDriverPick">+ Seleccionar dimensión driver</button>`}
            </div>`;

        const pickBtn = document.getElementById("btnDriverPick");
        if (pickBtn) pickBtn.addEventListener("click", async () => {
            const dimId = await UI.openDimensionPickerModal({ dimensionsList: this.dimensions });
            if (!dimId) return;
            step.driver.dimensionId = dimId;
            step.driver.valores = [];
            this.renderDriverTab(body, step);
            this.renderStepsPart();
        });
        const clearBtn = document.getElementById("btnDriverClear");
        if (clearBtn) clearBtn.addEventListener("click", () => {
            step.driver = { dimensionId: null, valores: [] };
            this.renderDriverTab(body, step);
            this.renderStepsPart();
        });
        const valuesBtn = document.getElementById("btnDriverValues");
        if (valuesBtn) valuesBtn.addEventListener("click", async () => {
            const keyCol = Provider.toIdentifier(dim.DIMENSION);
            const result = await UI.openDimensionValuesPickerModal({ project: this.project, dim, keyCol, selected: step.driver.valores });
            if (result === null) return;
            step.driver.valores = result;
            this.renderDriverTab(body, step);
            this.renderStepsPart();
        });
        body.querySelectorAll("[data-remove-driver-val]").forEach(a => {
            a.addEventListener("click", (e) => {
                e.preventDefault();
                step.driver.valores = step.driver.valores.filter(v => v !== a.dataset.removeDriverVal);
                this.renderDriverTab(body, step);
                this.renderStepsPart();
            });
        });
    },

    // -------------------- Variables del paso --------------------
    renderVariablesTab(body, step) {
        body.innerHTML = `
            <div class="flow-step-group">
                <div class="flow-step-group-title">
                    Variables de valor único
                    <button class="btn btn-secondary btn-sm" id="btnAddStepVar" style="float:right;">+ Nueva variable</button>
                </div>
                ${step.variables.length ? `
                    <div class="data-list">
                        <table>
                            <thead><tr><th>Nombre</th><th>Etiqueta</th><th>Tipo</th><th></th></tr></thead>
                            <tbody>
                                ${step.variables.map((v, idx) => `
                                    <tr>
                                        <td><strong>${UI.escapeHtml(v.name)}</strong></td>
                                        <td>${UI.escapeHtml(v.label)}</td>
                                        <td><span class="table-tag">${UI.escapeHtml(v.type)}</span></td>
                                        <td>
                                            <div class="row-actions">
                                                <button data-edit-var="${idx}" title="Editar">✎</button>
                                                <button data-del-var="${idx}" class="danger" title="Eliminar">🗑</button>
                                            </div>
                                        </td>
                                    </tr>`).join("")}
                            </tbody>
                        </table>
                    </div>` : `<div class="module-empty module-empty--inline">Este paso todavía no tiene variables.</div>`}
            </div>`;

        document.getElementById("btnAddStepVar").addEventListener("click", async () => {
            const v = await UI.openWorkflowVariableModal({});
            if (!v) return;
            step.variables.push({ id: Provider.newId(), ...v });
            this.renderVariablesTab(body, step);
        });
        body.querySelectorAll("[data-edit-var]").forEach(btn => {
            btn.addEventListener("click", async () => {
                const idx = parseInt(btn.dataset.editVar, 10);
                const v = await UI.openWorkflowVariableModal({ current: step.variables[idx] });
                if (!v) return;
                step.variables[idx] = { ...step.variables[idx], ...v };
                this.renderVariablesTab(body, step);
            });
        });
        body.querySelectorAll("[data-del-var]").forEach(btn => {
            btn.addEventListener("click", () => {
                const idx = parseInt(btn.dataset.delVar, 10);
                step.variables.splice(idx, 1);
                this.renderVariablesTab(body, step);
            });
        });
    },

    // -------------------- Tareas (bloques) --------------------
    renderTareasTab(body, step) {
        body.innerHTML = `
            <div class="flow-step-group">
                <button class="btn btn-secondary btn-sm" id="btnAddBlock">+ Añadir bloque</button>
                <div id="wfBlocksWrap" style="margin-top:12px;">
                    ${step.bloques.length ? step.bloques.map((b, bIdx) => this.blockHtml(b, bIdx)).join("") : `<div class="module-empty module-empty--inline">Todavía no hay bloques de tareas en este paso.</div>`}
                </div>
            </div>`;

        document.getElementById("btnAddBlock").addEventListener("click", () => {
            step.bloques.push({ id: Provider.newId(), titulo: "Nuevo bloque", tareas: [] });
            this.renderTareasTab(body, step);
        });

        this.bindBlockEvents(body, step);
    },

    blockHtml(block, bIdx) {
        return `
            <div class="flow-screen-block flow-screen-block--frame" data-block-idx="${bIdx}" style="margin-bottom:14px;">
                <div class="flow-frame-header">
                    <span class="modal-title-editable" contenteditable="true" spellcheck="false" data-block-title="${bIdx}">${UI.escapeHtml(block.titulo)}</span>
                    <span class="load-fn-toolbar-spacer"></span>
                    <button type="button" class="btn btn-secondary btn-sm" data-add-task="${bIdx}">+ Tarea</button>
                    <button type="button" class="flow-chain-card-remove" data-remove-block="${bIdx}" title="Eliminar bloque">✕</button>
                </div>
                <div class="flow-frame-vars">
                    ${block.tareas.length ? block.tareas.map((t, tIdx) => this.taskHtml(block, bIdx, t, tIdx)).join("") : `<p class="form-hint">Sin tareas todavía.</p>`}
                </div>
            </div>`;
    },

    taskHtml(block, bIdx, task, tIdx) {
        const typeInfo = this.TASK_TYPES[task.tipo] || { label: task.tipo, icon: "•" };
        const canAddCustomVar = task.tipo === "PLANTILLA" || task.tipo === "FUNCION" || task.tipo === "HTML";
        return `
            <div class="flow-screen-block" data-task-idx="${tIdx}" style="margin-bottom:8px;">
                <div class="flow-frame-header">
                    <span>${typeInfo.icon} <strong>${UI.escapeHtml(typeInfo.label)}</strong> — ${UI.escapeHtml(task.refNombre || task.nombre || "(sin referencia)")}</span>
                    <span class="load-fn-toolbar-spacer"></span>
                    ${canAddCustomVar ? `<button type="button" class="btn btn-secondary btn-sm" data-add-taskvar="${bIdx}:${tIdx}">+ Variable</button>` : ""}
                    <button type="button" class="flow-chain-card-remove" data-remove-task="${bIdx}:${tIdx}" title="Eliminar tarea">✕</button>
                </div>
                ${task.valores.length ? `
                    <div class="flow-mapping-vars-list">
                        ${task.valores.map((v, vIdx) => `
                            <div class="flow-target-row" data-edit-taskval="${bIdx}:${tIdx}:${vIdx}" style="cursor:pointer;">
                                <span class="flow-target-label">${UI.escapeHtml(v.etiqueta || v.clave)}</span>
                                <span class="table-tag">${v.tipo === "variable" ? "= " + UI.escapeHtml(v.valor || "—") : UI.escapeHtml(v.valor || "(vacío)")}</span>
                                ${v.ocultar ? `<span class="table-tag">Oculta</span>` : ""}
                            </div>`).join("")}
                    </div>` : `<p class="form-hint">${task.tipo === "PARAMETRIZACION" ? "Esta tarea no necesita variables adicionales." : "Sin variables asignadas."}</p>`}
            </div>`;
    },

    bindBlockEvents(body, step) {
        document.querySelectorAll("[data-block-title]").forEach(el => {
            el.addEventListener("blur", () => {
                const idx = parseInt(el.dataset.blockTitle, 10);
                step.bloques[idx].titulo = el.textContent.trim() || "Bloque";
            });
            el.addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); el.blur(); } });
        });

        document.querySelectorAll("[data-remove-block]").forEach(btn => {
            btn.addEventListener("click", async () => {
                const idx = parseInt(btn.dataset.removeBlock, 10);
                const ok = await UI.confirm("Eliminar bloque", `Se eliminará el bloque <strong>${UI.escapeHtml(step.bloques[idx].titulo)}</strong> y todas sus tareas.`);
                if (!ok) return;
                step.bloques.splice(idx, 1);
                this.renderTareasTab(body, step);
                this.renderStepsPart();
            });
        });

        document.querySelectorAll("[data-add-task]").forEach(btn => {
            btn.addEventListener("click", async () => {
                const bIdx = parseInt(btn.dataset.addTask, 10);
                await this.addTask(step, step.bloques[bIdx]);
                this.renderTareasTab(body, step);
                this.renderStepsPart();
            });
        });

        document.querySelectorAll("[data-remove-task]").forEach(btn => {
            btn.addEventListener("click", () => {
                const [bIdx, tIdx] = btn.dataset.removeTask.split(":").map(n => parseInt(n, 10));
                step.bloques[bIdx].tareas.splice(tIdx, 1);
                this.renderTareasTab(body, step);
                this.renderStepsPart();
            });
        });

        document.querySelectorAll("[data-add-taskvar]").forEach(btn => {
            btn.addEventListener("click", async () => {
                const [bIdx, tIdx] = btn.dataset.addTaskvar.split(":").map(n => parseInt(n, 10));
                const clave = await UI.openTextPromptModal({ title: "Nueva variable de la tarea", label: "Nombre de la variable", placeholder: "ej. ruta_fichero" });
                if (!clave || !clave.trim()) return;
                const task = step.bloques[bIdx].tareas[tIdx];
                const result = await UI.openWorkflowValueModal({ title: "Asignar valor", targetLabel: clave.trim(), stepVariables: step.variables });
                if (!result || result === "remove") return;
                task.valores.push({ clave: clave.trim(), etiqueta: clave.trim(), tipo: result.type, valor: result.value, ocultar: !!result.hidden });
                this.renderTareasTab(body, step);
            });
        });

        document.querySelectorAll("[data-edit-taskval]").forEach(row => {
            row.addEventListener("click", async () => {
                const [bIdx, tIdx, vIdx] = row.dataset.editTaskval.split(":").map(n => parseInt(n, 10));
                const task = step.bloques[bIdx].tareas[tIdx];
                const v = task.valores[vIdx];
                const result = await UI.openWorkflowValueModal({
                    title: "Asignar valor", targetLabel: v.etiqueta || v.clave,
                    stepVariables: step.variables,
                    current: { type: v.tipo, value: v.valor, hidden: v.ocultar }
                });
                if (result === null) return;
                if (result === "remove") { task.valores.splice(vIdx, 1); }
                else { v.tipo = result.type; v.valor = result.value; v.ocultar = !!result.hidden; }
                this.renderTareasTab(body, step);
            });
        });
    },

    async addTask(step, block) {
        const choice = await UI.choiceModal("Nueva tarea", "¿Qué tipo de tarea quieres añadir a este bloque?", [
            { key: "FLUJO_MANUAL", label: "Flujo manual" },
            { key: "PLANTILLA", label: "Plantilla" },
            { key: "FUNCION", label: "Función" },
            { key: "PARAMETRIZACION", label: "Actualizar tabla de parametrización" },
            { key: "HTML", label: "Página HTML", style: "primary" }
        ]);
        if (!choice) return;

        if (choice === "FLUJO_MANUAL") {
            const flow = await UI.openFlowManualPickerModal({ flows: this.manualFlows });
            if (!flow) return;
            const screenVars = await this.flowScreenVariables(flow.id);
            block.tareas.push({
                id: Provider.newId(), tipo: "FLUJO_MANUAL", nombre: flow.name, refId: flow.id, refNombre: flow.name,
                valores: screenVars.map(v => ({ clave: v.NOMBRE, etiqueta: v.ETIQUETA || v.NOMBRE, tipo: "constante", valor: "", ocultar: false }))
            });
            return;
        }

        if (choice === "PARAMETRIZACION") {
            const dimId = await UI.openDimensionPickerModal({ dimensionsList: this.dimensions });
            if (!dimId) return;
            const dim = this.dimensionById(dimId);
            block.tareas.push({ id: Provider.newId(), tipo: "PARAMETRIZACION", nombre: dim.DIMENSION, refId: dim.DIMENSION_ID, refNombre: dim.DIMENSION, valores: [] });
            return;
        }

        // PLANTILLA / FUNCION / HTML — todavía sin catálogo propio: referencia libre
        const labels = { PLANTILLA: "la plantilla", FUNCION: "la función", HTML: "la página HTML" };
        const name = await UI.openTextPromptModal({ title: "Nueva tarea", label: `Nombre de ${labels[choice]}`, placeholder: "Ej. " + (this.TASK_TYPES[choice] || {}).label });
        if (!name || !name.trim()) return;
        block.tareas.push({ id: Provider.newId(), tipo: choice, nombre: name.trim(), refId: null, refNombre: name.trim(), valores: [] });
    },

    // ------------------------------------------------------------
    // Guardado
    // ------------------------------------------------------------
    async save() {
        const wf = this.editing;
        if (!wf.name) { UI.toast("El workflow necesita un nombre.", "error"); return; }
        for (const s of wf.steps) {
            if (!s.name) { UI.toast("Todos los pasos necesitan un nombre.", "error"); return; }
        }

        const btn = document.getElementById("wfFormSave");
        if (btn) { btn.disabled = true; btn.textContent = "Guardando…"; }
        try {
            await this.persist();
            const name = wf.name;
            this.editingIsNew = false;
            this.closeForm();
            await this.loadList();
            this.renderList();
            if (window.Draco && Draco.renderProgress) Draco.renderProgress();
            UI.toast(`Workflow "${name}" guardado.`, "success");
        } catch (err) {
            UI.toast("Error al guardar el workflow: " + err.message, "error");
        } finally {
            if (btn) { btn.disabled = false; btn.textContent = "Guardar workflow"; }
        }
    },

    async persist() {
        const wf = this.editing;
        const id = wf.id;
        const pid = this.project.PROYECTO_ID;

        // 1) Cabecera --------------------------------------------------
        if (this.editingIsNew) {
            await Provider.runQuery(`INSERT INTO ${Provider.qualifyControl("WORKFLOWS")}
                (WORKFLOW_ID, PROYECTO_ID, WORKFLOW, DESCRIPCION, USUARIO, FECHA_CREACION, FECHA_MODIFICACION)
                VALUES ('${Provider.esc(id)}', '${Provider.esc(pid)}', '${Provider.esc(wf.name)}', '${Provider.esc(wf.description || "")}',
                        ${Provider.currentUserExpr()}, CURRENT_TIMESTAMP(), CURRENT_TIMESTAMP())`);
        } else {
            await Provider.runQuery(`UPDATE ${Provider.qualifyControl("WORKFLOWS")}
                SET WORKFLOW = '${Provider.esc(wf.name)}', DESCRIPCION = '${Provider.esc(wf.description || "")}', FECHA_MODIFICACION = CURRENT_TIMESTAMP()
                WHERE WORKFLOW_ID = '${Provider.esc(id)}'`);
        }

        // 2) Limpieza total de la configuración anterior de pasos -------
        const oldPasoIds = (await Provider.runQuery(
            `SELECT PASO_ID FROM ${Provider.qualifyControl("WORKFLOWS_PASOS")} WHERE WORKFLOW_ID = '${Provider.esc(id)}'`
        )).map(r => r.PASO_ID);
        if (oldPasoIds.length) {
            const oldTareaIds = (await Provider.runQuery(
                `SELECT TAREA_ID FROM ${Provider.qualifyControl("WORKFLOWS_PASOS_TAREAS")} WHERE PASO_ID IN (${oldPasoIds.map(p => `'${Provider.esc(p)}'`).join(",")})`
            )).map(r => r.TAREA_ID);
            if (oldTareaIds.length) {
                await Provider.runQuery(`DELETE FROM ${Provider.qualifyControl("WORKFLOWS_PASOS_TAREAS_VALORES")} WHERE TAREA_ID IN (${oldTareaIds.map(t => `'${Provider.esc(t)}'`).join(",")})`);
            }
        }
        for (const table of ["WORKFLOWS_PASOS_TAREAS", "WORKFLOWS_PASOS_BLOQUES", "WORKFLOWS_PASOS_VARIABLES", "WORKFLOWS_PASOS_DRIVER_VALORES"]) {
            await Provider.runQuery(`DELETE FROM ${Provider.qualifyControl(table)} WHERE PASO_ID IN (SELECT PASO_ID FROM ${Provider.qualifyControl("WORKFLOWS_PASOS")} WHERE WORKFLOW_ID = '${Provider.esc(id)}')`);
        }
        await Provider.runQuery(`DELETE FROM ${Provider.qualifyControl("WORKFLOWS_PASOS")} WHERE WORKFLOW_ID = '${Provider.esc(id)}'`);

        if (!wf.steps.length) return;

        // 3) Pasos -------------------------------------------------------
        const pasoVals = wf.steps.map((s, idx) => `(
            '${Provider.esc(s.id)}', '${Provider.esc(id)}', '${Provider.esc(pid)}', '${Provider.esc(s.name)}', ${idx},
            '${Provider.esc(s.inicio.tipo)}', '${Provider.esc(s.inicio.fecha || "")}', ${s.revision ? "TRUE" : "FALSE"},
            '${Provider.esc(s.fin.tipo)}', '${Provider.esc(s.fin.fecha || "")}',
            ${s.driver.dimensionId ? `'${Provider.esc(s.driver.dimensionId)}'` : "NULL"},
            '${s.driver.dimensionId ? (s.driver.valores.length ? "VALORES" : "TODOS") : "NINGUNO"}'
        )`).join(",\n");
        await Provider.runQuery(`INSERT INTO ${Provider.qualifyControl("WORKFLOWS_PASOS")}
            (PASO_ID, WORKFLOW_ID, PROYECTO_ID, PASO, ORDEN, INICIO_TIPO, INICIO_FECHA, REVISION, FIN_TIPO, FIN_FECHA, DRIVER_DIMENSION_ID, DRIVER_MODO)
            VALUES ${pasoVals}`);

        // 4) Valores del driver -------------------------------------------
        const driverVals = [];
        wf.steps.forEach(s => (s.driver.valores || []).forEach(v => driverVals.push(`('${Provider.esc(pid)}', '${Provider.esc(s.id)}', '${Provider.esc(v)}')`)));
        if (driverVals.length) {
            await Provider.runQuery(`INSERT INTO ${Provider.qualifyControl("WORKFLOWS_PASOS_DRIVER_VALORES")} (PROYECTO_ID, PASO_ID, VALOR) VALUES ${driverVals.join(",\n")}`);
        }

        // 5) Variables del paso ---------------------------------------------
        const varVals = [];
        wf.steps.forEach(s => (s.variables || []).forEach((v, idx) =>
            varVals.push(`('${Provider.esc(pid)}', '${Provider.esc(s.id)}', '${Provider.esc(v.id || Provider.newId())}', '${Provider.esc(v.name)}', '${Provider.esc(v.label)}', '${Provider.esc(v.type)}', ${idx})`)));
        if (varVals.length) {
            await Provider.runQuery(`INSERT INTO ${Provider.qualifyControl("WORKFLOWS_PASOS_VARIABLES")} (PROYECTO_ID, PASO_ID, VARIABLE_ID, NOMBRE, ETIQUETA, TIPO, ORDEN) VALUES ${varVals.join(",\n")}`);
        }

        // 6) Bloques + tareas + valores de tareas ---------------------------
        const bloqueVals = [], tareaVals = [], tareaValVals = [];
        wf.steps.forEach(s => {
            (s.bloques || []).forEach((b, bIdx) => {
                bloqueVals.push(`('${Provider.esc(pid)}', '${Provider.esc(s.id)}', '${Provider.esc(b.id)}', '${Provider.esc(b.titulo)}', ${bIdx})`);
                (b.tareas || []).forEach((t, tIdx) => {
                    tareaVals.push(`('${Provider.esc(pid)}', '${Provider.esc(s.id)}', '${Provider.esc(b.id)}', '${Provider.esc(t.id)}', '${Provider.esc(t.tipo)}', '${Provider.esc(t.nombre || "")}', ${t.refId ? `'${Provider.esc(t.refId)}'` : "NULL"}, '${Provider.esc(t.refNombre || "")}', ${tIdx})`);
                    (t.valores || []).forEach(v => {
                        tareaValVals.push(`('${Provider.esc(pid)}', '${Provider.esc(t.id)}', '${Provider.esc(v.clave)}', '${Provider.esc(v.etiqueta || v.clave)}', '${Provider.esc(v.tipo === "variable" ? "VARIABLE" : "CONSTANTE")}', '${Provider.esc(v.valor || "")}', ${v.ocultar ? "TRUE" : "FALSE"})`);
                    });
                });
            });
        });
        if (bloqueVals.length) {
            await Provider.runQuery(`INSERT INTO ${Provider.qualifyControl("WORKFLOWS_PASOS_BLOQUES")} (PROYECTO_ID, PASO_ID, BLOQUE_ID, TITULO, ORDEN) VALUES ${bloqueVals.join(",\n")}`);
        }
        if (tareaVals.length) {
            await Provider.runQuery(`INSERT INTO ${Provider.qualifyControl("WORKFLOWS_PASOS_TAREAS")} (PROYECTO_ID, PASO_ID, BLOQUE_ID, TAREA_ID, TIPO, NOMBRE, REF_ID, REF_NOMBRE, ORDEN) VALUES ${tareaVals.join(",\n")}`);
        }
        if (tareaValVals.length) {
            await Provider.runQuery(`INSERT INTO ${Provider.qualifyControl("WORKFLOWS_PASOS_TAREAS_VALORES")} (PROYECTO_ID, TAREA_ID, CLAVE, ETIQUETA, TIPO, VALOR, OCULTAR) VALUES ${tareaValVals.join(",\n")}`);
        }
    }
};

/**
 * ============================================================
 * DRACO PLANNING — JERARQUÍAS DE DIMENSIÓN
 * ============================================================
 * Permite definir jerarquías (ej. Geografía: País → Región →
 * Ciudad) arrastrando atributos de la dimensión a una lista de
 * niveles ordenada, con vista previa en vivo sobre los datos reales.
 */
const Hierarchies = {
    TABLE: "JERARQUIAS",
    ID_COL: "JERARQUIA_ID",
    list: [],
    editingId: null,
    editLevels: [], // array de colId, nivel superior primero
    dragFromIdx: null,

    async render(container, project, dim, onBack) {
        this.container = container;
        this.project = project;
        this.dim = dim;
        this.onBack = onBack;
        this.fields = Dimensions.parseFields(dim).map(f => ({ ...f, colId: Provider.toIdentifier(f.name) }));
        this.fullTable = Provider.qualify(project.DATASET, dim.TABLA);

        container.innerHTML = `
            <div class="module-header">
                <div>
                    <button class="btn-back" id="btnBackToDims2">← Dimensiones</button>
                    <h3>Jerarquías: ${UI.escapeHtml(dim[Dimensions.NAME_COL])}</h3>
                    <p>Define agrupaciones de niveles sobre los atributos de esta dimensión.</p>
                </div>
                <button class="btn btn-primary btn-sm" id="btnNewHier">+ Nueva jerarquía</button>
            </div>
            <div id="hierListWrap"><span class="spinner"></span></div>
            <div id="hierEditorWrap"></div>
        `;

        document.getElementById("btnBackToDims2").addEventListener("click", () => this.onBack());
        document.getElementById("btnNewHier").addEventListener("click", () => this.openEditor(null));
        await this.loadList();
    },

    async loadList() {
        const wrap = document.getElementById("hierListWrap");
        try {
            const sql = `SELECT ${this.ID_COL}, JERARQUIA, NIVELES_JSON
                         FROM ${Provider.qualifyControl(this.TABLE)}
                         WHERE DIMENSION_ID = '${Provider.esc(this.dim[Dimensions.ID_COL])}'
                         ORDER BY JERARQUIA`;
            this.list = await Provider.runQuery(sql);

            if (!this.list.length) {
                wrap.innerHTML = `<div class="module-empty">Todavía no hay jerarquías para esta dimensión.</div>`;
                return;
            }

            wrap.innerHTML = `
                <div class="hier-chip-list">
                    ${this.list.map(h => {
                        const levels = this.safeParse(h.NIVELES_JSON);
                        const names = levels.map(colId => (this.fields.find(f => f.colId === colId) || {}).name || colId).join(" → ");
                        return `
                            <div class="hier-chip-card">
                                <div>
                                    <strong>${UI.escapeHtml(h.JERARQUIA)}</strong>
                                    <span class="hier-chip-levels">${UI.escapeHtml(names || "(sin niveles)")}</span>
                                </div>
                                <div class="row-actions">
                                    <button data-edit-hier="${h[this.ID_COL]}" title="Editar">✎</button>
                                    <button data-del-hier="${h[this.ID_COL]}" class="danger" title="Eliminar">🗑</button>
                                </div>
                            </div>`;
                    }).join("")}
                </div>`;

            wrap.querySelectorAll("[data-edit-hier]").forEach(btn =>
                btn.addEventListener("click", () => this.openEditor(btn.dataset.editHier)));
            wrap.querySelectorAll("[data-del-hier]").forEach(btn =>
                btn.addEventListener("click", () => this.remove(btn.dataset.delHier)));
        } catch (err) {
            wrap.innerHTML = `<div class="module-empty">Error al cargar jerarquías: ${UI.escapeHtml(err.message)}</div>`;
        }
    },

    safeParse(json) {
        try { return JSON.parse(json || "[]"); } catch (e) { return []; }
    },

    openEditor(id) {
        const editing = id ? this.list.find(h => h[this.ID_COL] === id) : null;
        this.editingId = id;
        this.editLevels = editing ? this.safeParse(editing.NIVELES_JSON) : [];

        const wrap = document.getElementById("hierEditorWrap");
        wrap.innerHTML = `
            <div class="hier-editor-box">
                <div class="form-row" style="align-items:flex-end;">
                    <div class="form-group" style="flex:1;">
                        <label>Nombre de la jerarquía</label>
                        <input type="text" id="hierNameInput" placeholder="Ej. Geografía" value="${editing ? UI.escapeHtml(editing.JERARQUIA) : ""}">
                    </div>
                    <button class="btn btn-secondary btn-sm" id="btnCancelHier">Cancelar</button>
                    <button class="btn btn-primary btn-sm" id="btnSaveHier">Guardar jerarquía</button>
                </div>

                <div class="hierarchy-editor">
                    <div class="hierarchy-pool-col">
                        <div class="hierarchy-col-label">Atributos disponibles (clic para añadir)</div>
                        <div id="hierPoolList" class="hierarchy-pool-list"></div>
                    </div>
                    <div class="hierarchy-levels-col">
                        <div class="hierarchy-col-label">Niveles: superior → inferior (arrastra para reordenar)</div>
                        <div id="hierLevelsList" class="hierarchy-levels-list"></div>
                    </div>
                    <div class="hierarchy-preview-col">
                        <div class="hierarchy-col-label">Vista previa con datos reales</div>
                        <div id="hierPreviewTree" class="hierarchy-preview-tree"></div>
                    </div>
                </div>
            </div>
        `;

        document.getElementById("btnCancelHier").addEventListener("click", () => { wrap.innerHTML = ""; });
        document.getElementById("btnSaveHier").addEventListener("click", () => this.save());

        this.renderPoolAndLevels();
        this.updatePreview();
    },

    renderPoolAndLevels() {
        const poolEl = document.getElementById("hierPoolList");
        const levelsEl = document.getElementById("hierLevelsList");

        const available = this.fields.filter(f => !this.editLevels.includes(f.colId));
        poolEl.innerHTML = available.length
            ? available.map(f => `<button type="button" class="hier-chip" data-add="${f.colId}">${UI.escapeHtml(f.name)} <span>＋</span></button>`).join("")
            : `<div class="hierarchy-pool-empty">Todos los atributos están ya en la jerarquía.</div>`;

        poolEl.querySelectorAll("[data-add]").forEach(btn => {
            btn.addEventListener("click", () => {
                this.editLevels.push(btn.dataset.add);
                this.renderPoolAndLevels();
                this.updatePreview();
            });
        });

        levelsEl.innerHTML = this.editLevels.length
            ? this.editLevels.map((colId, idx) => {
                const f = this.fields.find(x => x.colId === colId);
                return `
                    <div class="hier-level-chip" draggable="true" data-idx="${idx}">
                        <span class="hier-level-num">${idx + 1}</span>
                        <span class="hier-level-name">${UI.escapeHtml(f ? f.name : colId)}</span>
                        <button type="button" class="hierarchy-chip-btn-remove" data-remove-level="${idx}" title="Quitar">✕</button>
                    </div>`;
            }).join("")
            : `<div class="hierarchy-levels-empty">Añade atributos desde la columna de la izquierda.</div>`;

        levelsEl.querySelectorAll("[data-remove-level]").forEach(btn => {
            btn.addEventListener("click", () => {
                this.editLevels.splice(parseInt(btn.dataset.removeLevel, 10), 1);
                this.renderPoolAndLevels();
                this.updatePreview();
            });
        });

        this.bindDrag(levelsEl);
    },

    bindDrag(levelsEl) {
        const chips = levelsEl.querySelectorAll(".hier-level-chip");
        chips.forEach(chip => {
            chip.addEventListener("dragstart", () => {
                this.dragFromIdx = parseInt(chip.dataset.idx, 10);
                chip.classList.add("dragging");
            });
            chip.addEventListener("dragend", () => chip.classList.remove("dragging"));
            chip.addEventListener("dragover", (e) => e.preventDefault());
            chip.addEventListener("drop", (e) => {
                e.preventDefault();
                const toIdx = parseInt(chip.dataset.idx, 10);
                if (this.dragFromIdx === null || this.dragFromIdx === toIdx) return;
                const [moved] = this.editLevels.splice(this.dragFromIdx, 1);
                this.editLevels.splice(toIdx, 0, moved);
                this.dragFromIdx = null;
                this.renderPoolAndLevels();
                this.updatePreview();
            });
        });
    },

    async updatePreview() {
        const box = document.getElementById("hierPreviewTree");
        if (!box) return;
        if (!this.editLevels.length) {
            box.innerHTML = `<div class="hierarchy-pool-empty">Añade al menos un nivel para ver la vista previa.</div>`;
            return;
        }

        box.innerHTML = `<span class="spinner"></span>`;
        try {
            const cols = this.editLevels.join(", ");
            const sql = `SELECT DISTINCT ${cols} FROM ${this.fullTable} ORDER BY ${cols} LIMIT 300`;
            const rows = await Provider.runQuery(sql);
            box.innerHTML = rows.length
                ? this.buildTreeHtml(rows, this.editLevels, 0)
                : `<div class="hierarchy-pool-empty">La tabla todavía no tiene datos que previsualizar.</div>`;
        } catch (err) {
            box.innerHTML = `<div class="hierarchy-pool-empty">No se pudo generar la vista previa: ${UI.escapeHtml(err.message)}</div>`;
        }
    },

    /** Agrupa filas planas en un árbol anidado según el orden de niveles */
    buildTreeHtml(rows, levels, depth) {
        if (depth >= levels.length) return "";
        const col = levels[depth];
        const groups = new Map();
        rows.forEach(r => {
            const key = r[col] ?? "(vacío)";
            if (!groups.has(key)) groups.set(key, []);
            groups.get(key).push(r);
        });

        const items = Array.from(groups.entries()).slice(0, 50).map(([key, subRows]) => `
            <li class="hier-preview-node">
                <div class="hier-preview-label"><span class="hier-preview-level-tag">N${depth + 1}</span><span class="hier-preview-value">${UI.escapeHtml(key)}</span></div>
                ${depth + 1 < levels.length ? `<ul class="hier-preview-children">${this.buildTreeHtml(subRows, levels, depth + 1)}</ul>` : ""}
            </li>`).join("");

        return `<ul class="hier-preview-root">${items}</ul>`;
    },

    async save() {
        const name = document.getElementById("hierNameInput").value.trim();
        if (!name) {
            UI.toast("Indica un nombre para la jerarquía.", "error");
            return;
        }
        if (!this.editLevels.length) {
            UI.toast("Añade al menos un nivel.", "error");
            return;
        }

        const levelsJson = JSON.stringify(this.editLevels).replace(/\\/g, "\\\\").replace(/'/g, "\\'");

        try {
            if (this.editingId) {
                const sql = `UPDATE ${Provider.qualifyControl(this.TABLE)}
                    SET JERARQUIA = '${Provider.esc(name)}', NIVELES_JSON = '${levelsJson}', FECHA_MODIFICACION = CURRENT_TIMESTAMP()
                    WHERE ${this.ID_COL} = '${Provider.esc(this.editingId)}'`;
                await Provider.runQuery(sql);
                UI.toast(`Jerarquía "${name}" actualizada.`, "success");
            } else {
                const id = Provider.newId();
                const sql = `INSERT INTO ${Provider.qualifyControl(this.TABLE)}
                    (${this.ID_COL}, DIMENSION_ID, PROYECTO_ID, JERARQUIA, NIVELES_JSON, USUARIO, FECHA_CREACION, FECHA_MODIFICACION)
                    VALUES ('${Provider.esc(id)}', '${Provider.esc(this.dim[Dimensions.ID_COL])}', '${Provider.esc(this.project.PROYECTO_ID)}',
                            '${Provider.esc(name)}', '${levelsJson}', ${Provider.currentUserExpr()}, CURRENT_TIMESTAMP(), CURRENT_TIMESTAMP())`;
                await Provider.runQuery(sql);
                UI.toast(`Jerarquía "${name}" creada.`, "success");
            }

            document.getElementById("hierEditorWrap").innerHTML = "";
            await this.loadList();
        } catch (err) {
            UI.toast("Error al guardar la jerarquía: " + err.message, "error");
        }
    },

    async remove(id) {
        const hier = this.list.find(h => h[this.ID_COL] === id);
        if (!hier) return;
        const ok = await UI.confirm("Eliminar jerarquía", `Se eliminará la jerarquía <strong>${UI.escapeHtml(hier.JERARQUIA)}</strong>.`);
        if (!ok) return;

        try {
            await Provider.runQuery(`DELETE FROM ${Provider.qualifyControl(this.TABLE)} WHERE ${this.ID_COL} = '${Provider.esc(id)}'`);
            UI.toast(`Jerarquía "${hier.JERARQUIA}" eliminada.`, "success");
            await this.loadList();
        } catch (err) {
            UI.toast("Error al eliminar la jerarquía: " + err.message, "error");
        }
    }
};

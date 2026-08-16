/**
 * ============================================================
 * DRACO PLANNING — VALORES DE LA DIMENSIÓN
 * ============================================================
 * Rejilla editable de los datos físicos de una dimensión:
 *  - Añadir filas, pegar bloques copiados de Excel.
 *  - Exportar a CSV / Excel.
 *  - Importar desde archivo (CSV/XLSX), sustituyendo todo o
 *    fusionando de forma incremental por clave.
 *  - "Guardar" vuelca la rejilla completa a la tabla física
 *    (TRUNCATE + INSERT). Es decir: lo que ves es lo que se guarda.
 */
const DimensionData = {
    MAX_ROWS_LOAD: 2000,

    async render(container, project, dim, onBack) {
        this.container = container;
        this.project = project;
        this.dim = dim;
        this.onBack = onBack;
        this.fields = Dimensions.parseFields(dim).map(f => ({ ...f, colId: Provider.toIdentifier(f.name) }));
        this.fullTable = Provider.qualify(project.DATASET, dim.TABLA);
        this.rows = []; // array de objetos { COLID: valor, ... }

        container.innerHTML = `
            <div class="module-header">
                <div>
                    <button class="btn-back" id="btnBackToDims">← Dimensiones</button>
                    <h3>Valores: ${UI.escapeHtml(dim[Dimensions.NAME_COL])}</h3>
                    <p>Tabla ${dim.TABLA} · ${this.fields.length} columna(s)</p>
                </div>
            </div>
            <div class="values-toolbar">
                <button class="btn btn-secondary btn-sm" id="btnAddRow">+ Añadir fila</button>
                <button class="btn btn-secondary btn-sm" id="btnExportCsv">Exportar CSV</button>
                <button class="btn btn-secondary btn-sm" id="btnExportXlsx">Exportar Excel</button>
                <button class="btn btn-secondary btn-sm" id="btnImportFile">Importar archivo</button>
                <input type="file" id="importFileInput" accept=".csv,.xlsx,.xls" style="display:none;">
                <span class="values-toolbar-spacer"></span>
                <span class="values-row-count" id="valuesRowCount"></span>
                <button class="btn btn-primary btn-sm" id="btnSaveValues">Guardar cambios</button>
            </div>
            <p class="form-hint">Pega bloques de celdas directamente desde Excel (Ctrl+V sobre una celda). "Guardar" sustituye por completo el contenido de la tabla con lo que ves aquí.</p>
            <div class="values-grid-wrap" id="valuesGridWrap"><span class="spinner"></span></div>
        `;

        document.getElementById("btnBackToDims").addEventListener("click", () => this.onBack());
        document.getElementById("btnAddRow").addEventListener("click", () => { this.addEmptyRow(); this.renderGrid(); });
        document.getElementById("btnExportCsv").addEventListener("click", () => this.exportCsv());
        document.getElementById("btnExportXlsx").addEventListener("click", () => this.exportXlsx());
        document.getElementById("btnImportFile").addEventListener("click", () => document.getElementById("importFileInput").click());
        document.getElementById("importFileInput").addEventListener("change", (e) => this.handleImportFile(e));
        document.getElementById("btnSaveValues").addEventListener("click", () => this.save());

        await this.loadData();
    },

    async loadData() {
        const wrap = document.getElementById("valuesGridWrap");
        try {
            const cols = this.fields.map(f => f.colId).join(", ");
            const sql = `SELECT ${cols} FROM ${this.fullTable} LIMIT ${this.MAX_ROWS_LOAD}`;
            this.rows = await Provider.runQuery(sql);
            this.renderGrid();
        } catch (err) {
            wrap.innerHTML = `<div class="module-empty">Error al cargar los valores: ${UI.escapeHtml(err.message)}</div>`;
        }
    },

    addEmptyRow() {
        const row = {};
        this.fields.forEach(f => { row[f.colId] = ""; });
        this.rows.push(row);
    },

    renderGrid() {
        const wrap = document.getElementById("valuesGridWrap");
        document.getElementById("valuesRowCount").textContent = `${this.rows.length} fila(s)`;

        if (!this.rows.length) this.addEmptyRow();

        const header = this.fields.map(f =>
            `<th>${UI.escapeHtml(f.name)}${f.key ? ' <span class="key-dot" title="Clave">🔑</span>' : ""}<br><span class="col-type">${f.type}</span></th>`
        ).join("");

        const bodyRows = this.rows.map((row, rIdx) => `
            <tr data-row-idx="${rIdx}">
                ${this.fields.map((f, cIdx) => `
                    <td><input type="text" data-col-idx="${cIdx}" data-col="${f.colId}" value="${UI.escapeHtml(row[f.colId] ?? "")}"></td>
                `).join("")}
                <td class="values-row-remove"><button type="button" data-remove-row="${rIdx}" title="Eliminar fila">✕</button></td>
            </tr>
        `).join("");

        wrap.innerHTML = `
            <table class="values-grid">
                <thead><tr>${header}<th></th></tr></thead>
                <tbody id="valuesGridBody">${bodyRows}</tbody>
            </table>`;

        const tbody = document.getElementById("valuesGridBody");

        tbody.querySelectorAll("input").forEach(input => {
            input.addEventListener("change", () => {
                const tr = input.closest("tr");
                const rIdx = parseInt(tr.dataset.rowIdx, 10);
                const col = input.dataset.col;
                this.rows[rIdx][col] = input.value;
            });
        });

        tbody.querySelectorAll("[data-remove-row]").forEach(btn => {
            btn.addEventListener("click", () => {
                this.syncRowsFromDom();
                this.rows.splice(parseInt(btn.dataset.removeRow, 10), 1);
                this.renderGrid();
            });
        });

        tbody.addEventListener("paste", (e) => this.handlePaste(e));
    },

    /** Vuelca lo que hay actualmente en los <input> del DOM a this.rows, por si hay cambios sin "change" disparado */
    syncRowsFromDom() {
        const tbody = document.getElementById("valuesGridBody");
        if (!tbody) return;
        tbody.querySelectorAll("tr").forEach(tr => {
            const rIdx = parseInt(tr.dataset.rowIdx, 10);
            if (!this.rows[rIdx]) return;
            tr.querySelectorAll("input").forEach(input => {
                this.rows[rIdx][input.dataset.col] = input.value;
            });
        });
    },

    handlePaste(e) {
        const target = e.target;
        if (!target.matches("input")) return;
        const text = (e.clipboardData || window.clipboardData).getData("text");
        if (!text || (!text.includes("\t") && !text.includes("\n"))) return; // pegado simple de una celda: comportamiento nativo

        e.preventDefault();
        this.syncRowsFromDom();

        const pastedRows = text.replace(/\r/g, "").split("\n").filter((r, i, arr) => !(i === arr.length - 1 && r === ""));
        const startRowIdx = parseInt(target.closest("tr").dataset.rowIdx, 10);
        const startColIdx = parseInt(target.dataset.colIdx, 10);

        pastedRows.forEach((rowText, rOffset) => {
            const cells = rowText.split("\t");
            const rowIdx = startRowIdx + rOffset;
            while (rowIdx >= this.rows.length) this.addEmptyRow();
            cells.forEach((cellText, cOffset) => {
                const colIdx = startColIdx + cOffset;
                if (colIdx >= this.fields.length) return;
                this.rows[rowIdx][this.fields[colIdx].colId] = cellText;
            });
        });

        this.renderGrid();
    },

    // ------------------------------------------------------------
    // Export
    // ------------------------------------------------------------
    toAoa() {
        this.syncRowsFromDom();
        const header = this.fields.map(f => f.name);
        const body = this.rows.map(r => this.fields.map(f => r[f.colId] ?? ""));
        return [header, ...body];
    },

    toCsv(aoa) {
        return aoa.map(row => row.map(cell => {
            const s = String(cell ?? "");
            return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
        }).join(",")).join("\n");
    },

    exportCsv() {
        const csv = this.toCsv(this.toAoa());
        UI.downloadBlob(`${this.dim.TABLA}.csv`, "\uFEFF" + csv, "text/csv;charset=utf-8");
    },

    exportXlsx() {
        const ws = XLSX.utils.aoa_to_sheet(this.toAoa());
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Datos");
        XLSX.writeFile(wb, `${this.dim.TABLA}.xlsx`);
    },

    // ------------------------------------------------------------
    // Import
    // ------------------------------------------------------------
    async handleImportFile(e) {
        const file = e.target.files[0];
        e.target.value = ""; // permite reimportar el mismo archivo dos veces seguidas
        if (!file) return;

        try {
            const aoa = await this.parseFileToAoa(file);
            if (!aoa.length) {
                UI.toast("El archivo está vacío.", "error");
                return;
            }

            const [headerRow, ...dataRows] = aoa;
            const headerMap = headerRow.map(h => String(h).trim().toUpperCase());
            const importedObjs = dataRows
                .filter(r => r.some(c => String(c ?? "").trim() !== ""))
                .map(r => {
                    const obj = {};
                    this.fields.forEach(f => {
                        const idx = headerMap.indexOf(f.name.toUpperCase()) !== -1
                            ? headerMap.indexOf(f.name.toUpperCase())
                            : headerMap.indexOf(f.colId);
                        obj[f.colId] = idx !== -1 ? String(r[idx] ?? "") : "";
                    });
                    return obj;
                });

            if (!importedObjs.length) {
                UI.toast("No se encontraron filas de datos en el archivo.", "error");
                return;
            }

            const choice = await UI.choiceModal(
                "Importar valores",
                `Se han leído <strong>${importedObjs.length}</strong> fila(s) de <strong>${UI.escapeHtml(file.name)}</strong>. ¿Cómo quieres aplicarlas a la rejilla?`,
                [
                    { key: "incremental", label: "Incremental (por clave)", style: "primary" },
                    { key: "replace", label: "Sustituir todo", style: "secondary" }
                ]
            );
            if (!choice) return;

            this.syncRowsFromDom();

            if (choice === "replace") {
                this.rows = importedObjs;
            } else {
                this.mergeIncremental(importedObjs);
            }

            this.renderGrid();
            UI.toast(`Importación aplicada (${choice === "replace" ? "sustitución total" : "incremental"}). Pulsa "Guardar cambios" para confirmarlo en la base de datos.`, "success");
        } catch (err) {
            UI.toast("Error al leer el archivo: " + err.message, "error");
        }
    },

    parseFileToAoa(file) {
        return new Promise((resolve, reject) => {
            const isCsv = /\.csv$/i.test(file.name);
            const reader = new FileReader();
            reader.onerror = () => reject(new Error("No se pudo leer el archivo."));
            reader.onload = () => {
                try {
                    const wb = isCsv
                        ? XLSX.read(reader.result, { type: "string" })
                        : XLSX.read(new Uint8Array(reader.result), { type: "array" });
                    const ws = wb.Sheets[wb.SheetNames[0]];
                    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false });
                    resolve(aoa);
                } catch (err) {
                    reject(err);
                }
            };
            if (isCsv) reader.readAsText(file, "utf-8");
            else reader.readAsArrayBuffer(file);
        });
    },

    /** Fusiona filas importadas en this.rows: si la clave coincide, sobrescribe; si no, añade */
    mergeIncremental(importedObjs) {
        const keyCols = this.fields.filter(f => f.key).map(f => f.colId);
        const keyCols_ = keyCols.length ? keyCols : [this.fields[0].colId];
        const sig = (obj) => keyCols_.map(c => String(obj[c] ?? "").trim().toUpperCase()).join("||");

        const index = new Map();
        this.rows.forEach((r, i) => index.set(sig(r), i));

        importedObjs.forEach(imp => {
            const s = sig(imp);
            if (index.has(s)) {
                this.rows[index.get(s)] = { ...this.rows[index.get(s)], ...imp };
            } else {
                this.rows.push(imp);
                index.set(s, this.rows.length - 1);
            }
        });
    },

    // ------------------------------------------------------------
    // Save (TRUNCATE + INSERT del contenido completo de la rejilla)
    // ------------------------------------------------------------
    formatValue(raw, type) {
        const v = (raw ?? "").toString().trim();
        if (v === "") return "NULL";

        if (["INTEGER", "FLOAT", "NUMERIC"].includes(type)) {
            const normalized = v.replace(",", ".");
            return /^-?\d+(\.\d+)?$/.test(normalized) ? normalized : "NULL";
        }
        if (type === "BOOLEAN") {
            const truthy = ["true", "1", "sí", "si", "yes", "verdadero"];
            return truthy.includes(v.toLowerCase()) ? "TRUE" : "FALSE";
        }
        return `'${Provider.esc(v)}'`;
    },

    async save() {
        this.syncRowsFromDom();

        const keyCols = this.fields.filter(f => f.key).map(f => f.colId);
        const keyCols_ = keyCols.length ? keyCols : [this.fields[0].colId];
        const validRows = this.rows.filter(r => keyCols_.every(c => String(r[c] ?? "").trim() !== ""));
        const skipped = this.rows.length - validRows.length;

        if (!validRows.length) {
            UI.toast("No hay filas válidas que guardar (revisa que la clave no esté vacía).", "error");
            return;
        }

        const btn = document.getElementById("btnSaveValues");
        btn.disabled = true;
        btn.textContent = "Guardando...";

        try {
            await Provider.runQuery(`TRUNCATE TABLE ${this.fullTable}`);

            const colList = this.fields.map(f => f.colId).join(", ");
            const chunkSize = 200;
            for (let i = 0; i < validRows.length; i += chunkSize) {
                const chunk = validRows.slice(i, i + chunkSize);
                const values = chunk.map(row =>
                    `(${this.fields.map(f => this.formatValue(row[f.colId], f.type)).join(", ")})`
                ).join(", ");
                await Provider.runQuery(`INSERT INTO ${this.fullTable} (${colList}) VALUES ${values}`);
            }

            UI.toast(`Guardado: ${validRows.length} fila(s)${skipped ? ` (se omitieron ${skipped} por clave vacía)` : ""}.`, "success");
            await this.loadData();
        } catch (err) {
            UI.toast("Error al guardar los valores: " + err.message, "error");
        } finally {
            btn.disabled = false;
            btn.textContent = "Guardar cambios";
        }
    }
};

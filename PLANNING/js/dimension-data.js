/**
 * ============================================================
 * DRACO PLANNING — VALORES DE LA DIMENSIÓN
 * ============================================================
 * Rejilla editable de los datos físicos de una dimensión:
 *  - Añadir filas, pegar bloques copiados de Excel.
 *  - Exportar a CSV / Excel.
 *  - Importar desde archivo (CSV/XLSX), sustituyendo todo o
 *    fusionando de forma incremental por clave. El mapeo de columnas
 *    del archivo importado se hace por ORDEN/POSICIÓN, no por el
 *    nombre de la cabecera: la primera fila del archivo se descarta
 *    siempre (se asume cabecera) pero su contenido no se usa para
 *    decidir qué columna va a qué campo.
 *  - Selección múltiple de filas (individual / todas las visibles)
 *    para borrado masivo, buscador que filtra por cualquier columna
 *    y orden ascendente/descendente pulsando en la cabecera.
 *  - "Guardar" vuelca la rejilla completa a la tabla física
 *    (TRUNCATE + INSERT). Es decir: lo que ves es lo que se guarda.
 */
const DimensionData = {
    MAX_ROWS_LOAD: 2000,

    async render(project, dim) {
        this.project = project;
        this.dim = dim;
        this.fields = Dimensions.parseFields(dim).map(f => ({ ...f, colId: Provider.toIdentifier(f.name) }));
        this.fullTable = Provider.qualify(project.DATASET, dim.TABLA);
        this.rows = []; // array de objetos { __uid, COLID: valor, ... }
        this.selected = new Set();  // __uid de las filas seleccionadas
        this.searchText = "";
        this.sortCol = null;
        this.sortDir = "asc";
        this._uidSeq = 0;

        let overlay = document.getElementById("valuesModal");
        if (!overlay) {
            overlay = document.createElement("div");
            overlay.className = "modal-overlay";
            overlay.id = "valuesModal";
            document.body.appendChild(overlay);
        }

        overlay.innerHTML = `
            <div class="modal-box modal-full">
                <div class="modal-header">
                    <div>
                        <h3>Valores: ${UI.escapeHtml(dim[Dimensions.NAME_COL])}</h3>
                        <span class="modal-subtitle">Tabla ${dim.TABLA} · ${this.fields.length} columna(s)</span>
                    </div>
                    <button class="modal-close" id="btnCloseValuesModal">&times;</button>
                </div>
                <div class="modal-body modal-body-flush">
                    <div class="values-toolbar">
                        <button class="btn btn-secondary btn-sm" id="btnAddRow">+ Añadir fila</button>
                        <button class="btn btn-secondary btn-sm" id="btnExportCsv">Exportar CSV</button>
                        <button class="btn btn-secondary btn-sm" id="btnExportXlsx">Exportar Excel</button>
                        <button class="btn btn-secondary btn-sm" id="btnImportFile">Importar archivo</button>
                        <input type="file" id="importFileInput" accept=".csv,.xlsx,.xls" style="display:none;">
                        <button class="btn btn-danger btn-sm" id="btnDeleteSelected" disabled>Eliminar seleccionadas</button>
                        <span class="values-toolbar-spacer"></span>
                        <input type="search" id="valuesSearchInput" class="values-search-input" placeholder="Buscar...">
                        <span class="values-row-count" id="valuesRowCount"></span>
                        <button class="btn btn-primary btn-sm" id="btnSaveValues">Guardar cambios</button>
                    </div>
                    <p class="form-hint">Pega bloques de celdas directamente desde Excel (Ctrl+V sobre una celda). Al importar un archivo, el orden de las columnas del fichero debe coincidir con el de esta rejilla (la cabecera del archivo se descarta y no se usa para el mapeo). Marca filas con la casilla para borrarlas en bloque, usa el buscador para filtrar y pulsa una cabecera de columna para ordenar. "Guardar" sustituye por completo el contenido de la tabla con lo que ves aquí. La clave debe ser única: las filas con clave repetida se marcan en rojo y bloquean el guardado.</p>
                    <div class="values-grid-wrap values-grid-wrap--modal" id="valuesGridWrap"><span class="spinner"></span></div>
                </div>
            </div>`;

        document.getElementById("btnCloseValuesModal").addEventListener("click", () => this.close());
        document.getElementById("btnAddRow").addEventListener("click", () => { this.syncRowsFromDom(); this.addEmptyRow(); this.renderGrid(); });
        document.getElementById("btnExportCsv").addEventListener("click", () => this.exportCsv());
        document.getElementById("btnExportXlsx").addEventListener("click", () => this.exportXlsx());
        document.getElementById("btnImportFile").addEventListener("click", () => document.getElementById("importFileInput").click());
        document.getElementById("importFileInput").addEventListener("change", (e) => this.handleImportFile(e));
        document.getElementById("btnDeleteSelected").addEventListener("click", () => this.deleteSelected());
        document.getElementById("btnSaveValues").addEventListener("click", () => this.save());
        document.getElementById("valuesSearchInput").addEventListener("input", (e) => {
            this.syncRowsFromDom();
            this.searchText = e.target.value.trim();
            this.renderGrid();
        });

        overlay.classList.add("visible");
        await this.loadData();
    },

    close() {
        const overlay = document.getElementById("valuesModal");
        if (overlay) overlay.classList.remove("visible");
    },

    nextUid() {
        this._uidSeq = (this._uidSeq || 0) + 1;
        return "r" + this._uidSeq;
    },

    async loadData() {
        const wrap = document.getElementById("valuesGridWrap");
        try {
            const cols = this.fields.map(f => f.colId).join(", ");
            const sql = `SELECT ${cols} FROM ${this.fullTable} LIMIT ${this.MAX_ROWS_LOAD}`;
            const data = await Provider.runQuery(sql);
            this.rows = data.map(r => ({ ...r, __uid: this.nextUid() }));
            this.selected = new Set();
            this.renderGrid();
        } catch (err) {
            wrap.innerHTML = `<div class="module-empty">Error al cargar los valores: ${UI.escapeHtml(err.message)}</div>`;
        }
    },

    addEmptyRow() {
        const row = { __uid: this.nextUid() };
        this.fields.forEach(f => { row[f.colId] = ""; });
        this.rows.push(row);
    },

    // ------------------------------------------------------------
    // Filtro (buscador) + orden por columna, sin tocar this.rows
    // (que sigue siendo la fuente de la verdad en su orden original)
    // ------------------------------------------------------------
    getVisibleRows() {
        let rows = this.rows;

        if (this.searchText) {
            const q = this.searchText.toLowerCase();
            rows = rows.filter(r => this.fields.some(f => String(r[f.colId] ?? "").toLowerCase().includes(q)));
        }

        if (this.sortCol) {
            const col = this.sortCol;
            const dir = this.sortDir === "desc" ? -1 : 1;
            rows = [...rows].sort((a, b) => {
                const av = a[col] ?? "";
                const bv = b[col] ?? "";
                const an = parseFloat(String(av).replace(",", "."));
                const bn = parseFloat(String(bv).replace(",", "."));
                const bothNumeric = String(av).trim() !== "" && String(bv).trim() !== "" && !isNaN(an) && !isNaN(bn);
                const cmp = bothNumeric
                    ? (an - bn)
                    : String(av).localeCompare(String(bv), "es", { sensitivity: "base", numeric: true });
                return cmp * dir;
            });
        }

        return rows;
    },

    updateDeleteSelectedBtn() {
        const btn = document.getElementById("btnDeleteSelected");
        if (!btn) return;
        btn.disabled = this.selected.size === 0;
        btn.textContent = this.selected.size ? `Eliminar seleccionadas (${this.selected.size})` : "Eliminar seleccionadas";
    },

    renderGrid() {
        const wrap = document.getElementById("valuesGridWrap");
        if (!this.rows.length) this.addEmptyRow();

        const visibleRows = this.getVisibleRows();
        document.getElementById("valuesRowCount").textContent = this.searchText
            ? `${visibleRows.length} de ${this.rows.length} fila(s)`
            : `${this.rows.length} fila(s)`;

        const allVisibleSelected = visibleRows.length > 0 && visibleRows.every(r => this.selected.has(r.__uid));
        const sortArrow = (colId) => this.sortCol === colId ? (this.sortDir === "asc" ? " ▲" : " ▼") : "";

        const header = `
            <th class="values-col-check"><input type="checkbox" id="valuesSelectAll" ${allVisibleSelected ? "checked" : ""} title="Seleccionar todas"></th>
            ${this.fields.map(f => `
                <th class="values-col-sortable" data-sort-col="${f.colId}" title="Ordenar por ${UI.escapeHtml(f.name)}">
                    ${UI.escapeHtml(f.name)}${f.key ? ' <span class="key-dot" title="Clave">🔑</span>' : ""}<span class="sort-arrow">${sortArrow(f.colId)}</span><br><span class="col-type">${f.type}</span>
                </th>`).join("")}
            <th></th>`;

        const bodyRows = visibleRows.length
            ? visibleRows.map(row => `
                <tr data-uid="${row.__uid}" class="${this.selected.has(row.__uid) ? "row-selected" : ""}">
                    <td class="values-col-check"><input type="checkbox" class="row-select" data-uid="${row.__uid}" ${this.selected.has(row.__uid) ? "checked" : ""}></td>
                    ${this.fields.map((f, cIdx) => `
                        <td><input type="text" data-col-idx="${cIdx}" data-col="${f.colId}" value="${UI.escapeHtml(row[f.colId] ?? "")}"></td>
                    `).join("")}
                    <td class="values-row-remove"><button type="button" data-remove-uid="${row.__uid}" title="Eliminar fila">✕</button></td>
                </tr>
            `).join("")
            : `<tr><td colspan="${this.fields.length + 2}" class="values-empty">No hay filas que coincidan con la búsqueda.</td></tr>`;

        wrap.innerHTML = `
            <table class="values-grid">
                <thead><tr>${header}</tr></thead>
                <tbody id="valuesGridBody">${bodyRows}</tbody>
            </table>`;

        const tbody = document.getElementById("valuesGridBody");

        tbody.querySelectorAll('input[type="text"]').forEach(input => {
            input.addEventListener("change", () => {
                const uid = input.closest("tr").dataset.uid;
                const row = this.rows.find(r => r.__uid === uid);
                if (row) row[input.dataset.col] = input.value;
            });
        });

        tbody.querySelectorAll("[data-remove-uid]").forEach(btn => {
            btn.addEventListener("click", () => {
                this.syncRowsFromDom();
                const uid = btn.dataset.removeUid;
                this.rows = this.rows.filter(r => r.__uid !== uid);
                this.selected.delete(uid);
                this.renderGrid();
            });
        });

        tbody.querySelectorAll(".row-select").forEach(cb => {
            cb.addEventListener("change", () => {
                this.syncRowsFromDom();
                if (cb.checked) this.selected.add(cb.dataset.uid);
                else this.selected.delete(cb.dataset.uid);
                cb.closest("tr").classList.toggle("row-selected", cb.checked);
                const selectAll = document.getElementById("valuesSelectAll");
                if (selectAll) selectAll.checked = this.getVisibleRows().every(r => this.selected.has(r.__uid));
                this.updateDeleteSelectedBtn();
            });
        });

        const selectAllCb = document.getElementById("valuesSelectAll");
        if (selectAllCb) {
            selectAllCb.addEventListener("change", (e) => {
                this.syncRowsFromDom();
                const vis = this.getVisibleRows();
                vis.forEach(r => e.target.checked ? this.selected.add(r.__uid) : this.selected.delete(r.__uid));
                this.renderGrid();
            });
        }

        wrap.querySelectorAll("[data-sort-col]").forEach(th => {
            th.addEventListener("click", () => {
                this.syncRowsFromDom();
                const col = th.dataset.sortCol;
                if (this.sortCol === col) this.sortDir = this.sortDir === "asc" ? "desc" : "asc";
                else { this.sortCol = col; this.sortDir = "asc"; }
                this.renderGrid();
            });
        });

        tbody.addEventListener("paste", (e) => this.handlePaste(e));

        this.updateDeleteSelectedBtn();
    },

    /** Vuelca lo que hay actualmente en los <input> del DOM a this.rows, por si hay cambios sin "change" disparado */
    syncRowsFromDom() {
        const tbody = document.getElementById("valuesGridBody");
        if (!tbody) return;
        tbody.querySelectorAll("tr[data-uid]").forEach(tr => {
            const row = this.rows.find(r => r.__uid === tr.dataset.uid);
            if (!row) return;
            tr.querySelectorAll('input[type="text"]').forEach(input => {
                row[input.dataset.col] = input.value;
            });
        });
    },

    async deleteSelected() {
        if (!this.selected.size) return;
        const ok = await UI.confirm(
            "Eliminar filas seleccionadas",
            `Se eliminarán <strong>${this.selected.size}</strong> fila(s) de la rejilla. Esto no afecta a la base de datos hasta que pulses "Guardar cambios".`
        );
        if (!ok) return;

        this.syncRowsFromDom();
        this.rows = this.rows.filter(r => !this.selected.has(r.__uid));
        this.selected.clear();
        this.renderGrid();
    },

    handlePaste(e) {
        const target = e.target;
        if (!target.matches('input[type="text"]')) return;
        const text = (e.clipboardData || window.clipboardData).getData("text");
        if (!text || (!text.includes("\t") && !text.includes("\n"))) return; // pegado simple de una celda: comportamiento nativo

        e.preventDefault();
        this.syncRowsFromDom();

        const pastedRows = text.replace(/\r/g, "").split("\n").filter((r, i, arr) => !(i === arr.length - 1 && r === ""));
        const startUid = target.closest("tr").dataset.uid;
        const startColIdx = parseInt(target.dataset.colIdx, 10);

        // Si hay un filtro de búsqueda u orden activo, la posición visible de
        // las filas no coincide con this.rows, así que no se pueden crear
        // filas nuevas "al final" de forma fiable: el pegado solo sobrescribe
        // filas ya visibles.
        const filtering = !!this.searchText || !!this.sortCol;
        const visibleRows = this.getVisibleRows();
        const startVisibleIdx = visibleRows.findIndex(r => r.__uid === startUid);
        if (startVisibleIdx === -1) return;

        let clipped = false;
        pastedRows.forEach((rowText, rOffset) => {
            const cells = rowText.split("\t");
            const visibleIdx = startVisibleIdx + rOffset;
            let targetRow;
            if (visibleIdx < visibleRows.length) {
                targetRow = visibleRows[visibleIdx];
            } else if (!filtering) {
                this.addEmptyRow();
                targetRow = this.rows[this.rows.length - 1];
            } else {
                clipped = true;
                return;
            }
            cells.forEach((cellText, cOffset) => {
                const colIdx = startColIdx + cOffset;
                if (colIdx >= this.fields.length) return;
                targetRow[this.fields[colIdx].colId] = cellText;
            });
        });

        if (clipped) {
            UI.toast('Quita el buscador/orden para poder pegar filas nuevas al final de la rejilla.', "info");
        }

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

            // La primera fila se descarta por ser la cabecera, pero su
            // contenido (nombres de columna) NO se usa para el mapeo: los
            // valores se asignan a los campos de la dimensión por ORDEN,
            // es decir, la 1ª columna del archivo va al 1er campo de la
            // rejilla, la 2ª al 2º, etc., independientemente de cómo se
            // llame la cabecera en el archivo.
            const [, ...dataRows] = aoa;
            const importedObjs = dataRows
                .filter(r => r.some(c => String(c ?? "").trim() !== ""))
                .map(r => {
                    const obj = { __uid: this.nextUid() };
                    this.fields.forEach((f, idx) => {
                        obj[f.colId] = idx < r.length ? String(r[idx] ?? "") : "";
                    });
                    return obj;
                });

            if (!importedObjs.length) {
                UI.toast("No se encontraron filas de datos en el archivo.", "error");
                return;
            }

            const choice = await UI.choiceModal(
                "Importar valores",
                `Se han leído <strong>${importedObjs.length}</strong> fila(s) de <strong>${UI.escapeHtml(file.name)}</strong> (columnas mapeadas por orden). ¿Cómo quieres aplicarlas a la rejilla?`,
                [
                    { key: "incremental", label: "Incremental (por clave)", style: "primary" },
                    { key: "replace", label: "Sustituir todo", style: "secondary" }
                ]
            );
            if (!choice) return;

            this.syncRowsFromDom();
            this.selected.clear();
            this.searchText = "";
            this.sortCol = null;
            const searchInput = document.getElementById("valuesSearchInput");
            if (searchInput) searchInput.value = "";

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
                    if (isCsv) {
                        // OJO: el CSV se parsea a mano en vez de pasarlo por SheetJS.
                        // SheetJS autodetecta fechas/números en el texto y los
                        // reformatea según su propio locale/formato por defecto (p.ej.
                        // "2023-01-01" podía acabar convertido a un serial de fecha u
                        // otro formato de texto). Parseando el CSV directamente se
                        // preserva el valor tal cual viene escrito en el archivo.
                        resolve(this.parseCsvText(reader.result));
                    } else {
                        const wb = XLSX.read(new Uint8Array(reader.result), { type: "array" });
                        const ws = wb.Sheets[wb.SheetNames[0]];
                        const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false });
                        resolve(aoa);
                    }
                } catch (err) {
                    reject(err);
                }
            };
            if (isCsv) reader.readAsText(file, "utf-8");
            else reader.readAsArrayBuffer(file);
        });
    },

    /**
     * Parseo manual de CSV a array-de-arrays de strings, sin ninguna
     * conversión de tipo (fechas, números...). Soporta campos entre
     * comillas dobles (con comas/saltos de línea dentro y "" como
     * comilla escapada), autodetecta si el separador es "," o ";" mirando
     * la primera línea, y quita el BOM inicial si lo hay.
     */
    parseCsvText(text) {
        if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);

        const firstLine = text.split(/\r\n|\n/, 1)[0] || "";
        const delimiter = (firstLine.split(";").length > firstLine.split(",").length) ? ";" : ",";

        const rows = [];
        let row = [];
        let field = "";
        let inQuotes = false;
        let sawAny = false;

        for (let i = 0; i < text.length; i++) {
            const ch = text[i];

            if (inQuotes) {
                if (ch === '"') {
                    if (text[i + 1] === '"') { field += '"'; i++; }
                    else inQuotes = false;
                } else {
                    field += ch;
                }
                continue;
            }

            if (ch === '"') { inQuotes = true; sawAny = true; continue; }
            if (ch === delimiter) { row.push(field); field = ""; sawAny = true; continue; }
            if (ch === "\r") continue; // el salto de línea real lo marca \n
            if (ch === "\n") {
                row.push(field);
                rows.push(row);
                row = []; field = ""; sawAny = false;
                continue;
            }
            field += ch;
            sawAny = true;
        }
        if (sawAny || field !== "" || row.length) {
            row.push(field);
            rows.push(row);
        }

        return rows;
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
                const i = index.get(s);
                // Conserva el __uid de la fila existente para no romper la
                // selección/orden de la rejilla que ya estuviera en pantalla.
                this.rows[i] = { ...this.rows[i], ...imp, __uid: this.rows[i].__uid };
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

    highlightDuplicateRows(validRows, sigOf, dupSignatures) {
        const tbody = document.getElementById("valuesGridBody");
        if (!tbody) return;
        tbody.querySelectorAll("tr").forEach(tr => tr.classList.remove("duplicate-row"));

        validRows.forEach(r => {
            if (!dupSignatures.has(sigOf(r))) return;
            const tr = tbody.querySelector(`tr[data-uid="${r.__uid}"]`);
            if (tr) tr.classList.add("duplicate-row");
        });
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

        // Validación de clave única: ni BigQuery ni Snowflake la fuerzan a nivel
        // de motor, así que la comprobamos aquí antes de guardar.
        const sigOf = (r) => keyCols_.map(c => String(r[c] ?? "").trim().toUpperCase()).join(" | ");
        const seen = new Map();
        const dupSignatures = new Set();
        validRows.forEach((r, i) => {
            const s = sigOf(r);
            if (seen.has(s)) dupSignatures.add(s);
            else seen.set(s, i);
        });

        if (dupSignatures.size) {
            // Quita cualquier filtro de búsqueda para que las filas duplicadas
            // (que se van a marcar en rojo) sean visibles siempre.
            if (this.searchText) {
                this.searchText = "";
                const searchInput = document.getElementById("valuesSearchInput");
                if (searchInput) searchInput.value = "";
            }
            this.renderGrid();
            this.highlightDuplicateRows(validRows, sigOf, dupSignatures);
            const sample = Array.from(dupSignatures).slice(0, 5).join(" · ");
            UI.toast(
                `No se puede guardar: hay ${dupSignatures.size} valor(es) de clave repetidos (${sample}${dupSignatures.size > 5 ? "…" : ""}). Corrige o elimina las filas duplicadas (marcadas en rojo).`,
                "error"
            );
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

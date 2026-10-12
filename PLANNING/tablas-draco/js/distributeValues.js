/* ==========================================================================
 * DISTRIBUIR VALORES  (Planning Panel estilo SAP Analytics Cloud)
 * --------------------------------------------------------------------------
 * Funcionalidad autocontenida que opera directamente sobre las celdas ya
 * pintadas del informe en la hoja de Excel (no depende del motor de
 * generación de SQL ni de ReportState): usa la selección del usuario y el
 * formato `indentLevel` que el propio "Draco" ya aplica a las celdas del
 * informe para reconocer la jerarquía (ver commands.js -> paintear filas
 * "range.format.indentLevel = ...").
 *
 * Flujo (igual que el Planning Panel de SAC):
 *   1) El usuario selecciona la celda ORIGEN (el total a repartir) y pulsa
 *      "Usar selección actual" en el bloque Origen.
 *   2) Selecciona el RANGO DESTINO (las celdas hijas/hermanas) y pulsa
 *      "Usar selección actual" en el bloque Destino. Si detecta que las
 *      filas del destino tienen indentLevel = indentLevel(origen) + 1 lo
 *      marca como "Hijos detectados automáticamente".
 *   3) Elige el Driver: Valores manuales / Proporcionalmente / Equitativamente.
 *   4) Elige la Operación: Distribuir (desde Origen) / Redistribuir (entre
 *      las propias celdas Destino) y la Acción: Sobrescribir / Añadir.
 *   5) Vista previa en vivo -> Aplicar.
 * ========================================================================== */

const DistributeValuesPanel = {

    // Estado interno del panel
    state: {
        source: null,   // { address, sheetName, row, col, value, indentLevel }
        targets: [],    // [{ address, row, col, value, indentLevel }]
        driver: "equal", // "manual" | "proportional" | "equal"
        operation: "distribute", // "distribute" | "redistribute"
        action: "overwrite",     // "overwrite" | "append"
        manualValues: {},        // address -> valor manual introducido
        updateSourceTotal: true  // al Distribuir, vuelca la suma del destino en el Origen
    },

    /* ---------------------------------------------------------------- *
     * Apertura / cierre del modal
     * ---------------------------------------------------------------- */

    open() {
        const modal = document.getElementById("distributeValuesModal");
        if (!modal) return;
        this.resetState();
        modal.style.display = "flex";
        this.render();
        // Al abrir, intenta tomar automáticamente la celda seleccionada
        // en Excel como Origen, para ahorrar un clic.
        this.captureSource(true);
    },

    close() {
        const modal = document.getElementById("distributeValuesModal");
        if (modal) modal.style.display = "none";
    },

    resetState() {
        this.state = {
            source: null,
            targets: [],
            driver: "equal",
            operation: "distribute",
            action: "overwrite",
            manualValues: {},
            updateSourceTotal: true
        };
    },

    /* ---------------------------------------------------------------- *
     * Lectura de la selección de Excel
     * ---------------------------------------------------------------- */

    async _readSelectionRows() {
        return await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load(["address", "rowCount", "columnCount", "values"]);
            const sheet = range.worksheet;
            sheet.load("name");
            await context.sync();

            const cells = [];
            for (let r = 0; r < range.rowCount; r++) {
                const rowRange = range.getRow(r);
                rowRange.load(["rowIndex", "columnIndex"]);
                rowRange.format.load("indentLevel");
                cells.push({ rowRange, r });
            }
            await context.sync();

            const out = [];
            for (let r = 0; r < range.rowCount; r++) {
                for (let c = 0; c < range.columnCount; c++) {
                    const raw = range.values[r][c];
                    out.push({
                        row: cells[r].rowRange.rowIndex + 1,          // 1-based
                        col: cells[r].rowRange.columnIndex + c + 1,   // 1-based
                        address: this._addressFromRC(cells[r].rowRange.rowIndex + 1, cells[r].rowRange.columnIndex + c + 1),
                        sheetName: sheet.name,
                        value: typeof raw === "number" ? raw : (parseFloat(String(raw).replace(",", ".")) || 0),
                        indentLevel: cells[r].rowRange.format.indentLevel || 0
                    });
                }
            }
            return { sheetName: sheet.name, cells: out };
        });
    },

    _addressFromRC(row1, col1) {
        // Reutiliza el helper compartido si existe (ReportDesignerUtils), si no, uno local.
        if (window.ReportDesignerUtils && typeof window.ReportDesignerUtils.addressFromRC === "function") {
            return window.ReportDesignerUtils.addressFromRC(row1, col1);
        }
        let col = col1, letters = "";
        while (col > 0) {
            const rem = (col - 1) % 26;
            letters = String.fromCharCode(65 + rem) + letters;
            col = Math.floor((col - 1) / 26);
        }
        return `${letters}${row1}`;
    },

    async captureSource(silent) {
        try {
            const { sheetName, cells } = await this._readSelectionRows();
            if (!cells.length) return;
            const first = cells[0];
            this.state.source = { ...first, sheetName };
            this.render();
        } catch (err) {
            if (!silent) alert("No se pudo leer la celda de origen: " + (err.message || err));
            console.warn("[DistributeValues] captureSource:", err);
        }
    },

    async captureTargets() {
        try {
            const { sheetName, cells } = await this._readSelectionRows();
            if (!cells.length) {
                alert("Selecciona primero, en la hoja, el rango de celdas destino.");
                return;
            }
            this.state.targets = cells.map(c => ({ ...c, sheetName }));
            // Conserva los valores manuales ya introducidos para direcciones que sigan presentes
            const keep = {};
            this.state.targets.forEach(t => {
                if (this.state.manualValues[t.address] !== undefined) keep[t.address] = this.state.manualValues[t.address];
            });
            this.state.manualValues = keep;
            this.render();
        } catch (err) {
            alert("No se pudo leer el rango destino: " + (err.message || err));
            console.warn("[DistributeValues] captureTargets:", err);
        }
    },

    /* ---------------------------------------------------------------- *
     * ¿Los destinos son hijos directos del origen? (mismo indentLevel + 1,
     * filas contiguas). Solo es una pista visual para el usuario.
     * ---------------------------------------------------------------- */
    targetsLookLikeChildren() {
        const src = this.state.source;
        const tgts = this.state.targets;
        if (!src || !tgts.length) return false;
        return tgts.every(t => t.indentLevel === (src.indentLevel + 1));
    },

    /* ---------------------------------------------------------------- *
     * Cálculo de reparto (equivalente a los "Driver" del Planning Panel)
     * ---------------------------------------------------------------- */
    computeNewValues() {
        const { targets, driver, operation, action, manualValues } = this.state;
        if (!targets.length) return [];

        // Importe a repartir: en Distribuir, el valor de Origen; en
        // Redistribuir, la suma actual de las propias celdas Destino.
        let amount;
        if (operation === "redistribute") {
            amount = targets.reduce((s, t) => s + (t.value || 0), 0);
        } else {
            amount = this.state.source ? (this.state.source.value || 0) : 0;
        }

        let allocated;
        if (driver === "manual") {
            allocated = targets.map(t => {
                const v = manualValues[t.address];
                return v === undefined || v === "" ? 0 : (parseFloat(String(v).replace(",", ".")) || 0);
            });
        } else if (driver === "equal") {
            const share = targets.length ? amount / targets.length : 0;
            allocated = targets.map(() => share);
        } else { // proportional
            const total = targets.reduce((s, t) => s + Math.abs(t.value || 0), 0);
            if (total === 0) {
                // Sin proporción existente: se reparte a partes iguales como alternativa segura.
                const share = targets.length ? amount / targets.length : 0;
                allocated = targets.map(() => share);
            } else {
                allocated = targets.map(t => amount * (Math.abs(t.value || 0) / total));
            }
        }

        return targets.map((t, i) => {
            const delta = allocated[i];
            const newValue = action === "append" ? (t.value || 0) + delta : delta;
            return { ...t, delta, newValue };
        });
    },

    /* ---------------------------------------------------------------- *
     * Aplicar a la hoja
     * ---------------------------------------------------------------- */
    async apply() {
        const rows = this.computeNewValues();
        if (!rows.length) {
            alert("No hay celdas destino seleccionadas.");
            return;
        }

        const btn = document.getElementById("btnDistribuirAplicar");
        try {
            if (btn) { btn.disabled = true; btn.innerText = "Aplicando…"; }

            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getItem(rows[0].sheetName);
                rows.forEach(r => {
                    const cell = sheet.getRangeByIndexes(r.row - 1, r.col - 1, 1, 1);
                    cell.values = [[this._round(r.newValue)]];
                });

                // En modo "Distribuir", si el usuario lo pide, se vuelca en el
                // Origen la suma de los importes repartidos (equivalente a
                // dejar el total consistente con sus hijos, como hace SAC).
                if (this.state.operation === "distribute" && this.state.updateSourceTotal && this.state.source) {
                    const total = rows.reduce((s, r) => s + r.newValue, 0);
                    const srcCell = sheet.getRangeByIndexes(this.state.source.row - 1, this.state.source.col - 1, 1, 1);
                    srcCell.values = [[this._round(total)]];
                }

                await context.sync();
            });

            // Si el informe tiene un refresco activo, no lo disparamos aquí:
            // esta acción escribe directamente sobre valores ya pintados,
            // igual que una edición manual del usuario en la celda.
            this.close();
        } catch (err) {
            console.error("[DistributeValues] apply:", err);
            alert("Error al aplicar la distribución: " + (err.message || err));
        } finally {
            if (btn) { btn.disabled = false; btn.innerText = "Aplicar"; }
        }
    },

    _round(n) {
        return Math.round((Number(n) || 0) * 1e6) / 1e6;
    },

    /* ---------------------------------------------------------------- *
     * Render
     * ---------------------------------------------------------------- */
    render() {
        this.renderSourceBox();
        this.renderTargetBox();
        this.renderPreview();
        this.updateApplyEnabled();
    },

    renderSourceBox() {
        const el = document.getElementById("dvSourceInfo");
        if (!el) return;
        const s = this.state.source;
        const opDistribute = this.state.operation === "distribute";
        el.innerHTML = s
            ? `<span class="range-address">${s.address}</span><span class="dv-source-value">${this._fmt(s.value)}</span>`
            : `<span class="dv-empty">Selecciona una celda en la hoja y pulsa "Usar selección"</span>`;

        const wrap = document.getElementById("dvSourceBox");
        if (wrap) wrap.style.opacity = opDistribute ? "1" : "0.45";
        const btn = document.getElementById("btnDvUseSource");
        if (btn) btn.disabled = !opDistribute;
    },

    renderTargetBox() {
        const el = document.getElementById("dvTargetInfo");
        if (!el) return;
        const t = this.state.targets;
        if (!t.length) {
            el.innerHTML = `<span class="dv-empty">Selecciona el rango destino en la hoja y pulsa "Usar selección"</span>`;
            return;
        }
        const hint = this.targetsLookLikeChildren()
            ? `<span class="badge success">Hijos directos detectados por sangría</span>`
            : `<span class="badge">Rango libre (${t.length} celdas)</span>`;
        el.innerHTML = `<span class="range-address">${t[0].sheetName}!${t[0].address}${t.length > 1 ? " … " + t[t.length - 1].address : ""}</span>${hint}`;
    },

    renderPreview() {
        const body = document.getElementById("dvPreviewBody");
        if (!body) return;
        const rows = this.computeNewValues();

        if (this.state.driver === "manual") {
            body.innerHTML = rows.map(r => `
                <tr>
                    <td>${r.address}</td>
                    <td class="dv-num">${this._fmt(r.value)}</td>
                    <td>
                        <input type="text" class="dv-manual-input" data-address="${r.address}"
                               value="${this.state.manualValues[r.address] ?? ""}" placeholder="0">
                    </td>
                    <td class="dv-num dv-new-value">${this._fmt(r.newValue)}</td>
                </tr>`).join("");

            body.querySelectorAll(".dv-manual-input").forEach(inp => {
                inp.addEventListener("input", (ev) => {
                    this.state.manualValues[ev.target.dataset.address] = ev.target.value;
                    this.renderPreview();
                });
            });
        } else {
            body.innerHTML = rows.map(r => `
                <tr>
                    <td>${r.address}</td>
                    <td class="dv-num">${this._fmt(r.value)}</td>
                    <td class="dv-num dv-delta">${this._fmt(r.delta)}</td>
                    <td class="dv-num dv-new-value">${this._fmt(r.newValue)}</td>
                </tr>`).join("");
        }

        const totalEl = document.getElementById("dvPreviewTotal");
        if (totalEl) {
            const total = rows.reduce((s, r) => s + r.newValue, 0);
            totalEl.innerText = this._fmt(total);
        }
    },

    updateApplyEnabled() {
        const btn = document.getElementById("btnDistribuirAplicar");
        if (!btn) return;
        const needsSource = this.state.operation === "distribute";
        const ok = this.state.targets.length > 0 && (!needsSource || !!this.state.source);
        btn.disabled = !ok;
    },

    _fmt(n) {
        const num = Number(n) || 0;
        return num.toLocaleString("es-ES", { maximumFractionDigits: 2 });
    },

    /* ---------------------------------------------------------------- *
     * Wiring de eventos (llamado una vez desde taskpane.js)
     * ---------------------------------------------------------------- */
    initEvents() {
        const byId = (id) => document.getElementById(id);

        byId("btnCloseDistribuirModal")?.addEventListener("click", () => this.close());
        byId("btnDistribuirCancelar")?.addEventListener("click", () => this.close());
        byId("btnDistribuirAplicar")?.addEventListener("click", () => this.apply());

        byId("btnDvUseSource")?.addEventListener("click", () => this.captureSource(false));
        byId("btnDvUseTarget")?.addEventListener("click", () => this.captureTargets());

        document.querySelectorAll('input[name="dvOperation"]').forEach(r => {
            r.addEventListener("change", (ev) => {
                this.state.operation = ev.target.value;
                this.render();
            });
        });

        document.querySelectorAll('input[name="dvDriver"]').forEach(r => {
            r.addEventListener("change", (ev) => {
                this.state.driver = ev.target.value;
                const manualHeader = byId("dvManualColHeader");
                if (manualHeader) manualHeader.style.display = ev.target.value === "manual" ? "" : "none";
                this.render();
            });
        });

        document.querySelectorAll('input[name="dvAction"]').forEach(r => {
            r.addEventListener("change", (ev) => {
                this.state.action = ev.target.value;
                this.render();
            });
        });

        byId("dvUpdateSourceTotal")?.addEventListener("change", (ev) => {
            this.state.updateSourceTotal = ev.target.checked;
        });
    }
};

window.DistributeValuesPanel = DistributeValuesPanel;

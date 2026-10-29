/**
 * ReportStore
 * ------------------------------------------------------------------------
 * Permite tener VARIOS informes en el mismo libro (antes solo existía uno,
 * cuyo diseño se escribía directamente en la hoja EDIT_REPORT).
 *
 * Todo lo que antes se escribía en EDIT_REPORT como "diseño" (filtros,
 * filas, columnas, Estático/Dinámico, rangos H10/N10, opciones de campo,
 * propiedades del informe) se guarda ahora en UNA sola clave JSON dentro
 * de Office.context.document.settings — el mismo mecanismo de "roaming
 * settings" que ya usa SemanticModelStore — así que no aparece nunca como
 * hoja ni pestaña oculta.
 *
 * La hoja EDIT_REPORT se SIGUE creando y usando (ver semantic_model.js /
 * ensureTechnicalSheets), pero ahora solo para lo que de verdad tiene que
 * vivir en una celda física:
 *   - D1 / E1: hoja y celda activas al pulsar "Editar informe" (para saber
 *     dónde estaba el usuario).
 *   - X1 / Y1: último SQL y último JSON devueltos por BigQuery.
 *   - A5 y el resto del "picker" de miembros: comunicación por doble clic
 *     con el XLAM.
 *   - B1: flag de reconocimiento de miembros.
 *
 * Para que el resto de commands.js (que lee EDIT_REPORT con coordenadas
 * fila/columna fijas: H10, N10, H12, N12, C15:F, H15:M, N15:S) siga
 * funcionando sin reescribir toda esa lógica, getReportGrid() genera un
 * "grid" {values, startRow:0, startCol:0} con esas mismas coordenadas,
 * construido a partir del JSON. commands.js combina ese grid con la
 * lectura física de EDIT_REPORT (ver getEditReportGrid en commands.js):
 * las coordenadas de diseño las gana siempre el JSON; el resto (D1, E1,
 * X1, Y1, B1, A5...) lo gana la hoja física.
 */

(function () {

    const RS_KEY = "epm_reports";              // JSON: { [reportId]: reportObj }
    const RS_ACTIVE_KEY = "epm_activeReportId"; // id del informe activo
    const RS_SEQ_KEY = "epm_reportSeq";         // último id secuencial usado
    const RS_CLIPBOARD_KEY = "epm_reportClipboard"; // JSON: informe copiado (ver copyReportToClipboard/pasteReportFromClipboard)

    const DEFAULT_REPORT_PROPERTIES = {
        reportName: "",
        suppressZeroRows: false,
        suppressZeroCols: false,
        subtotalsOnTop: false,
        hierarchyNodesOnTop: true,
        overwriteFormats: true,
        overwriteStaticFormats: false,
        autoFitColumns: false,
        planningReport: false
    };

    /* ---------------------------------------------------------------
     * Persistencia bruta (Office roaming settings)
     * ------------------------------------------------------------- */

    function _readStoreRaw() {
        try {
            const raw = Office.context.document.settings.get(RS_KEY);
            if (!raw) return {};
            const parsed = JSON.parse(raw);
            return (parsed && typeof parsed === "object") ? parsed : {};
        } catch (e) {
            console.error("ReportStore: JSON de epm_reports corrupto, se reinicia.", e);
            return {};
        }
    }

    function _saveAsync() {
        return new Promise((resolve, reject) => {
            Office.context.document.settings.saveAsync((asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    reject(asyncResult.error);
                } else {
                    resolve();
                }
            });
        });
    }

    async function _writeStoreRaw(storeObj) {
        Office.context.document.settings.set(RS_KEY, JSON.stringify(storeObj));
        await _saveAsync();
    }

    /* ---------------------------------------------------------------
     * Utilidades de grid (copia local, igual que en commands.js/
     * excelService.js) para poder reutilizar MODEL_HIER al expandir
     * jerarquías en filas/columnas.
     * ------------------------------------------------------------- */

    function cellValue(grid, row1, col1) {
        const r = row1 - 1 - grid.startRow;
        const c = col1 - 1 - grid.startCol;
        if (r < 0 || c < 0 || r >= grid.values.length) return "";
        const rowArr = grid.values[r];
        if (!rowArr || c >= rowArr.length) return "";
        const v = rowArr[c];
        return v === null || v === undefined ? "" : v;
    }

    function lastRowInColumnValues(grid, col1) {
        const c = col1 - 1 - grid.startCol;
        for (let r = grid.values.length - 1; r >= 0; r--) {
            const rowArr = grid.values[r];
            const v = rowArr && c >= 0 && c < rowArr.length ? rowArr[c] : "";
            if (v !== null && v !== undefined && String(v) !== "") {
                return r + 1 + grid.startRow;
            }
        }
        return grid.startRow;
    }

    function existsHierarchy(hierGrid, dimension, hierarchy) {
        const lastRow = lastRowInColumnValues(hierGrid, 1);
        for (let R = 2; R <= lastRow; R++) {
            if (String(cellValue(hierGrid, R, 4)).toUpperCase() === String(dimension).toUpperCase()
                && String(cellValue(hierGrid, R, 2)).toUpperCase() === String(hierarchy).toUpperCase()) {
                return true;
            }
        }
        return false;
    }

    /* ---------------------------------------------------------------
     * API pública: informe a informe
     * ------------------------------------------------------------- */

    function getAllReports() {
        return _readStoreRaw();
    }

    // Lista ordenada por id de creación (orden en el que se fueron
    // añadiendo), tal y como se deben mostrar en el listbox "Informe".
    function listReports() {
        const store = _readStoreRaw();
        return Object.keys(store)
            .map(id => Number(id))
            .sort((a, b) => a - b)
            .map(id => {
                const r = store[String(id)];
                return { id, name: r.name, semanticModelName: r.semanticModelName || "" };
            });
    }

    function getReport(reportId) {
        const store = _readStoreRaw();
        return store[String(reportId)] || null;
    }

    async function _writeReport(reportId, reportObj) {
        const store = _readStoreRaw();
        store[String(reportId)] = reportObj;
        await _writeStoreRaw(store);
    }

    function _nextSequence() {
        const raw = Office.context.document.settings.get(RS_SEQ_KEY);
        // Nunca reutilizar un id existente aunque el contador falte o vaya por detrás
        // (libros antiguos, informes pegados...): si no, createReport sobrescribiría
        // un informe sin avisar.
        const maxId = Object.keys(_readStoreRaw()).reduce((m, k) => Math.max(m, Number(k) || 0), 0);
        const next = Math.max(Number(raw) || 0, maxId) + 1;
        Office.context.document.settings.set(RS_SEQ_KEY, next);
        return next;
    }

    function _pad3(n) {
        return String(n).padStart(3, "0");
    }

    /**
     * Crea un informe nuevo y VACÍO (sin filtros/filas/columnas), lo marca
     * como activo y lo devuelve. El id interno es el secuencial (nunca se
     * reutiliza, aunque se borren informes). El nombre por defecto es
     * "Informe - 0XX".
     *
     * resultSheetName (opcional): hoja donde se va a pintar ESTE informe.
     * El add-in YA NO crea una pestaña nueva por informe: addReport() en
     * taskpane.js captura la hoja activa de Excel en el momento de pulsar
     * "Añadir informe" (ver captureActiveEditContext) y la pasa aquí, así
     * que el informe se queda anclado en la pestaña/celda donde estaba el
     * cursor. Si no se recibe (compatibilidad con llamadas antiguas), se
     * guarda vacío y el código que lee la hoja de resultados
     * (resultSheetNameFromGrid / getDracoResultSheetName en commands.js)
     * usa como último recurso la celda física compartida EDIT_REPORT!D1.
     */
    async function createReport(semanticModelName, resultSheetName) {
        const id = _nextSequence();
        const name = `Informe - ${_pad3(id)}`;
        const report = {
            id,
            name,
            semanticModelName: semanticModelName || "",
            design: {
                filters: [],
                rows: [],
                columns: [],
                expandedRows: [],
                expandedCols: [],
                rowsStatic: false,
                colsStatic: false,
                rrAddress: "",
                rcAddress: "",
                resultSheetName: resultSheetName || "",
                fieldOptions: {}
            },
            reportProperties: Object.assign({}, DEFAULT_REPORT_PROPERTIES, { reportName: name }),
            memberRecognition: false
        };
        await _writeReport(id, report);
        await _saveAsync();
        await setActiveReportId(id);
        return report;
    }

    /**
     * Hoja de resultados guardada para ESTE informe ("" si es un informe
     * antiguo, creado antes de que esto se guardara por informe: en ese
     * caso el llamante debe usar el fallback físico de EDIT_REPORT!D1, ver
     * resultSheetNameFromGrid en commands.js).
     */
    function getReportResultSheetName(reportId) {
        const report = getReport(reportId);
        return report && report.design ? (report.design.resultSheetName || "") : "";
    }

    /**
     * Fija la hoja de resultados de un informe concreto (llamado desde
     * "Editar report" en el taskpane, cuando el usuario reubica manualmente
     * dónde se pinta). Queda guardada solo para ESE informe: ya no se
     * comparte con el resto a través de una única celda física.
     */
    async function setReportResultSheetName(reportId, sheetName) {
        const report = getReport(reportId);
        if (!report) return;
        report.design = report.design || {};
        report.design.resultSheetName = sheetName || "";
        await _writeReport(reportId, report);
    }

    /* ---------------------------------------------------------------
     * "Copiar informe" / "Pegar informe" (botones del ribbon, ver
     * copiarInforme/pegarInforme en commands.js): un portapapeles con UN
     * único informe (el último copiado), guardado en Office roaming
     * settings -sobrevive a cerrar y reabrir el libro, y es visible desde
     * cualquier contexto (ribbon o taskpane), igual que el resto de
     * ReportStore-. "Pegar" crea un informe NUEVO (mismo camino que
     * createReport: id nuevo, anclado en la hoja que esté activa en ESE
     * momento), copiando el diseño (filtros/filas/columnas/Estático/
     * opciones de campo) y las Propiedades del informe del copiado, pero
     * NO su nombre ni su hoja/rangos de resultados -eso es propio de
     * cada informe, no algo que tenga sentido copiar-.
     * ------------------------------------------------------------- */

    function hasReportClipboard() {
        try {
            return !!Office.context.document.settings.get(RS_CLIPBOARD_KEY);
        } catch (e) {
            return false;
        }
    }

    function getReportClipboard() {
        try {
            const raw = Office.context.document.settings.get(RS_CLIPBOARD_KEY);
            if (!raw) return null;
            const parsed = JSON.parse(raw);
            return (parsed && typeof parsed === "object") ? parsed : null;
        } catch (e) {
            console.error("ReportStore: JSON de epm_reportClipboard corrupto.", e);
            return null;
        }
    }

    async function copyReportToClipboard(reportId) {
        const report = getReport(reportId);
        if (!report) return null;

        const design = report.design || {};
        const clip = {
            sourceName: report.name,
            semanticModelName: report.semanticModelName || "",
            design: {
                filters: design.filters || [],
                rows: design.rows || [],
                columns: design.columns || [],
                rowsStatic: !!design.rowsStatic,
                colsStatic: !!design.colsStatic,
                fieldOptions: design.fieldOptions || {},
                // Posición de los ejes (rangos de Filas y Columnas) y hoja de
                // origen: al pegar se recolocan respecto a la celda activa y,
                // si un eje es Estático, se copian sus celdas.
                rrAddress: design.rrAddress || "",
                rcAddress: design.rcAddress || "",
                sourceSheet: design.resultSheetName || ""
            },
            reportProperties: Object.assign({}, DEFAULT_REPORT_PROPERTIES, report.reportProperties || {})
        };

        Office.context.document.settings.set(RS_CLIPBOARD_KEY, JSON.stringify(clip));
        await _saveAsync();
        return { id: reportId, name: report.name };
    }

    /**
     * Crea un informe a partir del portapapeles. resultSheetName: hoja
     * activa en el momento de pulsar "Pegar informe" (igual que
     * createReport/addReport). Devuelve el informe recién creado, o null
     * si no había nada copiado.
     */
    /* ---- direcciones A1 (para recolocar los ejes al pegar) ---- */
    function _colToNum(c) { let n = 0; for (const ch of c.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64); return n; }
    function _numToCol(n) { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
    function _parseA1(a) {
        const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(String(a || "").trim());
        return m ? { col: _colToNum(m[1]), row: Number(m[2]) } : null;
    }
    function _rangeTopLeft(addr) { return addr ? _parseA1(String(addr).split("!").pop().split(":")[0]) : null; }
    function _shiftRange(addr, dRow, dCol) {
        if (!addr) return "";
        return String(addr).split("!").pop().split(":").map(part => {
            const p = _parseA1(part);
            return p ? _numToCol(Math.max(1, p.col + dCol)) + Math.max(1, p.row + dRow) : part;
        }).join(":");
    }

    /**
     * Pega el informe copiado. anchorAddress: celda activa al pegar (p. ej.
     * "H20"). Los ejes se recolocan con el MISMO desplazamiento: la esquina
     * superior izquierda del informe copiado (fila de Columnas, columna de
     * Filas) pasa a esa celda, y Filas/Columnas conservan su posición
     * relativa y su tamaño. Antes no se copiaban los rangos y el informe
     * pegado no sabía dónde pintarse ni dónde leer sus ejes Estáticos.
     * Devuelve el informe con __staticCopies: [{ fromSheet, from, to }] con
     * los ejes Estáticos cuyas celdas hay que copiar (lo hace commands.js).
     */
    async function pasteReportFromClipboard(resultSheetName, anchorAddress) {
        const clip = getReportClipboard();
        if (!clip) return null;

        const report = await createReport(clip.semanticModelName || "", resultSheetName);

        const d0 = clip.design || {};
        const rrTL = _rangeTopLeft(d0.rrAddress), rcTL = _rangeTopLeft(d0.rcAddress);
        const anchor = _parseA1(anchorAddress);
        let rrAddress = d0.rrAddress || "", rcAddress = d0.rcAddress || "";
        const staticCopies = [];
        if (anchor && (rrTL || rcTL)) {
            const top = Math.min(rrTL ? rrTL.row : Infinity, rcTL ? rcTL.row : Infinity);
            const left = Math.min(rrTL ? rrTL.col : Infinity, rcTL ? rcTL.col : Infinity);
            const dRow = anchor.row - top, dCol = anchor.col - left;
            rrAddress = _shiftRange(d0.rrAddress, dRow, dCol);
            rcAddress = _shiftRange(d0.rcAddress, dRow, dCol);
        }
        if (d0.sourceSheet) {
            if (d0.rowsStatic && d0.rrAddress && rrAddress) staticCopies.push({ fromSheet: d0.sourceSheet, from: d0.rrAddress, to: rrAddress });
            if (d0.colsStatic && d0.rcAddress && rcAddress) staticCopies.push({ fromSheet: d0.sourceSheet, from: d0.rcAddress, to: rcAddress });
        }

        // saveDesign ya expande jerarquías (expandedRows/expandedCols) a
        // partir de filters/rows/columns/fieldOptions -mismo camino que
        // sigue el taskpane al autoguardar el diseño-, así que no hace
        // falta duplicar esa lógica aquí; y (con el fix de más arriba)
        // conserva el resultSheetName que acaba de fijar createReport.
        await saveDesign(report.id, {
            filters: clip.design.filters,
            rows: clip.design.rows,
            columns: clip.design.columns,
            rowsStatic: clip.design.rowsStatic,
            colsStatic: clip.design.colsStatic,
            fieldOptions: clip.design.fieldOptions,
            rrAddress,
            rcAddress
        });

        // Propiedades del informe (subtotales, jerarquía arriba, suprimir
        // ceros...) del copiado, conservando el NOMBRE autogenerado del
        // informe nuevo ("Informe - 0XX"), no el del copiado -para no
        // acabar con dos informes con el mismo nombre-.
        const props = Object.assign({}, clip.reportProperties || {});
        delete props.reportName;
        await saveReportProperties(report.id, props);

        const out = getReport(report.id);
        if (out) out.__staticCopies = staticCopies;
        return out;
    }

    async function deleteReport(reportId) {
        const store = _readStoreRaw();
        delete store[String(reportId)];
        await _writeStoreRaw(store);

        if (Number(Office.context.document.settings.get(RS_ACTIVE_KEY)) === Number(reportId)) {
            const remaining = Object.keys(store).map(Number).sort((a, b) => a - b);
            await setActiveReportId(remaining[0] || null);
        }
    }

    /**
     * Fija algunos campos de la definición de un informe sin pasar por el
     * taskpane (saveDesign reconstruye la definición entera a partir del
     * estado del panel). Lo usan los informes vinculados (commands.js:
     * design.link y el eje que se copia del principal). Devuelve el informe.
     */
    async function updateDesignFields(reportId, fields) {
        const report = getReport(reportId);
        if (!report) return null;
        report.design = Object.assign({}, report.design || {}, fields || {});
        await _writeReport(reportId, report);
        await _saveAsync();
        return report;
    }

    async function renameReport(reportId, newName) {
        const report = getReport(reportId);
        if (!report) return;
        report.name = newName;
        report.reportProperties = report.reportProperties || {};
        report.reportProperties.reportName = newName;
        await _writeReport(reportId, report);
    }

    /**
     * Informe "activo": el que el taskpane está editando ahora mismo. Si el
     * marcado como activo ya no existe, se usa el primero (por orden de
     * creación) que quede, o null si no hay ninguno (taskpane vacío).
     */
    function getActiveReportId() {
        const raw = Office.context.document.settings.get(RS_ACTIVE_KEY);
        const active = raw ? Number(raw) : null;
        const store = _readStoreRaw();
        if (active && store[String(active)]) return active;

        const ids = Object.keys(store).map(Number).sort((a, b) => a - b);
        return ids.length > 0 ? ids[0] : null;
    }

    async function setActiveReportId(reportId) {
        Office.context.document.settings.set(RS_ACTIVE_KEY, reportId || "");
        await _saveAsync();
    }

    function getActiveReport() {
        const id = getActiveReportId();
        return id ? getReport(id) : null;
    }

    // Propiedades del informe activo (o los valores por defecto si todavía
    // no hay ningún informe). Sustituye a la antigua clave global
    // "draco_reportProperties": accesible en sync desde cualquier
    // contexto (taskpane o commands.html/ribbon).
    function getActiveReportProperties() {
        const report = getActiveReport();
        if (!report) return Object.assign({}, DEFAULT_REPORT_PROPERTIES);
        return Object.assign({}, DEFAULT_REPORT_PROPERTIES, report.reportProperties || {});
    }

    async function saveReportProperties(reportId, props) {
        const report = getReport(reportId);
        if (!report) return;
        report.reportProperties = Object.assign({}, DEFAULT_REPORT_PROPERTIES, report.reportProperties, props);
        if (props && props.reportName) report.name = props.reportName;
        await _writeReport(reportId, report);
    }

    async function setSemanticModel(reportId, semanticModelName) {
        const report = getReport(reportId);
        if (!report) return;
        report.semanticModelName = semanticModelName || "";
        await _writeReport(reportId, report);
    }

    /* ---------------------------------------------------------------
     * Diseño (filtros/filas/columnas/Estático/rangos/opciones de campo):
     * expande jerarquías (vía MODEL_HIER) EN EL MOMENTO DE GUARDAR, igual
     * que hacía antes excelService.saveEditReportDesign al escribir en la
     * hoja, para que getReportGrid() no tenga que repetir el trabajo cada
     * vez que se lee.
     * ------------------------------------------------------------- */

    /* ---------------------------------------------------------------
     * "Mostrar ID y texto": el MODELO define el atributo de texto (dimensión:
     *   field.textAttribute, marcado con "Texto" junto a "Key"; jerarquía:
     *   levels[i].textAttribute) y el INFORME elige el modo por campo
     *   (fieldOptions["zona|dim|campo"].displayMode = "ID" | "ID_TEXT" | "TEXT").
     *   "TEXT" = mismo SQL que "ID_TEXT" (se agrupa por ID y texto) pero la
     *   columna de ID NO se pinta: sus entradas llevan "HIDE" en la columna de
     *   orden (M/S de EDIT_REPORT), que commands.js lee como oculta.
     *   - Dimensión: al colocar en Filas/Columnas el atributo CLAVE de la
     *     dimensión, con "ID_TEXT" se añade justo detrás el atributo de texto: es EXACTAMENTE
     *     igual que si el usuario hubiese arrastrado los dos campos (dos
     *     columnas), así que el motor no necesita saber nada de esto.
     *   - Jerarquía: levels[i].textAttribute de cada nivel. Con "ID_TEXT"
     *     se añade, tras los niveles de ID, una SEGUNDA jerarquía con los
     *     atributos de texto de cada nivel (misma sangría/nodos, otra
     *     columna). Va marcada con el sufijo TEXT_SUFFIX en el nombre de
     *     jerarquía para que commands.js sepa que sus niveles van "pegados" a
     *     los de ID (mismos grupos de GROUPING SETS, ver
     *     computeGroupingSetDimensionsAndHierarchies).
     * ------------------------------------------------------------- */

    const TEXT_SUFFIX = "~TXT";

    const _up = v => String(v === null || v === undefined ? "" : v).trim().toUpperCase();

    function _modelForReport(report) {
        const store = window.SemanticModelStore;
        if (!store) return { model: null, name: undefined };
        const wanted = report && report.semanticModelName;
        if (wanted && store.getModel(wanted)) return { model: store.getModel(wanted), name: wanted };
        const active = store.getActiveModelName();
        return { model: active ? store.getModel(active) : null, name: undefined };
    }

    function _dimensionField(model, dimension) {
        if (!model || !Array.isArray(model.fields)) return null;
        return model.fields.find(f => f.enabled && f.type === "DIMENSION" && _up(f.name) === _up(dimension)) || null;
    }

    // Atributo de texto a añadir tras el atributo `attrName` de la dimensión
    // (solo si es su clave y la dimensión está en "ID y texto"); null si no aplica.
    function _flatTextAttribute(field, attrName) {
        if (!field || !field.textAttribute) return null;
        const attrs = field.attributes || [];
        const key = attrs.find(a => a.isKey && a.enabled !== false);
        if (!key || _up(key.name) !== _up(attrName)) return null;
        const txt = attrs.find(a => a.enabled && _up(a.name) === _up(field.textAttribute));
        if (!txt || _up(txt.name) === _up(key.name)) return null;
        return txt.name;
    }

    // { [nivel]: atributoTexto } de una jerarquía en "ID y texto", o null si
    // no aplica (modo "ID", o algún nivel visible sin texto válido/repetido).
    function _hierarchyTextAttributes(field, hierName, visibleLevels) {
        if (!field) return null;
        const hier = (field.hierarchies || []).find(h => _up(h.name) === _up(hierName));
        if (!hier) return null;

        const attrs = field.attributes || [];
        const levelAttrs = new Set((hier.levels || []).map(l => _up(l.attribute)));
        const used = new Set();
        const out = {};
        for (let i = 0; i < (hier.levels || []).length; i++) {
            const nivel = i + 1;
            if (visibleLevels && !visibleLevels.includes(nivel)) continue;
            const t = hier.levels[i].textAttribute;
            const attr = t ? attrs.find(a => a.enabled && _up(a.name) === _up(t)) : null;
            if (!attr || levelAttrs.has(_up(attr.name)) || used.has(_up(attr.name))) return null;
            used.add(_up(attr.name));
            out[nivel] = attr.name;
        }
        return Object.keys(out).length ? out : null;
    }

    /**
     * Expande una zona (Filas o Columnas) del diseño a la lista que lee el
     * motor: [dim, atributo, nivel, jerarquía, subtotal, orden] por campo.
     * ctx = { hierGrid, model }.
     */
    function _expandAxis(list, zoneId, fieldOptions, ctx) {
        const { hierGrid, model } = ctx;

        const subtotalAndOrder = (dimension, name) => {
            const opts = (fieldOptions || {})[`${zoneId}|${dimension}|${name}`] || {};
            const subtotal = opts.showTotals ? "X" : "";
            const order = opts.sortOrder === "asc" ? "UP" : (opts.sortOrder === "desc" ? "DOWN" : "");
            return [subtotal, order];
        };

        // Niveles de la jerarquía que se consultan y se pintan: los marcados
        // en «Mostrar niveles» (null = todos). «Expandir hasta nivel» NO
        // quita niveles: se consulta todo y, al pintar, la jerarquía sale
        // contraída en ese nivel (ver dracoExpandToLevelCollapse en commands.js).
        const visibleLevelsFor = (dimension, name) => {
            const opts = (fieldOptions || {})[`${zoneId}|${dimension}|${name}`] || {};
            return Array.isArray(opts.visibleLevels) ? opts.visibleLevels.map(Number) : null;
        };
        const levelIsVisible = (visibleLevels, nivel) => !visibleLevels || visibleLevels.includes(nivel);

        // ¿Ya está el atributo `name` de `dimension` como campo suelto en esta zona?
        const displayModeOf = (dimension, name) => {
            const opts = (fieldOptions || {})[`${zoneId}|${dimension}|${name}`] || {};
            return opts.displayMode === "ID_TEXT" || opts.displayMode === "TEXT" ? opts.displayMode : "ID";
        };
        const wantsIdText = (dimension, name) => displayModeOf(dimension, name) !== "ID";

        const hasFlat = (dimension, name) => list.some(it => !it.isHierarchy
            && _up(it.dimension) === _up(dimension) && _up(it.name) === _up(name));

        const out = [];
        list.forEach(item => {
            if (item.isHierarchy && existsHierarchy(hierGrid, item.dimension, item.name)) {
                const visibleLevels = visibleLevelsFor(item.dimension, item.name);
                const lastRow = lastRowInColumnValues(hierGrid, 4);
                // Niveles de la jerarquía, en orden, con los ocultos fuera.
                const defs = [];
                for (let R = 2; R <= lastRow; R++) {
                    if (String(cellValue(hierGrid, R, 4)).toUpperCase() === item.dimension.toUpperCase()
                        && String(cellValue(hierGrid, R, 2)).toUpperCase() === item.name.toUpperCase()) {
                        const nivel = Number(cellValue(hierGrid, R, 3));
                        if (!levelIsVisible(visibleLevels, nivel)) continue;
                        defs.push({ nivel, dim: cellValue(hierGrid, R, 4), attr: cellValue(hierGrid, R, 5), hier: cellValue(hierGrid, R, 2) });
                    }
                }
                defs.sort((a, b) => a.nivel - b.nivel);
                // Se renumeran 1, 2, 3... SIN huecos: el motor solo reconoce
                // una jerarquía si sus niveles son consecutivos (ver
                // pushHierarchyRunLengths en commands.js). Quitar niveles da
                // así lo mismo que una jerarquía creada sin ellos (antes, con
                // un hueco, los niveles se pintaban como campos sueltos).
                const niveles = defs.map(d => d.nivel); // nº original (para los textos)
                const hostStart = out.length;
                defs.forEach((d, i) => {
                    out.push([
                        d.dim,        // DIMENSION
                        d.attr,       // ATRIBUTO
                        i + 1,        // NIVEL (consecutivo)
                        d.hier,       // JERARQUIA
                        "", ""        // subtotal/orden: no aplica a jerarquías
                    ]);
                });
                // "ID y texto": segunda jerarquía con los textos de cada nivel.
                const texts = wantsIdText(item.dimension, item.name)
                    ? _hierarchyTextAttributes(_dimensionField(model, item.dimension), item.name,
                        niveles.length ? niveles : null) : null;
                if (texts && niveles.every(n => texts[n])) {
                    niveles.forEach((n, i) => {
                        out.push([item.dimension, texts[n], i + 1, item.name + TEXT_SUFFIX, "", ""]);
                    });
                    // "Solo texto": los niveles de ID se consultan pero no se pintan.
                    if (displayModeOf(item.dimension, item.name) === "TEXT") {
                        for (let k = hostStart; k < hostStart + niveles.length; k++) out[k][5] = "HIDE";
                    }
                }
            } else {
                const [subtotal, order] = subtotalAndOrder(item.dimension, item.name);
                const idRow = [item.dimension, item.name, 1, "", subtotal, order];
                out.push(idRow);
                // "ID y texto": el atributo de texto va justo detrás de la clave.
                if (!item.isHierarchy && wantsIdText(item.dimension, item.name)) {
                    const txt = _flatTextAttribute(_dimensionField(model, item.dimension), item.name);
                    if (txt && !hasFlat(item.dimension, txt)) {
                        out.push([item.dimension, txt, 1, "", "", ""]);
                        // "Solo texto": el ID se consulta (agrupa) pero no se pinta.
                        if (displayModeOf(item.dimension, item.name) === "TEXT") {
                            idRow[5] = order ? order + "|HIDE" : "HIDE";
                        }
                    }
                }
            }
        });
        return out;
    }

    /**
     * ¿Se puede ofrecer "ID y texto" para este campo del informe? (lo usa el
     * panel "Opciones de campo"). Devuelve
     *   { applicable, available, reason }
     * applicable: el campo admite el modo (clave de la dimensión o jerarquía);
     * available : además el modelo tiene el/los atributo(s) de texto marcados.
     */
    function getDisplayTextInfo(reportId, dimension, name, isHierarchy) {
        const report = getReport(reportId);
        const field = _dimensionField(_modelForReport(report).model, dimension);
        if (!field) return { applicable: false, available: false, reason: "" };

        if (isHierarchy) {
            const hier = (field.hierarchies || []).find(h => _up(h.name) === _up(name));
            if (!hier) return { applicable: false, available: false, reason: "" };
            const ok = !!_hierarchyTextAttributes(field, name, null);
            return { applicable: true, available: ok, reason: ok ? "" : "Define el atributo de texto de todos los niveles de la jerarquía en el modelo semántico." };
        }

        const key = (field.attributes || []).find(a => a.isKey && a.enabled !== false);
        if (!key || _up(key.name) !== _up(name)) return { applicable: false, available: false, reason: "" };
        const ok = !!_flatTextAttribute(field, name);
        return { applicable: true, available: ok, reason: ok ? "" : "Marca un atributo como \"Texto\" en la dimensión del modelo semántico." };
    }

    function _expansionContext(report) {
        const { model, name } = _modelForReport(report);
        // Mismo modelo para la jerarquía y para los textos (antes: siempre el "activo").
        return { hierGrid: window.SemanticModelStore.getModelGridSync("MODEL_HIER", name), model };
    }

    /**
     * Nº de columnas/filas FÍSICAS extra que aportan los "ID y texto" de una
     * zona ("rows" | "columns") del diseño (list = campos de la zona): el
     * taskpane lo usa para ensanchar/alargar el rango del eje y que el texto
     * no se solape con el otro eje.
     */
    function countTextColumns(reportId, zoneId, list, fieldOptions) {
        const report = getReport(reportId);
        if (!report) return 0;
        const ctx = _expansionContext(report);
        if (!ctx.model) return 0;
        const stripped = {};
        Object.keys(fieldOptions || {}).forEach(k => {
            const o = Object.assign({}, fieldOptions[k]);
            delete o.displayMode;
            stripped[k] = o;
        });
        // columnas físicas = entradas de nivel 1 que se pintan (las "HIDE" no)
        const physical = rows => rows.filter(r => Number(r[2]) === 1 && String(r[5] || "").indexOf("HIDE") === -1).length;
        const withText = physical(_expandAxis(list || [], zoneId, fieldOptions, ctx));
        const without = physical(_expandAxis(list || [], zoneId, stripped, ctx));
        return Math.max(0, withText - without);
    }

    /**
     * Recalcula expandedRows/expandedCols de un informe a partir de su diseño
     * (rows/columns/fieldOptions) y del modelo semántico ACTUAL, para que un
     * cambio de "ID / ID y texto" (o de jerarquías) en el modelo se aplique
     * sin tener que tocar el informe. Solo escribe si algo ha cambiado.
     */
    async function refreshExpansion(reportId) {
        const report = getReport(reportId);
        const d = report && report.design;
        if (!d || !Array.isArray(d.rows) || !Array.isArray(d.columns)) return;

        const ctx = await _expansionContext(report);
        if (!ctx.model) return;

        const expandedRows = _expandAxis(d.rows, "rows", d.fieldOptions, ctx);
        const expandedCols = _expandAxis(d.columns, "columns", d.fieldOptions, ctx);

        if (JSON.stringify(expandedRows) === JSON.stringify(d.expandedRows || [])
            && JSON.stringify(expandedCols) === JSON.stringify(d.expandedCols || [])) return;

        d.expandedRows = expandedRows;
        d.expandedCols = expandedCols;
        await _writeReport(reportId, report);
    }

    async function saveDesign(reportId, state) {
        const report = getReport(reportId);
        if (!report) return;

        const ctx = await _expansionContext(report);

        // OJO: se parte del report.design ANTERIOR (spread) y solo se
        // sobrescriben los campos que gestiona esta función. Antes se
        // reconstruía el objeto entero desde cero, lo que BORRABA en
        // silencio cualquier campo que no gestiona saveDesign -en
        // concreto, resultSheetName (ver createReport/
        // setReportResultSheetName)-: cada autoguardado del taskpane
        // (con cualquier cambio de diseño) dejaba el informe sin la hoja
        // de resultados asignada, así que dejaba de encontrarse desde
        // reportIdForResultSheet (commands.js) en cuanto el usuario
        // tocaba algo. Con el spread, resultSheetName (y cualquier otro
        // campo futuro no gestionado aquí) se conserva.
        report.design = {
            ...(report.design || {}),
            filters: (state.filters || []).map(f => ({
                dimension: f.dimension, name: f.name, isHierarchy: f.isHierarchy,
                realAttribute: f.realAttribute,
                // El filtro (multi-valor / rango / incluir-excluir) se guarda
                // como JSON en la misma celda "Valor" que antes tenía un
                // valor simple; commands.js (parseFilterValue) sabe leer
                // ambos formatos.
                value: f.filter ? JSON.stringify(f.filter) : ""
            })),
            rows: (state.rows || []).map(r => ({ dimension: r.dimension, name: r.name, isHierarchy: r.isHierarchy })),
            columns: (state.columns || []).map(c => ({ dimension: c.dimension, name: c.name, isHierarchy: c.isHierarchy })),
            expandedRows: _expandAxis(state.rows || [], "rows", state.fieldOptions, ctx),
            expandedCols: _expandAxis(state.columns || [], "columns", state.fieldOptions, ctx),
            rowsStatic: !!state.rowsStatic,
            colsStatic: !!state.colsStatic,
            rrAddress: state.rrAddress || "",
            rcAddress: state.rcAddress || "",
            fieldOptions: state.fieldOptions || {}
        };

        await _writeReport(reportId, report);
    }

    async function setMemberRecognition(reportId, value) {
        const report = getReport(reportId);
        if (!report) return;
        report.memberRecognition = !!value;
        await _writeReport(reportId, report);
    }

    /**
     * Genera el grid virtual {values, startRow:0, startCol:0} con EXACTAMENTE
     * las mismas coordenadas que tenía el diseño en EDIT_REPORT:
     *   H10 / N10  -> rrAddress / rcAddress
     *   H12 / N12  -> rowsStatic / colsStatic ("X"/"")
     *   C15:F..    -> filtros (Dimension, Atributo real, Valor, Jerarquía)
     *   H15:M..    -> filas expandidas (Dim, Atributo, Nivel, Jerarquía, Subtotal, Orden)
     *   N15:S..    -> columnas expandidas (Dim, Atributo, Nivel, Jerarquía, Subtotal, Orden)
     * commands.js combina este grid con la lectura física de EDIT_REPORT
     * (ver getEditReportGrid/isDesignOwnedCell) para las coordenadas que
     * SÍ siguen viviendo en la hoja (D1/E1/X1/Y1/B1/A5...).
     */
    function getReportGrid(reportId) {
        const report = getReport(reportId);
        const design = report ? (report.design || {}) : {};

        const filters = design.filters || [];
        const expandedRows = design.expandedRows || [];
        const expandedCols = design.expandedCols || [];

        const maxRow = 14 + Math.max(filters.length, expandedRows.length, expandedCols.length, 0);
        const values = [];
        for (let r = 0; r < maxRow; r++) values.push(new Array(19).fill(""));

        const set = (row1, col1, v) => {
            if (row1 - 1 < values.length) values[row1 - 1][col1 - 1] = (v === null || v === undefined) ? "" : v;
        };

        set(10, 8, design.rrAddress || "");
        set(10, 14, design.rcAddress || "");
        set(12, 8, design.rowsStatic ? "X" : "");
        set(12, 14, design.colsStatic ? "X" : "");

        filters.forEach((f, i) => {
            const row = 15 + i;
            set(row, 3, f.dimension || "");
            set(row, 4, f.realAttribute || "");
            set(row, 5, f.value || "");
            set(row, 6, f.isHierarchy ? f.name : "");
        });

        expandedRows.forEach((cols, i) => {
            const row = 15 + i;
            set(row, 8, cols[0]); set(row, 9, cols[1]); set(row, 10, cols[2]); set(row, 11, cols[3]);
            set(row, 12, cols[4]); set(row, 13, cols[5]);
        });

        expandedCols.forEach((cols, i) => {
            const row = 15 + i;
            set(row, 14, cols[0]); set(row, 15, cols[1]); set(row, 16, cols[2]); set(row, 17, cols[3]);
            set(row, 18, cols[4]); set(row, 19, cols[5]);
        });

        return { values, startRow: 0, startCol: 0 };
    }

    window.ReportStore = {
        getAllReports,
        listReports,
        getReport,
        createReport,
        deleteReport,
        renameReport,
        getActiveReportId,
        setActiveReportId,
        getActiveReport,
        getActiveReportProperties,
        saveReportProperties,
        setSemanticModel,
        saveDesign,
        refreshExpansion,
        getDisplayTextInfo,
        countTextColumns,
        TEXT_SUFFIX,
        updateDesignFields,
        setMemberRecognition,
        getReportGrid,
        getReportResultSheetName,
        setReportResultSheetName,
        hasReportClipboard,
        getReportClipboard,
        copyReportToClipboard,
        pasteReportFromClipboard
    };

})();

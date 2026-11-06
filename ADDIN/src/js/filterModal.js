/* ==========================================================================
 * FilterModal — puente hacia el diálogo independiente de Office
 * "Filtrar campo" (filterDialog.html), usado desde la zona "Filtros" del
 * taskpane.
 *
 * Este fichero pinta un overlay DENTRO de taskpane.html? No: el selector
 * vive en su propia ventana (igual que memberPicker.html), lo que le da
 * sitio para un panel auxiliar a la derecha; ver js/filterDialog.js +
 * filterDialog.html para la UI real.
 *
 * Un diálogo de Office no tiene acceso al modelo de objetos de Excel, así
 * que aquí (en el taskpane, donde SÍ está disponible window.ExcelService)
 * se resuelve primero la lista de valores y luego se abre el diálogo,
 * enviándosela por mensaje en cuanto avisa que está listo — mismo patrón
 * que openMemberRecognitionPicker en commands.js.
 *
 * El contrato hacia quien llama (taskpane.js) NO ha cambiado:
 * FilterModal.open(fieldData) sigue devolviendo una Promise que resuelve a
 * un objeto "filter" o a null si se cancela. Formato actual (mode:"list"):
 * el usuario puede combinar libremente valores sueltos y rangos,
 * incluidos y excluidos, cada uno como una selección independiente (ver
 * cabecera de js/filterDialog.js para el detalle):
 *
 *   Dimensión plana:
 *     { mode:"list", attribute:"PAIS",
 *       values:["ES","FR"], ranges:[{from:"A",to:"M"}],
 *       excludeValues:["IT"], excludeRanges:[] }
 *
 *   Jerarquía (el rango no aplica a jerarquías):
 *     { mode:"list",
 *       items:[{attribute:"PAIS", value:"España"}, ...],
 *       excludeItems:[{attribute:"CONTINENTE", value:"África"}, ...] }
 *
 * taskpane.js guarda ese objeto tal cual (como JSON) en entry.filter, y
 * commands.js (buildWhere) lo interpreta para construir la condición SQL.
 * Formatos de filtro guardados con versiones ANTERIORES del diálogo
 * ("values", "range", "mixed") o de antes de la selección múltiple (una
 * simple cadena) se siguen interpretando igual, por compatibilidad — ver
 * parseFilterValue más abajo, y su duplicado parseStoredFilterValue en
 * commands.js.
 * ========================================================================== */

/**
 * Lee la respuesta de BigQuery (jobs.query) como { names, rows }, donde
 * rows es un array de filas y cada fila un array de valores (string o null).
 * ANTES se buscaba el literal "v":"..." con una expresión regular, que se
 * saltaba los {"v": null}: en las jerarquías irregulares (niveles inferiores
 * a NULL) eso desalineaba las columnas y el árbol volvía a pintarse desde el
 * nivel 1. Si el texto no es JSON válido se usa una lectura por expresión
 * regular que SÍ cuenta los null.
 */
function parseBqRows(json) {
    const text = String(json || "");
    try {
        const p = JSON.parse(text);
        const names = ((p && p.schema && p.schema.fields) || []).map(f => f.name);
        const rows = ((p && p.rows) || []).map(r => (r.f || []).map(c => {
            const v = c ? c.v : null;
            return (v === null || v === undefined) ? null : String(v);
        }));
        return { names, rows };
    } catch (e) { /* no es JSON: lectura por regex */ }

    const names = [...text.matchAll(/"name":\s*"([^"]+)"/g)].map(m => m[1]);
    const values = [...text.matchAll(/"v":\s*(null|"((?:[^"\\]|\\.)*)")/g)].map(m => {
        if (m[1] === "null") return null;
        try { return JSON.parse('"' + m[2] + '"'); } catch (e2) { return m[2]; }
    });
    const rows = [];
    const n = names.length;
    if (n > 0) {
        for (let i = 0; i + n <= values.length; i += n) rows.push(values.slice(i, i + n));
    }
    return { names, rows };
}

/**
 * Construye el árbol de miembros (deduplicado por RUTA, no por "valor
 * anterior") y lo devuelve aplanado en orden de recorrido en profundidad:
 * [{ attribute, value, depth, label? }]. Un NULL en un nivel corta la rama:
 * ni se pinta ese nivel ni los siguientes (jerarquías irregulares).
 * layout.levels: [{ idCol, txtCol, attribute }] (txtCol -1 = sin texto).
 */
function buildMemberTree(parsed, layout) {
    const names = parsed.names || [];
    const levels = (layout && layout.levels) || names.map((nm, i) => ({ idCol: i, txtCol: -1, attribute: nm }));
    const root = { children: new Map() };

    for (const row of (parsed.rows || [])) {
        let node = root;
        for (let i = 0; i < levels.length; i++) {
            const L = levels[i];
            const raw = row[L.idCol];
            if (raw === null || raw === undefined || String(raw) === "") break;
            if (L.empty && L.empty.has(String(raw).trim())) break;    // «valor vacío» de la jerarquía = NULL
            const v = String(raw);
            let child = node.children.get(v);
            if (!child) {
                child = {
                    attribute: names[L.idCol] || L.attribute,
                    value: v,
                    label: null,
                    depth: i,
                    children: new Map()
                };
                node.children.set(v, child);
            }
            if (L.txtCol >= 0 && child.label === null) {
                const t = row[L.txtCol];
                if (t !== null && t !== undefined && String(t) !== "" && !(L.empty && L.empty.has(String(t).trim()))) child.label = String(t);
            }
            node = child;
        }
    }

    const out = [];
    const walk = (node) => {
        node.children.forEach(ch => {
            const item = { attribute: ch.attribute, value: ch.value, depth: ch.depth };
            if (ch.label !== null) item.label = ch.label;
            out.push(item);
            walk(ch);
        });
    };
    walk(root);
    return out;
}

/**
 * Compatibilidad (la usa calcEditor.js): lista {text, attribute, value,
 * depth} con el texto sangrado como antes, ya sin el fallo de los NULL.
 */
function loadJsonTree(json) {
    return buildMemberTree(parseBqRows(json), null)
        .map(it => Object.assign({ text: " ".repeat(it.depth * 4) + it.value }, it));
}

/**
 * Interpreta la cadena guardada en la columna "Valor" de un filtro.
 * Formato actual: JSON { mode:"list", values|items, ranges,
 * excludeValues|excludeItems, excludeRanges }. Formatos anteriores
 * ("values"/"range"/"mixed" con include/valuesInclude/rangeInclude) y el
 * formato "antiguo" (una cadena simple, tratada como igualdad de un único
 * valor) se aceptan igual, por compatibilidad. Debe coincidir con
 * parseStoredFilterValue de commands.js.
 */
function parseFilterValue(raw) {
    const s = String(raw || "").trim();
    if (s === "") return null;

    if (s[0] === "{") {
        try {
            const parsed = JSON.parse(s);
            if (parsed && typeof parsed === "object" && parsed.mode) return parsed;
        } catch (e) {
            // No era JSON válido: se interpreta como valor simple (ver abajo).
        }
    }

    return { mode: "values", include: true, values: [s] };
}

/**
 * Resumen legible de una lista de valores, con su prefijo de
 * incluir ("") / excluir ("≠ ").
 */
function describeValuesList(values, include) {
    if (!values || values.length === 0) return "";
    const prefix = include ? "" : "≠ ";
    if (values.length <= 2) return prefix + values.join(", ");
    return prefix + `${values.slice(0, 2).join(", ")} (+${values.length - 2} más)`;
}

/**
 * Resumen legible de un rango, con su prefijo de incluir/excluir.
 */
function describeRange(from, to, include) {
    const desde = from || "…";
    const hasta = to || "…";
    return (include ? "" : "≠ ") + `${desde} – ${hasta}`;
}

/**
 * Construye un resumen legible de un filtro (para el "tag" del taskpane).
 * Acepta el formato actual (mode:"list": valores/rangos incluidos y
 * excluidos, combinables libremente) y, por compatibilidad, los formatos
 * de versiones anteriores del diálogo ("values", "range", "mixed") y el
 * formato "antiguo" (cadena simple).
 *
 * Orden del resumen, igual que el que usa commands.js para componer el
 * WHERE: valores incluidos, rangos incluidos, valores excluidos, rangos
 * excluidos.
 */
function describeFilter(filter) {
    if (!filter) return "";

    if (typeof filter === "string") return filter; // formato antiguo

    if (filter.mode === "range") {
        return describeRange(filter.from, filter.to, filter.include);
    }

    if (filter.mode === "mixed") {
        const partes = [
            describeValuesList(filter.values, filter.valuesInclude),
            describeRange(filter.from, filter.to, filter.rangeInclude)
        ].filter(Boolean);
        return partes.join(" + ");
    }

    if (filter.mode === "list") {
        const incValues = filter.items ? filter.items.map(it => it.value) : (filter.values || []);
        const excValues = filter.excludeItems ? filter.excludeItems.map(it => it.value) : (filter.excludeValues || []);

        const partes = [
            describeValuesList(incValues, true),
            ...(filter.ranges || []).map(r => describeRange(r.from, r.to, true)),
            describeValuesList(excValues, false),
            ...(filter.excludeRanges || []).map(r => describeRange(r.from, r.to, false))
        ].filter(Boolean);

        return partes.join(" + ");
    }

    // Formato de selección múltiple "clásico" (values / items, sin mezclar
    // incluidos y excluidos): un único incluir/excluir para todo el conjunto.
    const values = filter.items ? filter.items.map(it => it.value) : (filter.values || []);
    return describeValuesList(values, filter.include);
}

const FilterModal = {
    // Ya no hay overlay en el DOM del taskpane que inicializar; se mantiene
    // como no-op para no tener que tocar la llamada existente en
    // taskpane.js (FilterModal.init() al arrancar).
    init() {},

    /**
     * Abre el diálogo de filtro para el campo indicado y devuelve una
     * Promise que resuelve al objeto "filter" (ver cabecera del fichero),
     * o null si el usuario cancela o el diálogo no se puede abrir.
     *
     * fieldData admite opcionalmente `currentFilter` (el filtro ya
     * aplicado anteriormente sobre este campo) para precargar la
     * selección al reabrir el diálogo.
     */
    open(fieldData) {
        return new Promise((resolve) => {
            // Antes: se esperaba aquí (await) a que terminase la consulta de
            // valores ANTES de abrir el diálogo. Con red lenta eso deja al
            // taskpane sin ninguna señal de actividad durante ese rato (el
            // "Esperando datos…" del diálogo apenas llega a verse, porque
            // para cuando la ventana se abre los datos ya están listos) —
            // parece que el selector se ha quedado bloqueado.
            //
            // Ahora el diálogo se abre INMEDIATAMENTE (el usuario ve
            // "Esperando datos…" desde el primer clic) y la consulta se
            // lanza en paralelo; se envía al diálogo en cuanto están listos
            // TANTO el diálogo (mensaje "ready") COMO los datos (lo que
            // termine último de los dos).
            const itemsPromise = (async () => {
                // Con columnas de texto (selector ID / Texto / ID y texto del
                // diálogo); el filtro se sigue guardando por ID.
                const query = await window.ExcelService.buildFilterValuesQuery(fieldData.dim, fieldData.name);
                if (!query || !query.sql) {
                    console.error("FilterModal: no se ha encontrado el atributo o jerarquía.", fieldData);
                    return null; // null = error, distinto de "sin resultados" ([])
                }
                // «Valor vacío» de las jerarquías del modelo ("---"): se trata como NULL al construir el árbol
                try {
                    const emp = (typeof dracoHierarchyEmptyTokens === "function") ? dracoHierarchyEmptyTokens(null) : new Map();
                    (query.levels || []).forEach(L => { L.empty = emp.get((String(fieldData.dim) + "|" + String(L.attribute)).toUpperCase()) || null; });
                } catch (e) { /* sin valor vacío */ }
                // Caché de miembros de la sesión (commands.js): mismo SQL y
                // conexión = sin volver a consultar; se vacía al cerrar sesión
                const json = (typeof window.dracoMemberQuery === "function")
                    ? await window.dracoMemberQuery(query.sql)
                    : await window.ExcelService.executeSQL(query.sql);
                const items = buildMemberTree(parseBqRows(json), query);
                // Cálculos entre miembros (del informe y del modelo, p. ej.
                // OIBDA = [Ingresos] - [Gastos]): se ofrecen como un miembro más
                // al final de la lista. Si se filtra, el cálculo solo sale
                // cuando está elegido (ver dracoCalcAllowedByFilters).
                try {
                    if (!fieldData.isHierarchy && typeof window.dracoCalcsForAttr === "function") {
                        const reportId = (typeof TaskPaneApp !== "undefined") ? TaskPaneApp.currentReportId : null;
                        const attr = (items[0] && items[0].attribute) || fieldData.name;
                        const have = new Set(items.map(i => String(i.value).toUpperCase()));
                        window.dracoCalcsForAttr(reportId, fieldData.dim, fieldData.name).forEach(c => {
                            const n = String(c.name || "").trim();
                            if (n && !have.has(n.toUpperCase())) items.push({ attribute: attr, value: n, depth: 0, calc: true });
                        });
                    }
                } catch (e) { console.warn("FilterModal: no se pudieron añadir los cálculos:", e); }
                return items;
            })();
            // Si displayDialogAsync fallara (ver más abajo) antes de que llegue
            // "ready", nadie más leería el resultado/error de itemsPromise —
            // este catch silencioso evita un "unhandled promise rejection" en
            // consola; el error real, si lo hay, ya se habrá logueado arriba.
            itemsPromise.catch(() => {});

            const dialogUrl = new URL("filterDialog.html", window.location.href).href;

            Office.context.ui.displayDialogAsync(
                dialogUrl,
                { height: 65, width: 48, displayInIframe: false },
                (asyncResult) => {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        console.error(
                            "FilterModal: displayDialogAsync ha fallado:",
                            asyncResult.error && asyncResult.error.code,
                            asyncResult.error && asyncResult.error.message
                        );
                        resolve(null);
                        return;
                    }

                    const dialog = asyncResult.value;
                    let settled = false;
                    const closeDialog = () => { try { dialog.close(); } catch (e) { /* ya cerrado */ } };

                    dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
                        let payload;
                        try {
                            payload = JSON.parse(arg.message);
                        } catch (err) {
                            console.error("FilterModal: mensaje del diálogo no es JSON válido:", err);
                            return;
                        }

                        if (payload.type === "ready") {
                            // El diálogo ya está mostrando "Esperando datos…"; en
                            // cuanto la consulta (lanzada en paralelo al abrir el
                            // diálogo) termine, le mandamos el resultado. Si ya
                            // había terminado antes de que llegase "ready", el
                            // .then() se resuelve al instante.
                            itemsPromise.then((items) => {
                                if (settled) return; // el usuario ya canceló mientras tanto
                                if (items === null) {
                                    // buildFilterValuesSQL no encontró el campo: no
                                    // hay nada que mostrar, se cierra el diálogo.
                                    settled = true;
                                    closeDialog();
                                    resolve(null);
                                    return;
                                }
                                dialog.messageChild(JSON.stringify({
                                    items,
                                    fieldData: { dim: fieldData.dim, name: fieldData.name, isHierarchy: fieldData.isHierarchy },
                                    currentFilter: fieldData.currentFilter || null,
                                    initialSearch: fieldData.initialSearch || ""
                                }));
                            }, (err) => {
                                console.error("Error cargando valores de filtro:", err);
                                if (settled) return;
                                settled = true;
                                closeDialog();
                                resolve(null);
                            });
                            return;
                        }

                        if (payload.type === "apply") {
                            settled = true;
                            closeDialog();
                            resolve(payload.filter);
                            return;
                        }

                        if (payload.type === "cancel") {
                            settled = true;
                            closeDialog();
                            resolve(null);
                        }
                    });

                    dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
                        // El usuario ha cerrado la ventana con la X del sistema
                        // operativo (no con nuestro botón, que ya manda "cancel").
                        if (!settled) resolve(null);
                    });
                }
            );
        });
    }
};

window.FilterModal = FilterModal;
window.describeFilter = describeFilter;
window.parseFilterValue = parseFilterValue;

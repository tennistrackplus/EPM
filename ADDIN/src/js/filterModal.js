/* ==========================================================================
 * FilterModal — puente hacia el diálogo independiente de Office
 * "Filtrar campo" (filterDialog.html), usado desde la zona "Filtros" del
 * taskpane.
 *
 * Antes este fichero pintaba un overlay DENTRO de taskpane.html. Ahora el
 * selector vive en su propia ventana (igual que memberPicker.html), lo que
 * le da sitio para un panel auxiliar a la derecha; ver
 * js/filterDialog.js + filterDialog.html para la UI real.
 *
 * Un diálogo de Office no tiene acceso al modelo de objetos de Excel, así
 * que aquí (en el taskpane, donde SÍ está disponible window.ExcelService)
 * se resuelve primero la lista de valores y luego se abre el diálogo,
 * enviándosela por mensaje en cuanto avisa que está listo — mismo patrón
 * que openMemberRecognitionPicker en commands.js.
 *
 * El contrato hacia quien llama (taskpane.js) NO ha cambiado:
 * FilterModal.open(fieldData) sigue devolviendo una Promise que resuelve a
 * un objeto "filter" o a null si se cancela:
 *
 *   Dimensión plana, varios valores:
 *     { mode: "values", include: true|false, attribute: "PAIS", values: ["ES","FR"] }
 *
 *   Dimensión plana, rango:
 *     { mode: "range", include: true|false, attribute: "IMPORTE", from: "0", to: "1000" }
 *
 *   Jerarquía, varios miembros (posiblemente de niveles/atributos distintos):
 *     { mode: "values", include: true|false, items: [{attribute:"PAIS", value:"España"}, ...] }
 *
 * taskpane.js guarda ese objeto tal cual (como JSON) en entry.filter, y
 * commands.js (buildWhere) lo interpreta para construir IN/NOT IN/BETWEEN.
 * Un valor de filtro "antiguo" (una simple cadena, sin JSON) se sigue
 * interpretando como igualdad simple, por compatibilidad — ver
 * parseFilterValue más abajo.
 * ========================================================================== */

/**
 * Traducción de LoadJson: separa los nombres de campo ("name":"...") y los
 * valores ("v":"...") del JSON crudo de BigQuery, y construye una lista
 * "en árbol" (indentada) deduplicando cada nivel frente al valor anterior,
 * igual que hacía el bucle de Ultimos()/Valores() en VBA.
 */
function loadJsonTree(json) {
    const fieldMatches = [...json.matchAll(/"name":\s*"([^"]+)"/g)];
    const campos = fieldMatches.map(m => m[1]);

    if (campos.length === 0) return [];

    const ultimos = campos.map(() => "");
    const valores = campos.map(() => "");

    const valueMatches = [...json.matchAll(/"v":\s*"([^"]*)"/g)];

    const items = [];
    let nivel = 0;

    for (const m of valueMatches) {
        valores[nivel] = m[1];
        nivel++;

        if (nivel > campos.length - 1) {
            for (let i = 0; i < campos.length; i++) {
                if (valores[i] !== ultimos[i]) {
                    items.push({
                        text: " ".repeat(i * 4) + valores[i],
                        attribute: campos[i],
                        value: valores[i]
                    });

                    ultimos[i] = valores[i];
                    for (let j = i + 1; j < campos.length; j++) ultimos[j] = "";
                }
            }
            nivel = 0;
        }
    }

    return items;
}

/**
 * Interpreta la cadena guardada en la columna "Valor" de un filtro.
 * Formato nuevo: JSON { mode: "values"|"range", include, values|items|from/to }.
 * Formato antiguo (informes previos a la selección múltiple): una cadena
 * simple, que se trata como igualdad de un único valor. Debe coincidir con
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
 * Construye un resumen legible de un filtro (para el "tag" del taskpane).
 * Acepta tanto el formato nuevo (objeto) como el antiguo (cadena simple),
 * por compatibilidad con informes guardados previamente.
 */
function describeFilter(filter) {
    if (!filter) return "";

    if (typeof filter === "string") return filter; // formato antiguo

    if (filter.mode === "range") {
        const desde = filter.from || "…";
        const hasta = filter.to || "…";
        return (filter.include ? "" : "≠ ") + `${desde} – ${hasta}`;
    }

    const values = filter.items ? filter.items.map(it => it.value) : (filter.values || []);
    if (values.length === 0) return "";

    const prefix = filter.include ? "" : "≠ ";
    if (values.length <= 2) return prefix + values.join(", ");
    return prefix + `${values.slice(0, 2).join(", ")} (+${values.length - 2} más)`;
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
        return new Promise(async (resolve) => {
            const container = null; // ya no hay contenedor propio: los mensajes de error van por consola.

            let items = [];
            try {
                const sql = await window.ExcelService.buildFilterValuesSQL(fieldData.dim, fieldData.name);

                if (!sql) {
                    console.error("FilterModal: no se ha encontrado el atributo o jerarquía.", fieldData);
                    resolve(null);
                    return;
                }

                const json = await window.ExcelService.executeSQL(sql);
                items = loadJsonTree(json);
            } catch (err) {
                console.error("Error cargando valores de filtro:", err);
                resolve(null);
                return;
            }

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
                            dialog.messageChild(JSON.stringify({
                                items,
                                fieldData: { dim: fieldData.dim, name: fieldData.name, isHierarchy: fieldData.isHierarchy },
                                currentFilter: fieldData.currentFilter || null,
                                initialSearch: fieldData.initialSearch || ""
                            }));
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

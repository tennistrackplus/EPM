/**
 * LkmlImport
 * ------------------------------------------------------------------------
 * Función inversa de LkmlExport (ver js/lkmlExport.js): a partir del texto
 * de un fichero .lkml, reconstruye el objeto { fact, fields } que espera
 * SemanticModelStore.saveModel() (mismo formato que fieldsState en
 * semantic_model.js).
 *
 * Es un parser "best effort" pensado sobre todo para reabrir ficheros
 * generados por el propio LkmlExport (views con sql_table_name /
 * dimension / measure, jerarquías documentadas como comentario "# Jerarquía
 * ..." y un explore con los joins fact-dimensión). Con LookML escrito a
 * mano que siga esa misma forma también debería funcionar razonablemente
 * bien, pero no es un parser completo del lenguaje LookML.
 *
 * Es una función PURA (no toca Excel ni Office.context.document): la usa
 * el diálogo independiente "Abrir modelo semántico" (openSemanticModel.html),
 * que no tiene acceso al modelo de objetos de Excel ni a
 * Office.context.document.settings.
 */
(function () {

    /**
     * A partir del índice de una '{' de apertura, devuelve el contenido
     * hasta su '}' de cierre correspondiente (contando anidamiento).
     */
    function extractBetweenBraces(text, openIdx) {

        let depth = 0;

        for (let i = openIdx; i < text.length; i++) {

            if (text[i] === "{") depth++;
            else if (text[i] === "}") {
                depth--;
                if (depth === 0) {
                    return { body: text.slice(openIdx + 1, i), end: i };
                }
            }

        }

        return { body: text.slice(openIdx + 1), end: text.length };

    }

    /**
     * Encuentra todos los bloques "keyword: nombre { ... }" de primer nivel
     * dentro de "content" (p.ej. keyword="view", keyword="dimension"...).
     */
    function findBlocks(content, keyword) {

        const blocks = [];
        const re = new RegExp(`\\b${keyword}\\s*:\\s*([a-zA-Z0-9_]+)\\s*\\{`, "g");
        let m;

        while ((m = re.exec(content)) !== null) {

            const openIdx = m.index + m[0].length - 1;
            const { body, end } = extractBetweenBraces(content, openIdx);

            blocks.push({ name: m[1], body });

            re.lastIndex = end;

        }

        return blocks;

    }

    function parseSqlTableName(body) {

        const m = body.match(/sql_table_name:\s*`([^`]+)`/);

        if (!m) return { project: "", dataset: "", table: "" };

        const parts = m[1].split(".");

        if (parts.length >= 3) return { project: parts[0], dataset: parts[1], table: parts.slice(2).join(".") };
        if (parts.length === 2) return { project: "", dataset: parts[0], table: parts[1] };

        return { project: "", dataset: "", table: parts[0] || "" };

    }

    function parseMeasures(body) {

        return findBlocks(body, "measure").map(blk => {

            const typeM = blk.body.match(/type:\s*([a-zA-Z0-9_]+)/);
            const sqlM = blk.body.match(/sql:\s*\$\{TABLE\}\.([a-zA-Z0-9_]+)\s*;;/);
            const fmtM = blk.body.match(/value_format_name:\s*([a-zA-Z0-9_]+)/);

            return {
                enabled: true,
                type: "MEASURE",
                name: sqlM ? sqlM[1] : blk.name,
                aggregation: typeM ? typeM[1] : "sum",
                format: fmtM ? fmtM[1] : ""
            };

        });

    }

    function parseAttributes(body) {

        return findBlocks(body, "dimension").map(blk => {

            const isKey = /primary_key:\s*yes/.test(blk.body);
            const typeM = blk.body.match(/#\s*tipo de origen:\s*(\S+)/);
            const sqlM = blk.body.match(/sql:\s*\$\{TABLE\}\.([a-zA-Z0-9_]+)\s*;;/);
            const fieldName = sqlM ? sqlM[1] : blk.name;

            return {
                name: fieldName,
                alias: blk.name && blk.name !== fieldName ? blk.name : fieldName,
                dataType: typeM ? typeM[1] : "",
                isKey: isKey,
                enabled: true
            };

        });

    }

    function parseHierarchies(body) {

        const hierarchies = [];
        const re = /#\s*Jerarquía\s*"([^"]+)":\s*\n((?:[ \t]*#\s*\d+\.\s*.+\n?)+)/g;
        let m;

        while ((m = re.exec(body)) !== null) {

            const name = m[1];
            const levels = [];
            const lvlRe = /#\s*\d+\.\s*(.+)/g;
            let lm;

            while ((lm = lvlRe.exec(m[2])) !== null) {
                levels.push({ attribute: lm[1].trim() });
            }

            if (levels.length > 0) {
                hierarchies.push({ name: name, levels: levels });
            }

        }

        return hierarchies;

    }

    // Devuelve { [dimViewName]: { factField, keyId } } a partir del bloque
    // "explore: modelo { ... join: dim { sql_on: ${fact.campo} = ${dim.clave} ;; } ... }"
    function parseJoins(exploreBody) {

        const joins = {};

        findBlocks(exploreBody, "join").forEach(blk => {

            const m = blk.body.match(
                /sql_on:\s*\$\{[a-zA-Z0-9_]+\.([a-zA-Z0-9_]+)\}\s*=\s*\$\{([a-zA-Z0-9_]+)\.([a-zA-Z0-9_]+)\}/
            );

            if (m) {
                joins[blk.name] = { factField: m[1], keyId: m[3] };
            }

        });

        return joins;

    }

    /**
     * @param {string} content Texto completo del fichero .lkml
     * @returns {{ fact: {project,dataset,table}, fields: Array }}
     */
    function parseContent(content) {

        if (!content || typeof content !== "string" || content.trim() === "") {
            throw new Error("El fichero está vacío.");
        }

        const views = findBlocks(content, "view");

        if (views.length === 0) {
            throw new Error("No se ha encontrado ningún bloque \"view\" en el fichero LookML.");
        }

        const explores = findBlocks(content, "explore");
        const explore = explores[0] || null;
        const joins = explore ? parseJoins(explore.body) : {};

        // La vista de hechos: la que declara measures o, si no hay ninguna,
        // la que referencia "view_name" dentro del explore.
        let factView = views.find(v => /\bmeasure\s*:/.test(v.body));

        if (!factView && explore) {
            const viewNameM = explore.body.match(/view_name:\s*([a-zA-Z0-9_]+)/);
            if (viewNameM) factView = views.find(v => v.name === viewNameM[1]);
        }

        if (!factView) factView = views[0];

        const fact = parseSqlTableName(factView.body);
        const measures = parseMeasures(factView.body);

        const dimensionFields = views
            .filter(v => v.name !== factView.name)
            .map(v => {

                const relTable = parseSqlTableName(v.body);
                const attributes = parseAttributes(v.body);
                const hierarchies = parseHierarchies(v.body);
                const join = joins[v.name];

                return {
                    enabled: true,
                    type: "DIMENSION",
                    name: join ? join.factField : v.name,
                    relProject: relTable.project,
                    relDataset: relTable.dataset,
                    relTable: relTable.table,
                    attributes: attributes,
                    hierarchies: hierarchies
                };

            });

        if (!fact.project && !fact.dataset && !fact.table && dimensionFields.length === 0 && measures.length === 0) {
            throw new Error("No se ha podido reconocer ninguna tabla de hechos, dimensión o medida en el fichero.");
        }

        return {
            fact: fact,
            fields: [...dimensionFields, ...measures]
        };

    }

    window.LkmlImport = { parseContent };

})();

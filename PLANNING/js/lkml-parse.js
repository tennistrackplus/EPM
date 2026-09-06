/**
 * LkmlParse — lee de vuelta el .lkml que genera js/lkml-export.js (el
 * mismo formato que ya exporta ADDIN/src/js/lkmlExport.js).
 *
 * No es un parser de LookML genérico: aprovecha que el formato lo
 * generamos nosotros mismos de forma determinista (ver lkml-export.js)
 * para extraer, con un escaneo simple por profundidad de llaves + unas
 * pocas expresiones regulares, exactamente lo que hace falta para
 * construir un informe:
 *
 *   {
 *     factTable: "`proyecto.schema.tabla`" (o null),
 *     measures: [{ name, column }],
 *     dimensions: [{
 *       id,                       // nombre de la view (colId sanitizado)
 *       table: "`proyecto.schema.tabla`",
 *       keyColumn,                // columna marcada primary_key: yes (o la primera)
 *       attributes: [{ colId }],
 *       hierarchies: [{ name, levels: [colId, ...] }]  // desde los comentarios
 *     }],
 *     joins: [{ dimId, fkColumn, refColumn }]   // desde el bloque explore
 *   }
 */
const LkmlParse = {

    // Divide `text` en los bloques de nivel superior "view: nombre { ... }"
    // / "explore: nombre { ... }", contando llaves para encontrar el cierre
    // correcto (soporta cualquier anidamiento, aunque nuestro generador no
    // anida más de un nivel).
    splitTopLevelBlocks(text) {
        return this.splitNamedBlocks(text, "(?:view|explore)", true);
    },

    // Genérico: encuentra todas las apariciones de "`keyword`: nombre {" en
    // `body` y devuelve [{ keyword, name, body }] (o [{name, body}] si
    // includeKeyword=false), resolviendo el cierre de cada bloque por
    // profundidad de llaves.
    splitNamedBlocks(body, keywordPattern, includeKeyword) {
        const blocks = [];
        const re = new RegExp(`(${keywordPattern}):\\s*(\\w+)\\s*\\{`, "g");
        let match;
        while ((match = re.exec(body))) {
            const startBody = match.index + match[0].length;
            let depth = 1, i = startBody;
            while (i < body.length && depth > 0) {
                if (body[i] === "{") depth++;
                else if (body[i] === "}") depth--;
                i++;
            }
            const entry = { name: match[2], body: body.slice(startBody, i - 1) };
            if (includeKeyword) entry.keyword = match[1];
            blocks.push(entry);
            re.lastIndex = i;
        }
        return blocks;
    },

    sqlTableName(body) {
        const m = body.match(/sql_table_name:\s*`([^`]+)`/);
        return m ? `\`${m[1]}\`` : null;
    },

    // Columna física referenciada en "sql: ${TABLE}.columna ;;"
    sqlColumn(body) {
        const m = body.match(/sql:\s*\$\{TABLE\}\.(\w+)\s*;;/);
        return m ? m[1] : null;
    },

    parseMeasures(factBody) {
        return this.splitNamedBlocks(factBody, "measure", false).map(b => ({
            name: b.name,
            column: this.sqlColumn(b.body) || b.name
        }));
    },

    parseAttributes(dimBody) {
        return this.splitNamedBlocks(dimBody, "dimension", false).map(b => ({
            colId: this.sqlColumn(b.body) || b.name,
            isKey: /primary_key:\s*yes/.test(b.body)
        }));
    },

    // Las jerarquías se documentan como comentario (LookML no tiene un
    // bloque nativo): "# Jerarquía \"Nombre\":" seguido de líneas
    // "#   N. Nombre de atributo" (ver lkml-export.js). El texto tras el
    // número es el NOMBRE del atributo (a.name), no la columna física, así
    // que hay que volver a derivar el identificador físico exactamente
    // igual que Provider.toIdentifier en BigQuery (mayúsculas, sin acentos)
    // para poder usarlo en el GROUP BY/JOIN de la consulta.
    bqIdentifier(text) {
        return String(text || "")
            .trim()
            .toUpperCase()
            .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
            .replace(/[^A-Z0-9_]+/g, "_")
            .replace(/^_+|_+$/g, "")
            .replace(/^(\d)/, "N$1");
    },

    parseHierarchies(dimBody) {
        const hierarchies = [];
        const hierRe = /#\s*Jerarquía\s*"([^"]+)":\s*\n((?:\s*#\s*\d+\.\s*.+\n?)+)/g;
        let m;
        while ((m = hierRe.exec(dimBody))) {
            const levels = [];
            const lvlRe = /#\s*\d+\.\s*(.+)/g;
            let lm;
            while ((lm = lvlRe.exec(m[2]))) {
                const label = lm[1].trim();
                levels.push({ label, colId: this.bqIdentifier(label) });
            }
            if (levels.length) hierarchies.push({ name: m[1], levels });
        }
        return hierarchies;
    },

    parseJoins(exploreBody) {
        return this.splitNamedBlocks(exploreBody, "join", false).map(b => {
            const m = b.body.match(/\$\{\w+\.(\w+)\}\s*=\s*\$\{\w+\.(\w+)\}/);
            return {
                dimId: b.name,
                fkColumn: m ? m[1] : null,
                refColumn: m ? m[2] : null
            };
        }).filter(j => j.fkColumn && j.refColumn);
    },

    parse(text) {
        const result = { factTable: null, measures: [], dimensions: [], joins: [] };
        if (!text) return result;

        const topBlocks = this.splitTopLevelBlocks(text);

        topBlocks.forEach(block => {
            if (block.keyword === "view") {
                const isFact = /measure:/.test(block.body);
                if (isFact) {
                    result.factTable = this.sqlTableName(block.body);
                    result.measures = this.parseMeasures(block.body);
                } else {
                    const attributes = this.parseAttributes(block.body);
                    const keyAttr = attributes.find(a => a.isKey) || attributes[0] || null;
                    result.dimensions.push({
                        id: block.name,
                        table: this.sqlTableName(block.body),
                        keyColumn: keyAttr ? keyAttr.colId : block.name,
                        attributes,
                        hierarchies: this.parseHierarchies(block.body)
                    });
                }
            } else if (block.keyword === "explore") {
                result.joins = this.parseJoins(block.body);
            }
        });

        return result;
    }
};

if (typeof module !== "undefined" && module.exports) module.exports = LkmlParse;

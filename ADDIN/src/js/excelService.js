/**
 * Servicio encargado de la comunicación directa con Excel mediante Office.js
 */
const ExcelService = {

    /**
     * Lee la información de la hoja DIM2 y la devuelve estructurada por Dimensiones
     */
    async readDim2Data() {
        return await Excel.run(async (context) => {
            const sheet = context.workbook.sheets.getItemOrNullObject("DIM2");
            await context.sync();

            if (sheet.isNullObject) {
                console.warn("Hoja DIM2 no encontrada.");
                return [];
            }

            const range = sheet.getUsedRange();
            range.load("values");
            await context.sync();

            const rows = range.values;
            if (!rows || rows.length <= 1) return [];

            // Obtenemos índices de columnas por encabezado
            const headers = rows[0].map(h => String(h).trim().toUpperCase());
            const dimIdx = headers.indexOf("DIMENSION");
            const attIdx = headers.indexOf("ATRIBUTE");
            const hierIdx = headers.indexOf("JERARQUIA");

            const dimensionsMap = {};

            for (let i = 1; i < rows.length; i++) {
                const row = rows[i];
                const dimName = row[dimIdx] ? String(row[dimIdx]).trim() : "";
                const attName = row[attIdx] ? String(row[attIdx]).trim() : "";
                const hierName = row[hierIdx] ? String(row[hierIdx]).trim() : "";

                if (!dimName) continue;

                if (!dimensionsMap[dimName]) {
                    dimensionsMap[dimName] = {
                        dimension: dimName,
                        hierarchies: [],
                        attributes: []
                    };
                }

                // Si hay jerarquía la añadimos
                if (hierName && !dimensionsMap[dimName].hierarchies.includes(hierName)) {
                    dimensionsMap[dimName].hierarchies.push(hierName);
                }

                // Si hay atributo lo añadimos
                if (attName && !dimensionsMap[dimName].attributes.includes(attName)) {
                    dimensionsMap[dimName].attributes.push(attName);
                }
            }

            return Object.values(dimensionsMap);
        });
    },

    /**
     * Lee los valores distintos de la hoja MET correspondientes a un campo específico
     */
    async readMetValuesForField(dimName, fieldName, isHierarchy) {
        return await Excel.run(async (context) => {
            const sheet = context.workbook.sheets.getItemOrNullObject("MET");
            await context.sync();

            if (sheet.isNullObject) {
                console.warn("Hoja MET no encontrada.");
                return [];
            }

            const range = sheet.getUsedRange();
            range.load("values");
            await context.sync();

            const rows = range.values;
            if (!rows || rows.length <= 1) return [];

            const uniqueValues = new Set();

            for (let i = 1; i < rows.length; i++) {
                const row = rows[i];
                const rowDim = row[0] ? String(row[0]).trim() : "";
                const rowAttOrHier1 = row[1] ? String(row[1]).trim() : "";
                const rowHier2 = row[2] ? String(row[2]).trim() : "";
                const rowVal = row[4] !== undefined && row[4] !== null ? String(row[4]).trim() : "";

                if (rowDim !== dimName) continue;

                if (isHierarchy) {
                    if (rowHier2 === fieldName && rowVal) {
                        uniqueValues.add(rowVal);
                    }
                } else {
                    if (rowAttOrHier1 === fieldName && rowVal) {
                        uniqueValues.add(rowVal);
                    }
                }
            }

            return Array.from(uniqueValues);
        });
    }
};
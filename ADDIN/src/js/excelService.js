/**
 * Estructura de datos local (MOCK_DATA) que reemplaza la lectura directa de las hojas DIM2 y MET de Excel.
 * Estructura: Dimension -> attributes -> matriz de valores
 *                        -> hierarchies -> matriz de objetos { level, value }
 */
const MOCK_DATA = {
    CECO: {
        attributes: {
            CECO: ["DE01", "ES01", "ES02", "ES03", "FR01", "FR02", "IT01", "IT02", "PT01", "UK01"],
            DESCRIPCION: ["Barcelona", "Berlín", "Lisboa", "Londres", "Lyon", "Madrid", "Milán", "París", "Roma", "Valencia"],
            CIUDAD: ["Barcelona", "Berlín", "Lisboa", "Londres", "Lyon", "Madrid", "Milán", "París", "Roma", "Valencia"],
            PAIS: ["Alemania", "España", "Francia", "Italia", "Portugal", "Reino Unido"],
            REGION: ["Europa"],
            SOCIEDAD: ["DE", "ES", "FR", "IT", "PT", "UK"],
            DIVISION: ["Comercial", "Operaciones"],
            AREA: ["Centro", "Levante", "Noreste", "Norte", "Oeste", "Sur"]
        },
        hierarchies: {
            HIER1: [
                { level: 1, value: "Europa" },
                { level: 2, value: "Alemania" },
                { level: 3, value: "Berlín" },
                { level: 4, value: "DE01" },
                { level: 2, value: "España" },
                { level: 3, value: "Barcelona" },
                { level: 4, value: "ES02" },
                { level: 3, value: "Madrid" },
                { level: 4, value: "ES01" },
                { level: 3, value: "Valencia" },
                { level: 4, value: "ES03" },
                { level: 2, value: "Francia" },
                { level: 3, value: "Lyon" },
                { level: 4, value: "FR02" },
                { level: 3, value: "París" },
                { level: 4, value: "FR01" },
                { level: 2, value: "Italia" },
                { level: 3, value: "Milán" },
                { level: 4, value: "IT01" },
                { level: 3, value: "Roma" },
                { level: 4, value: "IT02" },
                { level: 2, value: "Portugal" },
                { level: 3, value: "Lisboa" },
                { level: 4, value: "PT01" },
                { level: 2, value: "Reino Unido" },
                { level: 3, value: "Londres" },
                { level: 4, value: "UK01" }
            ]
        }
    },
    CUENTA: {
        attributes: {
            CUENTA: ["700000", "701000", "702000", "800000", "801000", "810000", "811000", "820000", "830000", "840000"],
            DESCRIPCION: ["Alquileres", "Infraestructura IT", "Licencias Software", "Marketing", "Otros Ingresos", "Seguridad Social", "Sueldos", "Ventas Producto", "Ventas Servicios", "Viajes"],
            NIVEL1: ["Gastos", "Ingresos"],
            NIVEL2: ["Administración", "Comercial", "Otros", "Personal", "Tecnología", "Ventas"],
            NIVEL3: ["Alquileres", "Cargas Sociales", "Infraestructura", "Licencias", "Marketing", "Otros ingresos", "Producto", "Servicios", "Sueldos", "Viajes"],
            TIPO: ["EXPENSE", "INCOME"],
            SIGNO: ["-1", "1"]
        },
        hierarchies: {
            HIER1: [
                { level: 1, value: "Gastos" },
                { level: 2, value: "Administración" },
                { level: 3, value: "Alquileres" },
                { level: 4, value: "830000" },
                { level: 3, value: "Viajes" },
                { level: 4, value: "840000" },
                { level: 2, value: "Comercial" },
                { level: 3, value: "Marketing" },
                { level: 4, value: "820000" },
                { level: 2, value: "Personal" },
                { level: 3, value: "Cargas Sociales" },
                { level: 4, value: "801000" },
                { level: 3, value: "Sueldos" },
                { level: 4, value: "800000" },
                { level: 2, value: "Tecnología" },
                { level: 3, value: "Infraestructura" },
                { level: 4, value: "810000" },
                { level: 3, value: "Licencias" },
                { level: 4, value: "811000" },
                { level: 1, value: "Ingresos" },
                { level: 2, value: "Otros" },
                { level: 3, value: "Otros ingresos" },
                { level: 4, value: "702000" },
                { level: 2, value: "Ventas" },
                { level: 3, value: "Producto" },
                { level: 4, value: "700000" },
                { level: 3, value: "Servicios" },
                { level: 4, value: "701000" }
            ]
        }
    },
    ESCENARIO: {
        attributes: {
            ESCENARIO: ["BUDGET", "REAL"]
        },
        hierarchies: {}
    },
    YEAR: {
        attributes: {
            YEAR: ["2026", "2027"]
        },
        hierarchies: {}
    },
    PERIOD: {
        attributes: {
            PERIOD: ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"]
        },
        hierarchies: {}
    }
};

/**
 * Servicio adaptado para operar en memoria contra MOCK_DATA
 */
const ExcelService = {

    /**
     * Lee la información estructurada de MOCK_DATA por Dimensiones
     */
    async readDim2Data() {
        try {
            const dimensionsList = [];

            for (const [dimName, dimContent] of Object.entries(MOCK_DATA)) {
                const hierarchiesList = dimContent.hierarchies ? Object.keys(dimContent.hierarchies) : [];
                const attributesList = dimContent.attributes ? Object.keys(dimContent.attributes) : [];

                dimensionsList.push({
                    dimension: dimName,
                    hierarchies: hierarchiesList,
                    attributes: attributesList
                });
            }

            return { data: dimensionsList };
        } catch (error) {
            console.error("Error leyendo MOCK_DATA:", error);
            return { error: "Error al procesar la variable de datos local." };
        }
    },

    /**
     * Lee los valores almacenados en MOCK_DATA conservando la estructura de nivel si es jerarquía
     */
    async readMetValuesForField(dimName, fieldName, isHierarchy) {
        const dimObj = MOCK_DATA[dimName];
        if (!dimObj) return [];

        if (isHierarchy) {
            const hierarchyItems = dimObj.hierarchies ? dimObj.hierarchies[fieldName] : null;
            if (!hierarchyItems) return [];
            
            // Retorna los objetos { level, value } para mantener la estructura jerárquica visual
            return hierarchyItems;
        } else {
            const attributeItems = dimObj.attributes ? dimObj.attributes[fieldName] : null;
            if (!attributeItems) return [];
            
            const uniqueValues = Array.from(new Set(attributeItems));
            return uniqueValues.map(val => ({ level: 1, value: val }));
        }
    }
};

window.ExcelService = ExcelService;
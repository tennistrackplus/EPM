/**
 * lkml-bootstrap.js — sustituye a LkmlOpenBridge/openSemanticModelDialog
 * del add-in (que dejan elegir un .lkml a mano en un repo cualquiera).
 *
 * Aquí, en su lugar: por cada cubo del proyecto con modelo semántico
 * generado, se descarga y parsea su .lkml (GithubRepo + LkmlParse, ya
 * cargados en la página anfitriona) y se registra como un modelo más en
 * SemanticModelStore — exactamente el mismo almacén que ya rellenaba el
 * add-in al importar un fichero a mano, así que el resto (selector
 * "Modelo semántico", ExcelService.readDim2Data, generación de SQL...)
 * funciona sin ningún cambio.
 *
 * Se ejecuta una vez, al cargar el taskpane.
 */
(function () {

    function splitQualified(q) {
        const clean = String(q || "").replace(/`/g, "");
        const parts = clean.split(".");
        return { project: parts[0] || "", dataset: parts[1] || "", table: parts.slice(2).join(".") || "" };
    }

    function lkmlToSemanticModel(lkmlModel) {
        const fact = splitQualified(lkmlModel.factTable);
        const fields = [];

        lkmlModel.dimensions.forEach(dim => {
            const dimParts = splitQualified(dim.table);
            fields.push({
                enabled: true,
                type: "DIMENSION",
                name: dim.id,
                relProject: dimParts.project,
                relDataset: dimParts.dataset,
                relTable: dimParts.table,
                attributes: dim.attributes.map(a => ({
                    name: a.colId, alias: a.colId, dataType: a.type || "STRING", isKey: !!a.isKey, enabled: true
                })),
                hierarchies: dim.hierarchies.map(h => ({
                    name: h.name, levels: h.levels.map(l => ({ attribute: l.colId }))
                }))
            });
        });

        lkmlModel.measures.forEach(m => {
            fields.push({
                enabled: true,
                type: "MEASURE",
                name: m.column, // FACT_FIELD y MEASURE usan el mismo valor (ver semanticModelStore.js::buildMeasureRows)
                aggregation: "SUM",
                format: "#,##0.00"
            });
        });

        return { fact, fields };
    }

    async function bootstrapInner() {
        const host = window.parent;
        if (!host || !host.GithubRepo || !host.LkmlParse || !host.Provider || !host.WidgetTableEditor) {
            console.error("lkml-bootstrap: falta GithubRepo/LkmlParse/Provider/WidgetTableEditor en la ventana anfitriona.");
            return;
        }

        const project = host.WidgetTableEditor.project;
        let cubes = [];
        try {
            cubes = await host.Provider.runQuery(
                `SELECT CUBO_ID, CUBOS, MODELO_YAML_PATH FROM ${host.Provider.qualifyControl("CUBOS")}
                 WHERE PROYECTO_ID = '${host.Provider.esc(project.PROYECTO_ID)}' ORDER BY CUBOS`
            );
        } catch (err) {
            console.error("lkml-bootstrap: error al listar cubos:", err);
            return;
        }

        // El PAT de GitHub no vive en DracoConfig.semanticModelGithub.token
        // (config.js) — se lee de Google Secret Manager, igual que hace
        // SemanticModel.generateAndPushLkml() justo antes de subir el
        // .lkml. Se pide una sola vez aquí y se reutiliza para todos los
        // cubos del bucle.
        let repoConfig = null;
        try {
            const cfg = (typeof host.DracoConfig !== "undefined" && host.DracoConfig.semanticModelGithub) || {};
            const token = await host.BQ.getGithubPatFromSecretManager();
            repoConfig = { url: cfg.url, branch: cfg.branch, token };
        } catch (err) {
            console.error("lkml-bootstrap: no se pudo obtener el token de GitHub desde Secret Manager:", err);
            return;
        }

        const report = host.WidgetTableEditor.state.report || {};
        let activeName = null;

        for (const cube of cubes) {
            let path = cube.MODELO_YAML_PATH ? cube.MODELO_YAML_PATH.replace(/\.yaml$/i, ".lkml") : null;
            if (!path && host.SemanticModel) {
                const ident = host.Provider.toIdentifier(cube.CUBOS) || cube.CUBO_ID;
                path = host.SemanticModel.buildPath(project, ident).replace(/\.yaml$/i, ".lkml");
            }
            if (!path) continue;

            try {
                const text = await host.GithubRepo.getFile(path, repoConfig);
                if (text === null) { console.warn(`lkml-bootstrap: no existe ${path} para el cubo "${cube.CUBOS}".`); continue; }
                const lkmlModel = host.LkmlParse.parse(text);
                const model = lkmlToSemanticModel(lkmlModel);
                await window.SemanticModelStore.saveModel(cube.CUBOS, model);
                if (cube.CUBO_ID === report.cuboId) activeName = cube.CUBOS;
            } catch (err) {
                console.error(`lkml-bootstrap: error al leer/parsear el modelo de "${cube.CUBOS}":`, err);
            }
        }

        if (activeName) {
            await window.SemanticModelStore.setActiveModelName(activeName);
        }

        // El selector de la izquierda (#semanticModelSelector) ya se
        // rellena en el propio init de taskpane.js (TaskPaneApp), que muy
        // probablemente termina ANTES que este bootstrap (que depende de
        // red). Se vuelve a poblar aquí para que, en cuanto lleguen los
        // modelos, aparezcan sin recargar nada.
        if (typeof TaskPaneApp !== "undefined" && TaskPaneApp.populateModelSelectorMain) {
            await TaskPaneApp.populateModelSelectorMain();
        }
    }

    // bootstrapInner() puede terminar en varios "return" tempranos (falta
    // algo en el host, falla la consulta de cubos...) — este wrapper
    // garantiza que, pase lo que pase, se avisa a Office.onReady (ver
    // host-bridge.js) de que ya puede arrancar taskpane.js/commands.js,
    // para no dejar el taskpane colgado esperando indefinidamente.
    async function bootstrap() {
        try {
            await bootstrapInner();
        } catch (err) {
            console.error("lkml-bootstrap: error inesperado:", err);
        } finally {
            if (window.__wteSignalModelsReady) window.__wteSignalModelsReady();
        }
    }

    if (document.readyState === "loading") {
        document.addEventListener("DOMContentLoaded", bootstrap);
    } else {
        bootstrap();
    }

})();

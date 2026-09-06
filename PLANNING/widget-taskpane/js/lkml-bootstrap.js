/**
 * lkml-bootstrap.js — sustituye a LkmlOpenBridge/openSemanticModelDialog
 * del add-in (que dejan elegir un .lkml a mano en un repo cualquiera).
 *
 * Diseño (v2, mucho más simple que la primera versión): el selector de
 * "Modelo semántico" del taskpane va asociado 1:1 a los cubos del
 * proyecto — eso ya lo sabemos sin tocar GitHub para nada, con una sola
 * consulta a CUBOS. Así que:
 *
 *   1) Se registra cada cubo como un modelo "vacío" (sin campos) en
 *      SemanticModelStore — instantáneo, sin red, solo para que aparezca
 *      en el desplegable.
 *   2) Si el informe ya tenía un cubo activo (se está reabriendo un
 *      widget guardado), SU .lkml sí se descarga y parsea ya mismo (una
 *      sola petición, no todas), para que se vea de inmediato.
 *   3) Para cualquier OTRO cubo, el .lkml no se descarga hasta que el
 *      usuario lo elige de verdad en el desplegable — se intercepta
 *      TaskPaneApp.onModelSelectorChange (sin tocar taskpane.js) para
 *      resolverlo la primera vez que se selecciona.
 *
 * Con esto, abrir el taskpane ya no depende de red salvo, como mucho,
 * por el cubo activo — nada que ver con la versión anterior, que traía
 * el .lkml de TODOS los cubos del proyecto por adelantado.
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

    function lkmlPathFor(host, project, cube) {
        if (cube.MODELO_YAML_PATH) return cube.MODELO_YAML_PATH.replace(/\.yaml$/i, ".lkml");
        if (!host.SemanticModel) return null;
        const ident = host.Provider.toIdentifier(cube.CUBOS) || cube.CUBO_ID;
        return host.SemanticModel.buildPath(project, ident).replace(/\.yaml$/i, ".lkml");
    }

    let repoConfigPromise = null;
    function getRepoConfig(host) {
        if (!repoConfigPromise) {
            repoConfigPromise = (async () => {
                const cfg = (typeof host.DracoConfig !== "undefined" && host.DracoConfig.semanticModelGithub) || {};
                const token = await host.BQ.getGithubPatFromSecretManager();
                return { url: cfg.url, branch: cfg.branch, token };
            })();
        }
        return repoConfigPromise;
    }

    // cube.CUBOS (nombre, lo que se ve en el selector) -> { host, project, cube, resolved }
    const pending = {};

    async function resolveModel(name) {
        const entry = pending[name];
        if (!entry || entry.resolved) return;
        entry.resolved = true; // marca como "en curso/hecho" ya, para no lanzar dos descargas si el usuario cambia rápido
        const { host, project, cube } = entry;
        try {
            const path = lkmlPathFor(host, project, cube);
            if (!path) { console.warn(`lkml-bootstrap: el cubo "${name}" no tiene modelo semántico (.lkml).`); return; }
            const repoConfig = await getRepoConfig(host);
            const text = await host.GithubRepo.getFile(path, repoConfig);
            if (text === null) { console.warn(`lkml-bootstrap: no existe ${path} para el cubo "${name}".`); return; }
            const lkmlModel = host.LkmlParse.parse(text);
            const model = lkmlToSemanticModel(lkmlModel);
            await window.SemanticModelStore.saveModel(name, model);
        } catch (err) {
            entry.resolved = false; // si falla, se puede reintentar si lo vuelve a elegir
            console.error(`lkml-bootstrap: error al leer/parsear el modelo de "${name}":`, err);
        }
    }

    // Intercepta el cambio de selector SIN tocar taskpane.js: la búsqueda de
    // TaskPaneApp.onModelSelectorChange se hace en el momento del evento
    // (es un método, no una referencia capturada), así que basta con
    // sustituirlo en cualquier momento después de que taskpane.js se cargue.
    function patchModelSelectorChange() {
        if (typeof TaskPaneApp === "undefined" || TaskPaneApp.__wteLkmlPatched) return false;
        const original = TaskPaneApp.onModelSelectorChange.bind(TaskPaneApp);
        TaskPaneApp.onModelSelectorChange = async function (modelName) {
            if (modelName) await resolveModel(modelName);
            return original(modelName);
        };
        TaskPaneApp.__wteLkmlPatched = true;
        return true;
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

        // Paso 1: un "modelo" vacío por cubo — instantáneo, sin red, solo
        // para que el selector los liste. Cero peticiones a GitHub aquí.
        for (const cube of cubes) {
            pending[cube.CUBOS] = { host, project, cube, resolved: false };
            await window.SemanticModelStore.saveModel(cube.CUBOS, { fact: {}, fields: [] });
        }

        patchModelSelectorChange();

        // Paso 2: si el informe ya tenía un cubo elegido (se reabre un
        // widget guardado), ese SÍ se resuelve ya — es la única petición
        // de red que hace falta para que se vea de inmediato.
        const report = host.WidgetTableEditor.state.report || {};
        const activeCube = cubes.find(c => c.CUBO_ID === report.cuboId);
        if (activeCube) {
            await resolveModel(activeCube.CUBOS);
            await window.SemanticModelStore.setActiveModelName(activeCube.CUBOS);
        }

        if (typeof TaskPaneApp !== "undefined" && TaskPaneApp.populateModelSelectorMain) {
            await TaskPaneApp.populateModelSelectorMain();
        }
    }

    // bootstrapInner() puede terminar en varios "return" tempranos (falta
    // algo en el host, falla la consulta de cubos...) — este wrapper
    // garantiza que, pase lo que pase, se avisa a Office.onReady (ver
    // host-bridge.js) de que ya puede arrancar taskpane.js/commands.js,
    // para no dejar el taskpane colgado esperando indefinidamente. Como
    // ahora bootstrapInner ya no depende de red salvo por, como mucho, UN
    // cubo, esto debería resolverse siempre muy rápido.
    async function bootstrap() {
        try {
            await bootstrapInner();
        } catch (err) {
            console.error("lkml-bootstrap: error inesperado:", err);
        } finally {
            patchModelSelectorChange(); // por si bootstrapInner falló antes de llegar a parchear
            if (window.__wteSignalModelsReady) window.__wteSignalModelsReady();
        }
    }

    if (document.readyState === "loading") {
        document.addEventListener("DOMContentLoaded", bootstrap);
    } else {
        bootstrap();
    }

})();

/**
 * ============================================================
 * DRACO PLANNING — FLUJOS DE CARGA EN LOS INFORMES
 * ============================================================
 * Atiende las peticiones "fl.*" del bloque «Flujo de carga» de un informe
 * (informes-herramienta.html), igual que js/informes-datos.js con "tu.*".
 * La página del informe NO toca la base de datos ni el storage: pide aquí la
 * definición del flujo, lo lanza y consulta su estado, con la sesión del usuario.
 *
 * Reutiliza la lógica de la pantalla de ejecución (js/flow-run.js: carga de
 * la definición, opciones de las variables y llamada al orquestador), la
 * validación de js/ui.js y la subida de ficheros de js/storage.js, para que
 * el informe se comporte exactamente igual que Draco ▸ Flujos ▸ Ejecutar.
 *
 *   fl.list      -> flujos del proyecto                       [{id, name, title, type}]
 *   fl.describe  -> pantalla de variables + cadena de pasos   {id, name, title, type, chain, blocks, options}
 *   fl.run       -> valida, sube ficheros y lanza             {runId, name}
 *   fl.status    -> última ejecución y estado de cada paso     {run, steps}
 */
const InformesFlujos = {
    fail(code, message, details) {
        const err = new Error(message);
        err.code = code;
        if (details) err.details = details;
        return err;
    },

    async list(project) {
        await Informes.guard();
        const rows = await Provider.runQuery(`SELECT FLUJO_ID, FLUJO, TIPO, SCREEN_TITLE
            FROM ${Provider.qualifyControl("FLUJOS")}
            WHERE PROYECTO_ID = '${Provider.esc(project.PROYECTO_ID)}'
            ORDER BY FLUJO`);
        return rows.map(r => ({ id: r.FLUJO_ID, name: r.FLUJO, title: r.SCREEN_TITLE || "", type: r.TIPO === "MANUAL" ? "manual" : "automatico" }));
    },

    /** Instancia propia de FlowRun para un flujo (sin DOM), comprobando que es del proyecto actual. */
    async open(project, ref) {
        if (!ref) throw this.fail("BAD_INPUT", "No se ha indicado el flujo de carga.");
        await Informes.guard();
        const fr = Object.assign(Object.create(FlowRun), {
            flujoId: ref, flow: null, interfaces: [], project: null, dimensionsCache: [],
            varOptions: {}, selOptState: {}, screenValues: {}, lastRunId: null, pollHandle: null
        });
        try { await fr.load(); }
        catch (e) { throw this.fail("NOT_FOUND", "Este flujo de carga ya no existe o no se puede leer."); }
        if (!fr.project || fr.project.PROYECTO_ID !== project.PROYECTO_ID) {
            throw this.fail("NOT_FOUND", "Este flujo de carga no pertenece al proyecto actual.");
        }
        await fr.resolveScreenVarOptions();
        return fr;
    },

    async describe(project, ref) {
        const fr = await this.open(project, ref);
        return {
            id: fr.flow.id,
            name: fr.flow.name,
            title: fr.flow.screen.title || fr.flow.name,
            type: fr.flow.type,
            chain: fr.flow.chain.map(c => {
                const i = fr.interfaceById(c.interfaceId);
                return { id: c.id, name: i ? i.name : null, origin: i ? i.originType : null };
            }),
            blocks: fr.flow.screen.blocks,
            options: fr.varOptions
        };
    },

    /**
     * values: {NOMBRE: valor | {low, high} | true/false}; sel: {VARIABLE_ID: [{sign, option, low, high}]}
     * (varios valores / cualquiera); files: {NOMBRE: File}. Misma validación que la pantalla de Draco.
     */
    async run(project, ref, values, sel, files) {
        const fr = await this.open(project, ref);
        values = Object.assign({}, values || {});
        sel = sel || {};
        files = files || {};
        const vars = [];
        fr.forEachScreenVariable(v => vars.push(v));
        vars.forEach(v => {
            if (v.selectMode === "multiple" || v.selectMode === "cualquiera") {
                values[v.name] = (sel[v.id] || []).filter(r => (r.low || "").toString().trim() !== "").map(r => ({
                    sign: r.sign === "E" ? "E" : "I",
                    option: r.option || "EQ",
                    low: r.low || "",
                    high: fr.selOptNeedsHigh(r.option) ? (r.high || "") : ""
                }));
            } else if (v.type === "BOOLEAN") {
                values[v.name] = !!values[v.name];
            }
        });
        const { valid, errors } = UI.validateScreenVars(vars, values, fr.varOptions, sel);
        if (!valid) throw this.fail("INVALID", `Revisa ${errors.length} campo(s) antes de ejecutar.`, { errors });

        // 1) ficheros locales -> storage (el valor de la variable pasa a ser su ruta)
        for (const varName of Object.keys(files)) {
            const file = files[varName];
            if (!file || typeof file.name !== "string") continue;
            const storagePath = Storage.buildPath(ref, varName, file.name);
            await Storage.upload(file, storagePath);
            values[varName] = storagePath;
        }
        // 2) orquestador
        const runId = Provider.newId();
        await fr.callOrchestrator(ref, values, runId);
        return { runId, name: fr.flow.name };
    },

    async status(project, ref) {
        await Informes.guard();
        const own = await Provider.runQuery(`SELECT PROYECTO_ID FROM ${Provider.qualifyControl("FLUJOS")} WHERE FLUJO_ID = '${Provider.esc(ref)}'`);
        if (!own[0] || own[0].PROYECTO_ID !== project.PROYECTO_ID) throw this.fail("NOT_FOUND", "Este flujo de carga no pertenece al proyecto actual.");
        const runRows = await Provider.runQuery(`
            SELECT RUN_ID, ESTADO, MENSAJE, FECHA_INICIO, FECHA_FIN FROM ${Provider.qualifyControl("FLUJOS_RUNS")}
            WHERE FLUJO_ID = '${Provider.esc(ref)}' ORDER BY FECHA_INICIO DESC LIMIT 1`);
        const run = runRows[0];
        if (!run) return { run: null, steps: {} };
        const stepRows = await Provider.runQuery(`
            SELECT PASO_ID, ESTADO, FILAS, MENSAJE FROM ${Provider.qualifyControl("FLUJOS_RUN_STEPS")}
            WHERE RUN_ID = '${Provider.esc(run.RUN_ID)}'`);
        const steps = {};
        stepRows.forEach(s => { steps[s.PASO_ID] = { estado: s.ESTADO, filas: s.FILAS, mensaje: s.MENSAJE || "" }; });
        return {
            run: { id: run.RUN_ID, estado: run.ESTADO, mensaje: run.MENSAJE || "", inicio: run.FECHA_INICIO == null ? "" : String(run.FECHA_INICIO), fin: run.FECHA_FIN == null ? "" : String(run.FECHA_FIN) },
            steps
        };
    },

    async handle(project, op, p) {
        if (!project) throw this.fail("BAD_INPUT", "No hay ningún proyecto seleccionado.");
        p = p || {};
        try {
            switch (op) {
                case "fl.list": return await this.list(project);
                case "fl.describe": return await this.describe(project, p.ref);
                case "fl.run": return await this.run(project, p.ref, p.values, p.sel, p.files);
                case "fl.status": return await this.status(project, p.ref);
                default: throw this.fail("NOT_AVAILABLE", "Esta función todavía no está disponible.");
            }
        } catch (err) {
            throw Informes.asError(err);
        }
    }
};

window.InformesFlujos = InformesFlujos;

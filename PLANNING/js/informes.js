/**
 * ============================================================
 * DRACO PLANNING — INFORMES (acceso a datos)
 * ============================================================
 * Un "informe" es una página (KPIs, gráficos, tablas, textos) que se
 * diseña en generador.html con la herramienta de informes y se abre en
 * visor.html. Aquí solo vive el acceso a sus dos tablas de control:
 *
 *   DRACO_CONTROL.INFORMES            una fila por informe, con su
 *                                     definición (SPEC_JSON) y PROYECTO_ID
 *   DRACO_CONTROL.INFORMES_FAVORITOS  los favoritos de cada usuario
 *
 * Decisiones (acordadas):
 *  - Solo se guarda la DEFINICIÓN en JSON, nunca el HTML: el visor genera
 *    el HTML cada vez que abre el informe.
 *  - Todo va acotado al proyecto: cada consulta filtra por PROYECTO_ID.
 *  - Un único diseñador por informe: guardar sobrescribe (sin control de
 *    versiones). "Guardar como" crea un informe nuevo.
 *  - Permisos: los de DRACO_CONTROL (no hay módulo de roles todavía).
 *  - Todo pasa por Provider, así que funciona igual en BigQuery y Snowflake.
 *  - Tope de tamaño de la definición: MAX_BYTES (la consulta SQL que la
 *    guarda tiene un límite de tamaño en el motor, de en torno a 1 MB).
 *
 * Los errores llevan un `code` para que quien llama pueda reaccionar:
 *    NO_AUTH      sesión caducada (hay que volver a iniciar sesión)
 *    NAME_EXISTS  ya hay un informe con ese nombre en el proyecto
 *    TOO_BIG      la definición supera MAX_BYTES
 *    NOT_FOUND    el informe no existe (o es de otro proyecto)
 *    BAD_JSON     la definición guardada no se puede leer
 *    BAD_INPUT    datos de entrada no válidos
 */
const Informes = {
    TABLE: "INFORMES",
    FAV_TABLE: "INFORMES_FAVORITOS",
    MAX_BYTES: 500000,
    MAX_NAME: 120,
    MAX_DESC: 500,

    fail(code, message) {
        const err = new Error(message);
        err.code = code;
        return err;
    },

    /** Comprueba la sesión antes de cada operación (y renueva Snowflake si hace falta). */
    async guard() {
        if (Provider.key() === "snowflake" && typeof SF !== "undefined" && SF.ensureConnected) {
            try { await SF.ensureConnected(); } catch (e) { /* se resuelve abajo */ }
        }
        if (!Provider.isConnected()) {
            throw this.fail("NO_AUTH", `Tu sesión de ${Provider.label()} ha caducado.`);
        }
        this.ensureTables().catch(() => {});   // sin esperar: ver ensureTables() y q()
    },

    /**
     * Garantiza que existen INFORMES e INFORMES_FAVORITOS (sin bloquear: ver guard() y q()). Normalmente ya las
     * ha creado el bootstrap al iniciar sesión (index.html -> DracoSchema.bootstrap),
     * pero quien tenga una sesión viva cuando se despliega esta versión puede
     * abrir generador.html/visor.html sin pasar por el login. CREATE TABLE IF
     * NOT EXISTS es inocuo si ya existen; se hace una sola vez por pestaña y
     * motor. Si el usuario no tiene permiso para crear tablas, se ignora y las
     * consultas fallarán con el mensaje real del motor.
     */
    ensureP: null,
    ensureTables(force) {
        if (typeof DracoSchema === "undefined") return Promise.resolve();
        const key = `draco_informes_tablas_${Provider.key()}_${Provider.databaseLabel()}`;
        if (!force) {
            try { if (sessionStorage.getItem(key)) return Promise.resolve(); } catch (e) { /* sin sessionStorage: se comprueba cada vez */ }
            if (this.ensureP) return this.ensureP;
        }
        this.ensureP = (async () => {
            try {
                await Promise.all([this.TABLE, this.FAV_TABLE, "TABLAS_DRACO"].map(t => Provider.runQuery(DracoSchema.ddl(t))));
                try { sessionStorage.setItem(key, "1"); } catch (e) { /* nada */ }
            } catch (e) { /* ver nota arriba */ }
        })();
        return this.ensureP;
    },

    isMissingTable(err) { return /not found|does not exist|no existe|doesn't exist/i.test((err && err.message) || ""); },

    /**
     * Ejecuta una consulta de este módulo. La comprobación de tablas (ensureTables) ya
     * no frena cada apertura: si la consulta falla porque la tabla no existe, se
     * crean las tablas y se repite UNA vez.
     */
    async q(sql) {
        try { return await Provider.runQuery(sql); }
        catch (err) {
            if (!this.isMissingTable(err)) throw err;
            await this.ensureTables(true);
            return Provider.runQuery(sql);
        }
    },

    /** Convierte el error NO_AUTH de BigQuery (BQ.request) en el mismo código que usa guard(). */
    asError(err) {
        if (err && err.code) return err;
        const e = new Error((err && err.message) || "Error desconocido");
        e.details = err && err.details;
        return e;
    },

    /** Fecha de una fila: BigQuery devuelve epoch en segundos (se muestra en hora local); Snowflake, el texto tal cual. */
    tsExpr(col) {
        return Provider.key() === "snowflake"
            ? `TO_CHAR(${col}, 'YYYY-MM-DD HH24:MI')`
            : `UNIX_SECONDS(${col})`;
    },
    formatTs(v) {
        if (v === null || v === undefined || v === "") return "";
        if (Provider.key() === "snowflake") return String(v);
        const n = Number(v);
        if (!isFinite(n)) return String(v);
        const d = new Date(n * 1000);
        const p = (x) => String(x).padStart(2, "0");
        return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
    },

    rowToMeta(r) {
        return {
            id: r.INFORME_ID,
            nombre: r.NOMBRE,
            descripcion: r.DESCRIPCION || "",
            usuario: r.USUARIO || "",
            creado: this.formatTs(r.CREADO),
            modificado: this.formatTs(r.MODIFICADO)
        };
    },

    // ---------------------------------------------------------
    // Lectura
    // ---------------------------------------------------------
    /** Lista (sin la definición) los informes de un proyecto, los últimos modificados primero. */
    async list(projectId) {
        await this.guard();
        try {
            const sql = `SELECT INFORME_ID, NOMBRE, DESCRIPCION, USUARIO,
                                ${this.tsExpr("FECHA_CREACION")} AS CREADO,
                                ${this.tsExpr("FECHA_MODIFICACION")} AS MODIFICADO
                         FROM ${Provider.qualifyControl(this.TABLE)}
                         WHERE PROYECTO_ID = '${Provider.esc(projectId)}'
                         ORDER BY FECHA_MODIFICACION DESC`;
            const rows = await this.q(sql);
            return rows.map(r => this.rowToMeta(r));
        } catch (err) { throw this.asError(err); }
    },

    /** Datos de un informe SIN la definición, buscándolo en cualquier proyecto (para abrir enlaces directos). */
    async getMeta(id) {
        await this.guard();
        try {
            const sql = `SELECT INFORME_ID, PROYECTO_ID, NOMBRE, DESCRIPCION, USUARIO,
                                ${this.tsExpr("FECHA_CREACION")} AS CREADO,
                                ${this.tsExpr("FECHA_MODIFICACION")} AS MODIFICADO
                         FROM ${Provider.qualifyControl(this.TABLE)}
                         WHERE INFORME_ID = '${Provider.esc(id)}'`;
            const rows = await this.q(sql);
            if (!rows.length) throw this.fail("NOT_FOUND", "El informe no existe.");
            return { ...this.rowToMeta(rows[0]), proyectoId: rows[0].PROYECTO_ID };
        } catch (err) { throw this.asError(err); }
    },

    /** Un informe con su definición ya convertida a objeto. Si se pasa projectId, debe pertenecer a ese proyecto. */
    async get(id, projectId) {
        await this.guard();
        try {
            const sql = `SELECT INFORME_ID, PROYECTO_ID, NOMBRE, DESCRIPCION, USUARIO, SPEC_JSON,
                                ${this.tsExpr("FECHA_CREACION")} AS CREADO,
                                ${this.tsExpr("FECHA_MODIFICACION")} AS MODIFICADO
                         FROM ${Provider.qualifyControl(this.TABLE)}
                         WHERE INFORME_ID = '${Provider.esc(id)}'
                           ${projectId ? `AND PROYECTO_ID = '${Provider.esc(projectId)}'` : ""}`;
            const rows = await this.q(sql);
            if (!rows.length) throw this.fail("NOT_FOUND", "El informe no existe en este proyecto.");
            let spec;
            try { spec = JSON.parse(rows[0].SPEC_JSON); }
            catch (e) { throw this.fail("BAD_JSON", "La definición guardada de este informe no se puede leer."); }
            return { ...this.rowToMeta(rows[0]), proyectoId: rows[0].PROYECTO_ID, spec };
        } catch (err) { throw this.asError(err); }
    },

    // ---------------------------------------------------------
    // Escritura
    // ---------------------------------------------------------
    /**
     * Guarda un informe. Sin `id`, crea uno nuevo; con `id`, sobrescribe ese
     * (y si ya no existiera, lo vuelve a crear con ese mismo id para no
     * perder el trabajo). Devuelve { id, nombre, descripcion }.
     */
    async save({ id, projectId, nombre, descripcion, spec }) {
        if (!projectId) throw this.fail("BAD_INPUT", "Falta el proyecto.");
        nombre = String(nombre || "").trim();
        descripcion = String(descripcion || "").trim();
        if (!nombre) throw this.fail("BAD_INPUT", "El informe necesita un nombre.");
        if (nombre.length > this.MAX_NAME) throw this.fail("BAD_INPUT", `El nombre no puede pasar de ${this.MAX_NAME} caracteres.`);
        if (descripcion.length > this.MAX_DESC) throw this.fail("BAD_INPUT", `La descripción no puede pasar de ${this.MAX_DESC} caracteres.`);
        if (!spec || typeof spec !== "object" || !Array.isArray(spec.sections)) {
            throw this.fail("BAD_INPUT", "La definición del informe no es válida.");
        }

        const json = JSON.stringify(spec);
        const bytes = new TextEncoder().encode(json).length;
        if (bytes > this.MAX_BYTES) {
            throw this.fail("TOO_BIG", `El informe ocupa ${Math.round(bytes / 1024)} KB y el máximo es ${Math.round(this.MAX_BYTES / 1000)} KB.`);
        }

        await this.guard();
        try {
            const t = Provider.qualifyControl(this.TABLE);

            // Nombre único dentro del proyecto (sin distinguir mayúsculas)
            const dup = await this.q(`SELECT INFORME_ID FROM ${t}
                WHERE PROYECTO_ID = '${Provider.esc(projectId)}' AND UPPER(NOMBRE) = UPPER('${Provider.esc(nombre)}')`);
            if (dup.some(r => r.INFORME_ID !== id)) {
                throw this.fail("NAME_EXISTS", "Ya existe un informe con ese nombre en este proyecto.");
            }

            let exists = false;
            if (id) {
                const cur = await this.q(`SELECT INFORME_ID FROM ${t}
                    WHERE INFORME_ID = '${Provider.esc(id)}' AND PROYECTO_ID = '${Provider.esc(projectId)}'`);
                exists = cur.length > 0;
            }

            if (exists) {
                await this.q(`UPDATE ${t}
                    SET NOMBRE = '${Provider.esc(nombre)}',
                        DESCRIPCION = '${Provider.esc(descripcion)}',
                        SPEC_JSON = '${Provider.esc(json)}',
                        FECHA_MODIFICACION = CURRENT_TIMESTAMP()
                    WHERE INFORME_ID = '${Provider.esc(id)}' AND PROYECTO_ID = '${Provider.esc(projectId)}'`);
                return { id, nombre, descripcion };
            }

            const newId = id || Provider.newId();
            await this.q(`INSERT INTO ${t}
                (INFORME_ID, PROYECTO_ID, NOMBRE, DESCRIPCION, SPEC_JSON, USUARIO, FECHA_CREACION, FECHA_MODIFICACION)
                VALUES ('${Provider.esc(newId)}', '${Provider.esc(projectId)}', '${Provider.esc(nombre)}', '${Provider.esc(descripcion)}',
                        '${Provider.esc(json)}', ${Provider.currentUserExpr()}, CURRENT_TIMESTAMP(), CURRENT_TIMESTAMP())`);
            return { id: newId, nombre, descripcion };
        } catch (err) { throw this.asError(err); }
    },

    /** Elimina un informe y los favoritos que apuntaban a él. */
    async remove(id, projectId) {
        if (!id || !projectId) throw this.fail("BAD_INPUT", "Faltan datos para eliminar el informe.");
        await this.guard();
        try {
            await this.q(`DELETE FROM ${Provider.qualifyControl(this.FAV_TABLE)}
                WHERE INFORME_ID = '${Provider.esc(id)}' AND PROYECTO_ID = '${Provider.esc(projectId)}'`);
            await this.q(`DELETE FROM ${Provider.qualifyControl(this.TABLE)}
                WHERE INFORME_ID = '${Provider.esc(id)}' AND PROYECTO_ID = '${Provider.esc(projectId)}'`);
        } catch (err) { throw this.asError(err); }
    },

    // ---------------------------------------------------------
    // Favoritos (por usuario)
    // ---------------------------------------------------------
    /** Ids de los informes del proyecto que el usuario actual tiene en favoritos. */
    async favoriteIds(projectId) {
        await this.guard();
        try {
            const rows = await this.q(`SELECT INFORME_ID FROM ${Provider.qualifyControl(this.FAV_TABLE)}
                WHERE PROYECTO_ID = '${Provider.esc(projectId)}' AND USUARIO = ${Provider.currentUserExpr()}`);
            return new Set(rows.map(r => r.INFORME_ID));
        } catch (err) { throw this.asError(err); }
    },

    /** Marca o desmarca un favorito. Devuelve el estado final (true = es favorito). */
    async setFavorite(projectId, id, on) {
        await this.guard();
        try {
            const t = Provider.qualifyControl(this.FAV_TABLE);
            const where = `PROYECTO_ID = '${Provider.esc(projectId)}' AND INFORME_ID = '${Provider.esc(id)}' AND USUARIO = ${Provider.currentUserExpr()}`;
            if (!on) {
                await this.q(`DELETE FROM ${t} WHERE ${where}`);
                return false;
            }
            const cur = await this.q(`SELECT INFORME_ID FROM ${t} WHERE ${where}`);
            if (!cur.length) {
                await this.q(`INSERT INTO ${t} (PROYECTO_ID, INFORME_ID, USUARIO, FECHA_ALTA)
                    SELECT '${Provider.esc(projectId)}', '${Provider.esc(id)}', ${Provider.currentUserExpr()}, CURRENT_TIMESTAMP()`);
            }
            return true;
        } catch (err) { throw this.asError(err); }
    }
};

window.Informes = Informes;

#!/usr/bin/env python3
"""
sync-addin.py — genera PLANNING/tablas-draco/ (el diseñador y el motor de "Tablas Draco") a partir de
ADDIN/src, sin editar a mano una copia del addin.

  python3 tools/sync-addin.py            # regenera PLANNING/tablas-draco/
  python3 tools/sync-addin.py --check    # solo comprueba que los parches siguen encajando

Qué hace:
  1. Extrae de ADDIN/src/taskpane.html SOLO la vista del diseñador de informes (#view-informes; las vistas
     de conexión y de modelo semántico las sustituye Planning, que ya tiene su conexión y sus cubos).
  2. Copia los JS/CSS del addin que hacen falta y las páginas de diálogo (filtros, miembros).
  3. Aplica PARCHES mínimos a esos JS (lista PATCHES). Cada parche exige que su "ancla" exista exactamente
     una vez: si el addin cambia y un ancla desaparece, el script FALLA y dice cuál, en vez de generar un
     paquete roto en silencio. Para traer novedades del addin: actualizar ADDIN/src y volver a ejecutarlo.
  4. Copia las plantillas propias (tools/templates/): puente de Excel (host-bridge), de proveedor
     (provider-bridge), arranque del modelo (model-bootstrap) y API del motor (engine-api).
"""
import os, re, shutil, sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "ADDIN", "src")
OUT = os.path.join(ROOT, "PLANNING", "tablas-draco")
TPL = os.path.join(ROOT, "tools", "templates")
CHECK = "--check" in sys.argv
# Se añade a todos los <script src> y <link href> de las páginas generadas: sin ello el navegador puede seguir usando
# una copia antigua (p. ej. de model-bootstrap.js) tras actualizar. Súbelo cuando cambies tablas-draco/.
BUNDLE_VERSION = "20261021b"

# JS del addin que se copian tal cual (más los parches de abajo)
JS_COPY = ["semanticModelStore.js", "reportStore.js", "filterRangeStore.js", "excelService.js",
           "filterModal.js", "distributeValues.js", "variablesStore.js", "calcFormula.js", "commands.js", "taskpane.js",
           "calcEditor.js", "calcDialog.js", "variablesDialog.js", "filterDialog.js", "memberPicker.js"]
CSS_COPY = ["theme.css", "styles.css", "modal.css", "distributeValues.css", "lkmlDialog.css"]
# Páginas de diálogo (Office.context.ui.displayDialogAsync) que se abren desde el diseñador
DIALOGS = ["filterDialog.html", "memberPicker.html", "addFilterRange.html", "calcDialog.html", "variablesDialog.html"]
# Plantillas propias -> js/
TEMPLATES = ["host-bridge.js", "provider-bridge.js", "model-bootstrap.js", "engine-api.js"]

# (archivo, ancla EXACTA que debe aparecer UNA vez, reemplazo, etiqueta)
PATCHES = [
    # --- executeSQL: la conexión es la ya autenticada de Planning (BigQuery o Snowflake), en vez de pedir el
    #     token a localStorage con fetch propio. Se reutiliza el conversor a "pseudo JSON de BigQuery" del addin.
    ("js/commands.js",
     '''    if (Provider.key() === "snowflake") {
        const rows = await SF.runQuery(sql);
        return snowflakeRowsToPseudoBqJson(rows);
    }

    if (Provider.key() === "fabric") {
        // Mismo texto "pseudo JSON de BigQuery" que usan los parsers existentes''',
     '''    // [Tablas Draco] Planning ya tiene la conexión autenticada (BigQuery o Snowflake): se usa siempre ella.
    if (window.DracoEngineApi) {
        return snowflakeRowsToPseudoBqJson(await Provider.runQuery(sql));
    }

    if (Provider.key() === "snowflake") {
        const rows = await SF.runQuery(sql);
        return snowflakeRowsToPseudoBqJson(rows);
    }

    if (Provider.key() === "fabric") {
        // Mismo texto "pseudo JSON de BigQuery" que usan los parsers existentes''',
     "executeSQL usa la conexión de Planning"),
    # --- informe DINÁMICO: si hay SQL ya generado (guardado), no se vuelve a construir
    ("js/commands.js",
     '''        sql = (window.PowerBIQuery && window.PowerBIQuery.isActiveModel())
            ? window.PowerBIQuery.buildDynamicQuery(ReportState, measuresGrid, atributesGrid, subtotalsOnTop)
            : buildSQL(relGrid, measuresGrid, atributesGrid, subtotalsOnTop);
        tSub = draco_perfMark(reportId, "  loadReportDefinition + buildSQL (JS puro)", tSub);

        const queryText = typeof sql === "string" ? sql : sql.text;
        console.log("BuildSQL ->", queryText);''',
     '''        // [Tablas Draco] con SQL ya generado y guardado no se vuelve a construir (loadReportDefinition sigue
        // haciendo falta: deja listo ReportState, que usa el pintado).
        sql = (window.DracoEngineApi && window.DracoEngineApi.storedSql)
            ? window.DracoEngineApi.storedSql
            : (window.PowerBIQuery && window.PowerBIQuery.isActiveModel())
            ? window.PowerBIQuery.buildDynamicQuery(ReportState, measuresGrid, atributesGrid, subtotalsOnTop)
            : buildSQL(relGrid, measuresGrid, atributesGrid, subtotalsOnTop);
        tSub = draco_perfMark(reportId, "  loadReportDefinition + buildSQL (JS puro)", tSub);

        const queryText = typeof sql === "string" ? sql : sql.text;
        console.log("BuildSQL ->", queryText);''',
     "informe dinámico con SQL guardado"),
    # --- informe ESTÁTICO ("Fijo"): igual
    ("js/commands.js",
     '''        sql = await buildSQLFixed(context, editReportGrid, relGrid, measuresGrid, atributesGrid, csvGrid);
        tSub = draco_perfMark(reportId, "  loadReportDefinition + buildSQLFixed (JS puro)", tSub);

        await context.sync();
        draco_perfMark(reportId, "  Escribir/sincronizar", tSub);
    });
    tPerf = draco_perfMark(reportId, "Generando SQL", tPerf);''',
     '''        sql = (window.DracoEngineApi && window.DracoEngineApi.storedSql)
            ? window.DracoEngineApi.storedSql
            : await buildSQLFixed(context, editReportGrid, relGrid, measuresGrid, atributesGrid, csvGrid);
        tSub = draco_perfMark(reportId, "  loadReportDefinition + buildSQLFixed (JS puro)", tSub);

        await context.sync();
        draco_perfMark(reportId, "  Escribir/sincronizar", tSub);
    });
    tPerf = draco_perfMark(reportId, "Generando SQL", tPerf);''',
     "informe estático con SQL guardado"),
    # --- informe HÍBRIDO: igual (el plan de pintado sí se construye: hace falta para pintar)
    ("js/commands.js",
     '''            built = await buildSQLHybrid(context, editReportGrid, relGrid, measuresGrid, atributesGrid, csvGrid, reportId);
            await context.sync();''',
     '''            built = await buildSQLHybrid(context, editReportGrid, relGrid, measuresGrid, atributesGrid, csvGrid, reportId);
            // [Tablas Draco] con SQL ya generado y guardado se ejecuta ese (el plan de pintado sí hace falta)
            if (window.DracoEngineApi && window.DracoEngineApi.storedSql) built.sql = window.DracoEngineApi.storedSql;
            await context.sync();''',
     "informe híbrido con SQL guardado"),
]


def read(p):
    with open(p, encoding="utf-8", errors="surrogateescape", newline="") as f:
        return f.read()


def write(p, s):
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, "w", encoding="utf-8", errors="surrogateescape", newline="") as f:
        f.write(s)


def norm(s):
    return s.replace("\r\n", "\n")


def main():
    problems = []
    files = {}      # ruta relativa -> contenido
    # ---- JS
    for name in JS_COPY:
        files["js/" + name] = norm(read(os.path.join(SRC, "js", name)))
    # ---- parches
    for rel, anchor, repl, label in PATCHES:
        s = files[rel]
        n = s.count(anchor)
        if n != 1:
            problems.append(f"PARCHE «{label}» en {rel}: el ancla aparece {n} veces (debe ser 1). ¿Cambió el addin?")
            continue
        files[rel] = s.replace(anchor, repl, 1)
    # ---- CSS
    for name in CSS_COPY:
        files["css/" + name] = norm(read(os.path.join(SRC, "css", name)))
    # ---- taskpane.html: solo la vista del diseñador
    html = norm(read(os.path.join(SRC, "taskpane.html")))
    m = re.search(r"<style>.*?</style>", html, re.S)
    a, b = html.find('<div id="view-informes"'), html.find('<div id="view-conexion"')
    if not m or a < 0 or b < 0 or b < a:
        problems.append("taskpane.html: no se encuentra <style> o las vistas #view-informes / #view-conexion.")
    else:
        view = html[a:b].rstrip()
        if view.count("<div") != view.count("</div>"):
            problems.append("taskpane.html: la vista #view-informes no está equilibrada (<div> vs </div>).")
        files["taskpane.html"] = f'''<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Tablas Draco · Diseño del informe</title>
    <!-- GENERADO por tools/sync-addin.py desde ADDIN/src/taskpane.html (solo #view-informes). No editar a mano. -->
    <script src="js/host-bridge.js"></script>
    <script src="js/provider-bridge.js"></script>
    <link rel="stylesheet" href="css/theme.css">
    <link rel="stylesheet" href="css/styles.css">
    <link rel="stylesheet" href="css/modal.css">
    <link rel="stylesheet" href="css/distributeValues.css">
    {m.group(0)}
    <style id="dracoOverrides">
        /* Tablas Draco: una tabla = UN informe, y no hay Excel real: se ocultan los controles que no aplican */
        #reportSelector, #btnAddReport, #btnRefreshAll, #btnDistribuirValores, #btnSaveToBucket {{ display: none !important; }}
        /* Estático / Híbrido: SÍ disponibles. Las líneas del eje Estático se editan en la rejilla del diseñador
           (seleccionar + Supr, botones «Miembro» y «Fórmula», doble clic): ver js/draco-table-designer.js */
    </style>
</head>
<body>
{view}

    <script src="js/semanticModelStore.js"></script>
    <script src="js/reportStore.js"></script>
    <script src="js/filterRangeStore.js"></script>
    <script src="js/excelService.js"></script>
    <script src="js/filterModal.js"></script>
    <script src="js/distributeValues.js"></script>
    <script src="js/variablesStore.js"></script>
    <script src="js/calcFormula.js"></script>
    <script src="js/commands.js"></script>
    <script src="js/engine-api.js"></script>
    <script src="js/taskpane.js"></script>
    <script src="js/calcEditor.js"></script>
    <script src="js/model-bootstrap.js"></script>
</body>
</html>
'''
    # ---- versión en las referencias locales de las páginas (caché)
    def versioned(h):
        h = re.sub(r'(<script src="js/[A-Za-z0-9_.-]+\.js)(")', r'\1?v=%s\2' % BUNDLE_VERSION, h)
        return re.sub(r'(<link rel="stylesheet" href="css/[A-Za-z0-9_.-]+\.css)(")', r'\1?v=%s\2' % BUNDLE_VERSION, h)
    # (los diálogos y engine.html se versionan más abajo, cuando ya están construidos)
    # ---- diálogos: solo cambia office.js por el puente
    for name in DIALOGS:
        d = norm(read(os.path.join(SRC, name)))
        tag = '<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>'
        if d.count(tag) != 1:
            problems.append(f"{name}: no se encuentra la etiqueta de office.js (1 vez).")
        else:
            d = d.replace(tag, '<script src="js/host-bridge.js"></script>', 1)
        files[name] = d
    # ---- plantillas propias
    for name in TEMPLATES:
        p = os.path.join(TPL, name)
        if not os.path.exists(p):
            problems.append(f"falta la plantilla tools/templates/{name}")
        else:
            files["js/" + name] = norm(read(p))
    # ---- motor sin interfaz (engine.html): mismas piezas, sin la vista
    files["engine.html"] = '''<!DOCTYPE html>
<html lang="es"><head><meta charset="UTF-8"><title>Tablas Draco · motor</title>
<!-- GENERADO por tools/sync-addin.py. Motor sin interfaz: ejecuta un informe y deja el resultado en su hoja. -->
<script src="js/host-bridge.js"></script><script src="js/provider-bridge.js"></script>
<script>if (window.__wteSignalModelsReady) window.__wteSignalModelsReady();</script>
</head><body>
<script src="js/semanticModelStore.js"></script><script src="js/reportStore.js"></script><script src="js/filterRangeStore.js"></script>
<script src="js/excelService.js"></script><script src="js/variablesStore.js"></script><script src="js/calcFormula.js"></script><script src="js/commands.js"></script><script src="js/engine-api.js"></script>
</body></html>
'''
    for rel in list(files):
        if rel.endswith(".html"):
            files[rel] = versioned(files[rel])
    if problems:
        print("\n".join("✗ " + p for p in problems))
        sys.exit(1)
    if CHECK:
        print(f"✓ comprobado: {len(PATCHES)} parches encajan en ADDIN/src ({len(files)} archivos se generarían)")
        return
    if os.path.isdir(OUT):
        shutil.rmtree(OUT)
    for rel, s in files.items():
        write(os.path.join(OUT, rel), s)
    # recursos que referencian las páginas (iconos, etc.)
    refs = set()
    for rel, s in files.items():
        if rel.endswith((".html", ".css", ".js")):
            refs.update(re.findall(r"""['"(]((?:\.\./)?assets/[A-Za-z0-9_./-]+\.(?:png|svg|gif|jpg))""", s))
    copied = 0
    for r in sorted(refs):
        r = r.replace("../", "")
        src = os.path.join(os.path.dirname(SRC), r)           # ADDIN/assets/...
        if os.path.exists(src):
            dst = os.path.join(OUT, r)
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            shutil.copyfile(src, dst); copied += 1
    print(f"✓ generado {os.path.relpath(OUT, ROOT)}: {len(files)} archivos, {copied} recursos; {len(PATCHES)} parches aplicados")


main()

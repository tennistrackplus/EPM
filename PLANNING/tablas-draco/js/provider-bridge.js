/**
 * provider-bridge.js — sustituye a config.js + bigquery.js + provider.js del add-in.
 * Aquí se reutiliza SIEMPRE la conexión ya autenticada de la aplicación anfitriona (PLANNING).
 * Firma distinta a propósito: el Provider del add-in cualifica con (project, dataset, table); el de
 * PLANNING con (dataset, table) porque el proyecto activo ya lo resuelve él. Este puente traduce.
 * (Generado por tools/sync-addin.py desde tools/templates/.)
 */
const Provider = {
    key() { return window.parent.Provider.key(); },
    qualify(project, dataset, table) { return window.parent.Provider.qualify(dataset, table); },
    esc(v) { return window.parent.Provider.esc(v); },
    runQuery(sql) { return window.parent.Provider.runQuery(sql); },
    // Tipo de texto para CAST(...) (commands.js lo usa en jerarquías)
    stringType() { return "STRING"; }
};
// Por si algún resto de código llama a BQ en vez de a Provider: mismo puente.
// Token y proyecto: los de la sesión de Planning (mismo origen; el token vive en su sessionStorage).
const BQ = {
    runQuery: (sql) => window.parent.Provider.runQuery(sql),
    getBillingProject: () => null,
    getToken: () => window.parent.BQ.getToken(),
    isConnected: () => window.parent.BQ.isConnected(),
    canWriteBigQuery: () => (window.parent.BQ.canWriteBigQuery ? window.parent.BQ.canWriteBigQuery() : true),
    jobProjectFor: (sql, preferred) => preferred
        || ((/`([a-z][a-z0-9-]{4,61}[a-z0-9])\.[A-Za-z0-9_]+\.[^`]+`/.exec(String(sql || "")) || [])[1])
        || window.parent.BQ.getGcpProject()
};

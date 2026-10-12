/**
 * provider-bridge.js — sustituye a config.js + bigquery.js + snowflake.js + provider.js del add-in.
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
    stringType() { return window.parent.Provider.key() === "snowflake" ? "VARCHAR" : "STRING"; }
};
// Por si algún resto de código llama a BQ/SF en vez de a Provider: mismo puente.
const BQ = { runQuery: (sql) => window.parent.Provider.runQuery(sql), getBillingProject: () => null };
const SF = { runQuery: (sql) => window.parent.Provider.runQuery(sql) };

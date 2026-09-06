/**
 * provider-bridge.js — sustituye a config.js + bigquery.js + snowflake.js +
 * provider.js del add-in (que autenticaban por su cuenta contra BigQuery o
 * Snowflake). Aquí se reutiliza SIEMPRE la conexión ya autenticada de la
 * aplicación anfitriona (PLANNING, window.parent.Provider).
 *
 * Firma distinta a propósito: el Provider del add-in cualifica con
 * (project, dataset, table) porque el modelo semántico del add-in guarda
 * los tres por separado; el Provider de PLANNING cualifica con
 * (dataset/schema, table) porque el proyecto activo ya lo resuelve él
 * solo. Aquí se hace de puente entre ambas firmas.
 */
const Provider = {
    key() { return window.parent.Provider.key(); },
    qualify(project, dataset, table) { return window.parent.Provider.qualify(dataset, table); },
    esc(v) { return window.parent.Provider.esc(v); },
    runQuery(sql) { return window.parent.Provider.runQuery(sql); }
};

// Por si algún resto de código (copiado tal cual del add-in) llama
// directamente a BQ/SF en vez de a Provider: mismo puente.
const BQ = { runQuery: (sql) => window.parent.Provider.runQuery(sql) };
const SF = { runQuery: (sql) => window.parent.Provider.runQuery(sql) };

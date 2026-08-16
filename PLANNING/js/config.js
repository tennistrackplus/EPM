/**
 * ============================================================
 * DRACO PLANNING — CONFIGURACIÓN
 * ============================================================
 * Rellena aquí tus credenciales antes de desplegar la app.
 */
const DracoConfig = {
    // ---------------------------------------------------------
    // BigQuery (Google OAuth 2.0 — flujo implícito)
    // ---------------------------------------------------------
    googleClientId: "TU_GOOGLE_CLIENT_ID.apps.googleusercontent.com",
    googleScopes: [
        "https://www.googleapis.com/auth/bigquery",
        "https://www.googleapis.com/auth/userinfo.email"
    ].join(" "),

    // ---------------------------------------------------------
    // Snowflake (OAuth 2.0 — Authorization Code + PKCE, cliente público)
    // ---------------------------------------------------------
    // Client ID que Snowflake genera al crear la SECURITY INTEGRATION
    // (ver sql/01_snowflake_oauth_integration.sql). No hace falta client
    // secret: al ser un cliente público, PKCE sustituye al secreto.
    snowflakeClientId: "TU_SNOWFLAKE_OAUTH_CLIENT_ID",
    snowflakeScopes: "refresh_token",

    // Nombre del dataset/base de datos de control donde viven las tablas
    // maestras de Draco (PROYECTOS, DIMENSIONES, CUBOS).
    controlDataset: "DRACO_CONTROL",

    // Base de datos Snowflake donde vivirán todos los esquemas de Draco
    // (DRACO_CONTROL y DRACO_<PROYECTO> son esquemas dentro de ella).
    snowflakeDatabase: "DRACO",

    // Prefijo aplicado a los datasets/esquemas de cada proyecto Draco y a
    // las tablas físicas de dimensiones/cubos.
    prefix: "DRACO_"
};


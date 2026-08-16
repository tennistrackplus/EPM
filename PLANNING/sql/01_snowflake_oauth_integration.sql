-- ============================================================
-- DRACO PLANNING · Configuración de OAuth para Snowflake
-- ============================================================
-- Ejecuta esto UNA VEZ como ACCOUNTADMIN (o un rol con privilegio
-- CREATE INTEGRATION) antes de poder conectar Draco Planning a
-- Snowflake. Crea un cliente OAuth "público" (sin client secret,
-- protegido con PKCE) apto para una app de navegador.
--
-- Sustituye {REDIRECT_URI} por la URL real donde publiques Draco
-- Planning + "/auth-callback-snowflake.html", por ejemplo:
--   https://tuempresa.github.io/draco-planning/auth-callback-snowflake.html
-- ============================================================

CREATE SECURITY INTEGRATION IF NOT EXISTS DRACO_PLANNING_OAUTH
  TYPE = OAUTH
  OAUTH_CLIENT = CUSTOM
  OAUTH_CLIENT_TYPE = 'PUBLIC'
  OAUTH_REDIRECT_URI = '{REDIRECT_URI}'
  OAUTH_ISSUE_REFRESH_TOKENS = TRUE
  OAUTH_REFRESH_TOKEN_VALIDITY = 7776000  -- 90 días
  ENABLED = TRUE
  COMMENT = 'Cliente OAuth público (PKCE) para la app web Draco Planning';

-- Recupera el Client ID generado y pégalo en js/config.js -> snowflakeClientId
DESCRIBE SECURITY INTEGRATION DRACO_PLANNING_OAUTH;
-- Busca la propiedad OAUTH_CLIENT_ID en el resultado.

-- ============================================================
-- Notas importantes
-- ============================================================
-- 1. El usuario que inicie sesión debe tener asignado por defecto
--    (o poder seleccionar) un rol con privilegios para crear bases
--    de datos/esquemas/tablas (ej. SYSADMIN) y usar el warehouse
--    que indiques al conectar.
--
-- 2. Draco Planning usa la SQL API v2 de Snowflake
--    (https://<account>.snowflakecomputing.com/api/v2/statements).
--    Si al conectar ves errores de tipo "CORS" / "Failed to fetch"
--    en la consola del navegador, es porque ese endpoint no está
--    devolviendo cabeceras CORS para tu origen: necesitarás publicar
--    un pequeño proxy que reenvíe las peticiones (ver README.md,
--    sección "Snowflake · si te bloquea CORS").
--
-- 3. Si quieres restringir qué IPs pueden completar el intercambio
--    de token, puedes asociar una NETWORK POLICY a la integración
--    con el parámetro OAUTH_ALLOW_NON_TLS_REDIRECT_URI / NETWORK_POLICY.
-- ============================================================

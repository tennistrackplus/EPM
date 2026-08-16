# Draco Planning

App web funcional (HTML/CSS/JS puro, sin build ni frameworks) para modelar
proyectos, dimensiones y cubos sobre BigQuery **o Snowflake**, con el mismo
lenguaje visual que vuestra app anterior (EPM Data Studio).

## 1. Qué hay ya funcional

- **Login / landing (`index.html`)**: selector de conectores (BigQuery,
  Snowflake, Microsoft Fabric, Amazon, SAP Datasphere). **BigQuery y
  Snowflake están realmente conectados**; el resto son placeholders con
  "estará disponible próximamente".
- **Capa de proveedor (`js/provider.js`)**: todo el resto de la app
  (proyectos, dimensiones, cubos, progreso) habla con `Provider`, no
  directamente con BigQuery o Snowflake — por eso las mismas pantallas
  funcionan igual sobre cualquiera de los dos motores.
- **Conexión real a BigQuery**: OAuth 2.0 implícito (popup), token en
  `localStorage`, llamadas REST directas (`js/bigquery.js`).
- **Conexión real a Snowflake**: OAuth 2.0 *Authorization Code + PKCE*
  (cliente público, sin client secret) y SQL API v2 (`js/snowflake.js`).
  Requiere que un ACCOUNTADMIN ejecute una vez
  `sql/01_snowflake_oauth_integration.sql` para dar de alta el cliente
  OAuth de Draco Planning en tu cuenta de Snowflake.
- **Bootstrap automático del esquema de control**: al conectar (con
  cualquiera de los dos motores) y elegir tu "hogar" de trabajo (proyecto
  GCP, o cuenta+warehouse de Snowflake), la app crea si no existen el
  dataset/esquema `DRACO_CONTROL` y las 3 tablas maestras. El script SQL
  equivalente está en `sql/00_control_schema.sql`.
- **Panel principal (`app.html`)**: barra de proyecto (crear/eliminar),
  bloques **Administración** y **Mi Progreso** (maximizables/contraíbles),
  y dentro de Administración:
  - **Dimensiones**: la clave principal de la tabla es siempre el propio
    nombre de la dimensión (ej. dimensión `CUENTA` → columna `CUENTA`,
    clave). El resto de columnas son "Atributos", y cualquiera de ellos
    también se puede marcar como clave para soportar claves compuestas
    (ej. Clase de coste + Sociedad).
    - **Valores** (icono ▤): rejilla editable de los datos físicos —
      añadir filas, pegar bloques copiados de Excel, exportar a CSV/Excel,
      importar desde archivo (sustituyendo todo o de forma incremental
      por clave). "Guardar" vuelca la rejilla completa a la tabla.
    - **Jerarquías** (icono ⛓): arrastra atributos a una lista de niveles
      (superior → inferior) con vista previa en vivo sobre los datos
      reales de la dimensión.
  - **Cubos**: se seleccionan dimensiones ya creadas en el proyecto (cada
    una aporta su columna clave como FK) y se definen medidas libres
    (nombre + tipo), igual que antes.
  - El resto de módulos del menú (Cargas de datos, Flujos de carga,
    Funciones, Flujos de proceso, Workflows, Roles) son pantallas
    "próximamente".

El grid de valores usa **SheetJS** (cargado por CDN en `app.html`) para
leer/escribir CSV y Excel sin necesitar backend.

⚠️ **Importante**: al editar una dimensión o cubo, la tabla física se recrea
con `CREATE OR REPLACE TABLE`, lo que **borra los datos existentes** si
cambias las columnas. Es el comportamiento esperado en fase de modelado.

## 2. Puesta en marcha — BigQuery

1. Habilita la **BigQuery API** en tu proyecto de GCP.
2. Crea credenciales **OAuth 2.0 Client ID** (tipo *Aplicación web*):
   - **Orígenes de JavaScript autorizados**: el dominio donde publiques esta app.
   - **URI de redireccionamiento autorizados**: esa URL + `/auth-callback.html`.
3. Copia el Client ID en `js/config.js` (`googleClientId`).

## 3. Puesta en marcha — Snowflake

1. Como `ACCOUNTADMIN` (o rol con `CREATE INTEGRATION`), ejecuta
   `sql/01_snowflake_oauth_integration.sql` sustituyendo `{REDIRECT_URI}`
   por tu dominio + `/auth-callback-snowflake.html`. Esto crea un cliente
   OAuth **público** (protegido con PKCE, sin client secret — apto para
   una app de navegador).
2. Copia el `OAUTH_CLIENT_ID` que te devuelve `DESCRIBE SECURITY
   INTEGRATION DRACO_PLANNING_OAUTH;` en `js/config.js`
   (`snowflakeClientId`).
3. Al conectar desde la app, te pedirá: identificador de cuenta (ej.
   `xy12345.eu-west-1`), warehouse, base de datos (por defecto `DRACO`) y,
   opcionalmente, un rol.

### ⚠️ Snowflake · si te bloquea CORS

El dominio `*.snowflakecomputing.com` no siempre envía cabeceras
`Access-Control-Allow-Origin` para peticiones hechas directamente desde el
navegador (a diferencia de `googleapis.com`, pensado para clientes JS). Si
al conectar ves en la consola un error de tipo **"CORS" / "Failed to
fetch"**, necesitas un proxy ligero que reenvíe las peticiones añadiendo
esas cabeceras:

1. Despliega `proxy/cloudflare-worker.js` (instrucciones dentro del
   archivo — es gratis y tarda unos minutos).
2. En `js/snowflake.js`, cambia `apiBase: ""` por la URL de tu proxy.

No hace falta tocar nada más: el resto del código ya usa `this.base()` en
vez de la URL de Snowflake directamente, así que el cambio es de una sola
línea.

## 4. Puesta en marcha — común a ambos

Sirve la carpeta con cualquier servidor estático (no vale `file://`, hace
falta `http(s)://` para los popups de login y `postMessage`). Por ejemplo:
`npx serve .` o `python3 -m http.server`. Abre `index.html`, elige tu
conector, conecta y entra.

## 5. Modelo de datos de control

Dataset/esquema `DRACO_CONTROL` (uno por proyecto GCP o por base de datos
Snowflake):

| Tabla         | Columnas clave                                                                 |
|---------------|---------------------------------------------------------------------------------|
| `PROYECTOS`   | PROYECTO_ID, PROYECTO, DESCRIPCION, DATASET, USUARIO, FECHA_CREACION            |
| `DIMENSIONES` | DIMENSION_ID, PROYECTO_ID, DIMENSION, DESCRIPCION, TABLA, CAMPOS_JSON, ...      |
| `CUBOS`       | CUBO_ID, PROYECTO_ID, CUBOS, DESCRIPCION, TABLA, CAMPOS_JSON, ...               |
| `JERARQUIAS`  | JERARQUIA_ID, DIMENSION_ID, PROYECTO_ID, JERARQUIA, NIVELES_JSON, ...           |

`DIMENSIONES.CAMPOS_JSON` guarda un array; el primer elemento es siempre la
clave principal (`__isPrimaryName: true`, nombre = identificador de la
dimensión). `CUBOS.CAMPOS_JSON` guarda un objeto `{ dimensions: [...],
measures: [...] }`. `JERARQUIAS.NIVELES_JSON` guarda el array ordenado de
identificadores de columna, nivel superior primero.

Añadí `CAMPOS_JSON` (definición de columnas en JSON) sobre las 3 columnas
que pediste, para poder recargar el diseñador de campos al editar.

En BigQuery, cada proyecto Draco es un **dataset** `DRACO_<PROYECTO>`. En
Snowflake, cada proyecto Draco es un **esquema** `DRACO_<PROYECTO>` dentro
de la base de datos que elijas (por defecto `DRACO`). En ambos casos viven
ahí las tablas físicas `DRACO_<DIMENSION>` y `DRACO_<CUBO>`.

Los tipos de campo del diseñador (STRING, INTEGER, FLOAT, NUMERIC,
BOOLEAN, DATE, DATETIME, TIMESTAMP) se traducen automáticamente al tipo
físico de cada motor (`Provider.mapFieldType`) — por ejemplo FLOAT se
convierte en `FLOAT64` en BigQuery y se queda en `FLOAT` en Snowflake.

## 6. Estructura de archivos

```
index.html                    Landing / conectores
auth-callback.html            Callback OAuth de BigQuery
auth-callback-snowflake.html  Callback OAuth de Snowflake
app.html                      Panel principal
css/theme.css                 Design system (reutilizado de la app anterior)
css/modal.css                 Modales y formularios
css/app.css                   Layout de la app
js/config.js                  Client IDs y constantes editables
js/bigquery.js                Cliente REST de BigQuery
js/snowflake.js                Cliente OAuth+PKCE / SQL API v2 de Snowflake
js/provider.js                 Capa de abstracción BigQuery ↔ Snowflake
js/schema.js                    Bootstrap del esquema de control
js/auth.js                       Lógica de login (solo en index.html)
js/ui.js                          Toasts, modales, confirmaciones, maximizar/contraer
js/dimensions.js                   Módulo Dimensiones
js/cubes.js                         Módulo Cubos
js/app.js                            Controlador principal de app.html
sql/00_control_schema.sql             Script SQL de referencia (ambos motores)
sql/01_snowflake_oauth_integration.sql Alta del cliente OAuth en Snowflake
proxy/cloudflare-worker.js             Proxy CORS opcional para Snowflake
```

## 7. Siguientes pasos sugeridos

Dime por cuál seguimos: Cargas de datos, Flujos de carga, Funciones, Flujos
de proceso, Workflows o Roles, y lo construimos con la misma mecánica
(listado por proyecto + alta/edición + tabla de metadatos en
`DRACO_CONTROL`, funcionando sobre ambos motores vía `Provider`).


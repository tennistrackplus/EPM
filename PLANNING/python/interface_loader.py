"""
============================================================
DRACO PLANNING - MOTOR DE CARGA DE INTERFACES
============================================================
Lee la definicion de una interfaz (guardada por la app en las tablas
de control INTERFACES / INTERFACES_VALUES / INTERFACES_INPUT /
INTERFACES_INPUT_FILTERS / INTERFACES_MAPPING) y resuelve, registro a
registro, el mapeo hacia el cubo destino, sustituyendo variables por
su valor real.

Diseno pensado para escalar a varios motores (Snowflake, BigQuery,
Microsoft Fabric...):

  - El NUCLEO (`apply_mapping`, `MappingConfig`) es puro Python +
    pandas: no conoce Snowflake ni BigQuery. Recibe DataFrames ya
    leidos y devuelve un DataFrame ya mapeado.
  - Todo lo que toca al motor (leer tablas de control, leer la tabla
    origen, escribir en el cubo) vive en un "adaptador" con una
    interfaz minima (`EngineAdapter`). Este fichero incluye el
    adaptador de Snowflake/Snowpark; para BigQuery o Fabric basta con
    escribir un adaptador nuevo que cumpla la misma interfaz - el
    resto del codigo (la logica de mapeo) no cambia.

Contrato de las variables (dict `variables` / columna VALOR de
`df_variables`): una variable de pantalla marcada en modo "unico" sigue
llegando como un escalar (str/num/bool). Una variable en modo "rango",
"varios valores" o "cualquiera" llega en cambio como una LISTA de
select-options estilo SAP:

    [{"sign": "I"|"E", "option": "EQ"|"NE"|"GT"|"GE"|"LT"|"LE"|"BT"|"NB"|"CP"|"NP",
      "low": "...", "high": "..."}, ...]

Usa `matches_select_options(valor, filas)` para evaluar un valor de
campo contra esa tabla (ya se usa automaticamente en los filtros de
entrada de una interfaz, y esta disponible tambien dentro del codigo
Python que el usuario escribe en la app - mapeo por FUNCION, por
CODIGO y las transformaciones de entrada/salida).

Uso tipico desde Snowflake (Snowpark Python Stored Procedure / Task):

    from snowflake.snowpark import Session
    from interface_loader import SnowflakeEngine, run_interface

    def main(session: Session, interfaz_id: str,
             tabla_origen: str, tabla_variables: str,
             tabla_destino: str = None) -> str:
        engine = SnowflakeEngine(session, control_database="DRACO",
                                  control_schema="DRACO_CONTROL")
        n = run_interface(engine, interfaz_id, tabla_origen,
                           tabla_variables, tabla_destino)
        return f"{n} filas cargadas"
"""

from __future__ import annotations

import ast
import fnmatch
import json
import operator
from dataclasses import dataclass, field
from typing import Any, Callable, Dict, List, Optional, Protocol

import pandas as pd


# ============================================================
# 1) MODELO DE LA CONFIGURACION DE UNA INTERFAZ
# ============================================================

@dataclass
class InputField:
    campo: str
    tipo: str = "STRING"
    orden: int = 0
    filtro_tipo: Optional[str] = None   # 'VALOR' | 'VARIABLE' | None
    filtro_valor: Optional[str] = None  # constante, o nombre de variable


@dataclass
class OutputMapping:
    campo_destino: str
    tipo: str            # 'CAMPO' | 'CONSTANTE' | 'VARIABLE' | 'FORMULA' | 'FUNCION'
    valor: Optional[str] = None
    codigo: Optional[str] = None


@dataclass
class MappingConfig:
    interfaz_id: str
    proyecto_id: str
    nombre: str
    tipo_origen: str              # 'TABLA' | 'FICHERO'
    origen: str
    cubo_id: str
    mapping_mode: str             # 'VISUAL' | 'CODIGO'
    mapping_code: Optional[str]
    input_transform_code: Optional[str]
    output_transform_code: Optional[str]
    input_fields: List[InputField] = field(default_factory=list)
    output_mappings: List[OutputMapping] = field(default_factory=list)


# ============================================================
# 2) ADAPTADOR DE MOTOR (unico punto que cambia entre motores)
# ============================================================

class EngineAdapter(Protocol):
    """Contrato minimo que cualquier motor (Snowflake, BigQuery, Fabric...)
    debe implementar para poder usar `run_interface`."""

    def qualify_control(self, table: str) -> str:
        """Nombre totalmente cualificado de una tabla del esquema de control."""
        ...

    def run_control_query(self, sql: str) -> List[Dict[str, Any]]:
        """Ejecuta SQL contra el esquema de control y devuelve list[dict]."""
        ...

    def read_table(self, table: str) -> pd.DataFrame:
        """Lee una tabla/vista completa (origen o variables) como DataFrame."""
        ...

    def write_table(self, df: pd.DataFrame, table: str, mode: str = "append") -> None:
        """Escribe el DataFrame resultante en la tabla destino (cubo)."""
        ...

    def esc(self, value: Any) -> str:
        """Escapa un literal para usarlo en SQL (comillas simples)."""
        ...


class SnowflakeEngine:
    """Adaptador de referencia sobre Snowpark. Para BigQuery o Fabric,
    crea una clase con los mismos 4 metodos (mismas firmas) usando el
    cliente/SDK correspondiente (google-cloud-bigquery / pandas-gbq,
    o el conector de Fabric / PySpark)."""

    def __init__(self, session, control_database: str, control_schema: str = "DRACO_CONTROL"):
        self.session = session
        self.control_database = control_database
        self.control_schema = control_schema

    def qualify_control(self, table: str) -> str:
        return f"{self.control_database}.{self.control_schema}.{table}"

    def esc(self, value: Any) -> str:
        if value is None:
            return ""
        return str(value).replace("\\", "\\\\").replace("'", "\\'")

    def run_control_query(self, sql: str) -> List[Dict[str, Any]]:
        return [row.as_dict() for row in self.session.sql(sql).collect()]

    def read_table(self, table: str) -> pd.DataFrame:
        return self.session.table(table).to_pandas()

    def write_table(self, df: pd.DataFrame, table: str, mode: str = "append") -> None:
        sdf = self.session.create_dataframe(df)
        sdf.write.mode(mode).save_as_table(table)


# ============================================================
# 3) CARGA DE LA CONFIGURACION DESDE LAS TABLAS DE CONTROL
# ============================================================

def load_mapping_config(engine: EngineAdapter, interfaz_id: str) -> MappingConfig:
    rows = engine.run_control_query(
        f"SELECT * FROM {engine.qualify_control('INTERFACES')} "
        f"WHERE INTERFAZ_ID = '{engine.esc(interfaz_id)}'"
    )
    if not rows:
        raise ValueError(f"No existe la interfaz {interfaz_id}")
    r = rows[0]

    inputs_raw = engine.run_control_query(
        f"SELECT CAMPO, TIPO, ORDEN FROM {engine.qualify_control('INTERFACES_INPUT')} "
        f"WHERE INTERFAZ_ID = '{engine.esc(interfaz_id)}' ORDER BY ORDEN"
    )
    filters_raw = engine.run_control_query(
        f"SELECT CAMPO, TIPO, VALOR FROM {engine.qualify_control('INTERFACES_INPUT_FILTERS')} "
        f"WHERE INTERFAZ_ID = '{engine.esc(interfaz_id)}'"
    )
    filters_by_campo = {f["CAMPO"]: f for f in filters_raw}

    input_fields = [
        InputField(
            campo=i["CAMPO"],
            tipo=i.get("TIPO") or "STRING",
            orden=i.get("ORDEN") or 0,
            filtro_tipo=(filters_by_campo.get(i["CAMPO"]) or {}).get("TIPO"),
            filtro_valor=(filters_by_campo.get(i["CAMPO"]) or {}).get("VALOR"),
        )
        for i in inputs_raw
    ]

    mapping_raw = engine.run_control_query(
        f"SELECT CAMPO_DESTINO, TIPO, VALOR, CODIGO FROM {engine.qualify_control('INTERFACES_MAPPING')} "
        f"WHERE INTERFAZ_ID = '{engine.esc(interfaz_id)}'"
    )
    output_mappings = [
        OutputMapping(
            campo_destino=m["CAMPO_DESTINO"],
            tipo=m["TIPO"],
            valor=m.get("VALOR"),
            codigo=m.get("CODIGO"),
        )
        for m in mapping_raw
    ]

    return MappingConfig(
        interfaz_id=r["INTERFAZ_ID"],
        proyecto_id=r["PROYECTO_ID"],
        nombre=r["INTERFAZ"],
        tipo_origen=r["TIPO"],
        origen=r.get("ORIGEN"),
        cubo_id=r["CUBO_ID"],
        mapping_mode=r.get("MAPPING_MODE") or "VISUAL",
        mapping_code=r.get("MAPPING_CODE"),
        input_transform_code=r.get("INPUT_TRANSFORM_CODE"),
        output_transform_code=r.get("OUTPUT_TRANSFORM_CODE"),
        input_fields=input_fields,
        output_mappings=output_mappings,
    )


# ============================================================
# 4) NUCLEO DE MAPEO — agnostico de motor (pandas puro)
# ============================================================

# Evaluador de formulas restringido: solo operaciones aritmeticas y de
# texto sobre los campos de entrada, nunca `eval` libre.
_ALLOWED_BINOPS = {
    ast.Add: operator.add, ast.Sub: operator.sub, ast.Mult: operator.mul,
    ast.Div: operator.truediv, ast.Mod: operator.mod, ast.Pow: operator.pow,
}


def _safe_eval_formula(expr: str, variables: Dict[str, Any]) -> Any:
    """Evalua una formula sencilla (p.ej. "cantidad * precio_unit") usando
    solo los nombres presentes en `variables` (campos de la fila + variables
    de la interfaz). No permite llamadas a funciones ni imports."""

    def _eval(node):
        if isinstance(node, ast.Expression):
            return _eval(node.body)
        if isinstance(node, ast.Constant):
            return node.value
        if isinstance(node, ast.Name):
            if node.id not in variables:
                raise ValueError(f"Nombre no permitido en formula: {node.id}")
            return variables[node.id]
        if isinstance(node, ast.BinOp) and type(node.op) in _ALLOWED_BINOPS:
            return _ALLOWED_BINOPS[type(node.op)](_eval(node.left), _eval(node.right))
        if isinstance(node, ast.UnaryOp) and isinstance(node.op, ast.USub):
            return -_eval(node.operand)
        raise ValueError(f"Expresion de formula no soportada: {ast.dump(node)}")

    tree = ast.parse(expr, mode="eval")
    return _eval(tree)


# Subconjunto de builtins permitido en el codigo Python que el usuario
# escribe en la app (funciones de campo, mapeo por codigo, transformaciones).
# Sin acceso a import/open/eval/exec ni al resto de builtins peligrosos.
import builtins as _builtins  # noqa: E402

_SAFE_BUILTIN_NAMES = (
    "round", "len", "str", "int", "float", "bool", "abs", "min", "max",
    "sum", "sorted", "enumerate", "range", "zip", "list", "dict", "tuple",
)
_SAFE_BUILTINS = {name: getattr(_builtins, name) for name in _SAFE_BUILTIN_NAMES}


def _compile_funcion(codigo: str) -> Callable[[Dict[str, Any], Dict[str, Any]], Any]:
    """Compila el codigo Python de un mapeo tipo FUNCION. Debe definir una
    funcion `mapear(fila, variables)` que devuelva el valor del campo."""
    ns: Dict[str, Any] = {"matches_select_options": matches_select_options}
    exec(codigo, {"__builtins__": _SAFE_BUILTINS}, ns)  # noqa: S102 - codigo del usuario en la app
    if "mapear" not in ns:
        raise ValueError("El codigo de la funcion debe definir `def mapear(fila, variables): ...`")
    return ns["mapear"]


# ------------------------------------------------------------
# Select-options estilo SAP: cada variable de pantalla marcada como
# "rango", "varios valores" o "cualquiera" (SELECT_MODE en
# FLUJOS_SCREEN_VARIABLES) ya NO llega como un valor escalar en el dict
# `variables`, sino como una lista de filas:
#
#     [{"sign": "I"|"E", "option": "EQ"|"NE"|"GT"|"GE"|"LT"|"LE"|"BT"|"NB"|"CP"|"NP",
#       "low": "...", "high": "..."}, ...]
#
# ("sign" I = incluir, E = excluir; "option" es el operador, igual que
# en los SELECT-OPTIONS de ABAP: EQ igual, NE distinto, GT/GE/LT/LE
# comparaciones, BT/NB entre / no entre, CP/NP coincide / no coincide
# con un patron admitiendo comodines "*").
#
# `matches_select_options(valor, filas)` evalua un valor de campo contra
# esa tabla y esta disponible tanto para los filtros de entrada (ver
# `_row_passes_filters`) como dentro del codigo Python que el usuario
# escribe en la app (mapeo por FUNCION, por CODIGO, y las
# transformaciones de entrada/salida), donde se inyecta automaticamente
# en el espacio de nombres de ejecucion.
# ------------------------------------------------------------

def _try_num(value: Any) -> Any:
    if value is None or value == "":
        return value
    try:
        return float(value)
    except (TypeError, ValueError):
        return value


def _select_options_row_matches(value: Any, low: Any, high: Any, option: str) -> bool:
    option = (option or "EQ").upper()

    if option == "CP":
        return fnmatch.fnmatch(str(value), str(low))
    if option == "NP":
        return not fnmatch.fnmatch(str(value), str(low))

    v, lo, hi = _try_num(value), _try_num(low), _try_num(high)

    if option == "EQ":
        return v == lo
    if option == "NE":
        return v != lo
    if option == "GT":
        return v is not None and lo is not None and v > lo
    if option == "GE":
        return v is not None and lo is not None and v >= lo
    if option == "LT":
        return v is not None and lo is not None and v < lo
    if option == "LE":
        return v is not None and lo is not None and v <= lo
    if option == "BT":
        return v is not None and lo is not None and hi is not None and lo <= v <= hi
    if option == "NB":
        return not (v is not None and lo is not None and hi is not None and lo <= v <= hi)
    return False


def matches_select_options(value: Any, rows: Optional[List[Dict[str, Any]]]) -> bool:
    """Evalua `value` contra una tabla de select-options estilo SAP.

    Semantica igual que en ABAP: las filas "I" (incluir) se combinan con
    OR entre si (si no hay ninguna fila "I", se considera que incluye
    cualquier valor); las filas "E" (excluir) descartan el valor aunque
    alguna "I" haya encajado. Filas sin "option" se tratan como EQ.
    """
    if not rows:
        return True

    include_rows = [r for r in rows if (r.get("sign") or "I").upper() != "E"]
    exclude_rows = [r for r in rows if (r.get("sign") or "I").upper() == "E"]

    included = True
    if include_rows:
        included = any(
            _select_options_row_matches(value, r.get("low"), r.get("high"), r.get("option"))
            for r in include_rows
        )

    excluded = any(
        _select_options_row_matches(value, r.get("low"), r.get("high"), r.get("option"))
        for r in exclude_rows
    )

    return included and not excluded


def _coerce_objetivo(objetivo: Any) -> Any:
    """Si `objetivo` llega como un string JSON de lista (p.ej. porque el motor
    devuelve un VARIANT/JSON como texto), lo decodifica para poder tratarlo
    como tabla de select-options. Si no es JSON o no es una lista, se
    devuelve tal cual (valor escalar normal)."""
    if isinstance(objetivo, str) and objetivo.strip().startswith("["):
        try:
            parsed = json.loads(objetivo)
        except (ValueError, TypeError):
            return objetivo
        if isinstance(parsed, list):
            return parsed
    return objetivo


def _row_passes_filters(row: Dict[str, Any], config: MappingConfig, variables: Dict[str, Any]) -> bool:
    for f in config.input_fields:
        if not f.filtro_tipo:
            continue
        objetivo = _coerce_objetivo(variables.get(f.filtro_valor) if f.filtro_tipo == "VARIABLE" else f.filtro_valor)
        valor_campo = row.get(f.campo)
        if isinstance(objetivo, list):
            # Variable en modo rango / varios valores / cualquiera (select-options
            # estilo SAP): en vez de un escalar, la variable trae una tabla
            # [{sign, option, low, high}, ...] — ver matches_select_options().
            if not matches_select_options(valor_campo, objetivo):
                return False
        else:
            if str(valor_campo) != str(objetivo):
                return False
    return True


def apply_mapping(df_input: pd.DataFrame, df_variables: pd.DataFrame, config: MappingConfig) -> pd.DataFrame:
    """Nucleo puro: aplica filtros + mapeo campo a campo de una interfaz
    sobre un DataFrame de entrada, sustituyendo variables por su valor.

    df_input:      filas de la tabla/fichero origen (columnas = INTERFACES_INPUT.CAMPO)
    df_variables:  2 columnas, VARIABLE_NOMBRE / VALOR (una fila por variable resuelta)
    config:        definicion de la interfaz (ver load_mapping_config)
    """
    variables = dict(zip(df_variables["VARIABLE_NOMBRE"], df_variables["VALOR"]))

    # Modo "mapeo por codigo": el propio codigo decide todo el DataFrame de salida.
    if config.mapping_mode == "CODIGO" and config.mapping_code:
        ns: Dict[str, Any] = {"pd": pd, "matches_select_options": matches_select_options}
        exec(config.mapping_code, ns)  # noqa: S102 - codigo definido por el propio usuario en la app
        if "mapear" not in ns:
            raise ValueError("El mapeo por codigo debe definir `def mapear(df_input, variables): ...`")
        return ns["mapear"](df_input, variables)

    # Modo visual: mapeo campo a campo, registro a registro.
    output_rows: List[Dict[str, Any]] = []
    for _, row in df_input.iterrows():
        fila = row.to_dict()
        if not _row_passes_filters(fila, config, variables):
            continue

        salida: Dict[str, Any] = {}
        for m in config.output_mappings:
            if m.tipo == "CAMPO":
                salida[m.campo_destino] = fila.get(m.valor)
            elif m.tipo == "CONSTANTE":
                salida[m.campo_destino] = m.valor
            elif m.tipo == "VARIABLE":
                salida[m.campo_destino] = variables.get(m.valor)
            elif m.tipo == "FORMULA":
                salida[m.campo_destino] = _safe_eval_formula(m.valor, {**fila, **variables})
            elif m.tipo == "FUNCION":
                fn = _compile_funcion(m.codigo or "")
                salida[m.campo_destino] = fn(fila, variables)
            else:
                salida[m.campo_destino] = None
        output_rows.append(salida)

    return pd.DataFrame(output_rows)


# ============================================================
# 5) ORQUESTADOR — enlaza motor + nucleo
# ============================================================

def run_interface(
    engine: EngineAdapter,
    interfaz_id: str,
    tabla_origen: str,
    tabla_variables: str,
    tabla_destino: Optional[str] = None,
    write_mode: str = "append",
) -> int:
    """Ejecuta una interfaz de principio a fin:
      1) Lee la definicion (mapeo) desde las tablas de control.
      2) Lee la tabla de origen y la tabla de variables (nombre/valor).
      3) Aplica el mapeo registro a registro, sustituyendo variables.
      4) Inserta el resultado en la tabla destino (cubo).

    tabla_destino es opcional: si no se indica, deberia resolverse
    fuera (por ejemplo consultando CUBOS.TABLA con config.cubo_id) y
    pasarse ya cualificada (proyecto.dataset.tabla / base.esquema.tabla).
    Devuelve el numero de filas insertadas.
    """
    config = load_mapping_config(engine, interfaz_id)

    df_input = engine.read_table(tabla_origen)
    df_variables = engine.read_table(tabla_variables)

    if config.input_transform_code:
        ns: Dict[str, Any] = {"pd": pd, "matches_select_options": matches_select_options}
        exec(config.input_transform_code, ns)  # noqa: S102
        if "transformar" in ns:
            df_input = ns["transformar"](df_input)

    df_output = apply_mapping(df_input, df_variables, config)

    if config.output_transform_code:
        ns = {"pd": pd, "matches_select_options": matches_select_options}
        exec(config.output_transform_code, ns)  # noqa: S102
        if "transformar" in ns:
            df_output = ns["transformar"](df_output)

    destino = tabla_destino or config.cubo_id
    engine.write_table(df_output, destino, mode=write_mode)
    return len(df_output)

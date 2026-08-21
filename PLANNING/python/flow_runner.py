"""
============================================================
DRACO PLANNING - MOTOR DE EJECUCION DE FLUJOS
============================================================
Ejecuta un FLUJO de principio a fin: lee su cadena de interfaces (en
orden, sin bifurcaciones) desde las tablas de control FLUJOS /
FLUJOS_INTERFACES / FLUJOS_INTERFACES_TARGETS y va llamando, paso a
paso, al motor de `interface_loader.py`, sustituyendo en cada paso las
variables de pantalla (parametro de entrada de este modulo) por su
valor real.

Para los pasos de tipo FICHERO, antes de mapear se descarga el fichero
del storage a local con `storage_io.py` y se lee a un DataFrame; para
los de tipo TABLA se lee directamente la tabla de origen.

Cada paso queda registrado en FLUJOS_RUNS / FLUJOS_RUN_STEPS (estado,
filas, error, timestamps) para que `flow_run.html` pueda pintar el
monitor en tiempo real sondeando esas tablas.

Uso tipico desde Snowflake (Snowpark Python Stored Procedure), a
llamar desde el boton "Ejecutar" de flow_run.html via `CALL`:

    from snowflake.snowpark import Session
    from flow_runner import main

    # firma pensada para un stored procedure:
    #   CALL DRACO_CONTROL.SP_RUN_FLUJO(:flujo_id, :variables_json, :run_id)
    def main_proc(session: Session, flujo_id: str, variables_json: str, run_id: str = None) -> str:
        return main(session, flujo_id, variables_json, run_id)

`variables_json` es un objeto JSON plano {"nombre_variable": "valor", ...}
con TODAS las variables de la pantalla del flujo (incluidas las de tipo
FILE, cuyo valor debe ser ya la ruta de storage tras la subida hecha
por el navegador - ver js/storage.js).
"""

from __future__ import annotations

import json
import traceback
import uuid
from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import Any, Callable, Dict, List, Optional

import pandas as pd

from interface_loader import EngineAdapter, SnowflakeEngine, apply_mapping, load_mapping_config, matches_select_options
from storage_io import LocalFsStorage, SnowflakeStageStorage, StorageAdapter, load_interface_values, read_file_to_dataframe

ESTADO_PENDIENTE = "PENDIENTE"
ESTADO_EN_CURSO = "EN_CURSO"
ESTADO_OK = "OK"
ESTADO_ERROR = "ERROR"


# ============================================================
# 1) MODELO DE LA DEFINICION DE UN FLUJO
# ============================================================

@dataclass
class FlowTarget:
    tipo: str   # 'CONSTANTE' | 'VARIABLE'
    valor: str


@dataclass
class FlowStep:
    paso_id: str
    orden: int
    interfaz_id: str
    # targets['file' | 'filter' | 'mapping'][clave] = FlowTarget
    targets: Dict[str, Dict[str, FlowTarget]] = field(default_factory=lambda: {"file": {}, "filter": {}, "mapping": {}})


@dataclass
class FlowDefinition:
    flujo_id: str
    proyecto_id: str
    nombre: str
    tipo: str  # 'AUTOMATICO' | 'MANUAL'
    steps: List[FlowStep] = field(default_factory=list)


def load_flow_definition(engine: EngineAdapter, flujo_id: str) -> FlowDefinition:
    rows = engine.run_control_query(
        f"SELECT * FROM {engine.qualify_control('FLUJOS')} WHERE FLUJO_ID = '{engine.esc(flujo_id)}'"
    )
    if not rows:
        raise ValueError(f"No existe el flujo {flujo_id}")
    r = rows[0]

    chain_raw = engine.run_control_query(
        f"SELECT PASO_ID, INTERFAZ_ID, ORDEN FROM {engine.qualify_control('FLUJOS_INTERFACES')} "
        f"WHERE FLUJO_ID = '{engine.esc(flujo_id)}' ORDER BY ORDEN"
    )
    targets_raw = engine.run_control_query(
        f"SELECT PASO_ID, GRUPO, CLAVE, TIPO, VALOR FROM {engine.qualify_control('FLUJOS_INTERFACES_TARGETS')} "
        f"WHERE FLUJO_ID = '{engine.esc(flujo_id)}'"
    )
    targets_by_paso: Dict[str, Dict[str, Dict[str, FlowTarget]]] = {}
    for t in targets_raw:
        grupo = (t["GRUPO"] or "").lower()
        bucket = targets_by_paso.setdefault(t["PASO_ID"], {"file": {}, "filter": {}, "mapping": {}})
        bucket.setdefault(grupo, {})[t["CLAVE"]] = FlowTarget(tipo=t["TIPO"], valor=t.get("VALOR"))

    steps = [
        FlowStep(
            paso_id=c["PASO_ID"],
            orden=c.get("ORDEN") or 0,
            interfaz_id=c["INTERFAZ_ID"],
            targets=targets_by_paso.get(c["PASO_ID"], {"file": {}, "filter": {}, "mapping": {}}),
        )
        for c in chain_raw
    ]

    return FlowDefinition(
        flujo_id=r["FLUJO_ID"],
        proyecto_id=r["PROYECTO_ID"],
        nombre=r["FLUJO"],
        tipo=r["TIPO"],
        steps=steps,
    )


def _resolve(target: Optional[FlowTarget], variables: Dict[str, Any]) -> Any:
    if target is None:
        return None
    if target.tipo == "VARIABLE":
        return variables.get(target.valor)
    return target.valor


def _cubo_tabla(engine: EngineAdapter, cubo_id: str) -> str:
    rows = engine.run_control_query(
        f"SELECT TABLA FROM {engine.qualify_control('CUBOS')} WHERE CUBO_ID = '{engine.esc(cubo_id)}'"
    )
    if not rows:
        raise ValueError(f"No existe el cubo destino {cubo_id}")
    return rows[0]["TABLA"]


# ============================================================
# 2) PERSISTENCIA DE LA EJECUCION (para el monitor)
# ============================================================

def _now() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")


def _start_run(engine: EngineAdapter, run_id: str, flow: FlowDefinition, variables: Dict[str, Any]) -> None:
    variables_json = engine.esc(json.dumps(variables, default=str))
    engine.run_control_query(
        f"INSERT INTO {engine.qualify_control('FLUJOS_RUNS')} "
        f"(RUN_ID, PROYECTO_ID, FLUJO_ID, ESTADO, VARIABLES_JSON, MENSAJE, USUARIO, FECHA_INICIO, FECHA_FIN) "
        f"VALUES ('{engine.esc(run_id)}', '{engine.esc(flow.proyecto_id)}', '{engine.esc(flow.flujo_id)}', "
        f"'{ESTADO_EN_CURSO}', '{variables_json}', '', CURRENT_USER(), CURRENT_TIMESTAMP(), NULL)"
    )
    if flow.steps:
        vals = ",\n".join(
            f"('{engine.esc(run_id)}', '{engine.esc(s.paso_id)}', {s.orden}, '{engine.esc(s.interfaz_id)}', "
            f"'{ESTADO_PENDIENTE}', NULL, '', NULL, NULL)"
            for s in flow.steps
        )
        engine.run_control_query(
            f"INSERT INTO {engine.qualify_control('FLUJOS_RUN_STEPS')} "
            f"(RUN_ID, PASO_ID, ORDEN, INTERFAZ_ID, ESTADO, FILAS, MENSAJE, FECHA_INICIO, FECHA_FIN) VALUES {vals}"
        )


def _mark_step(engine: EngineAdapter, run_id: str, paso_id: str, estado: str,
               filas: Optional[int] = None, mensaje: str = "") -> None:
    sets = [f"ESTADO = '{estado}'", f"MENSAJE = '{engine.esc(mensaje)}'"]
    if filas is not None:
        sets.append(f"FILAS = {int(filas)}")
    if estado == ESTADO_EN_CURSO:
        sets.append("FECHA_INICIO = CURRENT_TIMESTAMP()")
    if estado in (ESTADO_OK, ESTADO_ERROR):
        sets.append("FECHA_FIN = CURRENT_TIMESTAMP()")
    engine.run_control_query(
        f"UPDATE {engine.qualify_control('FLUJOS_RUN_STEPS')} SET {', '.join(sets)} "
        f"WHERE RUN_ID = '{engine.esc(run_id)}' AND PASO_ID = '{engine.esc(paso_id)}'"
    )


def _finish_run(engine: EngineAdapter, run_id: str, estado: str, mensaje: str = "") -> None:
    engine.run_control_query(
        f"UPDATE {engine.qualify_control('FLUJOS_RUNS')} "
        f"SET ESTADO = '{estado}', MENSAJE = '{engine.esc(mensaje)}', FECHA_FIN = CURRENT_TIMESTAMP() "
        f"WHERE RUN_ID = '{engine.esc(run_id)}'"
    )


# ============================================================
# 3) EJECUCION DE UN PASO (fichero -> storage; tabla -> lectura directa)
# ============================================================

def _run_step(engine: EngineAdapter, storage: Optional[StorageAdapter], step: FlowStep,
              variables: Dict[str, Any]) -> int:
    config = load_mapping_config(engine, step.interfaz_id)

    if config.tipo_origen == "FICHERO":
        if storage is None:
            raise ValueError("Esta interfaz es de tipo FICHERO pero no se ha pasado un StorageAdapter a run_flow().")
        ruta_local = _resolve(step.targets.get("file", {}).get("ruta_local"), variables)
        ruta_storage = _resolve(step.targets.get("file", {}).get("ruta_storage"), variables)
        if not ruta_storage:
            raise ValueError("Paso de fichero sin 'ruta_storage' asignada (ni por constante ni por variable).")
        local_path = storage.download(ruta_storage, ruta_local or None)
        values = load_interface_values(engine, step.interfaz_id)
        df_input = read_file_to_dataframe(local_path, config.origen, values)
    else:
        df_input = engine.read_table(config.origen)

    # Variables de filtro/mapeo asignadas en el paso: la CLAVE es el nombre
    # de variable que espera la interfaz (filtro VARIABLE / mapeo VARIABLE);
    # se resuelve contra las variables de pantalla del flujo y se monta como
    # tabla de variables (VARIABLE_NOMBRE/VALOR) para apply_mapping().
    step_vars: Dict[str, Any] = dict(variables)
    for grupo in ("filter", "mapping"):
        for clave, target in (step.targets.get(grupo) or {}).items():
            step_vars[clave] = _resolve(target, variables)

    df_variables = pd.DataFrame({
        "VARIABLE_NOMBRE": list(step_vars.keys()),
        "VALOR": list(step_vars.values()),
    })

    if config.input_transform_code:
        ns: Dict[str, Any] = {"pd": pd, "matches_select_options": matches_select_options}
        exec(config.input_transform_code, ns)  # noqa: S102 - codigo definido por el usuario en la app
        if "transformar" in ns:
            df_input = ns["transformar"](df_input)

    df_output = apply_mapping(df_input, df_variables, config)

    if config.output_transform_code:
        ns = {"pd": pd, "matches_select_options": matches_select_options}
        exec(config.output_transform_code, ns)  # noqa: S102
        if "transformar" in ns:
            df_output = ns["transformar"](df_output)

    tabla_destino = _cubo_tabla(engine, config.cubo_id)
    engine.write_table(df_output, tabla_destino, mode="append")
    return len(df_output)


# ============================================================
# 4) ORQUESTADOR PRINCIPAL
# ============================================================

StepCallback = Callable[[FlowStep, str, Optional[int], str], None]


def run_flow(
    engine: EngineAdapter,
    flujo_id: str,
    variables: Dict[str, Any],
    storage: Optional[StorageAdapter] = None,
    run_id: Optional[str] = None,
    on_step: Optional[StepCallback] = None,
) -> Dict[str, Any]:
    """Ejecuta un flujo de principio a fin.

    engine:     adaptador de motor (ver interface_loader.EngineAdapter).
    flujo_id:   FLUJOS.FLUJO_ID a ejecutar.
    variables:  variables de pantalla del flujo, {"nombre": valor, ...}.
                Para variables de tipo FILE, el valor debe ser ya la ruta
                de storage (el navegador sube el fichero ANTES de llamar
                a esta funcion - ver js/storage.js).
    storage:    StorageAdapter para descargar los ficheros de los pasos
                de tipo FICHERO. Obligatorio si el flujo tiene algun paso
                de ese tipo.
    run_id:     identificador de esta ejecucion; se genera uno si no se indica.
    on_step:    callback opcional(step, estado, filas, mensaje) para feedback
                en vivo ademas de la persistencia en FLUJOS_RUN_STEPS.

    Devuelve {"run_id", "estado", "filas_totales", "pasos": [...]}. El
    estado y las filas de cada paso quedan tambien en FLUJOS_RUNS /
    FLUJOS_RUN_STEPS para que flow_run.html pueda monitorizar la
    ejecucion sondeando esas tablas.
    """
    run_id = run_id or str(uuid.uuid4())
    flow = load_flow_definition(engine, flujo_id)
    _start_run(engine, run_id, flow, variables)

    total_rows = 0
    pasos_resultado: List[Dict[str, Any]] = []

    for step in flow.steps:
        _mark_step(engine, run_id, step.paso_id, ESTADO_EN_CURSO)
        if on_step:
            on_step(step, ESTADO_EN_CURSO, None, "")
        try:
            n = _run_step(engine, storage, step, variables)
            total_rows += n
            _mark_step(engine, run_id, step.paso_id, ESTADO_OK, filas=n)
            if on_step:
                on_step(step, ESTADO_OK, n, "")
            pasos_resultado.append({"paso_id": step.paso_id, "estado": ESTADO_OK, "filas": n})
        except Exception as exc:  # noqa: BLE001 - se registra y se relanza el resumen
            mensaje = f"{exc}\n{traceback.format_exc(limit=3)}"
            _mark_step(engine, run_id, step.paso_id, ESTADO_ERROR, mensaje=str(exc))
            if on_step:
                on_step(step, ESTADO_ERROR, None, str(exc))
            pasos_resultado.append({"paso_id": step.paso_id, "estado": ESTADO_ERROR, "mensaje": str(exc)})
            _finish_run(engine, run_id, ESTADO_ERROR, mensaje=f"Fallo en el paso {step.orden + 1} ({step.interfaz_id}): {exc}")
            return {"run_id": run_id, "estado": ESTADO_ERROR, "filas_totales": total_rows, "pasos": pasos_resultado, "error": mensaje}

    _finish_run(engine, run_id, ESTADO_OK)
    return {"run_id": run_id, "estado": ESTADO_OK, "filas_totales": total_rows, "pasos": pasos_resultado}


# ============================================================
# 5) ENTRADA COMO STORED PROCEDURE (Snowpark) Y COMO SCRIPT
# ============================================================

def main(session, flujo_id: str, variables_json: str, run_id: Optional[str] = None,
         control_database: str = "DRACO", control_schema: str = "DRACO_CONTROL",
         stage_name: str = "@DRACO_LANDING") -> str:
    """Punto de entrada pensado para desplegarse como Snowflake Python
    Stored Procedure y ser llamado con `CALL` desde flow_run.html:

        CREATE OR REPLACE PROCEDURE DRACO_CONTROL.SP_RUN_FLUJO(
            FLUJO_ID STRING, VARIABLES_JSON STRING, RUN_ID STRING)
        RETURNS STRING
        LANGUAGE PYTHON
        RUNTIME_VERSION = '3.11'
        PACKAGES = ('snowflake-snowpark-python', 'pandas')
        IMPORTS = ('@DRACO_CONTROL.LIBS/interface_loader.py',
                   '@DRACO_CONTROL.LIBS/storage_io.py',
                   '@DRACO_CONTROL.LIBS/flow_runner.py')
        HANDLER = 'flow_runner.main';

    Como el procedure de arriba solo declara 3 argumentos, `stage_name`
    se queda en su valor por defecto ("@DRACO_LANDING"): debe ser el
    MISMO stage que DracoConfig.snowflakeUploadStage (js/config.js) y
    que STAGE_NAME al llamar a SP_FINALIZE_FILE_UPLOAD (ver
    sql/02_snowflake_file_upload.sql), o los flujos no encontrarán los
    ficheros que suba el navegador. Si usas otro stage, cambia el valor
    por defecto de `stage_name` aquí arriba (o añade STAGE_NAME como
    4º argumento del procedure) y el de snowflakeUploadStage a la vez.
    """
    engine = SnowflakeEngine(session, control_database=control_database, control_schema=control_schema)
    storage = SnowflakeStageStorage(session, stage_name=stage_name)
    variables = json.loads(variables_json) if variables_json else {}
    result = run_flow(engine, flujo_id, variables, storage=storage, run_id=run_id)
    return json.dumps(result, default=str)


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Ejecuta un flujo de carga de Draco Planning fuera de Snowflake (motor local de pruebas).")
    parser.add_argument("--flujo-id", required=True)
    parser.add_argument("--variables", default="{}", help="JSON plano {nombre: valor} con las variables de pantalla")
    parser.add_argument("--storage-dir", default=".", help="Carpeta que hace de 'storage' (LocalFsStorage)")
    parser.add_argument("--run-id", default=None)
    args = parser.parse_args()

    print(
        "Este bloque __main__ es solo de referencia: usa LocalFsStorage y necesita "
        "que le pases tu propio EngineAdapter (ej. SnowflakeEngine con una sesion real). "
        "En produccion, despliega flow_runner.main() como Stored Procedure (ver docstring)."
    )

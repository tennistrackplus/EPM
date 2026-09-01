"""
============================================================
DRACO PLANNING - ENDPOINT HTTP PARA BIGQUERY (Cloud Run / Cloud Functions)
============================================================
Equivalente, para BigQuery, al stored procedure DRACO_CONTROL.SP_RUN_FLUJO
de Snowflake: envuelve `flow_runner.run_flow()` detras de un endpoint
HTTP para que js/flow-run.js (DracoConfig.flowRunnerHttpEndpoint) pueda
lanzar flujos.

Contrato HTTP (coincide con js/flow-run.js::callOrchestrator):

    POST <flowRunnerHttpEndpoint>
    Content-Type: application/json
    { "flujo_id": "...", "variables": { ... }, "run_id": "..." }

    -> 200/500 (según si el flujo termina OK o ERROR)
       { "run_id": "...", "estado": "OK"|"ERROR", "filas_totales": N,
         "pasos": [ ... ] }

Despliegue - ver el paso a paso completo en README.md, seccion
"Interfaces de fichero en BigQuery": en resumen, sube esta carpeta
`python/` completa (interface_loader.py, flow_runner.py, storage_io.py,
cloud_run_main.py, requirements.txt) como:

  - Cloud Run (recomendado, sin limite de tiempo de 9 min de Functions):
        gcloud run deploy draco-flow-runner \\
            --source . \\
            --region europe-west1 \\
            --allow-unauthenticated \\
            --set-env-vars GCP_PROJECT_ID=<tu-proyecto>,CONTROL_DATASET=DRACO_CONTROL,UPLOAD_BUCKET=<tu-bucket>

  - o Cloud Functions Gen 2 (entry point `main`, runtime Python 3.11).

Variables de entorno:
  GCP_PROJECT_ID   Project ID de BigQuery. Si se omite, se usa el
                   proyecto por defecto del entorno de ejecucion (el que
                   ya trae Cloud Run/Functions).
  CONTROL_DATASET  Dataset de control (por defecto DRACO_CONTROL; debe
                   coincidir con DracoConfig.controlDataset en js/config.js).
  UPLOAD_BUCKET    Bucket de GCS donde sube los ficheros el navegador
                   (debe coincidir con DracoConfig.bigqueryUploadBucket).
                   Solo hace falta si vas a tener interfaces de tipo
                   FICHERO en tus flujos.
  DRACO_SHARED_KEY Opcional. Si se define, el endpoint exige la cabecera
                   `X-Draco-Key` con este mismo valor (proteccion minima
                   adicional a la de IAM/--allow-unauthenticated). Deja
                   vacio para no exigir nada.
  ALLOWED_ORIGIN   Opcional. Valor de Access-Control-Allow-Origin para las
                   respuestas (por defecto "*"). Ponlo al dominio exacto
                   donde sirves PLANNING/ en produccion.
"""

from __future__ import annotations

import json
import os
from typing import Any, Dict, Tuple

import functions_framework
from google.cloud import bigquery, storage

from flow_runner import run_flow
from interface_loader import BigQueryEngine
from storage_io import GCSStorage


def _cors_headers() -> Dict[str, str]:
    origin = os.environ.get("ALLOWED_ORIGIN", "*")
    return {
        "Access-Control-Allow-Origin": origin,
        "Access-Control-Allow-Methods": "POST, OPTIONS",
        "Access-Control-Allow-Headers": "Content-Type, X-Draco-Key",
        "Access-Control-Max-Age": "3600",
    }


def _json_response(payload: Dict[str, Any], status: int) -> Tuple[str, int, Dict[str, str]]:
    headers = {"Content-Type": "application/json"}
    headers.update(_cors_headers())
    return json.dumps(payload, default=str), status, headers


def _build_engine_and_storage():
    project_id = os.environ.get("GCP_PROJECT_ID") or None
    control_dataset = os.environ.get("CONTROL_DATASET", "DRACO_CONTROL")
    bucket_name = os.environ.get("UPLOAD_BUCKET")

    bq_client = bigquery.Client(project=project_id)
    engine = BigQueryEngine(bq_client, control_dataset=control_dataset)

    storage_adapter = None
    if bucket_name:
        gcs_client = storage.Client(project=project_id)
        storage_adapter = GCSStorage(gcs_client, bucket_name)

    return engine, storage_adapter


@functions_framework.http
def main(request):
    """Punto de entrada HTTP (Cloud Run / Cloud Functions Gen2)."""
    if request.method == "OPTIONS":
        # Preflight CORS del navegador.
        return ("", 204, _cors_headers())

    if request.method != "POST":
        return _json_response({"error": "Solo se acepta POST."}, 405)

    shared_key = os.environ.get("DRACO_SHARED_KEY")
    if shared_key and request.headers.get("X-Draco-Key") != shared_key:
        return _json_response({"error": "No autorizado."}, 401)

    body = request.get_json(silent=True) or {}
    flujo_id = body.get("flujo_id")
    variables = body.get("variables") or {}
    run_id = body.get("run_id")

    if not flujo_id:
        return _json_response({"error": "Falta 'flujo_id' en el cuerpo de la peticion."}, 400)

    try:
        engine, storage_adapter = _build_engine_and_storage()
        result = run_flow(engine, flujo_id, variables, storage=storage_adapter, run_id=run_id)
    except Exception as exc:  # noqa: BLE001 - se devuelve como error HTTP, no se relanza
        return _json_response({"error": str(exc)}, 500)

    status = 200 if result.get("estado") != "ERROR" else 500
    return _json_response(result, status)


if __name__ == "__main__":
    # Servidor local de pruebas: `python cloud_run_main.py` levanta un
    # Flask de desarrollo en localhost:8080 con el mismo handler `main`.
    # En Cloud Run/Functions NO se usa este bloque (lo invoca
    # functions-framework / el runtime directamente).
    from flask import Flask
    from flask import request as flask_request

    app = Flask(__name__)
    app.add_url_rule("/", view_func=lambda: main(flask_request), methods=["POST", "OPTIONS"])
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8080)))

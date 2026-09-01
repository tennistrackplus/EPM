"""
============================================================
DRACO PLANNING - IO DE FICHEROS PARA EL MOTOR DE FLUJOS
============================================================
Complementa a `interface_loader.py` para el caso `TIPO = 'FICHERO'`:

  - `StorageAdapter` es el contrato minimo (2 metodos) que cualquier
    backend de storage debe cumplir para poder usarse desde
    `flow_runner.py`. Igual que `EngineAdapter` en interface_loader.py,
    es el UNICO punto que cambia entre instalaciones.
  - Se incluyen dos adaptadores de referencia:
      * `LocalFsStorage`     - "storage" = otra carpeta del mismo
                                filesystem (disco compartido / NAS).
                                Util tambien para desarrollo y pruebas.
      * `SnowflakeStageStorage` - stage interno de Snowflake, pensado
                                para usarse DENTRO de un Stored
                                Procedure (Snowpark): `session.file.get`
                                / `session.file.put`.
    Para GCS, S3 o Azure Blob, crea una clase con los mismos 2 metodos
    usando el SDK correspondiente (google-cloud-storage / boto3 /
    azure-storage-blob) - el resto de `flow_runner.py` no cambia.
  - `read_file_to_dataframe` interpreta el fichero ya descargado en
    local (csv/xlsx/json) usando las caracteristicas guardadas en
    INTERFACES_VALUES (separadores, codificacion...).
"""

from __future__ import annotations

import os
import shutil
from typing import Any, Dict, Optional, Protocol

import pandas as pd


# ============================================================
# 1) CONTRATO DE STORAGE (unico punto que cambia entre instalaciones)
# ============================================================

class StorageAdapter(Protocol):
    """Contrato minimo que cualquier backend de storage (Snowflake stage,
    GCS, S3, disco compartido...) debe implementar."""

    def download(self, storage_path: str, local_path: Optional[str] = None) -> str:
        """Descarga `storage_path` a `local_path` (o a una ruta por defecto
        si no se indica) y devuelve la ruta local resultante."""
        ...

    def upload(self, local_path: str, storage_path: str) -> None:
        """Sube el fichero local `local_path` a `storage_path`. Lo usa
        normalmente el lado navegador (js/storage.js) antes de lanzar el
        flujo, pero se deja aqui tambien por si el propio Python necesita
        subir un resultado (ej. un fichero de errores)."""
        ...


class LocalFsStorage:
    """Adaptador de referencia: 'storage' es otra carpeta (puede ser un
    disco/NAS compartido, o simplemente un directorio de pruebas). Sirve
    tambien como implementacion por defecto para desarrollo local."""

    def __init__(self, base_dir: str = ""):
        self.base_dir = base_dir

    def _resolve(self, path: str) -> str:
        return os.path.join(self.base_dir, path) if self.base_dir and not os.path.isabs(path) else path

    def download(self, storage_path: str, local_path: Optional[str] = None) -> str:
        src = self._resolve(storage_path)
        if not os.path.exists(src):
            raise FileNotFoundError(f"No existe el fichero en storage: {src}")
        dst = local_path or src
        if os.path.abspath(dst) != os.path.abspath(src):
            os.makedirs(os.path.dirname(dst) or ".", exist_ok=True)
            shutil.copyfile(src, dst)
        return dst

    def upload(self, local_path: str, storage_path: str) -> None:
        dst = self._resolve(storage_path)
        os.makedirs(os.path.dirname(dst) or ".", exist_ok=True)
        shutil.copyfile(local_path, dst)


class SnowflakeStageStorage:
    """Adaptador sobre un stage interno de Snowflake (ej. `@DRACO_LANDING`).
    Pensado para ejecutarse DENTRO de un Stored Procedure Snowpark: el
    sandbox del proc tiene su propio filesystem temporal (`/tmp` por
    defecto), al que `session.file.get` descarga el fichero antes de
    leerlo con pandas.

    `storage_path` (el valor que guarda la variable de pantalla de tipo
    FILE, ver js/storage.js::buildPath) es SIEMPRE una ruta relativa
    dentro del stage (ej. "flows/<flujo_id>/<var>/<fichero>"), nunca
    incluye el nombre del stage: este adaptador es quien antepone
    `stage_name` en cada operación. Debe coincidir con el stage al que
    sube el navegador (DracoConfig.snowflakeUploadStage en js/config.js
    y STAGE_NAME en sql/02_snowflake_file_upload.sql — por defecto
    `@DRACO_LANDING` en ambos sitios)."""

    def __init__(self, session, stage_name: str = "@DRACO_LANDING", default_local_dir: str = "/tmp"):
        self.session = session
        self.stage_name = stage_name if stage_name.startswith("@") else f"@{stage_name}"
        self.default_local_dir = default_local_dir

    def _stage_path(self, storage_path: str) -> str:
        return f"{self.stage_name.rstrip('/')}/{storage_path.lstrip('/')}"

    def download(self, storage_path: str, local_path: Optional[str] = None) -> str:
        stage_path = self._stage_path(storage_path)
        target_dir = os.path.dirname(local_path) if local_path else self.default_local_dir
        os.makedirs(target_dir or ".", exist_ok=True)
        results = self.session.file.get(stage_path, target_dir)
        if not results:
            raise FileNotFoundError(f"No se pudo descargar del storage: {stage_path}")
        downloaded = os.path.join(target_dir, os.path.basename(results[0].file))
        if local_path and os.path.abspath(downloaded) != os.path.abspath(local_path):
            os.replace(downloaded, local_path)
            return local_path
        return downloaded

    def upload(self, local_path: str, storage_path: str) -> None:
        stage_path = self._stage_path(storage_path)
        stage_dir = stage_path.rsplit("/", 1)[0] if "/" in stage_path else stage_path
        self.session.file.put(local_path, stage_dir, auto_compress=False, overwrite=True)


try:  # dependencia opcional: solo hace falta si usas GCSStorage
    from google.cloud import storage as gcs
except ImportError:  # pragma: no cover
    gcs = None


class GCSStorage:
    """Adaptador de storage sobre un bucket de Google Cloud Storage.
    Hermano de `SnowflakeStageStorage`: implementa el mismo contrato
    `StorageAdapter` (download/upload) para que `flow_runner.py` funcione
    SIN CAMBIOS cuando el motor es BigQuery. Pensado para instanciarse
    DENTRO de la Cloud Function / Cloud Run que expone
    `DracoConfig.flowRunnerHttpEndpoint` (ver python/cloud_run_main.py).

    `storage_path` (el valor de la variable de pantalla tipo FILE) es
    SIEMPRE una ruta relativa DENTRO del bucket (ej.
    "flujo123__fichero__1700000000000_datos.csv" - ver
    js/storage.js::buildPath), nunca incluye el nombre del bucket: este
    adaptador es quien lo antepone. Debe ser el MISMO bucket al que sube
    el navegador (DracoConfig.bigqueryUploadBucket en js/config.js).
    """

    def __init__(self, client: "gcs.Client", bucket_name: str, default_local_dir: str = "/tmp"):
        if gcs is None:  # pragma: no cover
            raise ImportError(
                "Falta el paquete 'google-cloud-storage'. Instalalo con "
                "`pip install google-cloud-storage` (ya viene incluido en "
                "python/requirements.txt de la Cloud Function)."
            )
        self.bucket = client.bucket(bucket_name)
        self.default_local_dir = default_local_dir

    def download(self, storage_path: str, local_path: Optional[str] = None) -> str:
        blob = self.bucket.blob(storage_path.lstrip("/"))
        if not blob.exists():
            raise FileNotFoundError(f"No existe el fichero en el bucket: {storage_path}")
        dst = local_path or os.path.join(self.default_local_dir, os.path.basename(storage_path))
        os.makedirs(os.path.dirname(dst) or ".", exist_ok=True)
        blob.download_to_filename(dst)
        return dst

    def upload(self, local_path: str, storage_path: str) -> None:
        blob = self.bucket.blob(storage_path.lstrip("/"))
        blob.upload_from_filename(local_path)


# ============================================================
# 2) CARACTERISTICAS DEL ORIGEN (INTERFACES_VALUES) Y LECTURA DEL FICHERO
# ============================================================

def load_interface_values(engine, interfaz_id: str) -> Dict[str, str]:
    """Lee INTERFACES_VALUES (separadores, codificacion...) de una interfaz."""
    rows = engine.run_control_query(
        f"SELECT CARACTERISTICA, VALOR FROM {engine.qualify_control('INTERFACES_VALUES')} "
        f"WHERE INTERFAZ_ID = '{engine.esc(interfaz_id)}'"
    )
    return {r["CARACTERISTICA"]: r["VALOR"] for r in rows}


def read_file_to_dataframe(local_path: str, tipo_origen: str, values: Dict[str, str]) -> pd.DataFrame:
    """Lee un fichero ya descargado en local a un DataFrame, respetando las
    caracteristicas definidas en la interfaz (INTERFACES_VALUES).

    tipo_origen: 'csv' | 'txt' | 'xlsx' | 'xls' | 'json' (INTERFACES.ORIGEN
    cuando INTERFACES.TIPO = 'FICHERO').
    """
    tipo = (tipo_origen or "csv").lower()
    encoding = values.get("CODIFICACION") or "UTF-8"
    sep = values.get("SEPARADOR_CAMPO") or ","
    decimal = values.get("SEPARADOR_DECIMAL") or "."
    miles = values.get("SEPARADOR_MILES") or None

    if tipo in ("csv", "txt"):
        return pd.read_csv(local_path, sep=sep, decimal=decimal, thousands=miles, encoding=encoding)
    if tipo in ("xlsx", "xls"):
        return pd.read_excel(local_path)
    if tipo == "json":
        return pd.read_json(local_path, encoding=encoding)

    raise ValueError(
        f"Tipo de fichero de origen no soportado: '{tipo_origen}'. "
        "Anade el caso correspondiente en read_file_to_dataframe()."
    )

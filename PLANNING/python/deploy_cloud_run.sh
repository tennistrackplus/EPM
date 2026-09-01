#!/usr/bin/env bash
# ============================================================
# DRACO PLANNING - Despliegue automático del Cloud Run (BigQuery)
# ============================================================
# Hace en un solo paso lo que en README.md, sección "7ter" (Paso 2), se
# explica manualmente: habilita las APIs necesarias, despliega este
# directorio como servicio Cloud Run y da los permisos IAM que necesita
# su cuenta de servicio (BigQuery Data Editor, BigQuery Job User, y
# Storage Object Viewer sobre el bucket de subida).
#
# Uso (desde Cloud Shell, DENTRO de la carpeta PLANNING/python):
#
#   chmod +x deploy_cloud_run.sh
#   ./deploy_cloud_run.sh <nombre-del-bucket> [region]
#
# Ejemplo:
#   ./deploy_cloud_run.sh draco-landing europe-west1
#
# Si omites la región, usa europe-west1 por defecto. El Project ID y el
# dataset de control se detectan solos (gcloud config / DRACO_CONTROL).
# Al final imprime la URL que hay que copiar en js/config.js
# (flowRunnerHttpEndpoint).
# ============================================================
set -euo pipefail

BUCKET="${1:?Uso: ./deploy_cloud_run.sh <nombre-del-bucket> [region]}"
REGION="${2:-europe-west1}"
CONTROL_DATASET="${CONTROL_DATASET:-DRACO_CONTROL}"
SERVICE_NAME="draco-flow-runner"

if [[ ! -f "flow_runner.py" || ! -f "cloud_run_main.py" ]]; then
    echo "❌ Ejecuta este script DESDE la carpeta PLANNING/python (no encuentro flow_runner.py / cloud_run_main.py aquí)."
    exit 1
fi

# El builder de Google Cloud (Buildpacks) no sabe arrancar el servicio si
# no existe main.py/app.py: le decimos explícitamente qué comando lanzar
# (functions-framework apuntando a cloud_run_main.py::main). Se genera
# solo si no existe ya (por si el usuario quiere personalizarlo).
if [[ ! -f "Procfile" ]]; then
    echo "web: functions-framework --target=main --source=cloud_run_main.py --port=\$PORT" > Procfile
fi

PROJECT_ID="$(gcloud config get-value project 2>/dev/null)"
if [[ -z "$PROJECT_ID" ]]; then
    echo "❌ No hay ningún proyecto activo en gcloud. Ejecuta 'gcloud config set project <TU_PROJECT_ID>' primero."
    exit 1
fi

echo "▶ Proyecto:        $PROJECT_ID"
echo "▶ Región:          $REGION"
echo "▶ Bucket:          $BUCKET"
echo "▶ Dataset control: $CONTROL_DATASET"
echo

echo "1/4 · Habilitando APIs necesarias (puede tardar un minuto la primera vez)..."
gcloud services enable run.googleapis.com bigquery.googleapis.com storage.googleapis.com cloudbuild.googleapis.com --project "$PROJECT_ID"

echo
echo "2/4 · Desplegando el servicio Cloud Run '$SERVICE_NAME' (2-4 minutos)..."
gcloud run deploy "$SERVICE_NAME" \
    --source . \
    --project "$PROJECT_ID" \
    --region "$REGION" \
    --allow-unauthenticated \
    --set-build-env-vars GOOGLE_PYTHON_VERSION=3.13 \
    --set-env-vars "GCP_PROJECT_ID=${PROJECT_ID},CONTROL_DATASET=${CONTROL_DATASET},UPLOAD_BUCKET=${BUCKET}"

SERVICE_URL="$(gcloud run services describe "$SERVICE_NAME" --project "$PROJECT_ID" --region "$REGION" --format='value(status.url)')"

echo
echo "3/4 · Dando permisos a la cuenta de servicio de Cloud Run..."
PROJECT_NUMBER="$(gcloud projects describe "$PROJECT_ID" --format='value(projectNumber)')"
SA="${PROJECT_NUMBER}-compute@developer.gserviceaccount.com"

gcloud projects add-iam-policy-binding "$PROJECT_ID" \
    --member="serviceAccount:${SA}" --role="roles/bigquery.dataEditor" --condition=None >/dev/null

gcloud projects add-iam-policy-binding "$PROJECT_ID" \
    --member="serviceAccount:${SA}" --role="roles/bigquery.jobUser" --condition=None >/dev/null

gsutil iam ch "serviceAccount:${SA}:roles/storage.objectViewer" "gs://${BUCKET}"

echo
echo "4/4 · Listo ✅"
echo
echo "Copia esta URL en js/config.js -> flowRunnerHttpEndpoint:"
echo
echo "    ${SERVICE_URL}"
echo

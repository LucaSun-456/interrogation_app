#!/bin/bash
# Fix host directory permissions for Docker appuser (UID 999). Run on the VPS once.
set -euo pipefail
cd /opt/interrogation-app

mkdir -p logs data materials materials/prompts

if [ -d experiment_data.xlsx ]; then
    echo "[WARN] Removing directory experiment_data.xlsx (should be data/experiment_data.xlsx)"
    rm -rf experiment_data.xlsx
fi
if [ -f experiment_data.xlsx ] && [ ! -f data/experiment_data.xlsx ]; then
    mv experiment_data.xlsx data/experiment_data.xlsx
fi

# Debian slim image: appuser from useradd -r is typically UID 999
APP_UID="${APP_UID:-999}"
APP_GID="${APP_GID:-999}"

chown -R "${APP_UID}:${APP_GID}" logs data materials
chmod -R u+rwX logs data materials

echo "Permissions fixed for UID ${APP_UID}. Restarting containers..."
docker compose up -d --force-recreate app nginx

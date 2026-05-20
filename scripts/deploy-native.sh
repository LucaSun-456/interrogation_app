#!/bin/bash
# Deploy without Docker: Python venv + Gunicorn + (optional) host nginx
set -euo pipefail

APP_DIR="${APP_DIR:-/home/spe_avatar/interrogation-app}"
APP_PORT="${APP_PORT:-3003}"
PY="${PYTHON:-python3}"
SERVICE_NAME="${SERVICE_NAME:-interrogation-app}"
export GUNICORN_BIND="${GUNICORN_BIND:-0.0.0.0:${APP_PORT}}"

cd "$APP_DIR"

echo "=== Interrogation App — native deploy (no Docker) ==="

if ! command -v "$PY" &>/dev/null; then
    echo "ERROR: $PY not found. Install Python 3.11+ first."
    exit 1
fi

echo "[1/6] Python: $($PY --version)"

if [ ! -d .venv ]; then
    echo "[2/6] Creating venv..."
    "$PY" -m venv .venv
else
    echo "[2/6] Using existing .venv"
fi

# shellcheck disable=SC1091
source .venv/bin/activate
pip install -U pip -q
pip install -r requirements.txt -q

if [ ! -f .env ]; then
    echo "WARN: .env missing. Copy .env.example and fill API keys."
    cp -n .env.example .env 2>/dev/null || true
fi
chmod 600 .env 2>/dev/null || true

mkdir -p data logs materials materials/prompts
if [ -d experiment_data.xlsx ] && [ ! -f data/experiment_data.xlsx ]; then
    echo "WARN: remove directory experiment_data.xlsx and use data/experiment_data.xlsx"
fi

export LOG_DIR="$APP_DIR/logs"
export DATA_DIR="$APP_DIR/data"
export EXCEL_FILE="$APP_DIR/data/experiment_data.xlsx"
export GUNICORN_WORKERS="${GUNICORN_WORKERS:-1}"

echo "[3/6] Building combined materials (if sources exist)..."
python scripts/build_combined_materials.py 2>/dev/null || true
if [ ! -f materials/combined_materials.md ] && [ -f materials/combined_materials.md.example ]; then
    cp materials/combined_materials.md.example materials/combined_materials.md
    echo "WARN: Using template materials/combined_materials.md — replace with real IRB materials when ready."
fi
if [ ! -f materials/combined_materials.md ]; then
    echo "ERROR: materials/combined_materials.md missing. Upload it or add source .docx/.pdf files."
    exit 1
fi

echo "[4/6] Smoke test import..."
python -c "from app import app; print('OK:', app.name)"

if systemctl list-unit-files 2>/dev/null | grep -q "^${SERVICE_NAME}.service"; then
    echo "[5/6] Restarting systemd service ${SERVICE_NAME}..."
    sudo systemctl restart "$SERVICE_NAME"
    sleep 2
    sudo systemctl status "$SERVICE_NAME" --no-pager || true
else
    echo "[5/6] No systemd unit '${SERVICE_NAME}'. Install scripts/interrogation-app.service.example"
    echo "      Or run manually: source .venv/bin/activate && gunicorn -c gunicorn_config.py wsgi:app"
fi

echo "[6/6] Health check..."
sleep 1
if curl -sf "http://127.0.0.1:${APP_PORT}/api/health" >/dev/null 2>&1; then
    curl -s "http://127.0.0.1:${APP_PORT}/api/health" | python -m json.tool
    echo ""
    echo "=== Done (port ${APP_PORT}) ==="
else
    echo "WARN: http://127.0.0.1:${APP_PORT}/api/health not reachable yet."
    echo "Start gunicorn or install systemd unit, then configure nginx."
fi

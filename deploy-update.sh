#!/bin/bash
set -euo pipefail

cd /opt/interrogation-app

echo "=== Redeploying Interrogation App ==="
echo ""

echo "[1/3] Stopping services..."
docker compose down

echo "[2/3] Rebuilding and starting..."
docker compose build --no-cache app
docker compose up -d

echo "[3/3] Checking health..."
sleep 5
curl -s http://localhost:8000/api/health | python3 -m json.tool

echo ""
echo "=== Deploy Complete ==="
echo "Check logs: docker compose logs -f app"

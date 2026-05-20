#!/bin/bash
# One-shot Ubuntu deploy: Gunicorn on 127.0.0.1:3003 + nginx → spe-avatar.com
#
# Usage (on a fresh Ubuntu server as root):
#   export GITHUB_REPO="https://github.com/LucaSun-456/interrogation_app.git"
#   curl -fsSL https://raw.githubusercontent.com/LucaSun-456/interrogation_app/main/scripts/deploy-ubuntu-spe-avatar.sh | bash
# Or after cloning:
#   sudo bash scripts/deploy-ubuntu-spe-avatar.sh
#
# Before running: upload .env to /home/spe_avatar/interrogation-app/.env (or set SKIP_ENV_CHECK=1 for dry run)

set -euo pipefail

APP_USER="${APP_USER:-spe_avatar}"
APP_DIR="${APP_DIR:-/home/${APP_USER}/interrogation-app}"
APP_PORT="${APP_PORT:-3003}"
DOMAIN="${DOMAIN:-spe-avatar.com}"
SERVER_IP="${SERVER_IP:-47.238.75.193}"
GITHUB_REPO="${GITHUB_REPO:-https://github.com/LucaSun-456/interrogation_app.git}"
SERVICE_NAME="${SERVICE_NAME:-interrogation-app}"
GUNICORN_BIND="0.0.0.0:${APP_PORT}"

if [ "$(id -u)" -ne 0 ]; then
    echo "Please run as root: sudo bash $0"
    exit 1
fi

echo "=== Deploy interrogation_app → ${GUNICORN_BIND}, nginx → ${DOMAIN} ==="

# ---- 1. System packages & user ----
echo "[1/8] Installing packages and user ${APP_USER}..."
apt-get update -qq
apt-get install -y -qq python3 python3-venv python3-pip nginx git curl ufw

if ! id "$APP_USER" &>/dev/null; then
    useradd -m -s /bin/bash "$APP_USER"
    echo "Created user ${APP_USER}"
fi

# ---- 2. Clone or update code ----
echo "[2/8] Application code in ${APP_DIR}..."
if [ ! -d "${APP_DIR}/.git" ]; then
    sudo -u "$APP_USER" git clone "$GITHUB_REPO" "$APP_DIR"
else
    sudo -u "$APP_USER" git -C "$APP_DIR" pull --ff-only || true
fi
chown -R "${APP_USER}:${APP_USER}" "/home/${APP_USER}"

# ---- 3. .env check ----
echo "[3/8] Checking .env..."
if [ ! -f "${APP_DIR}/.env" ]; then
    if [ "${SKIP_ENV_CHECK:-0}" = "1" ]; then
        echo "WARN: .env missing; copying .env.example"
        sudo -u "$APP_USER" cp "${APP_DIR}/.env.example" "${APP_DIR}/.env"
    else
        echo "ERROR: ${APP_DIR}/.env not found."
        echo "From your PC (PowerShell):"
        echo "  scp .env ${APP_USER}@${SERVER_IP}:${APP_DIR}/.env"
        echo "Then re-run: sudo bash ${APP_DIR}/scripts/deploy-ubuntu-spe-avatar.sh"
        exit 1
    fi
fi
chmod 600 "${APP_DIR}/.env"
chown "${APP_USER}:${APP_USER}" "${APP_DIR}/.env"

# Ensure domain in .env
if ! grep -q "^DOMAIN=" "${APP_DIR}/.env" 2>/dev/null; then
    echo "DOMAIN=${DOMAIN}" >> "${APP_DIR}/.env"
else
    sed -i "s/^DOMAIN=.*/DOMAIN=${DOMAIN}/" "${APP_DIR}/.env"
fi
if ! grep -q "^GUNICORN_BIND=" "${APP_DIR}/.env" 2>/dev/null; then
    echo "GUNICORN_BIND=${GUNICORN_BIND}" >> "${APP_DIR}/.env"
else
    sed -i "s|^GUNICORN_BIND=.*|GUNICORN_BIND=${GUNICORN_BIND}|" "${APP_DIR}/.env"
fi

# ---- 4. Python venv & deps ----
echo "[4/8] Python venv and dependencies..."
sudo -u "$APP_USER" bash -c "
set -e
cd '${APP_DIR}'
if [ ! -d .venv ]; then python3 -m venv .venv; fi
source .venv/bin/activate
pip install -U pip -q
pip install -r requirements.txt -q
mkdir -p data logs materials materials/prompts
EXCEL_PATH='${APP_DIR}/data/experiment_data.xlsx'
if [ -d \"\$EXCEL_PATH\" ]; then
  echo \"ERROR: \$EXCEL_PATH is a directory. Run: rm -rf \$EXCEL_PATH\"
  exit 1
fi
if [ -f \"\$EXCEL_PATH\" ]; then
  if ! python3 -c \"import zipfile; z=zipfile.ZipFile('${APP_DIR}/data/experiment_data.xlsx'); assert '[Content_Types].xml' in z.namelist()\" 2>/dev/null; then
    echo \"WARN: corrupt experiment_data.xlsx — quarantining\"
    mv \"\$EXCEL_PATH\" \"\${EXCEL_PATH}.corrupt.\$(date +%Y%m%d_%H%M%S)\"
  fi
fi
if grep -q '^GUNICORN_WORKERS=' '${APP_DIR}/.env' 2>/dev/null; then
  sed -i 's/^GUNICORN_WORKERS=.*/GUNICORN_WORKERS=1/' '${APP_DIR}/.env'
else
  echo 'GUNICORN_WORKERS=1' >> '${APP_DIR}/.env'
fi
export LOG_DIR='${APP_DIR}/logs'
export DATA_DIR='${APP_DIR}/data'
export EXCEL_FILE='${APP_DIR}/data/experiment_data.xlsx'
python scripts/build_combined_materials.py 2>/dev/null || true
python -c 'from app import app; print(\"Import OK:\", app.name)'
"

# ---- 5. systemd ----
echo "[5/8] systemd service ${SERVICE_NAME}..."
cat > "/etc/systemd/system/${SERVICE_NAME}.service" <<EOF
[Unit]
Description=Interrogation experiment app (Gunicorn :${APP_PORT})
After=network.target

[Service]
Type=simple
User=${APP_USER}
Group=${APP_USER}
WorkingDirectory=${APP_DIR}
EnvironmentFile=${APP_DIR}/.env
Environment=LOG_DIR=${APP_DIR}/logs
Environment=DATA_DIR=${APP_DIR}/data
Environment=EXCEL_FILE=${APP_DIR}/data/experiment_data.xlsx
Environment=GUNICORN_BIND=${GUNICORN_BIND}
Environment=GUNICORN_WORKERS=1
ExecStart=${APP_DIR}/.venv/bin/gunicorn -c gunicorn_config.py wsgi:app
Restart=on-failure
RestartSec=5

[Install]
WantedBy=multi-user.target
EOF

systemctl daemon-reload
systemctl enable "${SERVICE_NAME}"
systemctl restart "${SERVICE_NAME}"

# ---- 6. nginx reverse proxy (only spe-avatar.com; does not touch rogare.site / :3001) ----
echo "[6/8] nginx → ${DOMAIN} → ${GUNICORN_BIND} (rogare.site on :3001 unchanged)..."
if [ "$APP_PORT" = "3001" ]; then
    echo "ERROR: APP_PORT=3001 conflicts with rogare.site. Use default 3003."
    exit 1
fi
if ss -tlnp 2>/dev/null | grep -q ":${APP_PORT} "; then
  if ! curl -sf "http://127.0.0.1:${APP_PORT}/api/health" >/dev/null 2>&1; then
    echo "WARN: port ${APP_PORT} is in use by another process. Check: ss -tlnp | grep ${APP_PORT}"
  fi
fi

cat > "/etc/nginx/sites-available/${DOMAIN}" <<'NGINX_EOF'
server {
    listen 80;
    server_name spe-avatar.com www.spe-avatar.com;

    client_max_body_size 20m;

    location /.well-known/acme-challenge/ {
        root /var/www/certbot;
    }

    location / {
        proxy_pass http://127.0.0.1:3003;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection "upgrade";
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $http_x_forwarded_proto;
        proxy_set_header CF-Connecting-IP $http_cf_connecting_ip;
        proxy_buffering off;
        proxy_read_timeout 120s;
        proxy_send_timeout 120s;
    }

    location ~ /\. {
        deny all;
    }
}
NGINX_EOF

if [ "$APP_PORT" != "3003" ]; then
    sed -i "s|127.0.0.1:3003|127.0.0.1:${APP_PORT}|g" "/etc/nginx/sites-available/${DOMAIN}"
fi

ln -sf "/etc/nginx/sites-available/${DOMAIN}" "/etc/nginx/sites-enabled/${DOMAIN}"
# Do not remove other sites (e.g. rogare.site → :3001)
if [ -L /etc/nginx/sites-enabled/default ] && [ ! -f /etc/nginx/sites-enabled/rogare.site ]; then
    rm -f /etc/nginx/sites-enabled/default
fi
nginx -t
systemctl reload nginx

# ---- 7. Firewall ----
echo "[7/8] UFW (SSH + HTTP + HTTPS)..."
ufw allow 22/tcp comment "SSH" 2>/dev/null || true
ufw allow 80/tcp comment "HTTP" 2>/dev/null || true
ufw allow 443/tcp comment "HTTPS" 2>/dev/null || true
ufw allow "${APP_PORT}/tcp" comment "interrogation-app" 2>/dev/null || true
echo "y" | ufw enable 2>/dev/null || true

# ---- 8. Health check ----
echo "[8/8] Health check..."
sleep 2
if curl -sf "http://127.0.0.1:${APP_PORT}/api/health" >/dev/null; then
    echo "App OK on port ${APP_PORT}:"
    curl -s "http://127.0.0.1:${APP_PORT}/api/health" | python3 -m json.tool
else
    echo "WARN: health check failed. Logs:"
    journalctl -u "${SERVICE_NAME}" -n 40 --no-pager
    exit 1
fi

echo ""
echo "=== Deploy complete ==="
echo "  App:      http://127.0.0.1:${APP_PORT}"
echo "  By IP:    http://${SERVER_IP}:${APP_PORT}"
echo "  Domain:   http://${DOMAIN}  (DNS A → ${SERVER_IP}, nginx :80)"
echo ""
echo "Cloudflare (recommended):"
echo "  - A record @ and www → ${SERVER_IP}, proxy ON (orange cloud)"
echo "  - SSL/TLS mode: Flexible"
echo "  - Always Use HTTPS: ON"
echo ""
echo "Useful commands:"
echo "  sudo systemctl status ${SERVICE_NAME}"
echo "  sudo journalctl -u ${SERVICE_NAME} -f"
echo "  sudo -u ${APP_USER} bash -c 'cd ${APP_DIR} && git pull && bash scripts/deploy-native.sh'"

#!/bin/bash
set -euo pipefail

echo "=== Interrogation App - Server Setup ==="

# ---- System Updates ----
echo "[1/7] Updating system packages..."
apt-get update && apt-get upgrade -y

# ---- Install Docker ----
echo "[2/7] Installing Docker..."
if ! command -v docker &> /dev/null; then
    curl -fsSL https://get.docker.com | sh
    systemctl enable docker
    systemctl start docker
fi

# ---- Install Docker Compose ----
echo "[3/7] Installing Docker Compose..."
if ! docker compose version &> /dev/null; then
    apt-get install -y docker-compose-plugin
fi

# ---- Configure Firewall ----
echo "[4/7] Configuring firewall..."
ufw default deny incoming
ufw default allow outgoing
ufw allow 22/tcp comment "SSH"
ufw allow 80/tcp comment "HTTP"
ufw allow 443/tcp comment "HTTPS"
ufw --force enable
ufw status verbose

# ---- Setup .env ----
echo "[5/7] Setting up environment..."
if [ ! -f .env ]; then
    echo "[WARN] .env file not found. Copy .env.example to .env and fill in real values."
    cp -n .env.example .env 2>/dev/null || true
fi
chmod 600 .env

# ---- Writable dirs & data file (app container runs as UID 999) ----
mkdir -p logs data materials materials/prompts
# Remove Docker mistake: experiment_data.xlsx created as a directory when file was missing
if [ -d experiment_data.xlsx ]; then
    echo "[WARN] Removing erroneous directory experiment_data.xlsx (use data/experiment_data.xlsx)"
    rm -rf experiment_data.xlsx
fi
if [ -f experiment_data.xlsx ] && [ ! -f data/experiment_data.xlsx ]; then
    mv experiment_data.xlsx data/experiment_data.xlsx
    echo "[INFO] Moved experiment_data.xlsx -> data/experiment_data.xlsx"
fi
chown -R 999:999 logs data materials 2>/dev/null || true
chmod -R u+rwX logs data materials 2>/dev/null || true

# ---- Build images ----
echo "[6/7] Building Docker images..."
docker compose build --no-cache

# ---- Start services ----
echo "[7/7] Starting services..."
docker compose up -d

echo ""
echo "=== Setup Complete ==="
echo "App is running. Check status with: docker compose ps"
echo "View logs with: docker compose logs -f"
echo ""
echo "Next steps:"
echo "1. Edit .env with real API keys and passwords"
echo "2. Point your domain DNS A record to this server's IP"
echo "3. Run: bash certbot-setup.sh spe-avatar.com your@email.com"
echo "4. Update nginx/conf.d/app.conf with your domain, then:"
echo "   docker compose exec nginx nginx -s reload"

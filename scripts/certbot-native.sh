#!/bin/bash
# Let's Encrypt HTTPS for native Ubuntu + host nginx (no Docker).
# Prereqs: DNS A @ and www → this server; port 80 open; nginx site spe-avatar.com on :80.
#
# Usage:
#   sudo bash scripts/certbot-native.sh spe-avatar.com your@email.com

set -euo pipefail

DOMAIN="${1:-spe-avatar.com}"
EMAIL="${2:-}"
APP_DIR="${APP_DIR:-/home/spe_avatar/interrogation-app}"
APP_PORT="${APP_PORT:-3003}"

if [ "$(id -u)" -ne 0 ]; then
    echo "Run as root: sudo bash $0 spe-avatar.com your@email.com"
    exit 1
fi

if [ -z "$EMAIL" ]; then
    echo "Usage: sudo bash $0 <domain> <email>"
    exit 1
fi

echo "=== Certbot (native) for ${DOMAIN} → 127.0.0.1:${APP_PORT} ==="

apt-get update -qq
apt-get install -y -qq certbot python3-certbot-nginx

mkdir -p /var/www/certbot
chmod 755 /var/www/certbot

# Ensure HTTP vhost exists for ACME + proxy
if [ ! -f "/etc/nginx/sites-available/${DOMAIN}" ]; then
    cp "${APP_DIR}/scripts/nginx-spe-avatar.conf.example" "/etc/nginx/sites-available/${DOMAIN}"
    ln -sf "/etc/nginx/sites-available/${DOMAIN}" "/etc/nginx/sites-enabled/${DOMAIN}"
fi

nginx -t
systemctl reload nginx

CERT_DOMAINS=(-d "$DOMAIN")
if dig +short "www.${DOMAIN}" @8.8.8.8 2>/dev/null | grep -q .; then
    CERT_DOMAINS+=(-d "www.${DOMAIN}")
fi

certbot certonly --webroot -w /var/www/certbot \
    "${CERT_DOMAINS[@]}" \
    --email "$EMAIL" --agree-tos --no-eff-email \
    --non-interactive

cp "${APP_DIR}/scripts/nginx-spe-avatar-ssl.conf.example" \
    "/etc/nginx/sites-available/${DOMAIN}-ssl"
ln -sf "/etc/nginx/sites-available/${DOMAIN}-ssl" "/etc/nginx/sites-enabled/${DOMAIN}-ssl"

# Disable plain HTTP proxy site if SSL bundle includes redirect block
if grep -q "return 301 https" "/etc/nginx/sites-available/${DOMAIN}-ssl"; then
    rm -f "/etc/nginx/sites-enabled/${DOMAIN}"
fi

nginx -t
systemctl reload nginx

ufw allow 443/tcp comment "HTTPS" 2>/dev/null || true

echo ""
echo "=== Done ==="
echo "  https://${DOMAIN}"
echo "  Renew test: certbot renew --dry-run"

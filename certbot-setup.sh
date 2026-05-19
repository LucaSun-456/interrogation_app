#!/bin/bash
set -euo pipefail

DOMAIN="${1:-}"
if [ -z "$DOMAIN" ]; then
    echo "Usage: $0 <your-domain.com>"
    echo "Example: $0 interrogation.example.com"
    exit 1
fi

EMAIL="${2:-admin@$DOMAIN}"

echo "=== SSL Certificate Setup for $DOMAIN ==="
echo "Email: $EMAIL"
echo ""

cd /opt/interrogation-app

# Ensure nginx is running
docker compose up -d nginx

# Test with staging first
echo "[1/3] Testing with Let's Encrypt staging..."
docker compose run --rm certbot certonly \
    --webroot --webroot-path=/var/www/certbot \
    --staging \
    -d "$DOMAIN" -d "www.$DOMAIN" \
    --email "$EMAIL" --agree-tos --no-eff-email

echo ""
echo "[2/3] Obtaining real certificate..."
docker compose run --rm certbot certonly \
    --webroot --webroot-path=/var/www/certbot \
    -d "$DOMAIN" -d "www.$DOMAIN" \
    --email "$EMAIL" --agree-tos --no-eff-email

echo ""
echo "[3/3] Enabling HTTPS in nginx..."
cp -n nginx/conf.d/app-ssl.conf.example nginx/conf.d/app-ssl.conf 2>/dev/null || true
docker compose exec nginx nginx -t
docker compose exec nginx nginx -s reload
echo "Done. Visit https://$DOMAIN"
echo ""
echo "Certificate will auto-renew via the certbot container (checks every 12h)."

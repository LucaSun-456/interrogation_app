#!/bin/bash
set -euo pipefail

DOMAIN="${1:-}"
if [ -z "$DOMAIN" ]; then
    echo "Usage: $0 <your-domain.com> [email] [--staging]"
    echo "Example: $0 spe-avatar.com you@example.com"
    exit 1
fi

EMAIL="${2:-admin@$DOMAIN}"
USE_STAGING=false
if [[ "${3:-}" == "--staging" ]] || [[ "${2:-}" == "--staging" ]]; then
    USE_STAGING=true
    [[ "${2:-}" == "--staging" ]] && EMAIL="admin@$DOMAIN"
fi

echo "=== SSL Certificate Setup for $DOMAIN ==="
echo "Email: $EMAIL"
echo ""

cd /opt/interrogation-app

mkdir -p nginx/www nginx/ssl
docker compose up -d nginx app

# Only request domains that resolve to this server (skip www if no DNS)
CERT_DOMAINS=(-d "$DOMAIN")
if dig +short "www.$DOMAIN" @8.8.8.8 2>/dev/null | grep -q .; then
    CERT_DOMAINS+=(-d "www.$DOMAIN")
    echo "[INFO] www.$DOMAIN has DNS — including in certificate"
else
    echo "[WARN] www.$DOMAIN has no DNS — certificate will be for $DOMAIN only"
fi

CERTBOT_FLAGS=(certonly --webroot --webroot-path=/var/www/certbot
    "${CERT_DOMAINS[@]}"
    --email "$EMAIL" --agree-tos --no-eff-email
    --non-interactive --verbose)

if $USE_STAGING; then
    echo "[1/2] Staging test (optional, may be slow from China)..."
    docker compose run --rm certbot "${CERTBOT_FLAGS[@]}" --staging
fi

echo "[*] Obtaining production certificate (may take 1–3 min)..."
docker compose run --rm certbot "${CERTBOT_FLAGS[@]}"

echo ""
echo "[*] Enabling HTTPS in nginx..."
cp -n nginx/conf.d/app-ssl.conf.example nginx/conf.d/app-ssl.conf 2>/dev/null || true
docker compose exec nginx nginx -t
docker compose exec nginx nginx -s reload
echo "Done. Visit https://$DOMAIN"

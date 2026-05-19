#!/bin/bash
# Run once as root on Ubuntu: sudo bash scripts/ubuntu-setup-user.sh
set -euo pipefail

USERNAME="${USERNAME:-spe_avatar}"
APP_HOME="/home/${USERNAME}"
REPO_DIR="${APP_HOME}/interrogation-app"

if [ "$(id -u)" -ne 0 ]; then
    echo "Run as root: sudo bash $0"
    exit 1
fi

echo "=== Create user ${USERNAME} and prepare ${APP_HOME} ==="

if ! id "$USERNAME" &>/dev/null; then
    useradd -m -s /bin/bash "$USERNAME"
    echo "User ${USERNAME} created."
else
    echo "User ${USERNAME} already exists."
fi

apt-get update
apt-get install -y python3 python3-venv python3-pip nginx git curl

mkdir -p "$REPO_DIR"
chown -R "${USERNAME}:${USERNAME}" "$APP_HOME"

echo ""
echo "=== Done ==="
echo "Next (as ${USERNAME}):"
echo "  sudo -u ${USERNAME} -i"
echo "  cd ~/interrogation-app   # clone or upload project here"
echo "  # copy .env, then:"
echo "  bash scripts/deploy-native.sh"
echo ""
echo "Then as root:"
echo "  sudo cp ${REPO_DIR}/scripts/interrogation-app.service.example /etc/systemd/system/interrogation-app.service"
echo "  sudo systemctl enable --now interrogation-app"

#!/usr/bin/env bash
set -euo pipefail

APP_DIR="${APP_DIR:-$HOME/certificati-dpi}"

echo "==> Aggiornamento pacchetti"
sudo apt-get update

echo "==> Installazione Docker"
sudo apt-get install -y ca-certificates curl gnupg git ufw
sudo install -m 0755 -d /etc/apt/keyrings
if [ ! -f /etc/apt/keyrings/docker.asc ]; then
  curl -fsSL https://download.docker.com/linux/ubuntu/gpg | sudo gpg --dearmor -o /etc/apt/keyrings/docker.asc
  sudo chmod a+r /etc/apt/keyrings/docker.asc
fi
echo \
  "deb [arch=$(dpkg --print-architecture) signed-by=/etc/apt/keyrings/docker.asc] https://download.docker.com/linux/ubuntu \
  $(. /etc/os-release && echo "$VERSION_CODENAME") stable" | \
  sudo tee /etc/apt/sources.list.d/docker.list > /dev/null

sudo apt-get update
sudo apt-get install -y docker-ce docker-ce-cli containerd.io docker-buildx-plugin docker-compose-plugin
sudo usermod -aG docker "$USER" || true

echo "==> Apertura porte firewall locali"
sudo ufw allow 22/tcp || true
sudo ufw allow 80/tcp || true
sudo ufw allow 443/tcp || true
sudo ufw --force enable || true

echo "==> Preparazione cartelle dati"
mkdir -p "$APP_DIR/oracle-data"

echo "==> Avvio applicazione"
cd "$APP_DIR"
sudo docker compose -f docker-compose.oracle.yml up -d --build

echo
echo "Installazione completata."
echo "Se Docker e' stato appena installato, potresti dover uscire e rientrare nella sessione SSH."
echo "Controlla lo stato con:"
echo "  cd $APP_DIR && sudo docker compose -f docker-compose.oracle.yml ps"

#!/bin/bash
# =============================================================
#  Setup-Script: SIA 451 Parser auf Ubuntu VM
#  Ausführen als root oder mit sudo: bash setup_vm.sh
# =============================================================
set -e

APP_DIR="/app/crbxparser"
VENV_DIR="$APP_DIR/venv"
SERVICE="crbxparser"

echo "==> Venv erstellen..."
python3 -m venv "$VENV_DIR"

echo "==> Abhängigkeiten installieren..."
"$VENV_DIR/bin/pip" install --upgrade pip -q
"$VENV_DIR/bin/pip" install -r "$APP_DIR/requirements.txt" gunicorn -q

echo "==> Systemd-Service installieren..."
cp "$APP_DIR/crbxparser.service" /etc/systemd/system/

systemctl daemon-reload
systemctl enable "$SERVICE"
systemctl restart "$SERVICE"

echo ""
echo "==> Fertig. Status:"
systemctl status "$SERVICE" --no-pager

echo ""

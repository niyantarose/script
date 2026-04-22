#!/usr/bin/env bash
set -euo pipefail

APP_DIR="${APP_DIR:-/home/ubuntu/inventory_tool}"
PYTHON_BIN="${PYTHON_BIN:-python3}"
VENV_DIR="${VENV_DIR:-$APP_DIR/venv}"

echo "[1/6] パッケージ更新"
sudo apt update
sudo apt upgrade -y

echo "[2/6] Python / MySQL / Nginx をインストール"
sudo apt install -y \
  "${PYTHON_BIN}" \
  python3-pip \
  python3-venv \
  mysql-server \
  nginx

echo "[3/6] アプリ配置ディレクトリを準備"
mkdir -p "$APP_DIR"
mkdir -p "$APP_DIR/logs"

echo "[4/6] 仮想環境を準備"
if [ ! -d "$VENV_DIR" ]; then
  "$PYTHON_BIN" -m venv "$VENV_DIR"
fi
source "$VENV_DIR/bin/activate"
python -m pip install --upgrade pip
python -m pip install -r "$APP_DIR/requirements.txt"

echo "[5/6] systemd / nginx 設定の配置例"
echo "  sudo cp $APP_DIR/deploy/systemd/inventory-tool.service /etc/systemd/system/inventory-tool.service"
echo "  sudo cp $APP_DIR/deploy/nginx/inventory-tool.conf /etc/nginx/sites-available/inventory-tool"
echo "  sudo ln -sf /etc/nginx/sites-available/inventory-tool /etc/nginx/sites-enabled/inventory-tool"

echo "[6/6] 次の手順"
echo "  1. MySQLに inventory_db と inventory_user を作成"
echo "  2. .env を本番値で作成"
echo "  3. systemctl daemon-reload && systemctl enable --now inventory-tool"
echo "  4. nginx -t && systemctl restart nginx"


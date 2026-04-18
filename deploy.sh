#!/usr/bin/env bash
# deploy.sh — pull latest botpreza code on the server and restart the bot.
#
# Usage (on the server as root or a user with sudo):
#   sudo bash deploy.sh
#
# Assumptions:
#   - Repo is checked out at /opt/SMARTMONEY (override via SMARTMONEY_DIR env var)
#   - A Python venv lives at $SMARTMONEY_DIR/botpreza/venv
#   - systemd unit "botpreza" owns the running bot (install systemd/botpreza.service once)
#   - keys.env is kept on the server outside of git (never committed)

set -euo pipefail

REPO_DIR="${SMARTMONEY_DIR:-/opt/SMARTMONEY}"
BRANCH="${SMARTMONEY_BRANCH:-main}"
VENV_DIR="${SMARTMONEY_VENV:-$REPO_DIR/botpreza/venv}"
SERVICE_NAME="${SMARTMONEY_SERVICE:-botpreza}"

echo "▶ deploy: repo=$REPO_DIR branch=$BRANCH venv=$VENV_DIR service=$SERVICE_NAME"

if [ ! -d "$REPO_DIR/.git" ]; then
    echo "✖ $REPO_DIR is not a git checkout" >&2
    exit 1
fi

cd "$REPO_DIR"

echo "▶ fetching origin..."
git fetch --prune origin

echo "▶ checking out $BRANCH and fast-forwarding..."
git checkout "$BRANCH"
git reset --hard "origin/$BRANCH"

if [ ! -d "$VENV_DIR" ]; then
    echo "▶ creating venv at $VENV_DIR"
    python3 -m venv "$VENV_DIR"
fi

echo "▶ installing dependencies..."
"$VENV_DIR/bin/pip" install --upgrade pip
"$VENV_DIR/bin/pip" install -r "$REPO_DIR/botpreza/requirements.txt"

if [ -f "$REPO_DIR/systemd/botpreza.service" ]; then
    echo "▶ syncing systemd unit..."
    install -m 0644 "$REPO_DIR/systemd/botpreza.service" /etc/systemd/system/botpreza.service
    systemctl daemon-reload
    systemctl enable "$SERVICE_NAME" >/dev/null 2>&1 || true
fi

echo "▶ restarting $SERVICE_NAME..."
systemctl restart "$SERVICE_NAME"
sleep 1
systemctl --no-pager --full status "$SERVICE_NAME" | head -20 || true

echo "✔ deploy complete"

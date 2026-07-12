#!/bin/sh
set -eu

cd "$(dirname "$0")"
ENV_FILE=".env"
MODE_LINE="LICENSE_SERVER_MODE=local-pi-v1"

backup_old_env() {
  stamp="$(date +%Y%m%d-%H%M%S)"
  backup=".env.keygen-backup-$stamp"
  mv "$ENV_FILE" "$backup"
  echo "Đã backup env cũ thành $backup"
}

if [ -f "$ENV_FILE" ]; then
  if grep -q '^KEYGEN_' "$ENV_FILE" || ! grep -q "^$MODE_LINE$" "$ENV_FILE"; then
    backup_old_env
  elif grep -q '^ADMIN_TOKEN=change-' "$ENV_FILE"; then
    backup_old_env
  fi
fi

if [ ! -f "$ENV_FILE" ]; then
  token="310902"
  umask 077
  {
    echo "$MODE_LINE"
    echo "ADMIN_PORT=8080"
    echo "ADMIN_TOKEN=$token"
    echo "LICENSE_ACCOUNT_ID=6f1f56e8-3b6f-4a86-9a31-9e0e7f62c001"
  } > "$ENV_FILE"
  echo "Đã tạo .env mới."
  echo "ADMIN_TOKEN=$token"
else
  echo ".env hiện tại đã đúng format local-pi-v1; không ghi đè token."
fi


chmod 600 "$ENV_FILE"
mkdir -p releases license-data

docker compose -f docker-compose.pi.yml up -d --build
docker compose -f docker-compose.pi.yml ps

echo "Health: http://127.0.0.1:8080/health"
echo "Để xem lại token: grep '^ADMIN_TOKEN=' .env"

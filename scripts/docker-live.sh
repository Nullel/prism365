#!/bin/sh
# Live local install of Prism365 in Docker.
#
#   https://localhost:8443   the app, behind a local HTTPS proxy
#   PostgreSQL volume        sessions, collected sign-ins, baselines and findings; survives
#                            restarts, rebuilds and "down" (never run "down -v" unless you mean
#                            to delete them)
#
# Usage: scripts/docker-live.sh init|up|down|logs|status|backup|export-ca
set -eu
cd "$(dirname "$0")/.."

ENV_FILE="${PRISM_ENV_FILE:-.env.docker}"
PROJECT="${PRISM_PROJECT:-prism365}"

compose() {
  docker compose -p "$PROJECT" -f compose.yml -f compose.live.yml --env-file "$ENV_FILE" --profile live "$@"
}

value() { sed -n "s/^$1=//p" "$2" | tail -1; }
secret() { openssl rand -base64 64 | tr -dc 'A-Za-z0-9' | cut -c1-40; }
setting() { value "$1" "$ENV_FILE"; }

require_env() {
  if [ ! -f "$ENV_FILE" ]; then
    echo "No $ENV_FILE yet. Run: scripts/docker-live.sh init" >&2
    exit 1
  fi
}

case "${1:-help}" in
  init)
    if [ -f "$ENV_FILE" ]; then
      echo "$ENV_FILE already exists. Its passwords and key match the database volume; keep it." >&2
      exit 1
    fi
    source_env="${2:-.env}"
    for key in ENTRA_TENANT_ID ENTRA_CLIENT_ID ENTRA_ADMIN_OBJECT_IDS; do
      if [ -z "$(value "$key" "$source_env")" ]; then
        echo "$key is missing from $source_env." >&2
        exit 1
      fi
    done
    secrets_dir="${PRISM_SECRETS_DIR:-./secrets}"
    for file in entra-public.pem entra-private.pem; do
      if [ ! -f "$secrets_dir/$file" ]; then
        echo "$secrets_dir/$file is missing. Create the certificate first (docs/SETUP.md)." >&2
        exit 1
      fi
    done
    # Email alerts are optional. When SMTP_URL names a user, its password goes in the secrets
    # directory as smtp-password, like the certificate, and is mounted read-only.
    smtp_password_file=""
    smtp_url="$(value SMTP_URL "$source_env")"
    case "$smtp_url" in
      *://*@*)
        if [ ! -f "$secrets_dir/smtp-password" ]; then
          echo "SMTP_URL names a user, so $secrets_dir/smtp-password must hold its password." >&2
          exit 1
        fi
        smtp_password_file=/run/secrets/smtp-password
        ;;
    esac
    port="${PRISM_HTTPS_PORT:-8443}"
    umask 077
    cat > "$ENV_FILE" <<EOF
# Live Docker install, created by scripts/docker-live.sh init. Keep this file and back it up
# with the database: the passwords and session encryption key here belong to the data in the
# PostgreSQL volume. It is git-ignored and never copied into an image.
DATA_MODE=live
PUBLIC_ORIGIN=https://localhost:$port
PRISM_HTTPS_PORT=$port
ENTRA_TENANT_ID=$(value ENTRA_TENANT_ID "$source_env")
ENTRA_CLIENT_ID=$(value ENTRA_CLIENT_ID "$source_env")
ENTRA_ADMIN_OBJECT_IDS=$(value ENTRA_ADMIN_OBJECT_IDS "$source_env")
ENABLED_READ_FEATURES=$(value ENABLED_READ_FEATURES "$source_env")
SECRETS_DIR=$secrets_dir
ENTRA_CERTIFICATE_PATH=/run/secrets/entra-public.pem
ENTRA_PRIVATE_KEY_PATH=/run/secrets/entra-private.pem
DATABASE_ADMIN_PASSWORD=$(secret)
DATABASE_APP_PASSWORD=$(secret)
STATE_ENCRYPTION_KEY=$(openssl rand -hex 32)
SIGNINS_EXPECTED_COUNTRIES=$(value SIGNINS_EXPECTED_COUNTRIES "$source_env")
SIGNINS_TRUSTED_NETWORKS=$(value SIGNINS_TRUSTED_NETWORKS "$source_env")
SIGNINS_TIME_ZONE=$(value SIGNINS_TIME_ZONE "$source_env")
SMTP_URL=$smtp_url
SMTP_PASSWORD_FILE=$smtp_password_file
SMTP_FROM=$(value SMTP_FROM "$source_env")
ALERT_EMAIL_DOMAINS=$(value ALERT_EMAIL_DOMAINS "$source_env")
WEBHOOK_ALLOWED_PORTS=$(value WEBHOOK_ALLOWED_PORTS "$source_env")
EOF
    echo "Created $ENV_FILE. Next: scripts/docker-live.sh up"
    ;;
  up)
    require_env
    compose up -d --build
    origin="$(setting PUBLIC_ORIGIN)"
    printf "Waiting for %s " "$origin"
    attempt=0
    until curl -ks --max-time 3 "$origin/health" >/dev/null 2>&1; do
      attempt=$((attempt + 1))
      if [ "$attempt" -ge 60 ]; then
        echo
        echo "Prism365 did not answer within two minutes. Check: scripts/docker-live.sh logs" >&2
        exit 1
      fi
      printf "."
      sleep 2
    done
    echo
    echo "Prism365 is running at $origin"
    ;;
  down)
    require_env
    compose down
    echo "Stopped. The database volume, and everything collected, is kept."
    ;;
  logs)
    require_env
    compose logs -f --tail 100 app migrate proxy
    ;;
  status)
    require_env
    compose ps -a
    ;;
  backup)
    require_env
    mkdir -p backups
    chmod 700 backups
    file="backups/prism365-$(date +%Y%m%d-%H%M%S).dump"
    partial="$file.partial"
    umask 077
    # Written under a temporary name and read back before it gets its real name, so a failed or
    # truncated dump never sits in backups/ looking like a good one.
    if ! compose exec -T db pg_dump -U postgres -Fc prism365 > "$partial"; then
      rm -f "$partial"
      echo "The backup failed and nothing was kept. Is the stack running? Check: scripts/docker-live.sh status" >&2
      exit 1
    fi
    # Restoring to /dev/null reads every data block, so a truncated dump is caught, not only a
    # damaged table of contents.
    if ! compose exec -T db pg_restore -f /dev/null < "$partial" > /dev/null 2>&1; then
      rm -f "$partial"
      echo "The backup could not be read back, so it was discarded." >&2
      exit 1
    fi
    mv "$partial" "$file"
    echo "Wrote $file. It contains tenant data in plain form: keep it with $ENV_FILE, out of cloud sync."
    echo "Retention does not reach backups; delete old ones yourself."
    ;;
  export-ca)
    require_env
    container="$(compose ps -q proxy)"
    if [ -z "$container" ]; then
      echo "The proxy is not running. Run: scripts/docker-live.sh up" >&2
      exit 1
    fi
    docker cp "$container:/data/caddy/pki/authorities/local/root.crt" deploy/local/caddy-local-root.crt
    echo "Saved deploy/local/caddy-local-root.crt. To trust it for your user on this Mac:"
    echo "  security add-trusted-cert -r trustRoot -k ~/Library/Keychains/login.keychain-db deploy/local/caddy-local-root.crt"
    echo "It is not limited to localhost, and its key lives in the caddy_data volume. To undo:"
    echo "  security delete-certificate -c 'Caddy Local Authority - 2026 ECC Root' ~/Library/Keychains/login.keychain-db"
    ;;
  *)
    echo "Usage: scripts/docker-live.sh init|up|down|logs|status|backup|export-ca" >&2
    exit 1
    ;;
esac

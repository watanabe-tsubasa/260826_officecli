#!/usr/bin/env bash

set -euo pipefail

APP_NAME="ProtoOut Office Mail Agent"

echo "Azureへのログイン状態を確認しています..."

if ! az account show >/dev/null 2>&1; then
  echo "Azureへログインしていません。"
  echo "az login --use-device-code を実行してください。"
  exit 1
fi

echo "利用中のAzure環境:"
az account show \
  --query "{subscription:name, tenantId:tenantId, user:user.name}" \
  -o table

echo
echo "既存のアプリ登録を確認しています..."

APP_ID=$(
  az ad app list \
    --display-name "$APP_NAME" \
    --query "[0].appId" \
    -o tsv
)

if [ -z "$APP_ID" ]; then
  echo "Microsoft Graph用アプリを作成します..."

  APP_ID=$(
    az ad app create \
      --display-name "$APP_NAME" \
      --sign-in-audience PersonalMicrosoftAccount \
      --is-fallback-public-client true \
      --required-resource-accesses @config/graph-permissions.json \
      --query appId \
      -o tsv
  )
else
  echo "既存のアプリを利用します: $APP_ID"

  az ad app update \
    --id "$APP_ID" \
    --sign-in-audience PersonalMicrosoftAccount \
    --is-fallback-public-client true \
    --required-resource-accesses @config/graph-permissions.json \
    >/dev/null
fi

cat > .env <<ENVEOF
MICROSOFT_CLIENT_ID=$APP_ID
MICROSOFT_AUTHORITY=https://login.microsoftonline.com/consumers
MAIL_SUBJECT_PREFIX=[OFFICE-AI]
POLL_INTERVAL_MS=10000
ENVEOF

echo
echo "========================================"
echo "Microsoft Graphアプリの準備が完了しました"
echo "========================================"
echo "Application Client ID:"
echo "$APP_ID"
echo
echo ".env を作成しました"

#!/bin/bash
# Copyright 2026 Google LLC
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

set -e

# Change to the directory where the script is located
cd "$(dirname "$0")"

echo "============================================================"
echo "🚀 Gemini for Office 365 - Full Deployment Script"
echo "============================================================"

# 1. Load environment variables
if [ -f ".env" ]; then
    echo "📄 Loading configuration from .env..."
    # Export all variables from .env, ignoring comments
    export $(grep -v '^#' .env | xargs)
else
    echo "❌ ERROR: .env file not found. Please create one based on .env.example"
    exit 1
fi

# Ensure mandatory variables exist
if [ -z "$PROJECT_ID" ] || [ -z "$REGION" ] || [ -z "$SERVICE_ACCOUNT" ]; then
    echo "❌ ERROR: PROJECT_ID, REGION, and SERVICE_ACCOUNT must be defined in .env"
    exit 1
fi

# Respect pre-configured CLOUDSDK_CONFIG from environment or .env, or use standard user config if present
if [ -n "$CUSTOM_CLOUDSDK_CONFIG" ] && [ -d "$CUSTOM_CLOUDSDK_CONFIG" ]; then
    export CLOUDSDK_CONFIG="$CUSTOM_CLOUDSDK_CONFIG"
elif [ -d "$HOME/.config/gcloud-altostrat-deploy" ] && [ -w "$HOME/.config/gcloud-altostrat-deploy" ]; then
    export CLOUDSDK_CONFIG="$HOME/.config/gcloud-altostrat-deploy"
fi

echo ""
echo "☁️  Step 1: Deploying askgemini-proxy (Private Backend)..."
# --no-allow-unauthenticated only removes a public (allUsers) grant. --invoker-iam-check also turns
# Cloud Run's sign-in check back on if it was ever switched off, so only callers with run.invoker get in.
gcloud run deploy askgemini-proxy \
  --source ./geminiproxy \
  --project "$PROJECT_ID" \
  --region "$REGION" \
  --memory 1Gi \
  --service-account "$SERVICE_ACCOUNT" \
  --no-allow-unauthenticated \
  --invoker-iam-check \
  --set-env-vars "PROJECT_ID=${PROJECT_ID},GE_GCP_PROJECT_ID=${GE_GCP_PROJECT_ID},GE_GCP_LOCATION=${GE_GCP_LOCATION:-us},STREAM_ASSIST_ENDPOINT_LOCATION=${STREAM_ASSIST_ENDPOINT_LOCATION:-us},GEMINI_ENTERPRISE_APP_ID=${GEMINI_ENTERPRISE_APP_ID},BACKEND_MODE=streamassist,ENTERPRISE_COLLECTION_ID=${ENTERPRISE_COLLECTION_ID:-default_collection},ENTERPRISE_ASSISTANT_ID=${ENTERPRISE_ASSISTANT_ID:-default_assistant},ALLOW_SERVICE_ACCOUNT_FALLBACK=false" \
  --quiet

echo "🔒 Locking down askgemini-proxy to internal Service Account..."
gcloud run services add-iam-policy-binding askgemini-proxy \
  --region "$REGION" \
  --project "$PROJECT_ID" \
  --member="serviceAccount:$SERVICE_ACCOUNT" \
  --role="roles/run.invoker" \
  --quiet

ACTUAL_ASKGEMINI_URL=$(gcloud run services describe askgemini-proxy --project "$PROJECT_ID" --region "$REGION" --format="value(status.url)")

echo ""
echo "☁️  Step 2: Deploying auth-proxy..."
gcloud run deploy auth-proxy \
  --source ./authproxy \
  --project "$PROJECT_ID" \
  --region "$REGION" \
  --service-account "$SERVICE_ACCOUNT" \
  --allow-unauthenticated \
  --set-env-vars "MICROSOFT_ENTRA_APP_ID=${MICROSOFT_ENTRA_APP_ID},MICROSOFT_ENTRA_TENANT_ID=${MICROSOFT_ENTRA_TENANT_ID},DOWNSTREAM_BACKEND_URL=${ACTUAL_ASKGEMINI_URL},GE_GCP_PROJECT_ID=${GE_GCP_PROJECT_ID},GE_GCP_LOCATION=${GE_GCP_LOCATION:-us},USER_AUTH_MODE=${USER_AUTH_MODE:-cloud_identity},REQUIRE_ENTRA_AUTH=${REQUIRE_ENTRA_AUTH:-true},VERBOSE_LOGGING=${VERBOSE_LOGGING:-false},GOOGLE_OAUTH_CLIENT_ID=${GOOGLE_OAUTH_CLIENT_ID}" \
  --quiet

echo ""
echo "🔍 Querying actual Cloud Run URL for auth-proxy to inject into Webpack..."
if [ -n "$GEMINI_PROXY_URL" ]; then
    echo "Using pre-configured GEMINI_PROXY_URL from .env: $GEMINI_PROXY_URL"
    # Extract the base URL without /askGeminiEnterprise for the manifest generator
    ACTUAL_AUTH_PROXY_URL=$(echo "$GEMINI_PROXY_URL" | sed 's|/askGeminiEnterprise||')
else
    ACTUAL_AUTH_PROXY_URL=$(gcloud run services describe auth-proxy --project "$PROJECT_ID" --region "$REGION" --format="value(status.url)" 2>/dev/null || echo "")
fi

if [ -n "$ACTUAL_AUTH_PROXY_URL" ]; then
    echo "Found auth-proxy at: $ACTUAL_AUTH_PROXY_URL"
    export GEMINI_PROXY_URL="$ACTUAL_AUTH_PROXY_URL/askGeminiEnterprise"
else
    echo "⚠️  WARNING: Could not find auth-proxy. Have you deployed it yet? Proceeding with defaults."
fi

echo ""
echo "🏗️  Step 3: Building production Webpack bundle..."
cd microsoft-addin
if [ ! -d "node_modules" ] || [ ! -f "node_modules/.bin/webpack" ]; then
    echo "📦 Installing npm dependencies in microsoft-addin..."
    npm install --registry=https://registry.npmjs.org/
fi
npm run build
cd ..

echo ""
echo "☁️  Step 4: Deploying Frontend to Google Cloud Run..."

gcloud run deploy gemini-frontend \
  --source ./microsoft-addin \
  --project "$PROJECT_ID" \
  --region "$REGION" \
  --service-account "$SERVICE_ACCOUNT" \
  --allow-unauthenticated \
  --quiet

echo ""
echo "📝 Step 5: Fetching actual deployed URL and generating fresh Office Manifest..."
# Cloud Run URLs can have unpredictable hashes, so we extract the exact URL that was just deployed
if [ -n "$FRONTEND_URL" ]; then
    echo "Using pre-configured FRONTEND_URL from .env: $FRONTEND_URL"
    ACTUAL_FRONTEND_URL="$FRONTEND_URL"
else
    ACTUAL_FRONTEND_URL=$(gcloud run services describe gemini-frontend --project "$PROJECT_ID" --region "$REGION" --format="value(status.url)")
fi

if [ -z "$ACTUAL_FRONTEND_URL" ]; then
    echo "❌ ERROR: Could not retrieve the Cloud Run URL. Did the deployment fail?"
    exit 1
fi

echo "Deployed Frontend URL is: $ACTUAL_FRONTEND_URL"

# We pass it explicitly to the script so it doesn't try to guess it based on PROJECT_ID alone
if [ -n "$ACTUAL_AUTH_PROXY_URL" ]; then
    python3 scripts/generate_manifest.py --frontend-url "$ACTUAL_FRONTEND_URL" --auth-proxy-url "$ACTUAL_AUTH_PROXY_URL" --output microsoft-addin/manifest-gemini.xml
else
    python3 scripts/generate_manifest.py --frontend-url "$ACTUAL_FRONTEND_URL" --output microsoft-addin/manifest-gemini.xml
fi

echo ""
echo "✅ Deployment Complete!"
echo "============================================================"
echo "Your Add-in is live and correctly configured for your environment."
echo "Generated Manifests:"
echo "  • Office (Word, PowerPoint, Excel): microsoft-addin/manifest-gemini.xml"
echo "  • Outlook:                          microsoft-addin/manifest-gemini-outlook.xml"
echo "➡️  Next Step: Run './scripts/sideload_mac.sh' to install both manifests into Office & Outlook."

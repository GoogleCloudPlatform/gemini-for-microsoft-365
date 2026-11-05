#!/usr/bin/env bash
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

# ==============================================================================
# Rollback Script: Restore Baseline WIF & Agentspace-452714 Cloud Run Services
# Tag Snapshot: ca-snapshot-wif-baseline
# ==============================================================================

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "${SCRIPT_DIR}/.." && pwd)"
# Load configuration from .env if available
if [ -f "${REPO_ROOT}/.env" ]; then
  # export non-commented vars
  export $(grep -v '^#' "${REPO_ROOT}/.env" | xargs 2>/dev/null || true)
fi

PROJECT_ID="${1:-${PROJECT_ID:-agentspace-452714}}"
REGION="${2:-${REGION:-us-central1}}"

echo "======================================================================"
echo " Starting Rollback to Baseline (Snapshot: ca-snapshot-wif-baseline)   "
echo " Target Project: ${PROJECT_ID} (Region: ${REGION})                    "
echo "======================================================================"

CONFIG_DIR="${REPO_ROOT}/baseline-configs"

# 1. Verify baseline files exist
for svc in "auth-proxy" "askgemini-proxy" "gemini-frontend"; do
  file="${CONFIG_DIR}/${svc}-baseline.yaml"
  if [[ ! -f "${file}" ]]; then
    echo "ERROR: Baseline configuration file not found: ${file}"
    exit 1
  fi
done

# 2. Restore askgemini-proxy (Backend Proxy)
echo ""
echo "[1/3] Restoring Backend Proxy (askgemini-proxy)..."
gcloud run services replace "${CONFIG_DIR}/askgemini-proxy-baseline.yaml" \
  --project="${PROJECT_ID}" \
  --region="${REGION}"

# 3. Restore auth-proxy (Auth Gateway)
echo ""
echo "[2/3] Restoring Auth Gateway Proxy (auth-proxy)..."
gcloud run services replace "${CONFIG_DIR}/auth-proxy-baseline.yaml" \
  --project="${PROJECT_ID}" \
  --region="${REGION}"

# 4. Restore gemini-frontend (Frontend Web App)
echo ""
echo "[3/3] Restoring Frontend Add-in (gemini-frontend)..."
gcloud run services replace "${CONFIG_DIR}/gemini-frontend-baseline.yaml" \
  --project="${PROJECT_ID}" \
  --region="${REGION}"

echo ""
echo "======================================================================"
echo " SUCCESS: All Cloud Run services restored to baseline snapshot!       "
echo "======================================================================"
echo "Querying deployed service URLs for project ${PROJECT_ID}..."
AUTH_URL=$(gcloud run services describe auth-proxy --project="${PROJECT_ID}" --region="${REGION}" --format="value(status.url)" 2>/dev/null || echo "https://auth-proxy-${PROJECT_ID}.${REGION}.run.app")
BACKEND_URL=$(gcloud run services describe askgemini-proxy --project="${PROJECT_ID}" --region="${REGION}" --format="value(status.url)" 2>/dev/null || echo "https://askgemini-proxy-${PROJECT_ID}.${REGION}.run.app")
FRONTEND_URL=$(gcloud run services describe gemini-frontend --project="${PROJECT_ID}" --region="${REGION}" --format="value(status.url)" 2>/dev/null || echo "https://gemini-frontend-${PROJECT_ID}.${REGION}.run.app")

echo "Live Service URLs:"
echo " - auth-proxy:       ${AUTH_URL}"
echo " - askgemini-proxy:  ${BACKEND_URL}"
echo " - gemini-frontend:  ${FRONTEND_URL}"
echo "======================================================================"

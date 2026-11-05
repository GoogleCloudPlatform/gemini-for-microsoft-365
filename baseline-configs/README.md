# Baseline Cloud Run Configuration Snapshot & Rollback Guide
**Author:** Carlos Augusto, Principal Architect, Google  
**License:** Apache-2.0  

**Snapshot Tags:** `snapshot-wif-baseline` (and `ca-snapshot-wif-baseline`)  
**Snapshot Date:** August 25, 2026  
**Primary GCP Project:** `agentspace-452714` (Project Number: `16933400417`)  
**Target Gemini Enterprise Project:** `agentspace-wif` (Engine ID: `instance-demos1_1774616568648`)  
**Region:** `us-central1`  
**Unified Service Account:** `gemini-office365-sa@agentspace-452714.iam.gserviceaccount.com`

---

## 📦 Snapshot Inventory

This directory contains the exact, declarative YAML definitions of the production/WIF baseline Cloud Run services:

| Service | File | Description |
| :--- | :--- | :--- |
| **Backend Proxy** | [`askgemini-proxy-baseline.yaml`](askgemini-proxy-baseline.yaml) | Node.js 20 Express backend connecting to Discovery Engine StreamAssist on `agentspace-wif` |
| **Auth Gateway** | [`auth-proxy-baseline.yaml`](auth-proxy-baseline.yaml) | Python 3.11 FastAPI gateway handling Entra ID SSO token validation & Google S2S proxying |
| **Frontend Web App** | [`gemini-frontend-baseline.yaml`](gemini-frontend-baseline.yaml) | Nginx container hosting Office 365 add-in assets (PowerPoint, Word, Excel) |

---

## 🔄 Complete Rollback Instructions

If you have made experimental changes or pointed the services to another GCP project/engine, follow these steps to return everything back to this baseline state:

### Step 1: Roll Back the Codebase to Baseline

To return the entire code repository back to the exact baseline state:
```bash
git checkout snapshot-wif-baseline
```
*(Or alternatively: `git checkout ca-snapshot-wif-baseline`)*

---

### Step 2: Roll Back Live Cloud Run Services

You can restore all live Cloud Run services using either the automated script or manual `gcloud` commands:

#### Option A: One-Click Automated Script (Recommended)
From the root of the repository, execute:
```bash
./scripts/rollback_to_baseline.sh
```

#### Option B: Manual Declarative Restore via `gcloud`
Run the declarative `replace` command for each service:

```bash
# 1. Restore Backend Proxy (geminiproxy)
gcloud run services replace baseline-configs/askgemini-proxy-baseline.yaml \
  --project=agentspace-452714 \
  --region=us-central1

# 2. Restore Auth Gateway (authproxy)
gcloud run services replace baseline-configs/auth-proxy-baseline.yaml \
  --project=agentspace-452714 \
  --region=us-central1

# 3. Restore Frontend Add-in (microsoft-addin)
gcloud run services replace baseline-configs/gemini-frontend-baseline.yaml \
  --project=agentspace-452714 \
  --region=us-central1
```

---

### Step 3: Verify Live Endpoints

After rolling back, confirm the services are healthy:

```bash
# Verify Auth Proxy Health
curl -s https://auth-proxy-16933400417.us-central1.run.app/health | jq .

# Verify Backend Proxy Health
curl -s https://askgemini-proxy-16933400417.us-central1.run.app/ | jq .
```

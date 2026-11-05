# GE Office 365 Assistant Auth-Proxy Service
**Author:** Carlos Augusto, Principal Architect, Google  
**License:** Apache-2.0  

A decoupled, lightweight **FastAPI Python microservice** running on Google Cloud Run that validates **Microsoft Entra ID (Azure AD)** Single Sign-On JWT tokens and securely extracts authenticated end-user identities for the **GE Office 365 Assistant Add-On** (PowerPoint, Excel, Word).

---

## Key Features

- 🔐 **Microsoft Entra ID SSO Validation**: Automatically fetches and caches Microsoft public signing keys via JWKS (`https://login.microsoftonline.com/common/discovery/v2.0/keys`) to cryptographically verify RS256 token signatures and expiration timestamps.
- 🎯 **Decoupled Architecture**: Independent of any specific Office host application — handles requests uniformly across PowerPoint, Excel, Word, and web clients.
- 🛡️ **Dedicated Least-Privilege Service Account**: Runs under `gemini-office365-sa`, making authenticated Google Service-to-Service (S2S) IAM calls to private downstream Cloud Run services (`askgemini-proxy`).
- 👤 **End-User Identity Extraction & Forwarding**: Captures `preferred_username`, `email`, `upn`, `tenant_id`, and `oid` and forwards them to downstream engines for auditing, per-user rate limiting, and grounded AI context.
- 📊 **Native GCP Cloud Logging**: Formats all logs into GCP Structured JSON (`severity`, `message`, `structured_context`, `httpRequest`) for native indexing in Cloud Logging with support for a deep `VERBOSE_LOGGING=true` diagnostic mode.
- 🚀 **Cloud Run Ready**: Lightweight container built on `python:3.11-slim` with built-in health probes and CORS support.



---

## Deployment & Microsoft Entra ID Setup

For complete, step-by-step instructions on setting up your Entra ID App Registration, defining scopes, pre-authorizing Microsoft Office applications, and deploying to Google Cloud Run, see:

👉 **[Cloud Run Deployment Guide (Master Runbook)](../docs/deployment/cloud-run-deployment-guide.md)**  
👉 **[Developer Architecture Guide](../docs/architecture/developer-architecture-guide.md)**

---

## API Endpoints & Interactive Documentation

FastAPI automatically generates interactive OpenAPI Swagger documentation for the service:
- **Swagger UI**: `https://<YOUR_CLOUD_RUN_SERVICE_HOSTNAME>/docs`
- **ReDoc UI**: `https://<YOUR_CLOUD_RUN_SERVICE_HOSTNAME>/redoc`
- **OpenAPI JSON**: `https://<YOUR_CLOUD_RUN_SERVICE_HOSTNAME>/openapi.json`

| Method | Endpoint | Auth Required | Description |
| :--- | :--- | :---: | :--- |
| `GET` | `/` | No | Root health check probe for Cloud Run |
| `GET` | `/health` | No | Dedicated health & configuration probe |
| `GET` | `/api/auth/me` | **Yes (Bearer)** | Returns extracted claims and user profile |
| `POST` | `/api/auth/validate` | **Yes (Bearer)** | Explicit JWT validation endpoint |
| `POST` | `/askGeminiEnterprise` | **Yes (Bearer)** | Primary Office 365 add-in proxy endpoint |
| `POST` | `/api/gemini/chat` | **Yes (Bearer)** | Structured chat endpoint (alias to `/askGeminiEnterprise`) |
| `GET` | `/docs` | No | Interactive Swagger UI API documentation |
| `GET` | `/redoc` | No | Alternative ReDoc API documentation |


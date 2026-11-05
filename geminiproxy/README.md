# Gemini Enterprise Backend Proxy for Microsoft 365
**Author:** Carlos Augusto, Principal Architect, Google  
**License:** Apache-2.0  

A scalable, enterprise-grade Google Cloud Function (Gen 2) and Cloud Run microservice that bridges Microsoft 365 (Word, PowerPoint, Excel) with Google Cloud Vertex AI generative models.

---

## 🌟 Capabilities

- **Grounded Enterprise Intelligence:** Direct integration with Vertex AI Search Datastores for factual, enterprise-grounded document analysis and Q&A over quarterly reports, policies, and research filings.
- **Multimodal Visual Chart Generation:** Automatically generates high-resolution, flat 2D vector financial charts and infographics via `gemini-2.5-flash-image` (Nano Banana) and embeds them as base64 images.
- **Multi-Turn Session Memory:** In-memory caching and session context management for continuous multi-turn conversations.
- **Host-Optimized Output:** Built-in system instructions tuned specifically for Microsoft Word documents, PowerPoint widescreen presentations, and Excel spreadsheets.

---

## 🏗️ Architecture

```mermaid
graph LR
    O365[Microsoft 365 Add-in (Word / PPT / Excel)] -->|HTTPS POST JSON| Proxy[Gemini O365 Proxy Cloud Function]
    
    subgraph Google Cloud Vertex AI
        Proxy -->|Grounded Search| Search[Vertex AI Search Datastore]
        Proxy -->|Generative Text| Flash[Gemini 2.5 Flash]
        Proxy -->|Visual Charts| Imagen[Gemini 2.5 Flash Image]
    end
```

---

## 🚀 Quick Start & Local Setup

### 1. Prerequisites
- [Node.js](https://nodejs.org/) (v20+ recommended)
- [Google Cloud SDK (`gcloud`)](https://cloud.google.com/sdk/docs/install)
- A Google Cloud Project with Vertex AI API enabled (`aiplatform.googleapis.com`)

### 2. Environment Configuration
Copy the template configuration file:
```bash
cp .env.example .env
```

Configure the environment variables in `.env`:
```bash
GE_GCP_PROJECT_ID=your-google-cloud-project-id
GE_GCP_LOCATION=global
GEMINI_ENTERPRISE_APP_ID=instance-demos1_1774616568648
BACKEND_MODE=streamassist
ENTERPRISE_COLLECTION_ID=default_collection
ENTERPRISE_ASSISTANT_ID=default_assistant
ALLOW_SERVICE_ACCOUNT_FALLBACK=false
```

### 3. Environment Variables Reference

| Variable Name | Required / Optional | Default Value | Example Value | Description & Impact |
| :--- | :---: | :---: | :--- | :--- |
| `GE_GCP_PROJECT_ID` | **Required** | `process.env.GCP_PROJECT_ID` | `agentspace-wif` / `your-gemini-enterprise-project-id` | Google Cloud Project ID hosting Discovery Engine & Vertex AI. |
| `GEMINI_ENTERPRISE_APP_ID` | **Required** *(in `streamassist` mode)* | `""` | `instance-demos1_1774616568648` | Gemini Enterprise Search / Assist Engine ID. |
| `BACKEND_MODE` | Optional | `streamassist` | `streamassist` | Execution mode (`streamassist` for Discovery Engine, `vertex` for direct Vertex AI). |
| `GE_GCP_LOCATION` | Optional | `global` | `global` / `us` | Discovery Engine collection/engine location (`global`, `us`, `eu`). |
| `STREAM_ASSIST_ENDPOINT_LOCATION` | Optional | `global` | `global` / `us` | API endpoint subdomain routing prefix (`global` -> `discoveryengine.googleapis.com`). |
| `ENTERPRISE_COLLECTION_ID` | Optional | `default_collection` | `default_collection` | Discovery Engine collection resource ID. |
| `ENTERPRISE_ASSISTANT_ID` | Optional | `default_assistant` | `default_assistant` | Discovery Engine assistant resource ID. |
| `ALLOW_SERVICE_ACCOUNT_FALLBACK` | Optional | `false` | `false` (WIF) / `true` (GSuite) | When `true`, falls back to Service Account ADC if user token is absent. |
| `GCP_REGION` | Optional | `us-central1` | `us-central1` | GCP region for direct Vertex AI client (used in `vertex` mode). |
| `GEMINI_MODEL` | Optional | `gemini-2.5-flash` | `gemini-2.5-flash` | Generative text foundation model name for direct Vertex AI calls. |
| `GEMINI_IMAGE_MODEL` | Optional | `gemini-2.5-flash-image` | `gemini-2.5-flash-image` | Multimodal visual generation model for charts and diagrams. |
| `VERTEX_DATASTORE_ID` | Optional | `""` | `""` | Full resource name of Vertex AI Search Datastore in direct `vertex` mode. |
| `PORT` | Optional | `8080` | `8080` | Container listen port (injected by Cloud Run runtime). |

### 4. Install Dependencies & Run Locally
```bash
npm install
npm start
```
The server will start locally at `http://localhost:8080`.

---

## ☁️ Deployment to Google Cloud Run

```bash
gcloud run deploy askgemini-proxy \
  --source . \
  --project YOUR_GCP_PROJECT_ID \
  --region us-central1 \
  --memory 1Gi \
  --service-account gemini-office365-sa@YOUR_GCP_PROJECT_ID.iam.gserviceaccount.com \
  --set-env-vars "\
GE_GCP_PROJECT_ID=YOUR_PROJECT_ID,\
GEMINI_ENTERPRISE_APP_ID=YOUR_APP_ID,\
BACKEND_MODE=streamassist,\
GE_GCP_LOCATION=global,\
ALLOW_SERVICE_ACCOUNT_FALLBACK=false" \
  --no-allow-unauthenticated \
  --quiet
```

---

## 🔒 Security & Privacy Best Practices
- **No Hardcoded Secrets:** All project IDs, regions, and datastore configurations are strictly managed via environment variables.
- **Enterprise Grounding:** Only authenticated enterprise Vertex AI datastores are queried.
- **CORS Configured:** Permissive for enterprise add-in webviews while protecting internal execution pipelines.

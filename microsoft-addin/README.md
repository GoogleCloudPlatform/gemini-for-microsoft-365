# Gemini for Microsoft 365 (Word, PowerPoint, Excel)
**Author:** Carlos Augusto, Principal Architect, Google  
**License:** Apache-2.0  

An enterprise-ready Microsoft 365 Add-in that brings the power of **Google Cloud Vertex AI (Gemini 2.5 Flash & Multimodal Visual Generation)** directly into Microsoft Word, PowerPoint, and Excel.

---

## 🌟 Key Features

### 📄 Microsoft Word (`WordAdapter`)
- **Interactive Document Q&A:** Full document summarization, risk analysis, and action item extraction.
- **In-Document Execution:** Type `@gemini <prompt>` directly on any document line to generate content inline.
- **Live Text Transformation:** Highlight text to rewrite, summarize, professionalize, or convert into executive tables.

### 📊 Microsoft PowerPoint (`PPTAdapter`)
- **Executive Deck Generator:** Automatically converts unstructured text and financial briefings into structured multi-slide executive decks.
- **AI Visual Chart Injection:** Generates and embeds high-resolution 2D vector charts and comparison graphics natively onto slides.
- **Dual-Pipeline macOS Compatibility:** Optimized with `PowerPoint.run` slide selection synchronization and native `Office.context.document.setSelectedDataAsync` injection to bypass macOS WKWebView sandbox restrictions.
- **Intelligent Layouts:** Dedicated center-aligned Title slides, widescreen executive layouts (880px), and split-column visual layouts (420px text + 440x340px image).

### 📈 Microsoft Excel (`ExcelAdapter`)
- **Spreadsheet Intelligence:** Summarize worksheets, extract data patterns, and generate executive KPI metric cards.
- **Anomaly & Formula Risk Detection:** Identifies outliers, formatting inconsistencies, and formula risks.
- **In-Cell Triggers:** Type `@gemini <prompt>` or cell selection analysis.

---

## 🏗️ Architecture

```mermaid
graph TD
    User([User in Word / PPT / Excel]) --> Taskpane[Gemini Taskpane UI]
    Taskpane --> HostAdapter[Host Adapter Factory]
    
    HostAdapter -->|Word.run| Word[Word Host Adapter]
    HostAdapter -->|PowerPoint.run + Common API| PPT[PowerPoint Host Adapter]
    HostAdapter -->|Excel.run| Excel[Excel Host Adapter]
    
    Taskpane -->|HTTPS JSON| Proxy[Gemini O365 Backend Proxy]
    Proxy -->|Vertex AI Grounding| Vertex[Gemini 2.5 Flash + Search Datastore]
    Proxy -->|Image Generation| Imagen[Gemini 2.5 Flash Image]
```

---

## 🚀 Quick Start & Local Development

### 1. Prerequisites
- [Node.js](https://nodejs.org/) (v20+ recommended)
- Microsoft Office 365 Desktop (Word, PowerPoint, Excel) or Office Online

### 2. Environment Configuration
Copy the configuration template:
```bash
cp .env.example .env
```
Update `.env` with your `auth-proxy` Cloud Run endpoint:
```bash
GEMINI_PROXY_URL=https://auth-proxy-1062675944253.us-central1.run.app/askGeminiEnterprise
```

### 3. Environment Variables Reference

| Variable Name | Required / Optional | Default Value | Example Value | Description & Impact |
| :--- | :---: | :---: | :--- | :--- |
| `GEMINI_PROXY_URL` | Optional | `""` *(auto-derived from origin or `/api/config`)* | `https://auth-proxy-1062675944253.us-central1.run.app/askGeminiEnterprise` | URL of the `auth-proxy` gateway endpoint. Injected during webpack build or resolved dynamically at runtime. |
| `PORT` | Optional | `80` *(Cloud Run Nginx)* / `3000` *(dev-server)* | `80` | Web server HTTP listen port. |

### 4. Install & Build
```bash
npm install
npm run build
```

### 5. Local Development Server
```bash
npm run dev-server
```
Starts the local development server with self-signed HTTPS certificates at `https://localhost:3000`.

---

## ☁️ Production Deployment & Sideloading

### Deploy Frontend to Google Cloud Run
```bash
npm run build
gcloud run deploy gemini-frontend \
  --source . \
  --region us-central1 \
  --allow-unauthenticated \
  --service-account gemini-office365-sa@YOUR_GCP_PROJECT_ID.iam.gserviceaccount.com \
  --project YOUR_GCP_PROJECT_ID \
  --set-env-vars "GEMINI_PROXY_URL=https://auth-proxy-XXXXXXXXXX.us-central1.run.app/askGeminiEnterprise" \
  --quiet
```

### Sideloading into Microsoft Office
- **For Track 1 (WIF)**: Sideload `manifest-wif.xml`
- **For Track 2 (GSuite)**: Sideload `manifest-gsuite.xml`
- **On macOS**: `./scripts/sideload_mac.sh <manifest-file>`
- **On Web**: `Insert` > `Add-ins` > `Upload My Add-in` > Select the XML manifest.

---

## 🔒 Security & Compliance
- **No Hardcoded Secrets:** Production endpoints and credentials are isolated through environment variables.
- **Enterprise Isolation:** Documents and slide data remain strictly within your Google Cloud enterprise security boundary.

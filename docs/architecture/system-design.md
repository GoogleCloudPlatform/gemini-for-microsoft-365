# Architecture & System Design: Gemini for Microsoft 365
**Author:** Carlos Augusto, Principal Architect, Google  
**License:** Apache-2.0  

---

## 📑 Executive Summary

**Gemini for Microsoft 365** is an enterprise-grade AI integration platform that connects Microsoft Office client applications (**Word**, **PowerPoint**, **Excel**, and **Outlook**) with **Google Cloud Vertex AI**.

The system combines:
1. **Factual Enterprise Grounding** via **Vertex AI Search Datastores** (RAG over enterprise 10-K/10-Q filings, policies, and internal docs).
2. **Generative Intelligence** via **Gemini 2.5 Flash** for multi-turn conversational reasoning and content synthesis.
3. **Multimodal Visual Synthesis** via **Gemini 2.5 Flash Image** (Nano Banana) for dynamic 2D flat vector financial charts and infographics.
4. **Native Host Adapters** for Microsoft Word (inline `@gemini` insertion, text transformation), PowerPoint (multi-slide executive decks with native macOS WKWebView dual-pipeline image rendering), Excel (range analysis, anomaly detection, KPI metric cards), and Outlook (`OutlookAdapter.js` for email thread summarization, tone coaching, grounded reply drafting, meeting briefings, and action item extraction).

> 📖 **Decoupled Auth & S2S System Architecture:** For a deep dive into the decoupled Microsoft Entra ID authentication gateway, Google Cloud Service-to-Service IAM model, and implementation blueprints, see [developer-architecture-guide.md](developer-architecture-guide.md) and [../deployment/cloud-run-deployment-guide.md](../deployment/cloud-run-deployment-guide.md).

---

## 🏛️ Enterprise Multi-Tier Architecture (Milestones 2 & 3)

The enterprise deployment separates concerns into dedicated, independently scalable, least-privilege tiers across **Microsoft 365**, **Microsoft Entra ID**, and **Google Cloud Platform**:

```mermaid
graph TB
    subgraph OfficeClientTier ["1. Microsoft 365 Office Client Tier"]
        WordApp["Microsoft Word<br/>(Desktop & Web)"]
        PPTApp["Microsoft PowerPoint<br/>(Desktop & Web)"]
        ExcelApp["Microsoft Excel<br/>(Desktop & Web)"]
        OfficeSSO["Office.js SSO Runtime<br/>(Office.auth.getAccessToken)"]
        TaskpaneUI["Add-in Taskpane UI<br/>(HTML5 / Vanilla JS / CSS)"]
        GoogleDialog["Office Dialog API<br/>(google-auth.html / google-callback.html)"]
        
        WordApp <--> TaskpaneUI
        PPTApp <--> TaskpaneUI
        ExcelApp <--> TaskpaneUI
        TaskpaneUI <--> OfficeSSO
        TaskpaneUI <--> GoogleDialog
    end

    subgraph IdentityTier ["2. Microsoft Entra ID & Google Identity Tier"]
        EntraApp["Entra ID App Registration<br/>(App ID: e871aa77-...)"]
        AppUri["Application ID URI<br/>api://gemini-frontend-...run.app/e871aa77-..."]
        GoogleOAuth["Google Cloud OAuth 2.0 Web Client<br/>(Client ID: YOUR_GOOGLE_PROJECT_NUMBER-...apps.googleusercontent.com)"]
        
        OfficeSSO -->|"1. Silent SSO Token Request"| EntraApp
        EntraApp -->|"2. Signed Microsoft JWT"| OfficeSSO
        GoogleDialog -->|"3-Legged Consent Flow"| GoogleOAuth
        GoogleOAuth -->|"Google User OAuth Token ya29..."| GoogleDialog
    end

    subgraph GCPInfrastructure ["3. Google Cloud Platform (Project: agentspace-452714)"]
        FrontendRun["gemini-frontend (Cloud Run)<br/>Nginx Static Assets & Taskpane<br/>(Public HTTPS)"]
        
        subgraph AuthGatewayTier ["Auth & Token Translation Gateway"]
            AuthProxy["auth-proxy (Cloud Run)<br/>Python 3.11 / FastAPI<br/>Runtime SA: gemini-office365-sa"]
            ConfigEndpoint["Dynamic Config Gateway<br/>GET /api/config"]
            IdpDiscovery["Dynamic IdP Auto-Discovery<br/>GET /v1/.../aclConfig"]
        end
        
        subgraph PrivateBackendTier ["Private Core Inference Backend (Private S2S Only)"]
            BackendProxy["askgemini-proxy (Cloud Run)<br/>Node.js 20 Express Microservice<br/>(--no-allow-unauthenticated)"]
        end
        
        TaskpaneUI -->|"Loads Static Assets"| FrontendRun
        TaskpaneUI -->|"Fetches Dynamic Config"| ConfigEndpoint
        TaskpaneUI -->|"3. POST /askGeminiEnterprise + Entra Bearer JWT + X-End-User-Google-Token"| AuthProxy
        AuthProxy ---|"4. Discovers IdP Type (GSUITE vs THIRD_PARTY)"| IdpDiscovery
        AuthProxy -->|"5. Forward Request + Google S2S IAM Token + User Tokens"| BackendProxy
    end

    subgraph GeminiEnterpriseTier ["4. Google Cloud Discovery Engine / Gemini Enterprise"]
        StreamAssist["Discovery Engine API<br/>POST /v1/.../engines/gemini-enterprise-dummy-ap/assistants/default_assistant:streamAssist"]
        EnterpriseCorpus["Enterprise Grounding Corpus<br/>(Google Drive, GCS, Spanner, BigQuery)"]
        
        BackendProxy -->|"6. StreamAssist Request with User Context (Authorization: Bearer <ya29...>) + toolsSpec"| StreamAssist
        StreamAssist ---|"7. Semantic Grounding & User-Level Drive Search"| EnterpriseCorpus
        StreamAssist -.->|"8. Server-Sent Events (SSE) Stream"| BackendProxy
        BackendProxy -.->|"9. SSE Chunks"| AuthProxy
        AuthProxy -.->|"10. Stream to Client"| TaskpaneUI
    end

    style OfficeClientTier fill:#e8f0fe,stroke:#1a73e8,stroke-width:2px;
    style IdentityTier fill:#f3e8fd,stroke:#7b1fa2,stroke-width:2px;
    style GCPInfrastructure fill:#e6f4ea,stroke:#137333,stroke-width:2px;
    style AuthGatewayTier fill:#e0f2f1,stroke:#00796b,stroke-width:2px;
    style PrivateBackendTier fill:#fff3e0,stroke:#e65100,stroke-width:2px;
    style GeminiEnterpriseTier fill:#fef7e0,stroke:#f9ab00,stroke-width:2px;
```

---

## 🔄 End-to-End Execution Flows

The platform supports two distinct identity execution paths depending on enterprise IdP configuration.

### Track 1: Workforce Identity Federation (WIF) Silent Token Exchange

In Track 1, users experience seamless Single Sign-On with zero Google login prompts. Their Microsoft Entra ID JWT is exchanged on the fly for a Google STS federated token via RFC 8693.

```mermaid
sequenceDiagram
    autonumber
    actor User as Corporate User
    participant Office as Office Taskpane UI
    participant Entra as Microsoft Entra ID
    participant Auth as Cloud Run: auth-proxy
    participant STS as Google STS (sts.googleapis.com)
    participant Meta as GCP Metadata Server
    participant Backend as Cloud Run: askgemini-proxy
    participant Engine as Gemini / StreamAssist

    User->>Office: Submits prompt: "Summarize Q3 earnings"
    
    rect rgb(240, 248, 255)
        Note over Office,Entra: Phase 1: Silent Entra ID Token Acquisition
        Office->>Entra: Office.auth.getAccessToken({ forMSGraphAccess: false })
        Entra-->>Office: Returns Microsoft Entra ID JWT Bearer Token
    end

    rect rgb(255, 250, 235)
        Note over Office,Auth: Phase 2: Auth Gateway Validation & STS Exchange
        Office->>Auth: POST /askGeminiEnterprise<br/>Authorization: Bearer [Entra_JWT]<br/>Body: { prompt, history, sessionId }
        Auth->>Entra: Fetch/Match RS256 Public Key via JWKS cache
        Auth->>Auth: Verify signature, expiration (exp), audience (aud), tenant (tid)
        Auth->>Auth: Extract claims (user_id, email, name, tenant_id, oid)
        Auth->>STS: POST /v1/token (RFC 8693 Token Exchange)<br/>subject_token=[Entra_JWT], audience=[//iam.googleapis.com/workforcePools/...]
        STS-->>Auth: Returns Federated Google Access Token (ya29...)
    end

    rect rgb(235, 255, 235)
        Note over Auth,Backend: Phase 3: Google S2S IAM Token Exchange & Forwarding
        Auth->>Meta: GET /instance/service-accounts/default/identity?audience=https://askgemini-proxy-...
        Meta-->>Auth: Returns short-lived Google OIDC ID Token
        Auth->>Backend: POST /askGeminiEnterprise<br/>Authorization: Bearer [Google_ID_Token]<br/>Headers: X-End-User-Id, X-End-User-Email, X-End-User-Google-Token: ya29...<br/>Body: { prompt, sessionId, userPseudoId, authenticatedUser }
    end

    rect rgb(255, 240, 245)
        Note over Backend,Engine: Phase 4: Grounded Execution & Attribution
        Backend->>Engine: StreamAssist API Call with user attribution & STS token
        Engine-->>Backend: Grounded answer + citations
        Backend-->>Auth: HTTP 200 OK with AI result and citations
    end

    Auth-->>Office: HTTP 200 OK with authenticated AI payload
    Office-->>User: Renders formatted grounded response & citations
```

#### 💡 Deep Dive: How WIF Token Exchange & StreamAssist Grounding Work

```
[Microsoft Office 365 Client]
          │
          │ 1. Office.auth.getAccessToken()
          ▼
   [Microsoft Entra ID]
          │
          │ 2. Returns Entra ID JWT (claims: upn, email, oid, tid)
          ▼
   [Cloud Run: auth-proxy]
          │
          │ 3. POST https://sts.googleapis.com/v1/token (RFC 8693)
          │    subject_token = Entra_JWT
          │    audience = //iam.googleapis.com/.../workforcePools/{POOL}/providers/{PROVIDER}
          ▼
   [Google Security Token Service (STS)]
          │
          │ 4. Maps Microsoft UPN -> Google Principal & returns Google STS Token (ya29...)
          ▼
   [Cloud Run: auth-proxy]
          │
          │ 5. POST /askGeminiEnterprise to askgemini-proxy
          │    Authorization: Bearer <Google_S2S_IAM_Token> (Service Account token)
          │    X-End-User-Google-Token: ya29... (The Federated User Token)
          ▼
   [Cloud Run: askgemini-proxy]
          │
          │ 6. POST https://discoveryengine.googleapis.com/.../assistants/default_assistant:streamAssist
          │    Authorization: Bearer ya29... ◄── [FEDERATED END-USER TOKEN]
          ▼
[Google Discovery Engine / Gemini Enterprise]
```

1. **The Token Translation Chain (RFC 8693):**
   - In Track 1, enterprise users log into Microsoft 365 using corporate Entra ID accounts (`alex@contoso.com`). They do not maintain separate Google credentials.
   - `auth-proxy` validates the Entra JWT signature via Microsoft JWKS, then calls Google Security Token Service (`sts.googleapis.com/v1/token`) presenting the Entra JWT as a `subject_token`.
   - Google STS validates the Microsoft signature against Entra ID's OpenID discovery metadata, evaluates the Workforce Pool attribute mapping (`assertion.upn -> google.subject`), and returns a short-lived **Google STS Federated Access Token (`ya29...`)**.

2. **What Bearer Token is Passed to `streamAssist`?**
   - When `askgemini-proxy` invokes the Discovery Engine `streamAssist` API:
     ```http
     POST https://discoveryengine.googleapis.com/v1alpha/projects/PROJECT_ID/locations/global/collections/default_collection/engines/ENGINE_ID/assistants/default_assistant:streamAssist
     Authorization: Bearer ya29...
     X-Goog-User-Project: PROJECT_ID
     Content-Type: application/json
     ```
   - The token in the `Authorization: Bearer` header is the **Google STS Federated Access Token (`ya29...`)**.
   - It represents the authenticated workforce principal:
     `principal://iam.googleapis.com/locations/global/workforcePools/{POOL}/subject/alex@contoso.com`

3. **How Discovery Engine / StreamAssist Uses the End-User Token to Query:**
   - **Identity & License Verification:** Google Cloud's API gateway parses the workforce principal from the STS token, validates that the user holds `roles/discoveryengine.viewer` or `roles/discoveryengine.editor`, and consumes an assigned Gemini Enterprise seat license.
   - **User-Level ACL Grounding (Zero-Trust Retrieval):** When Discovery Engine searches connected enterprise data sources (SharePoint, Jira, Confluence, Salesforce, Google Drive, or indexed datastores), every document in the search index carries an Access Control List (ACL). Discovery Engine filters search results against the caller's verified identity before feeding retrieved context into Gemini. If `alex@contoso.com` lacks read access to a document, that document is filtered out prior to prompt assembly.
   - **Session Isolation & Audit Trail:** Discovery Engine isolates conversation state by workforce principal and logs queries with attributed user telemetry in Cloud Logging.

---

### Track 2: Cloud Identity 3-Legged Google User OAuth & S2S IAM

The sequence below illustrates the complete token acquisition, dynamic configuration retrieval, 3-legged Google user authorization, service-to-service IAM minting, and streaming grounding lifecycle:

```mermaid
sequenceDiagram
    autonumber
    actor User as Enterprise User (scim@...)
    participant PPT as Microsoft Office Client (PowerPoint/Word/Excel)
    participant FE as gemini-frontend (Cloud Run)
    participant Entra as Microsoft Entra ID (IdP)
    participant AP as auth-proxy Gateway (Cloud Run)
    participant Dialog as Office Dialog API (google-auth.html)
    participant GoogleAuth as accounts.google.com
    participant Callback as google-callback.html
    participant GP as askgemini-proxy (Private Cloud Run)
    participant GE as Gemini Enterprise (streamAssist)

    User->>PPT: Opens 'Gemini Assistant' in Office Ribbon
    PPT->>FE: Fetches taskpane.html, JS bundles & assets
    FE-->>PPT: Renders Taskpane UI (Status: Connecting...)
    
    rect rgb(240, 248, 255)
        Note over PPT,Entra: Phase 1: Microsoft 365 Office SSO Token Acquisition
        PPT->>PPT: Office.auth.getAccessToken({ forMSGraphAccess: false })
        PPT->>Entra: Requests access token for api://gemini-frontend-.../e871aa77-...
        Entra-->>PPT: Returns signed Microsoft Entra ID JWT (claims: upn, email, tid, aud)
        PPT->>PPT: Updates UI header: scim@jeansson.demo.altostrat.com [Logged in]
    end

    rect rgb(230, 245, 255)
        Note over PPT,AP: Phase 2: Dynamic Config & Google User Authorization (Cloud Identity Mode)
        PPT->>AP: GET /api/config
        AP-->>PPT: Returns { google_oauth_client_id: "YOUR_GOOGLE_PROJECT_NUMBER-...", user_auth_mode: "cloud_identity" }
        User->>PPT: Clicks "📁 Connect Drive"
        PPT->>Dialog: Office.context.ui.displayDialogAsync(google-auth.html?client_id=...&login_hint=scim@...)
        Dialog->>GoogleAuth: Interactive Consent for drive.readonly + cloud-platform scopes
        User->>GoogleAuth: Grants Consent
        GoogleAuth->>Callback: Redirects to /google-callback.html#access_token=ya29...
        Callback->>PPT: Office.context.ui.messageParent({ google_token: "ya29..." })
        Callback->>Callback: Closes dialog
        PPT->>PPT: Caches token (Status: ✅ Drive Linked)
    end

    User->>PPT: Submits prompt: "Do you see documents referring api://...?"
    
    rect rgb(240, 255, 240)
        Note over PPT,AP: Phase 3: Gateway Verification & Header Pass-Through
        PPT->>AP: POST /askGeminiEnterprise (Headers: Authorization: Bearer <Entra_JWT>, X-End-User-Google-Token: ya29...)
        AP->>AP: Verifies Microsoft signature via JWKS & validates aud/tid
        AP->>AP: Attaches verified user claims (X-End-User-Email, X-End-User-ID) and forwards X-End-User-Google-Token
    end

    rect rgb(255, 248, 240)
        Note over AP,GP: Phase 4: Service-to-Service Google IAM Authentication
        AP->>AP: Fetches Google OIDC ID token for audience: https://askgemini-proxy...
        AP->>GP: POST /askGeminiEnterprise (Headers: Authorization: Bearer <Google_IAM_Token>, X-End-User-Google-Token: ya29...)
        GP->>GP: Cloud Run IAM validates gemini-office365-sa has roles/run.invoker
    end

    rect rgb(255, 255, 240)
        Note over GP,GE: Phase 5: Gemini Enterprise Grounded Stream
        GP->>GE: POST /v1alpha/.../assistants/default_assistant:streamAssist (Headers: Authorization: Bearer <ya29...>, toolsSpec: { vertexAiSearchSpec: {} })
        GE->>GE: Grounded retrieval across user-accessible Google Drive files
        GE-->>GP: Streams SSE chunks with answers, citations, and grounding references
        GP-->>AP: Streams SSE response
        AP-->>PPT: Streams SSE response
    end

    PPT->>PPT: Renders answer incrementally in Taskpane UI
    PPT->>User: Displays formatted markdown, citations, and Insert Slide action buttons
```

#### 💡 Deep Dive: How Cloud Identity / Google Workspace OAuth & StreamAssist Grounding Work

```
[Microsoft Office 365 Client]
          │
          │ 1. Office.context.ui.displayDialogAsync(google-auth.html)
          ▼
   [accounts.google.com]
          │
          │ 2. Interactive Consent for drive.readonly + cloud-platform scopes
          │    User logs in with Google Workspace identity (e.g. scim@domain.com)
          ▼
   [Office Dialog Callback]
          │
          │ 3. messageParent({ google_token: "ya29..." }) -> Cached in client
          ▼
   [Office Taskpane Client]
          │
          │ 4. POST /askGeminiEnterprise
          │    Authorization: Bearer <Entra_JWT> (Perimeter Security)
          │    X-End-User-Google-Token: ya29... (Google User Token)
          ▼
   [Cloud Run: auth-proxy]
          │
          │ 5. Validates Entra JWT signature & forwards ya29... via S2S IAM channel
          ▼
   [Cloud Run: askgemini-proxy]
          │
          │ 6. POST https://discoveryengine.googleapis.com/.../assistants/default_assistant:streamAssist
          │    Authorization: Bearer ya29... ◄── [GOOGLE WORKSPACE USER TOKEN]
          │    toolsSpec: { vertexAiSearchSpec: {} }
          ▼
[Google Discovery Engine / Gemini Enterprise]
```

1. **The Dual-Boundary Authentication Model:**
   - **Perimeter Gatekeeper (Microsoft Entra ID):** Entra ID SSO ensures that only authorized corporate Microsoft 365 users can connect to the Cloud Run gateway.
   - **Data Grounding Gatekeeper (Google Cloud Identity / Workspace):** The 3-legged Google OAuth token authorizes Discovery Engine to search and read the specific user's Google Drive files and Google Workspace data.

2. **What Bearer Token is Passed to `streamAssist`?**
   - When `askgemini-proxy` invokes Discovery Engine `streamAssist`, it supplies the Google OAuth access token acquired from the user consent dialog:
     ```http
     POST https://discoveryengine.googleapis.com/v1alpha/projects/PROJECT_ID/locations/global/collections/default_collection/engines/ENGINE_ID/assistants/default_assistant:streamAssist
     Authorization: Bearer ya29...
     X-Goog-User-Project: PROJECT_ID
     Content-Type: application/json
     ```
   - The token carries the `https://www.googleapis.com/auth/drive.readonly` and `https://www.googleapis.com/auth/cloud-platform` scopes for `scim@domain.com`.

3. **How Discovery Engine / StreamAssist Uses the End-User Token to Query:**
   - **Delegated Google Drive Grounding:** Discovery Engine uses the user's `ya29...` token to execute on-the-fly semantic queries against the user's personal Google Drive, shared corporate drives, and shared-with-me documents.
   - **Native Google Workspace ACL Enforcement:** Discovery Engine only indexes and retrieves documents where `scim@domain.com` is an authorized viewer/editor in Google Drive. Private files belonging to other users or restricted teams remain completely invisible.
   - **User Attribution & History:** `askgemini-proxy` binds conversation sessions to `userPseudoId: scim@domain.com`, preserving chat history across Office sessions while maintaining strict multi-tenant isolation.

---

### 📊 Token & Security Comparison: WIF vs. Cloud Identity

| Dimension | Track 1: Workforce Identity Federation (WIF) | Track 2: Cloud Identity / Google Workspace |
| :--- | :--- | :--- |
| **End-User Login Experience** | **100% Silent SSO** (zero Google login popups). | **1-Click Google Sign-In** via Office Dialog API. |
| **Token in Office Add-in** | Microsoft Entra ID JWT only. | Microsoft Entra ID JWT + Google User OAuth Token (`ya29...`). |
| **Token Minting Mechanism** | Google STS RFC 8693 token exchange on `auth-proxy`. | Standard Google OAuth 2.0 Authorization Code flow. |
| **Bearer Token to `streamAssist`** | Google STS Federated Access Token (`ya29...`). | Google User 3-Legged OAuth Access Token (`ya29...`). |
| **Caller Identity in Google Cloud** | `principal://iam.googleapis.com/.../workforcePools/...` | `scim@company.com` (Google Workspace user account). |
| **Target Data Sources** | Datastores, SharePoint, Jira, Confluence, Salesforce, ACL-indexed repositories. | Personal & Shared Google Drive, Google Workspace docs, and Datastores. |
| **Enterprise Identity Authority** | Microsoft Entra ID (Single Source of Truth). | Dual Authority: Entra ID (Office client) + Cloud Identity (Google services). |

---

## 🔒 Key Architectural Principles & Invariants

### 1. The Office.js SSO Domain-Matching Rule
Microsoft Office Add-in SSO requires that the **`<SourceLocation>`** domain, the **`<WebApplicationInfo><Resource>`**, and the **Entra ID Application ID URI** match identically:
$$\text{Domain in } \langle\text{SourceLocation}\rangle \equiv \text{Domain in } \langle\text{Resource}\rangle \equiv \text{Entra ID Application ID URI}$$

* **Frontend Hosting URL:** `https://gemini-frontend-16933400417.us-central1.run.app/taskpane.html`
* **Manifest `<Resource>`:** `api://gemini-frontend-16933400417.us-central1.run.app/b990d644-e47b-4575-97b3-2067c488042b`
* **Entra ID Application ID URI:** `api://gemini-frontend-16933400417.us-central1.run.app/b990d644-e47b-4575-97b3-2067c488042b`

If the Application ID URI points to a backend or proxy domain (`auth-proxy`), Office detects a domain mismatch and immediately triggers **Office SSO Error 13007: Invalid resource Url specified in the manifest**.

### 2. Pre-Authorized Client Applications in Entra ID
To enable silent SSO without consent prompts across all Office platforms, the following 3 client IDs must be pre-authorized under **Expose an API**:
1. `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office on the Web - Word/PPT/Excel Online)
2. `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Office on the Web Companion / Outlook)
3. `00000002-0000-0ff1-ce00-000000000000` (Microsoft Office Desktop Client across macOS & Windows)

### 3. Decoupled Token Translation & Least Privilege
* **No Direct Internet Access to Core Backend:** `askgemini-proxy` is locked down with `--no-allow-unauthenticated`.
* **Runtime Service Account Identity:** `auth-proxy` runs as `gemini-office365-sa@agentspace-452714.iam.gserviceaccount.com` and only holds:
  - `roles/logging.logWriter` (Structured Cloud Logging)
  - `roles/discoveryengine.viewer` (Dynamic `aclConfig` auto-discovery)
  - `roles/run.invoker` on Cloud Run service `askgemini-proxy` (Private S2S communication)

---

```mermaid
graph TB
    subgraph ClientTier ["Microsoft 365 Client Tier (Desktop & Web)"]
        WordApp["Microsoft Word<br/>(WordAdapter)"]
        PPTApp["Microsoft PowerPoint<br/>(PPTAdapter)"]
        ExcelApp["Microsoft Excel<br/>(ExcelAdapter)"]
        TaskpaneUI["Add-in Taskpane Webview<br/>(Vanilla JS + Fluent UI CSS)"]
        
        WordApp <--> TaskpaneUI
        PPTApp <--> TaskpaneUI
        ExcelApp <--> TaskpaneUI
    end

    subgraph GCPInfrastructure ["Google Cloud Platform (genai-demo-catalog / us-central1)"]
        FrontendRun["Cloud Run: gemini-frontend<br/>(Nginx Container / Port 80)"]
        ProxyFunction["Cloud Function Gen 2: askGemini<br/>(Node.js 20 Microservice)"]
        
        TaskpaneUI -->|"1. Loads Static Assets & JS Bundle"| FrontendRun
        TaskpaneUI -->|"2. HTTPS POST /askGemini JSON"| ProxyFunction
    end

    subgraph VertexAIEngine ["Google Cloud Vertex AI Enterprise Backend"]
        SearchDS["Vertex AI Search Datastore<br/>(Enterprise 10-K/10-Q Docs in GCS)"]
        FlashModel["Gemini 2.5 Flash<br/>(Grounded Generative Text Model)"]
        ImageModel["Gemini 2.5 Flash Image<br/>(Nano Banana Visual Chart Generator)"]
        
        ProxyFunction -->|"A. Grounded RAG Query"| FlashModel
        FlashModel ---|"B. Semantic Retrieval & Citations"| SearchDS
        ProxyFunction -->|"C. Balanced Regex Extractor"| ImageModel
    end

    subgraph OutputPipeline ["Client Rendering & Document Injection"]
        ProxyFunction -->|"3. Grounded Text + Base64 PNGs"| TaskpaneUI
        TaskpaneUI -->|"Word.run / OOXML"| WordApp
        TaskpaneUI -->|"PowerPoint.run + setSelectedDataAsync"| PPTApp
        TaskpaneUI -->|"Excel.run"| ExcelApp
    end

    style ClientTier fill:#e8f0fe,stroke:#1a73e8,stroke-width:2px;
    style GCPInfrastructure fill:#e6f4ea,stroke:#137333,stroke-width:2px;
    style VertexAIEngine fill:#fef7e0,stroke:#f9ab00,stroke-width:2px;
    style OutputPipeline fill:#fce8e6,stroke:#c5221f,stroke-width:2px;
```

---

## 🔍 Deep Dive: Vertex AI Search Datastore & Grounding

Enterprise financial analysis and corporate decision-making require **zero hallucination**, **deterministic factual accuracy**, and **strict citation traceability**. The platform implements Google Cloud's native Vertex AI Search Grounding architecture.

```mermaid
sequenceDiagram
    autonumber
    actor User as Corporate User (Word/PPT/Excel)
    participant Addin as Taskpane UI (Office.js)
    participant Proxy as Cloud Function (askGemini)
    participant VertexAI as Vertex AI API (VertexAI SDK)
    participant Datastore as Vertex AI Search Datastore (GCS)
    participant Flash as Gemini 2.5 Flash
    participant Imagen as Gemini 2.5 Flash Image

    User->>Addin: Enters prompt: "Compare Q2 Cloud vs Services performance with charts"
    Addin->>Proxy: POST /askGemini { prompt, history, sessionId, enableGrounding: true }
    
    rect rgb(240, 248, 255)
        Note over Proxy,Datastore: 1. Enterprise Grounding & Retrieval Phase
        Proxy->>VertexAI: getGenerativeModel with datastore tool definition
        VertexAI->>Datastore: Vector Semantic Search across indexed 10-Q/10-K filings
        Datastore-->>VertexAI: Relevant text chunks & document snippets
        VertexAI->>Flash: System prompt + User prompt + Retrieved ground truth chunks
        Flash-->>Proxy: Grounded Markdown text + ![Chart](image: prompt) triggers + Citations
    end

    rect rgb(255, 245, 238)
        Note over Proxy,Imagen: 2. Multimodal Visual Synthesis Phase
        Proxy->>Proxy: Balanced Parenthesis Regex extracts image prompt: "Clean 2D bar chart..."
        Proxy->>Imagen: generateContent(imagePrompt)
        Imagen-->>Proxy: Candidate inlineData (Base64 PNG)
        Proxy->>Proxy: Inline Base64 Data URI into HTML container
    end

    Proxy-->>Addin: JSON Response { result, citations, groundingMetadata, sessionId }
    
    rect rgb(245, 255, 245)
        Note over Addin,User: 3. Office Host Injection Phase
        Addin->>Addin: Parse Markdown, Tables, Visuals into Presentation/Document Model
        Addin->>User: Renders Title slide, Widescreen text layout, and Native 2D Charts
    end
```

### 1. Grounding Configuration & Tool Definition
In `gemini-o365-proxy/index.js`, grounding is configured via the Vertex AI Generative Model `tools` schema:

```javascript
const modelConfig = {
  model: 'gemini-2.5-flash',
  systemInstruction: {
    parts: [{ text: SYSTEM_INSTRUCTION_TEXT }]
  },
  generationConfig: {
    temperature: 0.4,
    maxOutputTokens: 8192
  }
};

if (enableGrounding && DATASTORE_ID) {
  modelConfig.tools = [
    {
      retrieval: {
        vertexAiSearch: {
          datastore: process.env.VERTEX_DATASTORE_ID // Format: projects/PROJECT_ID/locations/global/collections/default_collection/dataStores/DATASTORE_ID
        }
      }
    }
  ];
}
```

### 2. How Grounding Guarantees Precision
1. **Document Ingestion:** Unstructured enterprise files (SEC 10-K, 10-Q filings, earnings releases, board presentations) are ingested from Google Cloud Storage (`gs://...`) into the **Vertex AI Search Datastore**.
2. **Chunking & Hybrid Search:** Vertex AI Search creates dense vector embeddings alongside sparse lexical search indexes.
3. **Retrieval-Augmented Generation (RAG):** When a user asks a question, Vertex AI automatically queries the datastore, extracts top-k relevant passages, and feeds them into the Gemini 2.5 Flash attention context.
4. **Citation Extraction:** Grounding metadata, chunk references, and URI citations are returned alongside the response and rendered as verifiable footnotes in the Office taskpane.

---

## 🎨 Multimodal Visual Pipeline (gemini-2.5-flash-image)

When users request visual charts, graphs, or executive diagrams, the system utilizes a dual-model generation pipeline:

```mermaid
graph LR
    UserPrompt["User Prompt: 'Show revenue chart'"] --> GeminiFlash["Gemini 2.5 Flash"]
    GeminiFlash -->|"Outputs Directive"| Directive["Directive: ![Chart](image: Prompt with params)"]
    Directive --> BalancedParser["Balanced Parenthesis Parser"]
    BalancedParser -->|"Clean Prompt"| NanoBanana["Gemini 2.5 Flash Image / Nano Banana"]
    NanoBanana -->|"Raw Base64 PNG"| DataUriEmbedder["HTML & Base64 Inliner"]
    DataUriEmbedder --> TaskpaneResponse["Taskpane / Host Document"]
```

### Balanced Parenthesis Prompt Parser
Financial chart prompts frequently contain embedded currency and accounting numbers in parentheses (e.g. `Google Services ($94.54B)`). Standard regular expressions fail on nested parentheses. The proxy backend implements a stack-based counter parser:

```javascript
// Stack-based balanced parenthesis prompt extraction
const startMarkerRegex = /!\[([^\]]*)\]\((?:image:|image-prompt:|imagen:)\s*/gi;
while ((match = startMarkerRegex.exec(processed)) !== null) {
  let openCount = 1;
  let i = contentStartIndex;
  while (i < processed.length && openCount > 0) {
    if (processed[i] === '(') openCount++;
    else if (processed[i] === ')') openCount--;
    i++;
  }
  if (openCount === 0) {
    const prompt = processed.substring(contentStartIndex, i - 1).trim();
    // Generate high-resolution 2D image via Vertex AI
  }
}
```

---

## 🖥️ Client-Side Host Adapter Architecture

The add-in uses a polymorphic adapter pattern (`HostAdapterFactory`) to dynamically bind to the running Office application:

```mermaid
classDiagram
    class BaseAdapter {
        <<abstract>>
        +insertText(text)
        +getSelectedText()
        +insertCard(title, body)
        +insertVisualChart(base64)
    }
    class WordAdapter {
        +insertText(text)
        +insertHeading(text, level)
        +insertTable(headers, rows)
        +insertImage(base64)
        +insertInlineGeminiTrigger()
    }
    class PPTAdapter {
        +createExecutiveDeck(slides)
        +insertTitleSlide(title, subtitle)
        +insertContentSlide(title, bullets)
        +insertSplitVisualSlide(title, bullets, base64)
        +injectImageNativeMac(base64, bounds)
    }
    class ExcelAdapter {
        +getSelectedRangeValues()
        +getSheetHeaders()
        +getSurroundingContext()
        +createNativeChart(spec)
        +modifyNativeChart(spec)
        +applyTaxonomyCategorization(spec)
        +applyRangeFormatting(spec)
        +insertFormula(formula)
        +insertSummarySheet(data)
        +extractTableInsights(range)
    }
    class OutlookAdapter {
        +getItemContext()
        +getSenderProfile()
        +insertReply(body, mode)
        +createCalendarInvite(event)
    }

    BaseAdapter <|-- WordAdapter
    BaseAdapter <|-- PPTAdapter
    BaseAdapter <|-- ExcelAdapter
    BaseAdapter <|-- OutlookAdapter
```

### 📈 Excel Client Architecture & Spreadsheet Intelligence

The Excel client integration provides native spreadsheet manipulation through high-performance `Excel.run` batch operations:

1. **Native Office Excel Chart Engine (`createNativeChart` & `modifyNativeChart`):**
   - Direct translation of Gemini chart directives (`ColumnClustered`, `Line`, `Pie`, `BarClustered`, `Area`) to native Office chart shapes.
   - Intelligent multi-column category vs series extraction and aggregation handling (`sum`, `average`, `count`).
   - Sticky chart modification listeners that enable multi-turn natural language adjustments to existing chart titles, dimensions, types, and legend placements.

2. **Taxonomy Categorization Engine (`applyTaxonomyCategorization`):**
   - High-throughput batch classification mapping unstructured column values (e.g. descriptions, merchant names, vendor strings) into normalized business categories.
   - Automatically inserts a new styled category column adjacent to the source column and applies header formatting.
   - Interactive visual taxonomy cards in taskpane with distribution breakdown badges.

3. **Table Insights & Statistical Analytics (`ExcelInsightsService.js`):**
   - Deterministic client-side computational engine that pre-calculates mean, median, min, max, variance, standard deviation, and Pearson correlation coefficients across numerical matrix selections.
   - Generates structured markdown executive summary tables and allows 1-click summary table insertion.

4. **Range Formatting & Formula Translation (`applyRangeFormatting`, `insertFormula`):**
   - Generates and applies custom fill colors, borders, typography, and number formats (`$#,##0.00`, `0.0%`, `YYYY-MM-DD`).
   - Translates complex business calculation requests into valid Excel formulas with auto-detected cell bounds and 1-click cell insertion.

5. **2-Row Contextual Header Tracker:**
   - Real-time sheet schema badge display (row/column bounds, sheet name, detected headers).
   - Dynamic prompt placeholder adaptation (`"Ask Gemini about selection A1:F50..."`).

---

### 🍎 The macOS WKWebView Dual-Pipeline Solution (PowerPoint)

#### The Problem:
On macOS, PowerPoint executes Office add-ins inside a sandboxed `WKWebView`. Passing large base64 image data (>50KB) through `PowerPoint.run` shape APIs triggers an XPC buffer serialization failure in the macOS window server, silently dropping the image.

#### The Dual-Pipeline Architecture:
1. **Slide Creation & Geometry:** Inside `PowerPoint.run`, the slide, header, and left-hand text shapes (width `420px`) are created and committed.
2. **Slide Selection Bridge:** The engine reads `newSlide.id` and synchronizes the active viewport:
   ```javascript
   context.presentation.setSelectedSlides([newSlide.id]);
   await context.sync();
   ```
3. **Native Office Common API Injection:** The engine immediately calls the lower-level C++ Office Common API:
   ```javascript
   Office.context.document.setSelectedDataAsync(base64Image, {
     coercionType: Office.CoercionType.Image,
     imageLeft: 480,
     imageTop: 110,
     imageWidth: 440,
     imageHeight: 340
   });
   ```
This bypasses macOS XPC serialization constraints and provides 100% reliable image insertion across all macOS and Windows Office versions.

---

## 🔒 Security, Compliance & Deployment Model

```mermaid
graph TD
    subgraph EnterpriseBoundary ["Google Cloud Enterprise Boundary"]
        subgraph IAM ["Cloud IAM & Policies"]
            ServiceAccount["Cloud Run / GCF Service Account<br/>(roles/aiplatform.user)"]
        end
        
        subgraph Services ["Deployed Services"]
            Frontend["gemini-frontend (Cloud Run)<br/>Region: us-central1"]
            Backend["askGemini (Cloud Functions Gen 2)<br/>Region: us-central1"]
            DatastoreRes["Vertex AI Search Datastore<br/>Collection: default_collection"]
        end
    end
    
    OfficeClient["Microsoft 365 Client<br/>(Word / PPT / Excel)"] -->|"HTTPS (TLS 1.3)"| Frontend
    OfficeClient -->|"HTTPS JSON (CORS)"| Backend
    Backend -->|"Native IAM Token"| DatastoreRes
```

1. **Zero Secret Storage:** No API keys or static credentials reside in client code or manifests. Authentication to Vertex AI and Gemini Enterprise is handled through Google Cloud IAM Service Account delegation and user-consented OAuth tokens.
2. **Data Residency:** All prompt tokens, document embeddings, and generated visuals remain strictly contained within the Google Cloud project (`your-gemini-enterprise-project-id`, `us-central1`).
3. **CORS Hardened:** Strict CORS headers allow legitimate Microsoft 365 webview origins while blocking unauthenticated third-party scrapers.

---

## ⚙️ Deployment Ordering & OAuth Client Dependencies

When deploying the solution for an environment with **Cloud Identity / Google Workspace Identity**:

```mermaid
flowchart LR
    Step1["1. Deploy Backend & Frontend<br/>(gemini-frontend, auth-proxy, geminiproxy)"] --> Step2["2. Obtain Frontend URL<br/>(https://gemini-frontend-...run.app)"]
    Step2 --> Step3["3. Create OAuth 2.0 Web Client<br/>(In Gemini Enterprise GCP Project)"]
    Step3 --> Step4["4. Set GOOGLE_OAUTH_CLIENT_ID<br/>(On auth-proxy Cloud Run service)"]
    Step4 --> Step5["5. Ready for Office 365 Users<br/>(Dynamic /api/config resolution)"]

    style Step1 fill:#e8f0fe,stroke:#1a73e8,stroke-width:2px;
    style Step2 fill:#fef7e0,stroke:#f9ab00,stroke-width:2px;
    style Step3 fill:#e6f4ea,stroke:#137333,stroke-width:2px;
    style Step4 fill:#f3e8fd,stroke:#7b1fa2,stroke-width:2px;
    style Step5 fill:#f0fdf4,stroke:#10b981,stroke-width:2px;
```

### Why Deployment Ordering Matters:
* **Redirect URI & Origin Dependency:** The Google OAuth 2.0 Web Client credentials must specify the exact JavaScript Origin (`https://gemini-frontend-...run.app`) and Redirect URI (`https://gemini-frontend-...run.app/google-callback.html`). These values are only known once the frontend Cloud Run service is deployed.
* **Dynamic Configuration Delivery:** By passing `GOOGLE_OAUTH_CLIENT_ID` to `auth-proxy` as an environment variable, the Office 365 add-in dynamically retrieves it at runtime via `/api/config`, preventing any static credential baking or build rebuilds when migrating across environments.

---

## 8. 🛠️ Optional Developer Mode: Zero-Auth / Service Account Fallback

> [!WARNING]
> **Non-Production & Testing Only**: This mode completely disables Microsoft Entra ID authentication and user-level ACL enforcement. It is designed **strictly for local development**, rapid prototyping (e.g. running PowerPoint locally on `localhost:3000`), or offline test environments where Microsoft Entra ID tenant registration is not yet configured.

```
[Local PowerPoint / Word / Webview]
          │
          │ 1. POST /askGeminiEnterprise (No Authorization Header)
          ▼
   [Cloud Run: auth-proxy]
          │ ⚙️ REQUIRE_ENTRA_AUTH=false
          │ ⚙️ USER_AUTH_MODE=service_account
          │
          │ 2. Ingests request as 'anonymous_dev_user'
          │ 3. Issues S2S IAM token for askgemini-proxy
          ▼
   [Cloud Run: askgemini-proxy]
          │ ⚙️ ALLOW_SERVICE_ACCOUNT_FALLBACK=true
          │
          │ 4. Detects absence of end-user Google token
          │ 5. Mints Google Cloud ADC access token from Service Account
          ▼
   [Discovery Engine streamAssist API]
          │ Authorization: Bearer <Service_Account_ADC_Token>
          ▼
 [Grounded Gemini Response Returned]
```

### 8.1 How It Works
1. **Perimeter Auth Bypass (`auth-proxy`):**
   When `REQUIRE_ENTRA_AUTH=false`, `auth-proxy` skips Microsoft JWKS signature verification if no `Authorization: Bearer` header is present. The request is assigned a default development profile (`anonymous_dev_user`).
2. **Identity Resolution Bypass:**
   When `USER_AUTH_MODE=service_account`, `auth-proxy` skips WIF token exchange and DWD minting, forwarding the request across the Google S2S IAM boundary without an `X-End-User-Google-Token` header.
3. **Service Account ADC Fallback (`askgemini-proxy`):**
   When `ALLOW_SERVICE_ACCOUNT_FALLBACK=true`, `askgemini-proxy` catches the missing user token, requests a standard Google Cloud Application Default Credentials (ADC) access token for the `gemini-office365-sa` service account via `google.auth.getClient()`, and invokes `streamAssist`.
4. **Office Add-in UI Badge:**
   The frontend add-in detects `user_auth_mode: "service_account"` via `/api/config` and renders `🤖 Service Account Active` on the top status badge, suppressing the Google Sign-In prompt.

### 8.2 Outcome of API Calls in This Mode

| Capability | Behavior in Service Account Fallback Mode |
| :--- | :--- |
| **Discovery Engine Datastores** (GCS, Web Search, BigQuery, Unstructured docs) | **Fully functional** — Grounding operates normally against all datastores attached to the engine that the Service Account has permissions to read. |
| **Multi-turn Chat & Conversational Memory** | **Fully functional** — Session continuity works across turns within the active session. |
| **Personal Google Drive Grounding** | **Bypassed / Not Accessible** — The Service Account cannot access individual users' private Google Drive files unless those files are explicitly shared with the service account email. |
| **Zero-Trust User ACL Filtering** | **Bypassed** — Results returned reflect the Service Account's global access permissions rather than individual user permissions. |
| **User Seat Licensing Attribution** | **Unattributed** — Queries are logged in Google Cloud under the Service Account identity rather than individual employee UPNs. |

### 8.3 Enabling & Disabling via Environment Variables

To enable Zero-Auth Dev Mode:
```bash
# 1. Disable Entra ID requirement and set auth mode on auth-proxy
gcloud run services update auth-proxy \
  --set-env-vars="REQUIRE_ENTRA_AUTH=false,USER_AUTH_MODE=service_account" \
  --region=us-central1

# 2. Allow Service Account ADC fallback on askgemini-proxy
gcloud run services update askgemini-proxy \
  --set-env-vars="ALLOW_SERVICE_ACCOUNT_FALLBACK=true" \
  --region=us-central1
```

To restore Full Enterprise Production Security:
```bash
# 1. Re-enable Entra ID verification on auth-proxy
gcloud run services update auth-proxy \
  --set-env-vars="REQUIRE_ENTRA_AUTH=true,USER_AUTH_MODE=auto" \
  --region=us-central1

# 2. Enforce end-user token requirement on askgemini-proxy
gcloud run services update askgemini-proxy \
  --set-env-vars="ALLOW_SERVICE_ACCOUNT_FALLBACK=false" \
  --region=us-central1
```


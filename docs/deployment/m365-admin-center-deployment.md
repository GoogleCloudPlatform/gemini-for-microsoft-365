# Centralized Enterprise Deployment Guide (Microsoft 365 Admin Center)
**Author:** Carlos Augusto, Principal Architect, Google  
**License:** Apache-2.0  

This guide provides step-by-step instructions for IT Administrators, Global Admins, and Exchange/Application Administrators to deploy and roll out the **Gemini Enterprise (Agentspace)** Microsoft 365 Add-in across an entire organization or targeted pilot groups using the **Microsoft 365 Admin Center**.

---

## 📋 Table of Contents
1. [Overview & Architecture](#1-overview--architecture)
2. [Prerequisites & Administrator Roles](#2-prerequisites--administrator-roles)
3. [Step-by-Step Centralized Deployment Procedure](#3-step-by-step-centralized-deployment-procedure)
4. [Granting Tenant-Wide Admin Consent in Microsoft Entra ID](#4-granting-tenant-wide-admin-consent-in-microsoft-entra-id)
5. [Managing the Deployment Lifecycle](#5-managing-the-deployment-lifecycle)
6. [End-User Experience & Client Availability](#6-end-user-experience--client-availability)
7. [Troubleshooting & Frequently Asked Questions](#7-troubleshooting--frequently-asked-questions)

---

## 1. Overview & Architecture

Centralized Deployment via the Microsoft 365 Admin Center is the official Microsoft-recommended method for deploying Office Add-ins to users and groups.

```
┌────────────────────────────────────────────────────────────────────────────────────────┐
│                              Microsoft 365 Admin Center                                │
│                                (admin.microsoft.com)                                   │
│                                                                                        │
│   1. Upload Manifest (manifest-ca.xml)                                                 │
│   2. Target Users: "Entire Organization" OR "Pilot Security Group"                     │
└──────────────────────────────────────────┬─────────────────────────────────────────────┘
                                           │
                                           │ Deploys via Exchange Online / Graph API
                                           ▼
┌────────────────────────────────────────────────────────────────────────────────────────┐
│                            Assigned Enterprise End Users                               │
│                                                                                        │
│  ┌─────────────────────────┐  ┌─────────────────────────┐  ┌────────────────────────┐  │
│  │    Microsoft Word       │  │  Microsoft PowerPoint   │  │    Microsoft Excel     │  │
│  │ (Desktop Win/Mac & Web) │  │ (Desktop Win/Mac & Web) │  │(Desktop Win/Mac & Web) │  │
│  │                         │  │                         │  │                        │  │
│  │  [✨ Gemini Assistant]  │  │  [✨ Gemini Assistant]  │  │ [✨ Gemini Assistant]  │  │
│  └────────────┬────────────┘  └────────────┬────────────┘  └────────────┬───────────┘  │
└───────────────┼────────────────────────────┼────────────────────────────┼──────────────┘
                │                            │                            │
                └────────────────────────────┼────────────────────────────┘
                                             │ Silent SSO (Office.auth.getAccessToken)
                                             ▼
                               ┌───────────────────────────┐
                               │  Microsoft Entra ID (SSO) │
                               └─────────────┬─────────────┘
                                             │ Authorization: Bearer <Entra_JWT>
                                             ▼
                               ┌───────────────────────────┐
                               │ Google Cloud Run Gateway  │
                               │        auth-proxy         │
                               └───────────────────────────┘
```

### Benefits of Centralized Deployment:
- **No Manual Sideloading Required:** Automatically pushed to user ribbons in Word, PowerPoint, and Excel.
- **Zero Client Installations:** The add-in is a cloud-native web application served over HTTPS.
- **Granular Group Targeting:** Target specific departments (e.g. Sales, Legal, Finance) or pilot security groups.
- **Single Sign-On (SSO):** Seamless Entra ID authentication with pre-authorized tenant consent (no user prompts).

---

## 2. Prerequisites & Administrator Roles

### Required Microsoft 365 Admin Roles
To deploy add-ins centrally, your account must have one of the following directory roles:
- **Global Administrator**
- **Exchange Administrator**
- **Cloud Application Administrator** (or **Application Administrator**)

### Tenant & User Requirements
- **Exchange Online:** Users must have an active Exchange Online mailbox (Exchange on-premises mailboxes are not supported for centralized deployment).
- **Supported Office Licenses:** Microsoft 365 E3, E5, Business Standard, Business Premium, or Office 365 Enterprise plans.
- **Supported Office Clients:**
  - **Word:** Desktop (Windows 10/11, macOS) and Word on the Web.
  - **PowerPoint:** Desktop (Windows 10/11, macOS) and PowerPoint on the Web.
  - **Excel:** Desktop (Windows 10/11, macOS) and Excel on the Web.

### Manifest Files & Hosted URLs
- **Track 1 (WIF Architecture):**
  - **Manifest File:** [`manifest-wif.xml`](manifest-wif.xml)
  - **Manifest URL:** `https://gemini-frontend-1062675944253.us-central1.run.app/manifest-wif.xml`
  - **Hosted Add-in UI:** `https://gemini-frontend-1062675944253.us-central1.run.app/taskpane.html`
- **Track 2 (GSuite / Cloud Identity Architecture):**
  - **Manifest File:** [`manifest-gsuite.xml`](manifest-gsuite.xml)
  - **Manifest URL:** `https://gemini-frontend-16933400417.us-central1.run.app/manifest-gsuite.xml`
  - **Hosted Add-in UI:** `https://gemini-frontend-16933400417.us-central1.run.app/taskpane.html`

---

## 3. Step-by-Step Centralized Deployment Procedure

### Step 3.1: Navigate to Integrated Apps
1. Open your browser and navigate to the **Microsoft 365 Admin Center**:
   👉 [https://admin.microsoft.com](https://admin.microsoft.com)
2. Sign in with your administrator credentials.
3. In the left navigation menu, expand **Settings** and select **Integrated apps**.
   *(Note: If you do not see Integrated apps, click **Show all** at the bottom of the navigation menu).*

```
Microsoft 365 admin center
├── Users
├── Teams & groups
├── Billing
└── Settings
    ├── Domains
    ├── Org settings
    └── 📌 Integrated apps  <-- CLICK HERE
```

---

### Step 3.2: Initiate Custom App Upload
1. On the **Integrated apps** page, click the **Upload custom apps** button in the top action bar.

```
┌────────────────────────────────────────────────────────────────────────────────────────┐
│ Integrated apps                                                                        │
│ Manage apps and add-ins deployed to your organization.                                 │
│                                                                                        │
│  [➕ Upload custom apps]   [🔍 Search apps]                                            │
└────────────────────────────────────────────────────────────────────────────────────────┘
```

---

### Step 3.3: Select App Type and Provide Manifest
1. In the **Upload App to deploy** panel:
   - **App type**: Select **Office Add-in**.
2. Under **Choose how to upload the app**, choose according to your architecture track:
   - **For WIF Deployment (Option A - URL):**
     ```text
     https://gemini-frontend-1062675944253.us-central1.run.app/manifest-wif.xml
     ```
     *(Or upload file [`manifest-wif.xml`](manifest-wif.xml))*
   - **For GSuite Deployment (Option A - URL):**
     ```text
     https://gemini-frontend-16933400417.us-central1.run.app/manifest-gsuite.xml
     ```
     *(Or upload file [`manifest-gsuite.xml`](manifest-gsuite.xml))*
3. Click **Validate**.

```
┌────────────────────────────────────────────────────────────────────────────────────────┐
│ Upload App to deploy                                                                   │
│                                                                                        │
│ App type:                                                                              │
│ (o) Office Add-in                                                                      │
│ ( ) Teams app                                                                          │
│                                                                                        │
│ Choose how to upload the app:                                                          │
│ (o) Provide link to manifest file                                                      │
│     [ https://gemini-frontend-1062675944253.us-central1.run.app/manifest-wif.xml       ] │
│ ( ) Upload manifest file (.xml) from device                                            │
│                                                                                        │
│                                                                 [ Validate ]           │
└────────────────────────────────────────────────────────────────────────────────────────┘
```

4. The portal validates the manifest schema, supported host applications (Word, PowerPoint, Excel), and single sign-on parameters.
5. Once validation succeeds (marked with a green checkmark), click **Next**.

---

### Step 3.4: Configure User Assignments
Choose who in your organization gets access to the Gemini Enterprise assistant:

```
┌────────────────────────────────────────────────────────────────────────────────────────┐
│ Assign Users                                                                           │
│ Specify who has access to this add-in.                                                 │
│                                                                                        │
│ ( ) Entire organization                                                                │
│ (o) Specific users/groups  <-- (Recommended for Pilot / Phased Rollout)                │
│     [ Search: "Gemini Enterprise Pilots", "Executive Team", "Sales"                 ]  │
│ ( ) Just me (Testing only)                                                             │
│                                                                                        │
│ Deployment method:                                                                     │
│ [x] Fixed (Default) - Add-in automatically deployed to user ribbons                    │
│                                                                                        │
│                                                            [ Back ]  [ Next ]          │
└────────────────────────────────────────────────────────────────────────────────────────┘
```

1. **Assignment Options**:
   - **Specific users/groups (Recommended for initial launch):** Select security groups, Microsoft 365 groups, or individual users to pilot the assistant before broad rollout.
   - **Entire organization:** Deploys the add-in to every licensed user across the tenant.
   - **Just me:** Deploys the add-in only to the active administrator account for validation.
2. Click **Next**.

---

### Step 3.5: Accept Permissions & Capabilities
1. The portal displays the requested permissions:
   - **ReadWriteDocument:** Allows the assistant to read selected text/context and insert generated summaries, slides, or tables.
   - **Single Sign-On (SSO):** Entra ID user identity delegation via `<WebApplicationInfo>`.
2. Click **Next**.

---

### Step 3.6: Review and Finish Deployment
1. Review your deployment summary:
   - **App Name:** `Gemini Enterprise (Agentspace)`
   - **Hosts:** `Word`, `PowerPoint`, `Excel`
   - **Assigned Users:** Selected groups / Organization
   - **Status:** Ready to deploy
2. Click **Finish deployment**.
3. The Admin Center will process the deployment and display a confirmation screen:
   `"Deployment completed. The app will be available to users shortly."`
4. Click **Done**.

---

## 4. Granting Tenant-Wide Admin Consent in Microsoft Entra ID

To ensure that end users experience **zero permission popups** when opening the Gemini assistant for the first time, grant tenant-wide admin consent in Microsoft Entra ID:

### Step 4.1: Open App Registration in Microsoft Entra Admin Center
1. Navigate to the **Microsoft Entra Admin Center**:
   👉 [https://entra.microsoft.com](https://entra.microsoft.com)
2. In the left navigation, go to **Applications** > **App registrations**.
3. Click the **All applications** tab and select the Gemini Add-in application:
   - **Application (client) ID:** `b990d644-e47b-4575-97b3-2067c488042b`

### Step 4.2: Grant Admin Consent for API Permissions
1. Under the **Manage** menu, select **API permissions**.
2. Verify the configured Microsoft Graph / API permissions:
   - `openid` (Sign users in)
   - `profile` (View user basic profile)
   - `email` (View user email address)
   - `User.Read` (Sign in and read user profile)
3. Click **Grant admin consent for [Your Organization]**.
4. Confirm by clicking **Yes**.
5. Verify that all permissions show a green status checkmark: `"Granted for [Your Organization]"`.

```
┌────────────────────────────────────────────────────────────────────────────────────────┐
│ Configured permissions                                                                 │
│                                                                                        │
│ [➕ Add a permission]  [🔑 Grant admin consent for Contoso]                           │
│                                                                                        │
│ API / Permissions Name       Type          Admin consent required   Status             │
│ Microsoft Graph (4)                                                                    │
│   email                      Delegated     No                       ✔ Granted          │
│   openid                     Delegated     No                       ✔ Granted          │
│   profile                    Delegated     No                       ✔ Granted          │
│   User.Read                  Delegated     No                       ✔ Granted          │
│                                                                                        │
│ Exposed APIs (1)                                                                       │
│   access_as_user             Delegated     Yes                      ✔ Granted          │
└────────────────────────────────────────────────────────────────────────────────────────┘
```

### Step 4.3: Expose API & Pre-Authorize Office Client Applications
1. In the left navigation of the app registration, select **Expose an API**.
2. Verify that the **Application ID URI** is configured to match the frontend hosting domain:
   ```text
   api://gemini-frontend-16933400417.us-central1.run.app/b990d644-e47b-4575-97b3-2067c488042b
   ```
   *(Ensure there is no trailing slash).*
3. Under **Authorized client applications**, ensure the 3 official Microsoft 365 client applications are pre-authorized with the `access_as_user` scope:

   | Client ID | Application / Surface |
   | :--- | :--- |
   | `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` | **Office on the Web** (Word, PowerPoint, Excel online) |
   | `d3590ed6-52b3-4102-aeff-aad2292ab01c` | **Office on the Web** (WAC / Outlook) |
   | `00000002-0000-0ff1-ce00-000000000000` | **Office Desktop** (Windows and macOS) |

---

## 5. Managing the Deployment Lifecycle

### Updating the Add-in
Because the frontend UI is hosted on Google Cloud Run (`https://gemini-frontend-16933400417.us-central1.run.app`), updates to web pages, UI components, prompts, and backend integrations **take effect immediately without needing to re-upload the manifest**.

When an XML manifest change *is* required (e.g. adding a new ribbon button, updating icon assets, or altering scopes):
1. In the Microsoft 365 Admin Center, go to **Settings** > **Integrated apps**.
2. Click **Gemini Enterprise (Agentspace)** from the list of deployed apps.
3. In the flyout pane, click **Update app** (or **Manage manifest**).
4. Select **Provide link to manifest file** or upload the updated XML file.
5. Click **Next** > **Finish update**.

### Modifying User Assignments
1. Go to **Settings** > **Integrated apps** > select **Gemini Enterprise (Agentspace)**.
2. In the flyout pane, select the **Users** tab.
3. Click **Edit assigned users** to add or remove security groups and departments.
4. Click **Save**.

### Turning Off or Removing the Add-in
1. Go to **Settings** > **Integrated apps** > select **Gemini Enterprise (Agentspace)**.
2. Click **Delete app** (or toggle **Active** to **Inactive**).
3. Confirm the removal. The add-in will be removed from all user ribbons on their next Office restart.

---

## 6. End-User Experience & Client Availability

### Rollout Timeline & Ribbon Propagation
- **Office for the Web (office.com):** Available **immediately** (under 5 minutes) upon browser page refresh.
- **Office Desktop (macOS & Windows):** Typically appears within **1 to 4 hours** after deployment. Users can accelerate this by restarting Word, PowerPoint, or Excel.

### Where Users Find the Add-in

#### 1. Microsoft Word
- **Home Tab Ribbon:** The **Gemini Assistant** icon appears in the `Gemini Enterprise` group.
- **In-Document Prompts:** Users can type `@gemini <prompt>` directly into paragraphs and click **⚡ Run @gemini**.

#### 2. Microsoft PowerPoint
- **Home Tab Ribbon:** The **Gemini Assistant** icon is available in the right section of the Home tab.
- **Slide Generator:** Generates full executive decks, widescreen slides, metric callouts, and embeds visual infographics.

#### 3. Microsoft Excel
- **Home Tab Ribbon:** The **Gemini Assistant** icon opens the Excel intelligence pane for sheet data analysis, anomaly detection, and formula explanations.

---

## 7. Troubleshooting & Frequently Asked Questions

### Q1: The add-in does not appear in a user's desktop ribbon after deployment.
**Solutions**:
1. Verify the user is a member of the assigned group in **Integrated apps** > **Gemini Enterprise** > **Users**.
2. Verify the user is signed into the Office desktop application with their corporate Microsoft 365 account (not a personal Microsoft account).
3. Force a ribbon cache refresh in Office:
   - On Windows: Go to **File** > **Account** > click **Update Options** > **Update Now**.
   - On macOS: Completely quit and reopen the application (`Cmd + Q`).
4. Check **Insert** > **Add-ins** > **Admin Managed** to see if the add-in is listed.

### Q2: User sees error `13001` or `13003` when launching the add-in.
**Cause:** The user's device is not signed into the corporate tenant, or third-party cookies/tokens are blocked.  
**Solution:** Ensure the user is signed into Office with their corporate Entra ID account, and verify that tenant admin consent was completed in Section 4.

### Q3: Discovery Engine / Gemini returns `HTTP 403 Permission Denied`.
**Cause:** The corporate user's identity is authenticated by Entra ID, but the user does not have an assigned license in Google Cloud Discovery Engine / Gemini Enterprise.  
**Solution:**
- Assign the user a Gemini Enterprise license in the Google Cloud Admin console, OR
- Enable `ALLOW_SERVICE_ACCOUNT_FALLBACK=true` on the `askgemini-proxy` backend service to permit service-account-grounded fallback during trial periods.

### Q4: Can I restrict deployment to specific countries or departments?
**Yes.** In Step 3.4, choose **Specific users/groups** and select Entra ID dynamic security groups (e.g. `Department - US Sales`, `Region - EMEA Legal`).

---

## 📞 Support & Administrator Contacts
- **Technical Lead:** Carlos Augusto (`admin@caugusto.altostrat.com`)
- **GCP Project:** `agentspace-452714` (Region: `us-central1`)
- **Auth Gateway URL:** `https://auth-proxy-16933400417.us-central1.run.app`
- **Frontend URL:** `https://gemini-frontend-16933400417.us-central1.run.app`

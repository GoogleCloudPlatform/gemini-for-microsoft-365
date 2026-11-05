/**
 * Copyright 2026 Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
 * Gemini for Microsoft 365 - Add-in Taskpane Controller
 * 
 * Manages the taskpane UI, chat history, selection toolbar, in-document triggers,
 * and host-adaptive document intelligence for Word, PowerPoint, and Excel.
 * Styled in accordance with Google Material 3 / Gemini in Google Docs design system.
 * 
 * @author Sathya AG, Principal Architect, Google
 */

import { askGeminiEnterprise, askGeminiEnterpriseStream, getActiveProxyUrl, setProxyUrlOverride, fetchEnterpriseSessions, deleteEnterpriseSession } from '../core/geminiClient.js';
import { 
  getOfficeAuthToken, 
  getUserProfile, 
  getLastAuthError, 
  initiateGoogleSignIn, 
  isGoogleTokenValid, 
  getGoogleAccessToken,
  setGoogleAccessToken,
  syncGoogleTokenFromBackend,
  getGoogleOAuthClientId,
  fetchAppConfig
} from '../core/authService.js';
import { parseMarkdown, normalizeEscapedNewlines } from '../core/markdownParser.js';
import { HostAdapterFactory } from '../adapters/HostAdapterFactory.js';
import { initPowerPointDiagnostics } from '../adapters/ppt/pptDiagnostics.js';
import { BUILDER_REVISION } from '../adapters/ppt/slideBuilder.js';
import excelHostModule from '../taskpanes/excel/index.js';
import { docChips as excelDocChips, selectionChips as excelSelectionChips } from '../taskpanes/excel/chips.js';
import { renderResponseActions as renderExcelResponseActions } from '../taskpanes/excel/responseActions.js';
import { SessionManager } from '../core/sessionManager.js';

let currentSessionId = null;
let chatHistoryState = [];
let askChatHistory = [];
let askSessionId = null;
let isProcessingInDocCommand = false;
let isGenerating = false;
let hostAdapter = null;
let currentSelectedText = "";
let userClearedSelection = false;
let attachedFileContent = null;
let attachedFileName = null;
let attachedFileMimeType = null;

/**
 * Clears prior chat interaction history and empties the prompt textarea above the send button.
 * Invoked whenever a user clicks any top action chip/button.
 */
export function clearChatAndInput() {
  const historyDiv = document.getElementById("chatHistory");
  if (historyDiv) {
    historyDiv.innerHTML = "";
  }
  const promptInput = document.getElementById("promptText");
  if (promptInput) {
    promptInput.value = "";
    promptInput.style.height = "auto";
  }
  console.log("🧹 [Chat] Cleared previous chat interactions and reset input box.");
}

/**
 * Paints the slide-builder revision into the header.
 */
function showBuildStamp() {
  try {
    const el = document.getElementById("buildStamp");
    if (el) el.textContent = BUILDER_REVISION;
  } catch (_) {}
}

async function initTaskpane(info = {}) {
  showBuildStamp();
  // Detect active Microsoft Office host (Word, PowerPoint, Excel) dynamically
  hostAdapter = HostAdapterFactory.getAdapter(info);
  const hostName = hostAdapter ? hostAdapter.name : 'Word';
  if (typeof window !== "undefined") {
    window.__activeOfficeHost = hostName;
  }

  // IMMEDIATELY adapt UI for the active host (PowerPoint, Word, Excel)
  adaptUIForHost(hostName);

  if (hostName === "PowerPoint") {
    initPowerPointDiagnostics();
  } else if (hostName === "Excel") {
    try {
      excelHostModule.onReady({
        adapter: hostAdapter,
        onSelectionChanged: (rawText) => handleSelectionChanged(rawText)
      });
    } catch (e) {
      console.warn("Excel host module onReady failed:", e);
    }
  }

  // Pre-fetch dynamic backend configuration (Google OAuth Client ID)
  fetchAppConfig().catch(e => console.warn("Background config fetch failed:", e));

  // Initialize Entra ID & Google Drive Identity in UI.
  // Deliberately not awaited: initAuthUI() waits on Office SSO, which can take
  // many seconds or never settle. Awaiting it here left Send, the slash-command
  // popup, upload, New chat, history and the suggestion prompts unwired until
  // SSO answered, so the Word UI looked present but did nothing.
  initAuthUI().catch((e) => console.warn("Auth UI init failed; continuing without identity badge:", e));

  // Initialize Collapsible Troubleshooting & Diagnostics Panel
  initTroubleshootPanel();

  if (hostAdapter && hostAdapter.name === "Outlook") {
    if (typeof document !== "undefined" && document.body) {
      document.body.classList.add("host-outlook");
    }
    const hostBadge = document.getElementById("hostBadge");
    if (hostBadge) {
      hostBadge.innerText = "Outlook";
      hostBadge.style.display = "inline-block";
    }
    const outlookContainer = document.getElementById("outlookHostContainer");
    const docContainer = document.getElementById("officeDocHostContainer");
    if (outlookContainer) outlookContainer.style.display = "none";
    if (docContainer) docContainer.style.display = "flex";
    setupOutlookUI();
  }

  // Wire capture listener on top buttons to clear chat history and input box on click
  const docToolsContainer = document.getElementById("docToolsChipsContainer");
  if (docToolsContainer) {
    docToolsContainer.addEventListener("click", (e) => {
      if (e.target.closest(".quick-chip")) {
        clearChatAndInput();
      }
    }, true);
  }

  // Wire click-to-zoom on preview images in chat history
  const chatHistoryDiv = document.getElementById("chatHistory");
  if (chatHistoryDiv) {
    chatHistoryDiv.addEventListener("click", (e) => {
      const insertNewBtn = e.target.closest(".img-action-btn-insert-new");
      const insertCurrentBtn = e.target.closest(".img-action-btn-insert-current");
      if (insertNewBtn) {
        e.stopPropagation();
        const src = insertNewBtn.getAttribute("data-img-src");
        if (src) {
          performDocumentInsertion(`<img src="${src}" />`, `![Image](${src})`, "insert_cursor", { imageOnly: true });
        }
        return;
      }
      if (insertCurrentBtn) {
        e.stopPropagation();
        const src = insertCurrentBtn.getAttribute("data-img-src");
        if (src) {
          performDocumentInsertion(`<img src="${src}" />`, `![Image](${src})`, "insert_current_slide", { imageOnly: true });
        }
        return;
      }

      const zoomBtn = e.target.closest(".img-zoom-btn, .img-action-btn-zoom");
      const clickedImg = e.target.tagName === "IMG" ? e.target : e.target.closest("img");
      if (zoomBtn) {
        e.stopPropagation();
        let src = zoomBtn.getAttribute("data-img-src");
        let alt = zoomBtn.getAttribute("data-img-alt") || "Generated Visual";
        if (!src) {
          const card = zoomBtn.closest(".office-visual-image-card, .rendered-chart-container, .chat-bubble");
          const img = card ? card.querySelector("img") : null;
          if (img && img.src) {
            src = img.src;
            alt = img.alt || alt;
          }
        }
        if (src) {
          openImageZoomModal(src, alt);
        }
      } else if (clickedImg && clickedImg.src) {
        e.stopPropagation();
        openImageZoomModal(clickedImg.src, clickedImg.alt);
      }
    });
  }

  // Log Add-in startup and host context to Cloud Logging
  sendDiagnosticLogToCloud(
    `Office Add-in initialized on host '${info?.host || hostName}' (${info?.platform || 'Unknown platform'})`,
    "INFO",
    "STARTUP",
    { host: info?.host || hostName, platform: info?.platform }
  );

  // Wire Google Drive 1-click connect button (used in GSuite / Cloud Identity mode)
  const googleDriveBtn = document.getElementById("googleDriveBtn");
  if (googleDriveBtn) {
    googleDriveBtn.onclick = async () => {
      if (googleDriveBtn.classList.contains("busy")) return;
      const config = await fetchAppConfig();
      const isWifMode = config?.user_auth_mode === 'wif' || (config?.user_auth_mode === 'auto' && !config?.google_oauth_client_id);
      if (isWifMode) {
        console.log("WIF SSO active: Google OAuth login not required.");
        return;
      }
      console.log("Triggering Google OAuth 3-legged sign-in flow...");
      // .busy swaps the Google "G" for a spinner (taskpane.html) until initAuthUI()
      // rewrites the class and label below.
      googleDriveBtn.classList.add("busy");
      googleDriveBtn.textContent = "Logging in...";
      try {
        const profile = getUserProfile();
        const res = await initiateGoogleSignIn(null, 'select_account');
        if (res.status === 'success') {
          if (hostAdapter && hostAdapter.name === "Outlook") {
            const historyDiv = document.getElementById("chatHistory");
            if (historyDiv) {
              historyDiv.querySelectorAll(".chat-bubble.system").forEach((b) => {
                if (b.innerText && (b.innerText.includes("Google Login Error") || b.innerText.includes("Google login successful"))) {
                  b.remove();
                }
              });
            }
          }
          appendBubble("✅ Google login successful! Gemini grounding is now active.", "system");
        } else {
          console.warn("Google Sign-In was not completed:", res.error);
          const errorMsg = res.error || "Google Sign-In was not completed.";
          showToast(`Google Sign-In: ${errorMsg}`, "error");
          appendBubble(`❌ Google Login Error: ${errorMsg}`, "system");
        }
      } catch (err) {
        const errorMsg = err && err.message ? err.message : String(err);
        console.warn("Google Sign-In could not be started:", errorMsg);
        showToast(`Google Sign-In Error: ${errorMsg}`, "error");
        appendBubble(`❌ Google Login Error: ${errorMsg}`, "system");
      } finally {
        await initAuthUI();
      }
    };
  }

  const runBtn = document.getElementById("run");
  if (runBtn) {
    runBtn.onclick = () => callGeminiProxy();
  }

  // Setup Slash Commands Popup, Autocomplete, and Auto-Expansion (also wires Enter key on #promptText)
  setupSlashCommands();

  // Document Upload (+) Button & File Input Handling
  const uploadDocBtn = document.getElementById("uploadDocBtn");
  const fileUploadInput = document.getElementById("fileUploadInput");
  const attachmentContainer = document.getElementById("attachmentContainer");
  const attachmentName = document.getElementById("attachmentName");
  const removeAttachmentBtn = document.getElementById("removeAttachmentBtn");

  if (uploadDocBtn && fileUploadInput) {
    uploadDocBtn.onclick = () => {
      fileUploadInput.value = "";
      fileUploadInput.click();
    };

    fileUploadInput.onchange = (e) => {
      const file = e.target.files && e.target.files[0];
      if (!file) return;

      attachedFileName = file.name;
      let mimeType = file.type;
      if (!mimeType) {
        const lowerName = file.name.toLowerCase();
        if (lowerName.endsWith(".pdf")) mimeType = "application/pdf";
        else if (lowerName.endsWith(".docx")) mimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
        else if (lowerName.endsWith(".doc")) mimeType = "application/msword";
        else if (lowerName.endsWith(".csv")) mimeType = "text/csv";
        else if (lowerName.endsWith(".json")) mimeType = "application/json";
        else if (lowerName.endsWith(".md")) mimeType = "text/markdown";
        else mimeType = "text/plain";
      }
      attachedFileMimeType = mimeType;

      if (attachmentName) attachmentName.innerText = file.name;
      if (attachmentContainer) attachmentContainer.style.display = "block";
      updateSendButtonState();

      const reader = new FileReader();
      reader.onload = (event) => {
        const dataUrl = String(event.target.result || "");
        attachedFileContent = dataUrl.includes(",") ? dataUrl.split(",")[1] : dataUrl;
        updateSendButtonState();
        showToast(`Attached ${file.name}`);
      };
      reader.onerror = () => {
        showToast("Could not read file");
      };
      reader.readAsDataURL(file);
    };
  }

  if (removeAttachmentBtn) {
    removeAttachmentBtn.onclick = (e) => {
      e.stopPropagation();
      attachedFileContent = null;
      attachedFileName = null;
      attachedFileMimeType = null;
      if (fileUploadInput) fileUploadInput.value = "";
      if (attachmentContainer) attachmentContainer.style.display = "none";
      updateSendButtonState();
      showToast("Attachment removed");
    };
  }
  const clearSessionBtn = document.getElementById("clearSession");
  if (clearSessionBtn) {
    clearSessionBtn.onclick = resetChatSession;
  }
  const newChatBtn = document.getElementById("newChatBtn");
  if (newChatBtn) {
    newChatBtn.onclick = resetChatSession;
  }

  const menuBtn = document.getElementById("menuBtn");
  if (menuBtn) {
    menuBtn.onclick = () => toggleRecentChats(true);
  }

  const historyBtn = document.getElementById("historyBtn");
  if (historyBtn) {
    historyBtn.onclick = () => toggleRecentChats(true);
  }

  const closeRecentChatsBtn = document.getElementById("closeRecentChatsBtn");
  if (closeRecentChatsBtn) {
    closeRecentChatsBtn.onclick = () => toggleRecentChats(false);
  }

  const closeRecentChatsCrossBtn = document.getElementById("closeRecentChatsCrossBtn");
  if (closeRecentChatsCrossBtn) {
    closeRecentChatsCrossBtn.onclick = () => toggleRecentChats(false);
  }

  const viewAllChatsLink = document.getElementById("viewAllChatsLink");
  if (viewAllChatsLink) {
    viewAllChatsLink.onclick = () => {
      showToast("All recent chats are displayed above.");
    };
  }

  const learnMoreLink = document.getElementById("learnMoreLink");
  if (learnMoreLink) {
    learnMoreLink.onclick = (e) => {
      e.preventDefault();
      showToast("Gemini protects enterprise data privacy and security.");
    };
  }

  const closeHistoryBtn = document.getElementById("closeHistoryBtn");
  if (closeHistoryBtn) {
    closeHistoryBtn.onclick = closeHistoryDrawer;
  }

  const historyOverlay = document.getElementById("historyOverlay");
  if (historyOverlay) {
    historyOverlay.onclick = closeHistoryDrawer;
  }

  const clearAllHistoryBtn = document.getElementById("clearAllHistoryBtn");
  if (clearAllHistoryBtn) {
    clearAllHistoryBtn.onclick = () => {
      SessionManager.clearAllSessions(hostName);
      resetChatSession();
      renderHistoryList();
      showToast("Cleared recent chats");
    };
  }

  const scanBtn = document.getElementById("scanInDoc");
  if (scanBtn) {
    scanBtn.onclick = () => checkForInDocumentCommands(true);
  }

  // Target Proxy Endpoint selector setup
  const endpointSelect = document.getElementById("endpointSelect");
  if (endpointSelect) {
    const savedUrl = window.localStorage ? window.localStorage.getItem('gemini_proxy_url') : "";
    if (savedUrl) {
      endpointSelect.value = savedUrl;
    }
    endpointSelect.onchange = (e) => {
      const selectedUrl = e.target.value;
      setProxyUrlOverride(selectedUrl);
      const debugStatus = document.getElementById("debugStatus");
      if (debugStatus) {
        if (hostAdapter && hostAdapter.name === "Outlook" && !selectedUrl) {
          debugStatus.innerText = "";
          debugStatus.style.display = "none";
        } else {
          debugStatus.innerText = selectedUrl ? `Target: Override Active` : `${hostAdapter ? hostAdapter.name : 'Office'} Ready`;
        }
      }
    };
  }

  // Setup Document Intelligence Chips (Feature 1: Full Document Q&A)
  setupDocToolsChips();

  // Setup Selection Quick Toolbar Chips (Feature 3: Inline Rewrite Toolbar)
  setupSelectionChips();

  // Setup Transform Doc to Deck Feature
  initDocToDeckFeature();

  // Wire Tip Banner Dismiss button
  const dismissTipBtn = document.getElementById("dismissTipBtn");
  if (dismissTipBtn) {
    dismissTipBtn.onclick = () => {
      const tipBanner = document.getElementById("tipBanner");
      if (tipBanner) tipBanner.style.display = "none";
    };
  }

  // Wire Context Attachment Pill Clear button
  const clearSelectionBtn = document.getElementById("clearSelectionBtn");
  if (clearSelectionBtn) {
    clearSelectionBtn.onclick = () => {
      userClearedSelection = true;
      currentSelectedText = "";
      renderAdaptiveActionChips(false);
    };
  }

  // Wire interactive sign-in click on user profile badge
  const userAuthBadge = document.getElementById("userAuthBadge");
  const userStatusDot = document.getElementById("userStatusDot");
  const userProfileBar = document.getElementById("userProfileBar");
  
  const handleAuthClick = async () => {
    console.log("Triggering explicit Office Entra ID SSO sign-in...");
    if (userAuthBadge) userAuthBadge.innerText = "Signing in...";
    const token = await getOfficeAuthToken(true);
    await initAuthUI();
    const lastErr = getLastAuthError();
    if (!token && lastErr) {
      console.warn("Explicit sign-in attempt did not yield token:", lastErr);
      if (lastErr.code === 13007) {
        const activeHost = (typeof window !== "undefined" && window.location?.host) ? window.location.host : "your-frontend-host";
        const entraId = (appConfig?.microsoft_entra_app_id || (typeof ENV_CONFIG !== "undefined" && ENV_CONFIG.MICROSOFT_ENTRA_APP_ID) || "your-entra-client-id");
        alert("Office SSO Error 13007: Application ID URI mismatch or client app not authorized in Entra ID.\n\n" +
              "1. In Entra ID App '" + entraId + "' -> Expose an API:\n" +
              "   Set App ID URI: api://" + activeHost + "/" + entraId + "\n" +
              "   Authorized Client IDs:\n" +
              "   - 00000002-0000-0ff1-ce00-000000000000 (Office Desktop)\n" +
              "   - ea5a67f6-b6f3-4338-b240-c655ddc3cc8e (Office Web)\n" +
              "   - d3590ed6-52b3-4102-aeff-aad2292ab01c (Office Web / WAC)\n" +
              "2. Restart Office.");
      } else if (lastErr.code === 13001) {
        alert("Office SSO Error 13001: You are not currently signed into Microsoft Office with a corporate Microsoft Entra ID account.");
      } else if (lastErr.code === 13002) {
        alert("Office SSO Error 13002: Sign-in or consent was cancelled.");
      } else if (lastErr.code === 13012) {
        alert("Office SSO Error 13012: SSO API is not supported on this platform version or requires Office restart.");
      } else if (lastErr.code) {
        alert(`Office SSO Code ${lastErr.code}: ${lastErr.message || JSON.stringify(lastErr)}`);
      }
    }
  };

  if (userAuthBadge) userAuthBadge.onclick = handleAuthClick;
  if (userStatusDot) userStatusDot.onclick = handleAuthClick;
  if (userProfileBar) {
    userProfileBar.style.cursor = "pointer";
    userProfileBar.onclick = handleAuthClick;
  }

  // Attach selection change handler for Word & Excel in-document detection & selection indicator
  try {
    if ((hostName === "Word" || hostName === "Excel") && typeof Office !== 'undefined' && Office.context && Office.context.document && typeof Office.context.document.addHandlerAsync === 'function') {
      Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        () => handleSelectionChanged()
      );
    }
  } catch (e) {
    console.warn("Could not attach selection handler:", e);
  }

  // Close menus when clicking anywhere in document
  document.addEventListener("click", () => {
    document.querySelectorAll(".dropdown-menu.show").forEach(menu => {
      menu.classList.remove("show");
    });
  });

  // Initialize rotating headline and suggestion prompts
  setupRotatingHeadline();
  setupSuggestionPrompts();
  updateSendButtonState();

  // Restore previous active session messages only if available for this exact document
  // and the document itself is not a brand-new blank document.
  const savedActiveId = SessionManager.getActiveSessionId(hostName);
  const savedSession = savedActiveId ? SessionManager.getSession(savedActiveId) : null;
  if (savedActiveId && savedSession && Array.isArray(savedSession.messages) && savedSession.messages.length > 0) {
    if (hostName === "Word" && hostAdapter && typeof hostAdapter.getFullDocumentText === "function") {
      updateChatDisplayState(false);
      hostAdapter.getFullDocumentText().then((docText) => {
        if (isGenerating || currentSessionId) return; // User already started interacting
        if (!docText || docText.trim().length === 0) {
          // Brand-new empty Word document: never carry over a previous document's chat session
          currentSessionId = null;
          SessionManager.clearActiveSession(hostName);
          updateChatDisplayState(false);
        } else {
          loadSessionIntoUI(savedActiveId);
        }
      }).catch(() => {
        if (!isGenerating && !currentSessionId) {
          updateChatDisplayState(false);
        }
      });
    } else {
      loadSessionIntoUI(savedActiveId);
    }
  } else {
    currentSessionId = null;
    SessionManager.clearActiveSession(hostName);
    updateChatDisplayState(false);
  }

  // Check if text is already highlighted on startup (for Excel, initial state is handled by excelHostModule)
  if (hostName !== "Excel") {
    handleSelectionChanged();
  } else {
    const selToolbar = document.getElementById("selectionToolbar");
    if (selToolbar) selToolbar.style.display = "none";
  }

  // Cross-Surface Auth Sync: When switching back to an already-open surface (e.g., Excel/Word/PPT -> Outlook or vice versa),
  // automatically pull the Google OAuth token from the backend if the user logged in on another surface.
  const handleWindowFocusAuthSync = async () => {
    if (!isGoogleTokenValid()) {
      const synced = await syncGoogleTokenFromBackend();
      if (synced) {
        await initAuthUI();
      }
    }
  };
  window.addEventListener("focus", handleWindowFocusAuthSync);
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible") {
      handleWindowFocusAuthSync();
    }
  });
}


Office.onReady((info) => {
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", () => initTaskpane(info));
  } else {
    initTaskpane(info);
  }
});

async function initAuthUI() {
  const userAuthBar = document.getElementById("userAuthBar");
  const userEmailText = document.getElementById("userEmailText");
  const userStatusDot = document.getElementById("userStatusDot");
  const googleDriveBtn = document.getElementById("googleDriveBtn");

  try {
    const token = await getOfficeAuthToken();
    const profile = getUserProfile();
    const lastErr = getLastAuthError();
    const config = await fetchAppConfig();
    const isWifMode = config?.user_auth_mode === 'wif' || (config?.user_auth_mode === 'auto' && !config?.google_oauth_client_id);

    // Cross-surface Google OAuth sync: If this surface's browser storage partition (e.g. outlook.office.com vs *.sharepoint.com)
    // does not yet have a Google token, check if the user already logged in on another Office surface via Entra ID SSO.
    if (!isWifMode && !isGoogleTokenValid() && token) {
      await syncGoogleTokenFromBackend();
    }

    // 1. Entra ID / Microsoft 365 status & WIF status
    if (token && profile.is_authenticated) {
      if (userEmailText) {
        userEmailText.innerText = profile.email || profile.name;
        userEmailText.className = "user-email-text";
        userEmailText.title = isWifMode
          ? `WIF SSO Active (Connected as ${profile.email || profile.name})`
          : `Connected to Microsoft 365 as ${profile.email || profile.name} (Tenant: ${profile.tenant_id || 'Entra ID'})`;
      }
      if (userStatusDot) {
        userStatusDot.className = "user-status-dot";
        userStatusDot.title = isWifMode
          ? `WIF SSO Active (Connected as ${profile.email || profile.name})`
          : `Connected to Microsoft 365 as ${profile.email || profile.name}`;
      }
      if (userAuthBar) {
        userAuthBar.title = isWifMode
          ? `WIF SSO Active (Connected as ${profile.email || profile.name})`
          : `Microsoft 365 Identity: ${profile.email || profile.name}`;
      }
    } else {
      const errHint = lastErr ? (lastErr.message || lastErr.code || JSON.stringify(lastErr)) : "Not connected to Microsoft 365";
      if (userEmailText) {
        userEmailText.innerText = profile.email && profile.email !== 'user@organization.com' ? profile.email : "Not Connected";
        userEmailText.className = "user-email-text offline";
        userEmailText.title = isWifMode ? `WIF SSO Inactive (${errHint})` : `SSO: ${errHint}`;
      }
      if (userStatusDot) {
        userStatusDot.className = "user-status-dot offline";
        userStatusDot.title = isWifMode ? `WIF SSO Inactive (${errHint})` : `SSO: ${errHint}`;
      }
      if (userAuthBar) {
        userAuthBar.title = isWifMode ? `WIF SSO Inactive (${errHint})` : `SSO: ${errHint}`;
      }
    }

    // 2. Google OAuth button (Used only in GSuite / Cloud Identity mode, hidden in WIF mode)
    if (googleDriveBtn) {
      if (isWifMode) {
        // In WIF mode, user identity is purely Microsoft Entra ID exchanged silently with Google STS.
        // No secondary badge or button is shown.
        googleDriveBtn.style.display = "none";
      } else {
        googleDriveBtn.style.display = "inline-flex";
        if (isGoogleTokenValid()) {
          googleDriveBtn.className = "google-drive-btn connected";
          googleDriveBtn.textContent = "Gemini Connected";
          googleDriveBtn.title = "Gemini grounding active (OAuth token valid)";
          googleDriveBtn.style.cursor = "pointer";
        } else {
          googleDriveBtn.className = "google-drive-btn";
          googleDriveBtn.textContent = "Login with Google";
          googleDriveBtn.title = "Click to sign into Google for Gemini grounding";
          googleDriveBtn.style.cursor = "pointer";
        }
      }
    }
  } catch (err) {
    console.warn("Auth UI init error:", err);
    if (userEmailText) {
      userEmailText.innerText = "Unauthenticated";
      userEmailText.className = "user-email-text offline";
      userEmailText.title = "Auth error: " + (err.message || String(err));
    }
    if (userStatusDot) {
      userStatusDot.className = "user-status-dot offline";
      userStatusDot.title = "Auth error: " + (err.message || String(err));
    }
  } finally {
    updateDiagnosticsPanel().catch(() => {});
  }
}

let headlineIntervalId = null;

function setupRotatingHeadline() {
  const isOutlook = hostAdapter && hostAdapter.name === "Outlook";
  const headlines = isOutlook
    ? [
        "Summarize this email thread",
        "Draft a grounded reply",
        "Coach my email tone & clarity",
        "Extract action items & deadlines"
      ]
    : [
        "Create custom images",
        "Summarize this document",
        "Create an outline for a pitch",
        "Brainstorm a list of ideas"
      ];
  let headlineIndex = 0;
  const headlineEl = document.getElementById("headlineText");
  if (!headlineEl) return;

  headlineEl.innerText = headlines[0];
  if (headlineIntervalId) {
    clearInterval(headlineIntervalId);
  }

  headlineIntervalId = setInterval(() => {
    headlineIndex = (headlineIndex + 1) % headlines.length;
    headlineEl.style.opacity = '0';
    setTimeout(() => {
      headlineEl.innerText = headlines[headlineIndex];
      headlineEl.style.opacity = '1';
    }, 300);
  }, 4500);
}

function setupSuggestionPrompts() {
  const container = document.getElementById("suggestionsContainer");
  if (!container) return;

  container.querySelectorAll(".suggestion-item").forEach(item => {
    item.onclick = () => {
      const outlookAction = item.getAttribute("data-outlook-action");
      if (outlookAction && hostAdapter && hostAdapter.name === "Outlook") {
        runOutlookChatAction(outlookAction);
        return;
      }
      const text = item.getAttribute("data-suggestion") || item.innerText.trim();
      applySuggestion(text);
    };
  });
}

function applySuggestion(text) {
  const textarea = document.getElementById("promptText");
  if (!textarea) return;
  textarea.value = text;
  textarea.focus();
  textarea.style.height = 'auto';
  textarea.style.height = Math.min(textarea.scrollHeight, 120) + 'px';
  updateSendButtonState();
}

function updateSendButtonState() {
  const runBtn = document.getElementById("run");
  const promptInput = document.getElementById("promptText");
  if (!runBtn) return;

  const hasText = promptInput && promptInput.value.trim().length > 0;
  const hasAttachment = Boolean(attachedFileContent);

  if ((hasText || hasAttachment) && !isGenerating) {
    runBtn.classList.add("active");
  } else {
    runBtn.classList.remove("active");
  }
}

function updateChatDisplayState(hasMessages) {
  const centerHero = document.getElementById("centerHero");
  const suggestionsContainer = document.getElementById("suggestionsContainer");
  const emptyState = document.getElementById("emptyStateContainer");
  const historyDiv = document.getElementById("chatHistory");
  const docToolsBar = document.getElementById("docToolsBar");
  if (!historyDiv) return;

  const isOutlook = hostAdapter && hostAdapter.name === "Outlook";
  const isOutlookReadMode = isOutlook && !(typeof hostAdapter.isComposeMode === "function" && hostAdapter.isComposeMode());

  if (hasMessages) {
    if (centerHero) centerHero.style.display = "none";
    if (suggestionsContainer) suggestionsContainer.style.display = "none";
    if (emptyState) emptyState.style.display = "none";
    if (isOutlookReadMode && docToolsBar) docToolsBar.style.display = "none";
    historyDiv.style.flex = "";
    historyDiv.style.order = "";
    historyDiv.style.paddingBottom = "";
    historyDiv.style.display = "flex";
  } else {
    if (centerHero) centerHero.style.display = "flex";
    if (suggestionsContainer) suggestionsContainer.style.display = "flex";
    if (emptyState) emptyState.style.display = "flex";
    if (isOutlook && docToolsBar) docToolsBar.style.display = "flex";
    historyDiv.style.flex = "";
    historyDiv.style.order = "";
    historyDiv.style.paddingBottom = "";
    historyDiv.style.display = "none";
    historyDiv.innerHTML = "";
  }
}

function adaptUIForHost(hostName) {
  const tipBannerText = document.getElementById("tipBannerText");
  const docToolsBar = document.getElementById("docToolsBar");
  const docToolsTitle = document.getElementById("docToolsTitle");
  const welcomeBubble = document.getElementById("welcomeSystemBubble");
  const debugStatus = document.getElementById("debugStatus");
  const excelSchemaPreview = document.getElementById("excelSchemaPreview");
  const selectionAttachmentPill = document.getElementById("selectionAttachmentPill");
  const promptInput = document.getElementById("promptText");
  const suggestionsContainer = document.getElementById("suggestionsContainer");

  if (hostName === "PowerPoint") {
    if (excelSchemaPreview) excelSchemaPreview.style.display = "none";
    if (docToolsBar) docToolsBar.style.display = "flex";
    if (tipBannerText) tipBannerText.innerHTML = `💡 Type <code>@gemini &lt;prompt&gt;</code> in Slide`;
    if (docToolsTitle) docToolsTitle.innerHTML = `📊 <strong>Chat with Slides:</strong>`;
    if (debugStatus) debugStatus.innerText = `PowerPoint Ready`;
    if (welcomeBubble) welcomeBubble.innerHTML = `Type a prompt below, click an action chip above, or select slide content!`;
    if (promptInput) promptInput.placeholder = "Ask Gemini";
    if (suggestionsContainer) {
      suggestionsContainer.innerHTML = `
        <div class="suggestion-item" data-suggestion="Summarize key takeaways from these slides">
          <span class="return-arrow"><svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><polyline points="9 10 4 15 9 20"></polyline><path d="M20 4v7a4 4 0 0 1-4 4H4"></path></svg></span>
          <span>Summarize key takeaways from these slides</span>
        </div>
        <div class="suggestion-item" data-suggestion="Create an outline for a pitch deck">
          <span class="return-arrow"><svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><polyline points="9 10 4 15 9 20"></polyline><path d="M20 4v7a4 4 0 0 1-4 4H4"></path></svg></span>
          <span>Create an outline for a pitch deck</span>
        </div>
        <div class="suggestion-item" data-suggestion="Brainstorm visual slide ideas">
          <span class="return-arrow"><svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><polyline points="9 10 4 15 9 20"></polyline><path d="M20 4v7a4 4 0 0 1-4 4H4"></path></svg></span>
          <span>Brainstorm visual slide ideas</span>
        </div>
      `;
      setupSuggestionPrompts();
    }
  } else if (hostName === "Excel") {
    if (excelSchemaPreview) excelSchemaPreview.style.display = "flex";
    if (docToolsBar) docToolsBar.style.display = "flex";
    if (selectionAttachmentPill) selectionAttachmentPill.style.display = "none";
    const selChipsRow = document.querySelector("#selectionToolbar .selection-chips-row");
    if (selChipsRow) {
      selChipsRow.innerHTML = `
        <button class="selection-action-chip" id="selTableInsights">📈 Table Insights</button>
        <button class="selection-action-chip" id="selFormulaBuilder">✨ Formula Builder</button>
      `;
    }
    if (tipBannerText) tipBannerText.innerHTML = `💡 Type <code>@gemini &lt;prompt&gt;</code> in Cell`;
    if (docToolsTitle) docToolsTitle.innerHTML = `📊 <strong>Spreadsheet Tools & Analysis:</strong>`;
    if (debugStatus) debugStatus.innerText = `Excel Ready`;
    if (welcomeBubble) welcomeBubble.innerHTML = `Type a prompt below, request formulas or native charts, or select cells/charts in Excel!`;
    if (promptInput) promptInput.placeholder = "Ask Gemini";
    if (suggestionsContainer) {
      suggestionsContainer.innerHTML = `
        <div class="suggestion-item" data-suggestion="Analyze formulas and anomalies in this sheet">
          <span class="return-arrow"><svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><polyline points="9 10 4 15 9 20"></polyline><path d="M20 4v7a4 4 0 0 1-4 4H4"></path></svg></span>
          <span>Analyze formulas and anomalies in this sheet</span>
        </div>
        <div class="suggestion-item" data-suggestion="Suggest key insights and trends">
          <span class="return-arrow"><svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><polyline points="9 10 4 15 9 20"></polyline><path d="M20 4v7a4 4 0 0 1-4 4H4"></path></svg></span>
          <span>Suggest key insights and trends</span>
        </div>
        <div class="suggestion-item" data-suggestion="Draft an executive summary of this data">
          <span class="return-arrow"><svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><polyline points="9 10 4 15 9 20"></polyline><path d="M20 4v7a4 4 0 0 1-4 4H4"></path></svg></span>
          <span>Draft an executive summary of this data</span>
        </div>
      `;
      setupSuggestionPrompts();
    }
  } else if (hostName === "Outlook") {
    if (excelSchemaPreview) excelSchemaPreview.style.display = "none";
    if (docToolsBar) docToolsBar.style.display = "flex";
    if (selectionAttachmentPill) selectionAttachmentPill.style.display = "none";
    const diagPptGroup = document.getElementById("diagPptGroup");
    if (diagPptGroup) diagPptGroup.style.display = "none";
    const isCompose = hostAdapter && typeof hostAdapter.isComposeMode === "function" && hostAdapter.isComposeMode();
    const isAppointment = typeof Office !== "undefined" && Office.context?.mailbox?.item?.itemType === "appointment";
    if (promptInput) promptInput.placeholder = "Ask Gemini or type / for commands";
    if (suggestionsContainer) {
      const returnArrowSvg = `<span class="return-arrow"><svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><polyline points="9 10 4 15 9 20"></polyline><path d="M20 4v7a4 4 0 0 1-4 4H4"></path></svg></span>`;
      if (isCompose) {
        suggestionsContainer.innerHTML = `
          <div class="suggestion-item" data-outlook-action="coach" data-suggestion="Coach my email draft for tone, clarity, and deal risk">
            ${returnArrowSvg}
            <span>Coach my email draft for tone, clarity, and deal risk</span>
          </div>
          <div class="suggestion-item" data-outlook-action="reply" data-suggestion="Draft a grounded reply citing contract and SLA terms">
            ${returnArrowSvg}
            <span>Draft a grounded reply citing contract and SLA terms</span>
          </div>
          <div class="suggestion-item" data-outlook-action="actions" data-suggestion="Extract action items, deliverables, and key dates">
            ${returnArrowSvg}
            <span>Extract action items, deliverables, and key dates</span>
          </div>
        `;
      } else {
        suggestionsContainer.innerHTML = `
          <div class="suggestion-item" data-outlook-action="summarize" data-suggestion="Summarize this email thread and key action items">
            ${returnArrowSvg}
            <span>Summarize this email thread and key action items</span>
          </div>
          <div class="suggestion-item" data-outlook-action="reply" data-suggestion="Draft a grounded reply citing contract and SLA terms">
            ${returnArrowSvg}
            <span>Draft a grounded reply citing contract and SLA terms</span>
          </div>
          <div class="suggestion-item" data-outlook-action="actions" data-suggestion="Extract action items, deliverables, and key dates">
            ${returnArrowSvg}
            <span>Extract action items, deliverables, and key dates</span>
          </div>
        `;
      }
      setupSuggestionPrompts();
    }
  } else {
    // Word (Default)
    if (excelSchemaPreview) excelSchemaPreview.style.display = "none";
    if (docToolsBar) docToolsBar.style.display = "none";
    if (tipBannerText) tipBannerText.innerHTML = `💡 Type <code>@gemini &lt;prompt&gt;</code> in Doc`;
    if (docToolsTitle) docToolsTitle.innerHTML = `📄 <strong>Chat with Document:</strong>`;
    if (debugStatus) debugStatus.innerText = `Word Ready`;
    if (welcomeBubble) welcomeBubble.innerHTML = `Type a prompt below, click an action chip above, or highlight text in Word!`;
    if (promptInput) promptInput.placeholder = "Ask Gemini";
  }

  renderAdaptiveActionChips(false);
}

function showToast(msg) {
  const toast = document.getElementById("toastMsg");
  if (!toast) return;
  toast.innerText = msg;
  toast.classList.add("show");
  setTimeout(() => { toast.classList.remove("show"); }, 2200);
}

function resetChatSession() {
  const hostName = hostAdapter ? hostAdapter.name : 'Word';
  currentSessionId = null;
  chatHistoryState = [];
  askChatHistory = [];
  askSessionId = null;
  ["summarizeResultCard", "replyResultCard", "coachResultContainer", "actionItemsResultCard", "askResultCard"].forEach((id) => {
    const el = document.getElementById(id);
    if (el) el.style.display = "none";
  });
  const replyIntent = document.getElementById("replyIntentText");
  if (replyIntent) replyIntent.value = "";
  const askInput = document.getElementById("askQuestionText");
  if (askInput) askInput.value = "";
  if (typeof window !== "undefined") {
    window.__isSummarizeSlidesAction = false;
    window.__forceSingleSlide = false;
    window.__lastUserPrompt = "";
  }
  if (hostAdapter && hostAdapter.name === "Excel") {
    try {
      excelHostModule.onSessionReset({
        adapter: hostAdapter,
        onSelectionChanged: (rawText) => handleSelectionChanged(rawText)
      });
    } catch (e) {
      console.warn("Excel host session reset failed:", e);
    }
  }
  SessionManager.clearActiveSession(hostName);

  attachedFileContent = null;
  attachedFileName = null;
  attachedFileMimeType = null;
  const fileUploadInput = document.getElementById("fileUploadInput");
  const attachmentContainer = document.getElementById("attachmentContainer");
  if (fileUploadInput) fileUploadInput.value = "";
  if (attachmentContainer) attachmentContainer.style.display = "none";

  const promptInput = document.getElementById("promptText");
  if (promptInput) {
    promptInput.value = "";
    promptInput.style.height = "auto";
  }
  updateSendButtonState();

  updateChatDisplayState(false);
  showToast("Started new conversation");
}

function loadSessionIntoUI(sessionId) {
  const historyDiv = document.getElementById("chatHistory");
  if (!historyDiv || !sessionId) return;

  const hostName = hostAdapter ? hostAdapter.name : 'Word';
  const hostSessions = SessionManager.getAllSessions(hostName);
  const session = hostSessions.find(s => s.id === sessionId) || null;
  if (!session || !session.messages || session.messages.length === 0) {
    currentSessionId = null;
    SessionManager.clearActiveSession(hostName);
    updateChatDisplayState(false);
    return;
  }

  currentSessionId = sessionId;
  SessionManager.setActiveSessionId(sessionId, hostName);

  updateChatDisplayState(true);

  historyDiv.innerHTML = "";
  chatHistoryState = [];
  let lastUserPrompt = "";
  session.messages.forEach(msg => {
    if (msg.role === 'user') {
      lastUserPrompt = msg.text || "";
      chatHistoryState.push({ role: 'user', parts: [{ text: lastUserPrompt }] });
      appendBubble(msg.text, "user");
    } else if (msg.role === 'assistant') {
      chatHistoryState.push({ role: 'model', parts: [{ text: msg.text || "" }] });
      appendAssistantBubble(msg.text, { citations: msg.citations || [] }, lastUserPrompt);
    } else if (msg.role === 'error') {
      appendBubble(msg.text, "error");
    } else {
      appendBubble(msg.text, "system");
    }
  });
}

// Adaptive Action Bar & Chips Renderer
function renderAdaptiveActionChips(isSelected = false, slideCount = 1, words = 0, isTextSelection = false) {
  const hostName = hostAdapter ? hostAdapter.name : "Word";
  const docToolsTitle = document.getElementById("docToolsTitle");
  const debugStatus = document.getElementById("debugStatus");
  const container = document.getElementById("docToolsChipsContainer");
  const selectionPill = document.getElementById("selectionAttachmentPill");
  const selectionPillLabel = document.getElementById("selectionPillLabel");

  if (!container) return;

  if (hostName === "PowerPoint") {
    if (isSelected) {
      const labelText = isTextSelection 
        ? `Selected Text (${words} words)` 
        : (slideCount > 1 ? `${slideCount} Slides (${words} words)` : `Selected Content (${words} words)`);
      const barTitle = isTextSelection 
        ? `✨ <strong>Selected Text:</strong>` 
        : `✨ <strong>Selected Slides (${slideCount}):</strong>`;

      if (docToolsTitle) docToolsTitle.innerHTML = barTitle;
      if (debugStatus) debugStatus.innerHTML = `<span style="color:#0b57d0; font-weight:600;">${labelText}</span>`;

      if (selectionPill) {
        if (selectionPillLabel) selectionPillLabel.innerText = `Attached: ${labelText}`;
        selectionPill.style.display = "flex";
      }

      container.innerHTML = `
        <button class="quick-chip" id="chipSummarize">📊 Summarize Slides</button>
        <button class="quick-chip" id="chipGenImage">🎨 Create Image</button>
        <button class="quick-chip" id="chipCreateChart">📈 Create Chart</button>
        <button class="quick-chip" id="chipRisks">⚠️ Key Risks</button>
        <button class="quick-chip" id="chipRewrite">🪄 Rewrite Slide</button>
        <button class="quick-chip" id="chipDocToDeck">📄 Doc to Deck</button>
        <button class="quick-chip" id="chipShorten">📉 Shorten</button>
        <button class="quick-chip" id="chipTable">📊 Table</button>
      `;

      const cSummarize = document.getElementById("chipSummarize");
      if (cSummarize) cSummarize.onclick = () => runPowerPointSlideAction("takeaways");

      const cGenImage = document.getElementById("chipGenImage");
      if (cGenImage) {
        cGenImage.onclick = () => handleCreateImageClick();
      }

      const cCreateChart = document.getElementById("chipCreateChart");
      if (cCreateChart) {
        cCreateChart.onclick = () => handleCreateChartClick();
      }

      const cRisks = document.getElementById("chipRisks");
      if (cRisks) cRisks.onclick = () => runPowerPointSlideAction("risks");

      const cRewrite = document.getElementById("chipRewrite");
      if (cRewrite) cRewrite.onclick = () => runSelectionPrompt("Rewrite and elevate the selected slide content into clear, high-impact executive presentation prose.");

      const cDocToDeck = document.getElementById("chipDocToDeck");
      if (cDocToDeck) {
        cDocToDeck.onclick = () => {
          const fileInput = document.getElementById("docToDeckFileInput");
          if (fileInput) fileInput.click();
        };
      }

      const cShorten = document.getElementById("chipShorten");
      if (cShorten) cShorten.onclick = () => runSelectionPrompt("Make this slide content significantly more concise and punchy, removing unnecessary fluff.");

      const cTable = document.getElementById("chipTable");
      if (cTable) cTable.onclick = () => runSelectionPrompt("Convert this slide's metrics and structured data into a clean, executive markdown table.");

    } else {
      if (docToolsTitle) docToolsTitle.innerHTML = `📊 <strong>Chat with Slides:</strong>`;
      if (debugStatus) debugStatus.innerText = `PowerPoint Ready`;

      if (selectionPill) selectionPill.style.display = "none";

      container.innerHTML = `
        <button class="quick-chip" id="chipDocToDeck">📄 Doc to Deck</button>
        <button class="quick-chip" id="chipSummarize">📊 Summarize Slides</button>
        <button class="quick-chip" id="chipGenImage">🎨 Create Image</button>
        <button class="quick-chip" id="chipCreateChart">📈 Create Chart</button>
      `;

      const cDocToDeck = document.getElementById("chipDocToDeck");
      if (cDocToDeck) {
        cDocToDeck.onclick = () => {
          const fileInput = document.getElementById("docToDeckFileInput");
          if (fileInput) fileInput.click();
        };
      }

      const cSum = document.getElementById("chipSummarize");
      if (cSum) cSum.onclick = () => runPowerPointSlideAction("takeaways");

      const cGenImage = document.getElementById("chipGenImage");
      if (cGenImage) {
        cGenImage.onclick = () => handleCreateImageClick();
      }

      const cCreateChart = document.getElementById("chipCreateChart");
      if (cCreateChart) {
        cCreateChart.onclick = () => handleCreateChartClick();
      }
    }
  } else if (hostName === "Excel") {
    const chipsToRender = isSelected ? excelSelectionChips : excelDocChips;
    container.innerHTML = "";

    if (isSelected) {
      if (docToolsTitle) docToolsTitle.innerHTML = `✨ <strong>Selected Cells / Range:</strong>`;
      if (debugStatus) debugStatus.innerHTML = `<span style="color:#0b57d0; font-weight:600;">Selection Active</span>`;
    } else {
      if (docToolsTitle) docToolsTitle.innerHTML = `📊 <strong>Spreadsheet Tools & Analysis:</strong>`;
      if (debugStatus) debugStatus.innerText = `Excel Ready`;
    }

    if (selectionPill) selectionPill.style.display = "none";

    chipsToRender.forEach(chip => {
      const btn = document.createElement("button");
      btn.className = "quick-chip selection-action-chip";
      if (chip.id) btn.id = chip.id;
      btn.innerHTML = chip.label;
      if (chip.title) btn.title = chip.title;

      btn.onclick = async () => {
        if (chip.isChartBuilder || chip.id === "chipChartGen" || chip.id === "selChart") {
          await handleExcelChartBuilderClick(chip);
          return;
        }
        if (chip.isFormulaGenerator || chip.id === "chipFormulaGen" || chip.id === "selFormula") {
          await handleExcelFormulaGeneratorClick(chip);
          return;
        }
        clearChatAndInput();
        const loadingText = document.getElementById("loading");
        if (loadingText) {
          loadingText.innerText = (chip.isInsights || chip.isExecSummary)
            ? "⚡ Pre-parsing Excel table data & calculating statistics..."
            : "⚡ Analyzing spreadsheet context...";
          loadingText.style.display = "block";
        }

        try {
          let fullPrompt = "";
          let displayBubble = "";

          if (chip.isInsights && typeof hostAdapter.getTableInsightsPrompt === 'function') {
            fullPrompt = await hostAdapter.getTableInsightsPrompt(chip.scope || (isSelected ? "selection" : "auto"));
            const scopeLabel = chip.scope === "selection" || isSelected ? "Highlighted Cells" : "Active Range / Sheet";
            displayBubble = `📊 [Table Insights] Scope: ${scopeLabel}`;
          } else if (chip.isExecSummary && typeof hostAdapter.getExecutiveSummaryPrompt === 'function') {
            fullPrompt = await hostAdapter.getExecutiveSummaryPrompt(chip.scope || (isSelected ? "selection" : "auto"));
            const scopeLabel = chip.scope === "selection" || isSelected ? "Highlighted Cells" : "Active Range / Sheet";
            displayBubble = `📑 [Executive Summary] Scope: ${scopeLabel}`;
          } else {
            const docContext = await hostAdapter.getFullDocumentText();
            const instruction = chip.prompt || chip.instruction || "Analyze the spreadsheet data.";
            fullPrompt = docContext && docContext.length > 20
              ? `Spreadsheet Context:\n"""\n${docContext.substring(0, 100000)}\n"""\n\nUser Instruction: ${instruction}`
              : instruction;
            displayBubble = `📊 [${chip.label.replace(/^[^\w]+/, '').trim()}] Analyzing spreadsheet...`;
          }

          await executeGeminiWorkflow(fullPrompt, displayBubble);
        } catch (err) {
          console.error("Excel chip execution error:", err);
          appendBubble("Error executing action: " + err.message, "system");
        } finally {
          if (loadingText) loadingText.style.display = "none";
        }
      };

      container.appendChild(btn);
    });
  } else if (hostName === "Outlook") {
    const isCompose = hostAdapter && typeof hostAdapter.isComposeMode === "function" && hostAdapter.isComposeMode();
    const isAppointment = typeof Office !== "undefined" && Office.context?.mailbox?.item?.itemType === "appointment";
    if (docToolsTitle) {
      docToolsTitle.innerHTML = isAppointment
        ? `📅 <strong>Meeting Assistant:</strong>`
        : (isCompose ? `✍️ <strong>Compose &amp; Coach:</strong>` : `📧 <strong>Email Assistant:</strong>`);
    }
    if (debugStatus) {
      debugStatus.innerText = "";
      debugStatus.style.display = "none";
    }
    if (selectionPill) selectionPill.style.display = "none";

    container.innerHTML = `
      <button class="quick-chip" id="chipOutlookSummarize" style="background-color:#e8f0fe; color:#0b57d0; border-color:#c2e7ff; font-weight:600;">📋 Summarize</button>
      <button class="quick-chip" id="chipOutlookReply" style="background-color:#f0fdf4; color:#137333; border-color:#bbf7d0; font-weight:600;">✉️ Draft Reply</button>
      <button class="quick-chip" id="chipOutlookCoach" style="background-color:#fef7e0; color:#b06000; border-color:#fde68a; font-weight:600; ${isCompose ? '' : 'display:none;'}">🎯 Coach</button>
      <button class="quick-chip" id="chipOutlookActions">✅ Action Items</button>
    `;

    const cSum = document.getElementById("chipOutlookSummarize");
    if (cSum) cSum.onclick = () => runOutlookChatAction("summarize");

    const cReply = document.getElementById("chipOutlookReply");
    if (cReply) cReply.onclick = () => runOutlookChatAction("reply");

    const cCoach = document.getElementById("chipOutlookCoach");
    if (cCoach) cCoach.onclick = () => runOutlookChatAction("coach");

    const cActions = document.getElementById("chipOutlookActions");
    if (cActions) cActions.onclick = () => runOutlookChatAction("actions");
  } else {
    // Word
    if (isSelected) {
      if (docToolsTitle) docToolsTitle.innerHTML = `✨ <strong>Selected Text:</strong>`;
      if (debugStatus) debugStatus.innerHTML = `<span style="color:#0b57d0; font-weight:600;">${words} words</span>`;
      if (selectionPill) {
        if (selectionPillLabel) selectionPillLabel.innerText = `Selected Text (${words} words)`;
        selectionPill.style.display = "flex";
      }
      container.innerHTML = `
        <button class="quick-chip" id="selRewrite">🪄 Rewrite</button>
        <button class="quick-chip" id="selProfessional">💼 Professional</button>
        <button class="quick-chip" id="selShorten">📉 Shorten</button>
        <button class="quick-chip" id="selBullets">🎯 Bullet Points</button>
        <button class="quick-chip" id="selTable">📊 Table</button>
      `;
      setupSelectionChips();
    } else {
      if (docToolsTitle) docToolsTitle.innerHTML = `📄 <strong>Chat with Document:</strong>`;
      if (debugStatus) debugStatus.innerText = `Word Ready`;
      if (selectionPill) selectionPill.style.display = "none";

      container.innerHTML = `
        <button class="quick-chip" id="chipSummarize">📄 Summarize Doc</button>
        <button class="quick-chip" id="chipRisks">⚠️ Key Risks</button>
        <button class="quick-chip" id="chipActionItems">✅ Action Items</button>
        <button class="quick-chip" id="chipExecBox">💡 Executive Card</button>
      `;
      setupDocToolsChips();
    }
  }
}

// Fallback / Initial Chip Bindings for Word & Excel
function setupDocToolsChips() {
  const hostName = hostAdapter ? hostAdapter.name : "Word";

  const chipSummarize = document.getElementById("chipSummarize");
  if (chipSummarize) {
    chipSummarize.onclick = () => {
      if (hostName === "PowerPoint") {
        runPowerPointSlideAction("summarize");
      } else if (hostName === "Excel") {
        runDocIntelligencePrompt("Analyze and summarize this spreadsheet data. Provide key patterns, trends, data anomalies, and an executive summary table.");
      } else {
        runDocIntelligencePrompt("Summarize this document thoroughly. Provide an executive overview, key takeaways, and a structured markdown summary table.");
      }
    };
  }

  const chipRisks = document.getElementById("chipRisks");
  if (chipRisks) {
    chipRisks.onclick = () => {
      if (hostName === "PowerPoint") {
        runPowerPointSlideAction("risks");
      } else if (hostName === "Excel") {
        runDocIntelligencePrompt("Analyze this spreadsheet data and extract all Key Financial / Operational Risks, Outliers, and Data Gaps with Mitigation Suggestions.");
      } else {
        runDocIntelligencePrompt("Analyze this document and extract all Key Risks, Ambiguities, and Operational Gaps. Present them in a structured markdown table with Severity and Mitigation Suggestions.");
      }
    };
  }

  const chipActionItems = document.getElementById("chipActionItems");
  if (chipActionItems) {
    chipActionItems.onclick = () => {
      if (hostName === "PowerPoint") {
        runPowerPointSlideAction("action_items");
      } else {
        const noun = hostName === "Excel" ? "spreadsheet" : "document";
        runDocIntelligencePrompt(`Extract all Action Items, Deliverables, and Next Steps from this ${noun}. Present them in a structured table with Task, Owner, and Priority.`);
      }
    };
  }

  const chipExecBox = document.getElementById("chipExecBox");
  if (chipExecBox) {
    chipExecBox.onclick = () => {
      if (hostName === "PowerPoint") {
        runPowerPointSlideAction("takeaways");
      } else if (hostName === "Excel") {
        runDocIntelligencePrompt("Generate an Executive Metrics Summary Card highlighting Key Financial / Operational KPIs and totals from this sheet.");
      } else {
        runDocIntelligencePrompt("Generate an Executive Summary Box for this document using a callout block (> [!NOTE] ...) highlighting Strategic Purpose, Key Metrics, and Impact.");
      }
    };
  }

  const chipDocToDeck = document.getElementById("chipDocToDeck");
  if (chipDocToDeck) {
    chipDocToDeck.onclick = () => {
      const fileInput = document.getElementById("docToDeckFileInput");
      if (fileInput) fileInput.click();
    };
  }
}

function switchToSession(sessionId) {
  loadSessionIntoUI(sessionId);
  toggleRecentChats(false);
  showToast("Loaded chat session");
}

function escapeHtml(str) {
  if (!str) return '';
  return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function isSessionToday(ts) {
  if (!ts) return false;
  const d = new Date(ts);
  const now = new Date();
  return d.toDateString() === now.toDateString();
}

function toggleRecentChats(show) {
  const recentView = document.getElementById("recentChatsView");
  const drawer = document.getElementById("historyDrawer");
  const overlay = document.getElementById("historyOverlay");
  const btn = document.getElementById("historyBtn");

  const target = recentView || drawer;
  if (!target) return;

  const shouldOpen = typeof show === 'boolean' ? show : !target.classList.contains("open");
  if (shouldOpen) {
    target.classList.add("open");
    if (overlay) overlay.classList.add("open");
    if (btn) btn.classList.add("active");
    renderHistoryList();
  } else {
    target.classList.remove("open");
    if (overlay) overlay.classList.remove("open");
    if (btn) btn.classList.remove("active");
  }
}

function closeHistoryDrawer() {
  toggleRecentChats(false);
}

async function renderHistoryList() {
  const historyList = document.getElementById("recentChatsList") || document.getElementById("historyList");
  const drawerTitle = document.getElementById("historyDrawerTitle");
  if (!historyList) return;

  const hostName = hostAdapter ? hostAdapter.name : 'Word';
  if (drawerTitle) {
    drawerTitle.innerText = `Recent chats (${hostName})`;
  }

  historyList.innerHTML = '<div style="text-align:center; padding:16px; color:#5f6368; font-size:12px;">Loading conversations...</div>';

  try {
    let remoteSessions = [];
    try {
      remoteSessions = await fetchEnterpriseSessions(20, hostName);
    } catch (e) {
      console.warn("Could not fetch remote Discovery Engine sessions:", e);
    }

    const remoteMetaMap = new Map();
    remoteSessions.forEach(rs => {
      if (!rs || !rs.sessionId) return;
      remoteMetaMap.set(rs.sessionId, rs);
      SessionManager.syncRemoteSession(rs);
    });

    // Surface Isolation: Strictly retrieve sessions belonging ONLY to the active host (Word, PowerPoint, Excel, Outlook)
    const hostSessions = SessionManager.getAllSessions(hostName);
    const mergedList = hostSessions
      .filter(ls => ls && ls.host && ls.host.toLowerCase() === hostName.toLowerCase())
      .map(ls => {
        const rs = remoteMetaMap.get(ls.id);
        let title = ls.title || 'Untitled Conversation';
        if (rs && rs.displayName && rs.displayName !== 'Office 365 Session' && (!title || title === 'New Conversation' || title === 'Conversation' || title === 'Untitled Conversation')) {
          title = rs.displayName;
        }
        return {
          id: ls.id,
          title,
          turnsCount: (rs && rs.turnsCount) || (ls.messages ? Math.max(1, Math.floor(ls.messages.length / 2)) : 1),
          updatedAt: ls.updatedAt || Date.now(),
          isLocal: !rs,
          host: hostName
        };
      });
    mergedList.sort((a, b) => b.updatedAt - a.updatedAt);

    if (mergedList.length === 0) {
      historyList.innerHTML = `<div style="text-align:center; padding:20px 16px; color:#5f6368; font-size:13px;">No previous ${hostName} chats found.</div>`;
      return;
    }

    historyList.innerHTML = "";

    const todayItems = mergedList.filter(s => isSessionToday(s.updatedAt));
    const earlierItems = mergedList.filter(s => !isSessionToday(s.updatedAt));

    function appendGroup(title, items) {
      if (items.length === 0) return;
      const label = document.createElement("div");
      label.className = "recent-section-label";
      label.innerText = title;
      historyList.appendChild(label);

      items.forEach(s => {
        const itemEl = document.createElement("div");
        const isActive = s.id === currentSessionId;
        itemEl.className = `recent-row ${isActive ? 'active' : ''}`;

        itemEl.innerHTML = `
          <span class="recent-row-icon">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="#444746">
              <path d="M4 6h16v2H4zm0 5h11v2H4zm0 5h6v2H4z"/>
            </svg>
          </span>
          <span class="recent-row-text" title="${escapeHtml(s.title)}">${escapeHtml(s.title)}</span>
          <button class="recent-row-del" title="Delete conversation" data-id="${s.id}">🗑️</button>
        `;

        itemEl.onclick = (e) => {
          if (e.target.closest(".recent-row-del")) return;
          switchToSession(s.id);
        };

        const delBtn = itemEl.querySelector(".recent-row-del");
        if (delBtn) {
          delBtn.onclick = async (e) => {
            e.stopPropagation();
            SessionManager.deleteSession(s.id);
            deleteEnterpriseSession(s.id);
            if (currentSessionId === s.id) {
              resetChatSession();
            }
            itemEl.remove();
            if (historyList.querySelectorAll(".recent-row").length === 0) {
              historyList.innerHTML = `<div style="text-align:center; padding:20px 16px; color:#5f6368; font-size:13px;">No previous ${hostName} chats found.</div>`;
            }
          };
        }

        historyList.appendChild(itemEl);
      });
    }

    appendGroup("Today", todayItems);
    appendGroup("Earlier", earlierItems);

  } catch (err) {
    historyList.innerHTML = `<div style="color:#b3261e; padding:12px; font-size:12px;">Error loading history: ${err.message}</div>`;
  }
}

// Setup Selection Quick Toolbar Chips (Feature: Inline Tone Rewrites & Polish)
function setupSelectionChips() {
  const isExcel = Boolean(hostAdapter && hostAdapter.name === "Excel");
  if (isExcel) {
    const selChipsRow = document.querySelector("#selectionToolbar .selection-chips-row");
    if (selChipsRow && !document.getElementById("selTableInsights")) {
      selChipsRow.innerHTML = `
        <button class="selection-action-chip" id="selTableInsights">📈 Table Insights</button>
        <button class="selection-action-chip" id="selFormulaBuilder">✨ Formula Builder</button>
      `;
    }

    const selTableInsights = document.getElementById("selTableInsights");
    if (selTableInsights) {
      selTableInsights.onclick = async () => {
        if (isGenerating || isProcessingInDocCommand) return;
        const selToolbar = document.getElementById("selectionToolbar");
        if (selToolbar) selToolbar.style.display = "none";

        const loadingText = document.getElementById("loading");
        if (loadingText) {
          loadingText.innerText = "⚡ Pre-parsing Excel table data & calculating statistics...";
          loadingText.style.display = "block";
        }

        try {
          let fullPrompt = "";
          if (hostAdapter && typeof hostAdapter.getTableInsightsPrompt === "function") {
            fullPrompt = await hostAdapter.getTableInsightsPrompt("selection");
          } else {
            const ctxText = hostAdapter ? await hostAdapter.getSelectedText() : currentSelectedText;
            fullPrompt = `Selected Spreadsheet Data:\n"""\n${ctxText}\n"""\n\nUser Instruction: Perform deep Table Insights on the highlighted cells with summary statistics, anomalies, key trends, and a visual graph directive.`;
          }
          const displayBubble = `📊 [Table Insights] Scope: Highlighted Cells`;
          await executeGeminiWorkflow(fullPrompt, displayBubble, false);
        } catch (err) {
          console.error("Excel Table Insights selection chip error:", err);
          appendBubble("Error executing Table Insights: " + err.message, "system");
        } finally {
          if (loadingText) loadingText.style.display = "none";
        }
      };
    }

    const selFormulaBuilder = document.getElementById("selFormulaBuilder");
    if (selFormulaBuilder) {
      selFormulaBuilder.onclick = async () => {
        if (isGenerating || isProcessingInDocCommand) return;
        const selToolbar = document.getElementById("selectionToolbar");
        if (selToolbar) selToolbar.style.display = "none";

        const loadingText = document.getElementById("loading");
        if (loadingText) {
          loadingText.innerText = "⚡ Analyzing spreadsheet context...";
          loadingText.style.display = "block";
        }

        try {
          const selContext = (hostAdapter && typeof hostAdapter.getSelectedText === "function")
            ? await hostAdapter.getSelectedText()
            : currentSelectedText;
          const docContext = (hostAdapter && typeof hostAdapter.getFullDocumentText === "function")
            ? await hostAdapter.getFullDocumentText()
            : "";
          const contextParts = [];
          if (selContext && selContext.trim().length > 0) {
            contextParts.push(`Selected Data / Range Context:\n"""\n${selContext.substring(0, 50000)}\n"""`);
          }
          if (docContext && docContext.trim().length > 0 && docContext !== selContext) {
            contextParts.push(`Spreadsheet Context:\n"""\n${docContext.substring(0, 100000)}\n"""`);
          }
          const instruction = "Generate an optimal Excel formula (XLOOKUP, SUMIFS, INDEX/MATCH, AVERAGEIFS, etc.) for the selected cells/range and active context based on the worksheet column schema and data. Return the raw formula formatted in a code block.";
          const fullPrompt = contextParts.length > 0
            ? `${contextParts.join("\n\n")}\n\nUser Instruction: ${instruction}`
            : instruction;
          const displayBubble = `✨ [Formula Builder] Scope: Highlighted Cells`;
          await executeGeminiWorkflow(fullPrompt, displayBubble, false);
        } catch (err) {
          console.error("Excel Formula Builder selection chip error:", err);
          appendBubble("Error executing Formula Builder: " + err.message, "system");
        } finally {
          if (loadingText) loadingText.style.display = "none";
        }
      };
    }

    const clearSelectionBtn = document.getElementById("clearSelectionToolbarBtn");
    if (clearSelectionBtn) {
      clearSelectionBtn.onclick = (e) => {
        e.stopPropagation();
        const selToolbar = document.getElementById("selectionToolbar");
        if (selToolbar) selToolbar.style.display = "none";
        currentSelectedText = "";
      };
    }
    return;
  }

  const selRewrite = document.getElementById("selRewrite");
  if (selRewrite) {
    selRewrite.onclick = () => runSelectionPrompt(
      "Rewrite and rephrase the selected text from scratch with a fresh structure, varied vocabulary, and enhanced clarity while preserving its original meaning.",
      "Rewrite",
      "text"
    );
  }

  const selFormal = document.getElementById("selFormal") || document.getElementById("selProfessional");
  if (selFormal) {
    selFormal.onclick = () => runSelectionPrompt(
      "Elevate the tone of the selected text to authoritative, formal executive business prose suitable for C-level leadership.",
      "Formal tone",
      "text"
    );
  }

  const selCasual = document.getElementById("selCasual");
  if (selCasual) {
    selCasual.onclick = () => runSelectionPrompt(
      "Rewrite the selected text in an approachable, friendly, and conversational tone that is engaging and easy to read while retaining all core points.",
      "Casual tone",
      "text"
    );
  }

  const selShorten = document.getElementById("selShorten");
  if (selShorten) {
    selShorten.onclick = () => runSelectionPrompt(
      "Condense and tighten the selected text to be concise, punchy, and high-impact, eliminating all unnecessary filler while keeping key decisions and metrics.",
      "Shorten",
      "text"
    );
  }

  const selElaborate = document.getElementById("selElaborate");
  if (selElaborate) {
    selElaborate.onclick = () => runSelectionPrompt(
      "Expand and elaborate on the selected text with supporting context, operational specifics, and analytical depth.",
      "Elaborate",
      "text"
    );
  }

  const selBullets = document.getElementById("selBullets");
  if (selBullets) {
    selBullets.onclick = () => runSelectionPrompt(
      "Convert the key information, arguments, and takeaways in the selected text into clean, high-impact bullet points (* or -).",
      "Bullet points",
      "bullets"
    );
  }

  const selTable = document.getElementById("selTable");
  if (selTable) {
    selTable.onclick = () => runSelectionPrompt(
      "Extract and organize the key comparative data, metrics, entities, and categories in the selected text into a structured markdown table with clear column headers.",
      "Table",
      "table"
    );
  }

  const clearSelectionBtn = document.getElementById("clearSelectionToolbarBtn");
  if (clearSelectionBtn) {
    clearSelectionBtn.onclick = (e) => {
      e.stopPropagation();
      const selToolbar = document.getElementById("selectionToolbar");
      if (selToolbar) selToolbar.style.display = "none";
      currentSelectedText = "";
    };
  }
}

async function runSelectionPrompt(instruction, displayLabel = null, optionType = "text") {
  if (isGenerating || isProcessingInDocCommand) return;
  isGenerating = true;
  const runButton = document.getElementById("run");
  if (runButton) runButton.disabled = true;

  try {
    if (!currentSelectedText && hostAdapter) {
      currentSelectedText = (hostAdapter.name === "Excel" && typeof hostAdapter.getSelectedRawText === "function")
        ? await hostAdapter.getSelectedRawText()
        : await hostAdapter.getSelectedText();
    }

    if (!currentSelectedText) {
      const hostName = hostAdapter ? hostAdapter.name : 'Word';
      appendBubble(`Please highlight text in ${hostName} first.`, "system");
      return;
    }

    let formatRules = "";
    if (optionType === "table") {
      // Output Format Option: Structured Markdown Table
      formatRules = [
        "CRITICAL FORMATTING INSTRUCTIONS (OUTPUT FORMAT):",
        "- Convert the selected content directly into a clean, structured markdown table with descriptive column headers.",
        "- Provide exactly ONE single markdown table ready for immediate document replacement.",
        "- Output ONLY the markdown table without conversational preamble (e.g., do NOT write 'Here is the table:'), introductory comments, or concluding remarks.",
        "- Do NOT provide multiple alternative tables or options."
      ].join("\n");
    } else if (optionType === "bullets") {
      // Output Format Option: Bullet Points List
      formatRules = [
        "CRITICAL FORMATTING INSTRUCTIONS (OUTPUT FORMAT):",
        "- Convert the selected content directly into clean, high-impact markdown bullet points (* or -).",
        "- Provide exactly ONE single, cohesive bulleted list ready for immediate document replacement.",
        "- Output ONLY the bullet points without conversational preamble (e.g., do NOT write 'Here are the bullet points:'), introductory comments, or concluding remarks.",
        "- Do NOT provide multiple alternative options or variations.",
        "- Do NOT generate a table."
      ].join("\n");
    } else {
      // Text Content & Tone Option: Prose rewrite, professional tone, shorten
      formatRules = [
        "CRITICAL FORMATTING INSTRUCTIONS (TEXT REWRITE):",
        "- Provide exactly ONE single, definitive drop-in replacement version in natural prose.",
        "- Output ONLY the revised replacement text ready for immediate document replacement.",
        "- Do NOT output multiple numbered options (e.g., do NOT write 'Option 1', 'Option 2', or alternative choices).",
        "- Do NOT include conversational greetings, preamble, quotes, callout boxes, or concluding remarks.",
        "- Do NOT format as bullet points or tables unless explicitly asked."
      ].join("\n");
    }

    const selectionContextLabel = hostAdapter && hostAdapter.name === "PowerPoint"
      ? "Selected Slide Context"
      : (hostAdapter && hostAdapter.name === "Excel" ? "Selected Spreadsheet Data" : "Selected Document Text");
    let fullPrompt = `${selectionContextLabel}:\n"${currentSelectedText}"\n\nTask: ${instruction}\n\n${formatRules}`;
    if (hostAdapter && hostAdapter.name === "PowerPoint") {
      const { enhancePromptForPowerPoint } = await import('../adapters/ppt/promptEnhancer.js');
      fullPrompt = enhancePromptForPowerPoint(fullPrompt, { userInstruction: instruction });
    }
    const displayUserBubble = displayLabel 
      ? `✨ ${displayLabel}\n📌 Context: "${currentSelectedText.substring(0, 70)}${currentSelectedText.length > 70 ? '...' : ''}"`
      : `✨ ${instruction}\n📌 Context: "${currentSelectedText.substring(0, 70)}${currentSelectedText.length > 70 ? '...' : ''}"`;

    // Hide the selection toolbar once clicked
    const selToolbar = document.getElementById("selectionToolbar");
    if (selToolbar) selToolbar.style.display = "none";

    await executeGeminiWorkflow(fullPrompt, displayUserBubble, false);
  } finally {
    isGenerating = false;
    if (runButton) runButton.disabled = false;
  }
}

// Handle Word & Excel Selection Changes Dynamically (PowerPoint reads on-demand)
async function handleSelectionChanged(prefetchedText = null) {
  if (!hostAdapter || hostAdapter.name === "PowerPoint") return;
  const selToolbar = document.getElementById("selectionToolbar");
  const previewText = document.getElementById("selectionPreviewText");
  const wordCountSpan = document.getElementById("selectionWordCount");

  try {
    const isExcel = hostAdapter.name === "Excel";
    const text = (typeof prefetchedText === "string")
      ? prefetchedText
      : (isExcel && typeof hostAdapter.getSelectedRawText === "function"
          ? await hostAdapter.getSelectedRawText()
          : await hostAdapter.getSelectedText());
    currentSelectedText = text ? text.trim() : "";

    const hasSelectedText = isExcel
      ? currentSelectedText.length > 0
      : currentSelectedText.length > 5;

    if (hasSelectedText) {
      const words = currentSelectedText.split(/\s+/).filter(w => w.length > 0).length;
      if (wordCountSpan) wordCountSpan.innerText = `(${words} word${words === 1 ? '' : 's'})`;
      if (previewText) {
        previewText.innerText = `"${currentSelectedText.substring(0, 80)}${currentSelectedText.length > 80 ? '...' : ''}"`;
      }
      if (selToolbar) selToolbar.style.display = "block";
    } else {
      if (selToolbar) selToolbar.style.display = "none";
    }

    // Check for in-document @gemini command
    await checkForInDocumentCommands(false);
  } catch (err) {
    console.warn("Selection change handler error:", err);
  }
}

// Document Intelligence Functions (Used by Slash Commands)
export function executeSummarizeDoc(isSlashCommand = false) {
  const hostName = hostAdapter ? hostAdapter.name : "Word";
  const label = isSlashCommand ? "⚡ /summarize-doc" : null;
  if (hostName === "PowerPoint") {
    runDocIntelligencePrompt("Summarize these presentation slides thoroughly. Provide an executive overview of all slides, key themes, and a structured summary table.", label);
  } else if (hostName === "Excel") {
    runDocIntelligencePrompt("Analyze and summarize this spreadsheet data. Provide key patterns, trends, data anomalies, and an executive summary table.", label);
  } else {
    runDocIntelligencePrompt("Summarize this document thoroughly. Provide an executive overview, key takeaways, and a structured markdown summary table.", label);
  }
}

export function executeIdentifyRisks(isSlashCommand = false) {
  const hostName = hostAdapter ? hostAdapter.name : "Word";
  const label = isSlashCommand ? "⚡ /identify-risks" : null;
  if (hostName === "PowerPoint") {
    runDocIntelligencePrompt("Analyze these presentation slides and extract all Key Strategic Risks, Challenges, and Operational Gaps with Mitigation Suggestions.", label);
  } else if (hostName === "Excel") {
    runDocIntelligencePrompt("Analyze this spreadsheet data and extract all Key Financial / Operational Risks, Outliers, and Data Gaps with Mitigation Suggestions.", label);
  } else {
    runDocIntelligencePrompt("Analyze this document and extract all Key Risks, Ambiguities, and Operational Gaps. Present them in a structured markdown table with Severity and Mitigation Suggestions.", label);
  }
}

export function executeIdentifyAction(isSlashCommand = false) {
  const hostName = hostAdapter ? hostAdapter.name : "Word";
  const noun = hostName === "PowerPoint" ? "presentation deck" : (hostName === "Excel" ? "spreadsheet" : "document");
  const label = isSlashCommand ? "⚡ /identify-action" : null;
  runDocIntelligencePrompt(`Extract all Action Items, Deliverables, and Next Steps from this ${noun}. Present them in a structured table with Task, Owner, and Priority.`, label);
}

export function executeExecCard(isSlashCommand = false) {
  const hostName = hostAdapter ? hostAdapter.name : "Word";
  const label = isSlashCommand ? "⚡ /exec-card" : null;
  if (hostName === "PowerPoint") {
    runDocIntelligencePrompt("Generate an Executive Slide Takeaway Callout Box highlighting Strategic Impact and Key Metrics for this deck.", label);
  } else if (hostName === "Excel") {
    runDocIntelligencePrompt("Generate an Executive Metrics Summary Card highlighting Key Financial / Operational KPIs and totals from this sheet.", label);
  } else {
    runDocIntelligencePrompt("Generate an Executive Summary Box for this document using a callout block (> [!NOTE] ...) highlighting Strategic Purpose, Key Metrics, and Impact.", label);
  }
}

async function runDocIntelligencePrompt(instruction, displayLabel = null) {
  if (isGenerating || isProcessingInDocCommand) return;
  isGenerating = true;
  const runButton = document.getElementById("run");
  if (runButton) runButton.disabled = true;

  try {
    const loadingText = document.getElementById("loading");
    if (loadingText) {
      loadingText.innerText = "⚡ Reading document context...";
      loadingText.style.display = "flex";
    }

    const fullDocText = await hostAdapter.getFullDocumentText();
    let fullPrompt = "";

    if (fullDocText && fullDocText.length > 20) {
      const docContextLabel = hostAdapter && hostAdapter.name === "PowerPoint"
        ? "Presentation Context"
        : (hostAdapter && hostAdapter.name === "Excel" ? "Spreadsheet Context" : "Active Document Context");
      fullPrompt = `${docContextLabel}:\n"""\n${fullDocText.substring(0, 100000)}\n"""\n\nUser Instruction: ${instruction}\n\nConstraint: Answer strictly based on the provided ${docContextLabel}. Do not reference external datastores.`;
    } else {
      fullPrompt = instruction;
    }

    const displayUserBubble = displayLabel || `📄 [Document Analysis] ${instruction.substring(0, 55)}...`;
    await executeGeminiWorkflow(fullPrompt, displayUserBubble, false);
  } finally {
    isGenerating = false;
    if (runButton) runButton.disabled = false;
  }
}

// Feature: Slash Commands Autocomplete & Execution
export const SLASH_COMMANDS = [
  {
    command: "/summarize-doc",
    label: "Summarize Doc",
    icon: "📑",
    description: "Executive overview, key takeaways, and summary table",
    action: () => executeSummarizeDoc(true)
  },
  {
    command: "/identify-risks",
    label: "Key Risks",
    icon: "⚠️",
    description: "Extract strategic, operational, and financial risks",
    action: () => executeIdentifyRisks(true)
  },
  {
    command: "/identify-action",
    label: "Action Items",
    icon: "✅",
    description: "Extract action items, deliverables, and next steps",
    action: () => executeIdentifyAction(true)
  },
  {
    command: "/exec-card",
    label: "Exec Card",
    icon: "💡",
    description: "Generate executive summary card with purpose & metrics",
    action: () => executeExecCard(true)
  }
];

export const OUTLOOK_SLASH_COMMANDS = [
  {
    command: "/summarize",
    label: "Summarize Thread",
    icon: "📋",
    description: "3-bullet executive briefing & assigned action items",
    action: () => runOutlookChatAction("summarize")
  },
  {
    command: "/reply",
    label: "Draft Grounded Reply",
    icon: "✍️",
    description: "Draft a reply grounded in contract & SLA terms",
    action: () => runOutlookChatAction("reply")
  },
  {
    command: "/coach",
    label: "Coach My Email",
    icon: "🎯",
    description: "Review draft tone & generate Executive + Diplomatic options",
    action: () => runOutlookChatAction("coach")
  },
  {
    command: "/actions",
    label: "Action Items",
    icon: "✅",
    description: "Extract action items, owners, and due dates table",
    action: () => runOutlookChatAction("actions")
  }
];

function getActiveSlashCommands() {
  return (hostAdapter && hostAdapter.name === "Outlook") ? OUTLOOK_SLASH_COMMANDS : SLASH_COMMANDS;
}

let activeSlashIndex = 0;
let currentMatchingCommands = [...SLASH_COMMANDS];

export function hideSlashPopup() {
  const popup = document.getElementById("slashCommandsPopup");
  if (popup) {
    popup.style.display = "none";
  }
}

// The slash menu opens upwards from the chat box and .main-canvas clips
// overflow, so on short task panes (laptop screens, or below PowerPoint's doc
// tools bar) the top of the menu would be cut off. Cap it at the room left
// above the chat box; the command list scrolls inside that height.
function fitSlashPopupToCanvas(popup) {
  const anchor = document.getElementById("chatBoxWrapper");
  const canvas = document.getElementById("mainCanvas");
  if (!anchor || !canvas) return;
  // 8px gap above the chat box (see .slash-popup) + 8px clearance at the top.
  const room = anchor.getBoundingClientRect().top - canvas.getBoundingClientRect().top - 16;
  popup.style.maxHeight = `${Math.max(Math.floor(room), 120)}px`;
}

export function showSlashPopup() {
  const popup = document.getElementById("slashCommandsPopup");
  if (popup) {
    popup.style.display = "flex";
    fitSlashPopupToCanvas(popup);
  }
}

function updateActiveSlashItem() {
  const list = document.getElementById("slashCommandsList");
  if (!list) return;
  const items = list.querySelectorAll(".slash-command-item");
  items.forEach((item, idx) => {
    if (idx === activeSlashIndex) {
      item.classList.add("active");
      item.setAttribute("aria-selected", "true");
      item.scrollIntoView({ block: "nearest" });
    } else {
      item.classList.remove("active");
      item.setAttribute("aria-selected", "false");
    }
  });
}

function renderSlashPopupItems(commands, activeIdx) {
  const list = document.getElementById("slashCommandsList");
  if (!list) return;

  // Secure DOM update using replaceChildren()
  list.replaceChildren();

  if (commands.length === 0) {
    const empty = document.createElement("div");
    empty.className = "slash-empty-item";
    empty.textContent = "No matching slash commands";
    list.appendChild(empty);
    return;
  }

  commands.forEach((cmd, idx) => {
    const item = document.createElement("div");
    item.className = `slash-command-item${idx === activeIdx ? " active" : ""}`;
    item.setAttribute("role", "option");
    item.setAttribute("aria-selected", idx === activeIdx ? "true" : "false");

    const iconSpan = document.createElement("span");
    iconSpan.className = "slash-command-icon";
    iconSpan.textContent = cmd.icon;

    const detailsDiv = document.createElement("div");
    detailsDiv.className = "slash-command-details";

    const topRow = document.createElement("div");
    topRow.className = "slash-command-top";

    const nameSpan = document.createElement("span");
    nameSpan.className = "slash-command-name";
    nameSpan.textContent = cmd.command;

    const labelSpan = document.createElement("span");
    labelSpan.className = "slash-command-label";
    labelSpan.textContent = cmd.label;

    topRow.appendChild(nameSpan);
    topRow.appendChild(labelSpan);

    const descDiv = document.createElement("div");
    descDiv.className = "slash-command-desc";
    descDiv.textContent = cmd.description;

    detailsDiv.appendChild(topRow);
    detailsDiv.appendChild(descDiv);

    item.appendChild(iconSpan);
    item.appendChild(detailsDiv);

    item.addEventListener("mousedown", (e) => {
      e.preventDefault();
    });

    item.onclick = (e) => {
      e.stopPropagation();
      executeSlashCommand(cmd);
    };

    item.onmouseenter = () => {
      activeSlashIndex = idx;
      updateActiveSlashItem();
    };

    list.appendChild(item);
  });
}

function executeSlashCommand(cmd) {
  if (isGenerating || isProcessingInDocCommand) return;
  hideSlashPopup();
  const promptInput = document.getElementById("promptText");
  if (promptInput) {
    promptInput.value = "";
    promptInput.style.height = "auto";
  }
  updateSendButtonState();
  if (cmd && typeof cmd.action === "function") {
    cmd.action();
  }
}

function setupSlashCommands() {
  const promptInput = document.getElementById("promptText");
  const popup = document.getElementById("slashCommandsPopup");
  if (!promptInput || !popup) return;

  promptInput.addEventListener("input", () => {
    promptInput.style.height = "auto";
    promptInput.style.height = Math.min(promptInput.scrollHeight, 120) + "px";
    updateSendButtonState();

    const val = promptInput.value;
    const cursorPos = promptInput.selectionStart;
    const textBeforeCursor = val.substring(0, cursorPos);

    const slashMatch = textBeforeCursor.match(/(?:^|\s)\/([a-zA-Z0-9\-_]*)$/);
    if (slashMatch) {
      const query = slashMatch[1].toLowerCase();
      const sourceCommands = getActiveSlashCommands();
      currentMatchingCommands = sourceCommands.filter(c => {
        const cmdName = c.command.toLowerCase().replace("/", "");
        const label = c.label.toLowerCase();
        return query === "" || cmdName.includes(query) || label.includes(query);
      });

      activeSlashIndex = 0;
      renderSlashPopupItems(currentMatchingCommands, activeSlashIndex);
      showSlashPopup();
    } else {
      hideSlashPopup();
    }
  });

  promptInput.addEventListener("keydown", (e) => {
    // Do nothing on Enter or Tab while a command is running
    if (isGenerating || isProcessingInDocCommand) {
      if (e.key === "Enter" || e.key === "Tab") {
        e.preventDefault();
      }
      return;
    }

    const isPopupVisible = popup.style.display !== "none";

    if (isPopupVisible) {
      if (e.key === "ArrowDown") {
        e.preventDefault();
        if (currentMatchingCommands.length > 0) {
          activeSlashIndex = (activeSlashIndex + 1) % currentMatchingCommands.length;
          updateActiveSlashItem();
        }
        return;
      }
      if (e.key === "ArrowUp") {
        e.preventDefault();
        if (currentMatchingCommands.length > 0) {
          activeSlashIndex = (activeSlashIndex - 1 + currentMatchingCommands.length) % currentMatchingCommands.length;
          updateActiveSlashItem();
        }
        return;
      }
      if (e.key === "Enter" || e.key === "Tab") {
        if (currentMatchingCommands.length > 0 && activeSlashIndex >= 0 && activeSlashIndex < currentMatchingCommands.length) {
          e.preventDefault();
          executeSlashCommand(currentMatchingCommands[activeSlashIndex]);
          return;
        }
      }
      if (e.key === "Escape") {
        e.preventDefault();
        hideSlashPopup();
        return;
      }
    }

    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      callGeminiProxy();
    }
  });

  document.addEventListener("click", (e) => {
    if (popup.style.display !== "none" && !popup.contains(e.target) && e.target !== promptInput) {
      hideSlashPopup();
    }
  });
}

// Universal In-Document / In-App Command Processor (@gemini <prompt> and /slash)
async function checkForInDocumentCommands(forceRun = false) {
  if (isProcessingInDocCommand || !hostAdapter || typeof hostAdapter.checkInDocumentCommands !== 'function') return;

  const debugStatus = document.getElementById("debugStatus");
  const scanBtn = document.getElementById("scanInDoc");
  const loadingText = document.getElementById("loading");

  await hostAdapter.checkInDocumentCommands(forceRun, {
    onStatus: (msg) => {
      if (debugStatus) {
        debugStatus.innerText = msg;
        if (hostAdapter && hostAdapter.name === "Outlook") {
          debugStatus.style.display = msg ? "block" : "none";
        }
      }
    },
    executePrompt: async (userPrompt) => {
      isProcessingInDocCommand = true;
      if (scanBtn) scanBtn.disabled = true;
      if (loadingText) {
        loadingText.innerText = "⚡ Generating response...";
        loadingText.style.display = "flex";
      }

      const isSlash = userPrompt.startsWith("/");
      const displayBubbleText = isSlash ? userPrompt : `@gemini ${userPrompt}`;
      appendBubble(displayBubbleText, "user");

      try {
        let promptToRun = userPrompt;
        let enableGrounding = true;
        const lowerPrompt = userPrompt.toLowerCase().trim();

        if (lowerPrompt.startsWith("/summarize-doc") || lowerPrompt === "/summarize") {
          promptToRun = "Summarize this document thoroughly. Provide an executive overview, key takeaways, and a structured markdown summary table.";
          enableGrounding = false;
        } else if (lowerPrompt.startsWith("/identify-risks") || lowerPrompt === "/risks") {
          promptToRun = "Analyze this document and extract all Key Risks, Ambiguities, and Operational Gaps. Present them in a structured markdown table with Severity and Mitigation Suggestions.";
          enableGrounding = false;
        } else if (lowerPrompt.startsWith("/identify-action") || lowerPrompt.startsWith("/identify-actions")) {
          promptToRun = "Extract all Action Items, Deliverables, and Next Steps from this document. Present them in a structured table with Task, Owner, and Priority.";
          enableGrounding = false;
        } else if (lowerPrompt.startsWith("/exec-card") || lowerPrompt === "/exec") {
          promptToRun = "Generate an Executive Summary Box for this document using a callout block (> [!NOTE] ...) highlighting Strategic Purpose, Key Metrics, and Impact.";
          enableGrounding = false;
        }

        let fullPrompt = promptToRun;
        if (!enableGrounding) {
          const fullDocText = await hostAdapter.getFullDocumentText();
          if (fullDocText && fullDocText.length > 20) {
            fullPrompt = `Active Document Context:\n"""\n${fullDocText.substring(0, 100000)}\n"""\n\nUser Instruction: ${promptToRun}\n\nConstraint: Answer strictly based on the provided Active Document Context. Do not reference external datastores.`;
          }
        }

        const hostName = hostAdapter ? hostAdapter.name : 'Word';
        const data = await askGeminiEnterprise(fullPrompt, chatHistoryState, currentSessionId, enableGrounding);
        currentSessionId = data.sessionId || currentSessionId || `local-${Date.now()}`;
        SessionManager.setActiveSessionId(currentSessionId, hostName);
        if (Array.isArray(data.history)) chatHistoryState = data.history;

        const aiResultText = data.result || "No content returned.";
        appendAssistantBubble(aiResultText, data, userPrompt);

        if (currentSessionId) {
          SessionManager.addMessage(currentSessionId, { role: 'user', text: displayBubbleText }, hostName);
          SessionManager.addMessage(currentSessionId, { role: 'assistant', text: aiResultText }, hostName);
        }

        const hasStructuredChart = /```(?:json|chart|pie|bar|column|line|doughnut|donut)?\s*\{[\s\S]*?(?:chartType|chart_type|pie|bar|column|line|doughnut)[\s\S]*?```/i.test(aiResultText) ||
                                   /\{\s*"(?:chartType|chart_type)"\s*:\s*"(?:pie|bar|line|doughnut|donut|column)"/i.test(aiResultText);
        const hasDistinctNonChartImageIntent = /(?:photo|photograph|portrait|illustration|logo|camera|scenery|picture of|image of a)/i.test((userPrompt || '').toLowerCase());

        let fullAiText = aiResultText;
        if (data && Array.isArray(data.images) && data.images.length > 0) {
          if (!hasStructuredChart || hasDistinctNonChartImageIntent) {
            for (const imgUrl of data.images) {
              if (imgUrl && !fullAiText.includes(imgUrl)) {
                fullAiText += `\n\n![Generated Image](${imgUrl})\n\n`;
              }
            }
          }
        }
        return parseMarkdown(fullAiText, { hasDistinctNonChartImageIntent });
      } catch (err) {
        console.error("In-document command execution error:", err);
        if (debugStatus) debugStatus.innerText = "Error: " + err.message;
        return null;
      } finally {
        isProcessingInDocCommand = false;
        if (scanBtn) scanBtn.disabled = false;
        if (loadingText) {
          loadingText.style.display = "none";
        }
      }
    }
  });
}

export function isSubstantiveText(text) {
  if (!text || typeof text !== "string") return false;
  const cleaned = text
    .replace(/\(Visual\s*\/\s*Slide\s*content\)/gi, "")
    .replace(/\[Slide\s*\d+\]:?/gi, "")
    .replace(/[-•*#_`~|\s\r\n\t]/g, "")
    .trim();
  return /[a-zA-Z0-9]/.test(cleaned) && cleaned.length >= 3;
}

// Dedicated Image Generation Action from Selection or User Query
async function handleCreateImageClick() {
  clearChatAndInput();
  const loadingText = document.getElementById("loading");
  if (loadingText) {
    loadingText.innerText = "⚡ Checking selection...";
    loadingText.style.display = "block";
  }

  let selectedText = "";
  let slideCount = 0;

  try {
    if (hostAdapter && typeof hostAdapter.getSelectedText === "function") {
      const raw = await hostAdapter.getSelectedText();
      selectedText = isSubstantiveText(raw) ? raw.trim() : "";
    }
    if (hostAdapter && typeof hostAdapter.getSelectedSlidesText === "function") {
      const slides = await hostAdapter.getSelectedSlidesText();
      if (slides) slideCount = slides.length;
    }
  } catch (err) {
    console.warn("Could not read selection for image generation:", err);
  } finally {
    if (loadingText) loadingText.style.display = "none";
  }

  // 1. If substantive text is selected AND not multiple slides: generate image for that selected content
  if (selectedText && selectedText.length > 0 && slideCount <= 1) {
    const prompt = `Create a high-quality, professional image visual illustration representing the following selected content:

"""
${selectedText}
"""

CRITICAL INSTRUCTIONS:
1. Trigger the native image generation tool directly to produce the visual image.
2. DO NOT output raw JSON, visual_request schemas, or code blocks.
3. Generate an impactful, visually stunning illustration for this presentation topic.`;
    const snippet = selectedText.replace(/\s+/g, " ").trim().substring(0, 75);
    const displayUserBubble = `🎨 Generate image for selected text:\n"${snippet}${selectedText.length > 75 ? '...' : ''}"`;
    await executeGeminiWorkflow(prompt, displayUserBubble, null, { isolateSession: true });
    return;
  }

  // 2. If selected content is blank/null OR multiple slides are selected: ask the user what the image should be created for
  const reasonText = slideCount > 1
    ? `${slideCount} slides are currently selected.`
    : `No specific text is currently selected.`;

  const html = `
    <div>
      <div style="font-weight:600; color:#0b57d0; margin-bottom:4px;">🎨 Create Image</div>
      <div style="font-size:12px; margin-bottom:6px;">${reasonText} What should the image be created for?</div>
      <div style="font-size:11px; color:#444746; margin-bottom:8px;">Choose a suggested concept below or describe what you want in the input box:</div>
      <div style="display:flex; flex-wrap:wrap; gap:6px; margin-bottom:4px;">
        <button class="quick-chip img-opt-btn">🏬 Retail Storefront</button>
        <button class="quick-chip img-opt-btn">📦 Supply Chain Logistics</button>
        <button class="quick-chip img-opt-btn">👥 Executive Strategy Team</button>
        <button class="quick-chip img-opt-btn">🌱 Corporate Sustainability</button>
      </div>
    </div>
  `;

  appendInteractiveBubble(html, (bubble) => {
    const optButtons = bubble.querySelectorAll(".img-opt-btn");
    optButtons.forEach((btn) => {
      btn.addEventListener("click", async () => {
        const concept = btn.innerText.replace(/^[^\w]+/, "").trim();
        optButtons.forEach((b) => (b.disabled = true));
        const prompt = `Create a high-quality, professional image visual illustration of: ${concept}.

CRITICAL INSTRUCTIONS:
1. Trigger the native image generation tool directly to produce the visual image.
2. DO NOT output raw JSON, visual_request schemas, or code blocks.`;
        const displayBubble = `🎨 Generate image: "${concept}"`;
        await executeGeminiWorkflow(prompt, displayBubble, null, { isolateSession: true });
      });
    });

    const promptInput = document.getElementById("promptText");
    if (promptInput) {
      promptInput.value = "Create an image of ";
      promptInput.focus();
      if (promptInput.setSelectionRange) {
        promptInput.setSelectionRange(promptInput.value.length, promptInput.value.length);
      }
    }
  });
}

// Dedicated Chart Creation Action from Selection or User Guidance
async function handleCreateChartClick() {
  clearChatAndInput();
  const loadingText = document.getElementById("loading");
  if (loadingText) {
    loadingText.innerText = "⚡ Checking selection for chart...";
    loadingText.style.display = "block";
  }

  let selectedText = "";
  let slideCount = 0;

  try {
    if (hostAdapter && typeof hostAdapter.getSelectedText === "function") {
      const raw = await hostAdapter.getSelectedText();
      selectedText = isSubstantiveText(raw) ? raw.trim() : "";
    }
    if (hostAdapter && typeof hostAdapter.getSelectedSlidesText === "function") {
      const slides = await hostAdapter.getSelectedSlidesText();
      if (slides) slideCount = slides.length;
    }
  } catch (err) {
    console.warn("Could not read selection for chart:", err);
  } finally {
    if (loadingText) loadingText.style.display = "none";
  }

  // Case 1: Specific text IS selected (and not multiple slides)
  // Requirement: "If the text is selected, ask the user what type of chart and pass both to streamassist"
  if (selectedText && selectedText.length > 0 && slideCount <= 1) {
    const preview = selectedText.replace(/\s+/g, " ").trim().substring(0, 80);
    const html = `
      <div>
        <div style="font-weight:600; color:#0b57d0; margin-bottom:4px;">📈 Create Chart for Selected Content</div>
        <div style="font-size:11.5px; color:#444746; margin-bottom:8px; font-style:italic;">"${escapeHtml(preview)}${selectedText.length > 80 ? '...' : ''}"</div>
        <div style="font-size:12px; margin-bottom:8px;">What type of chart would you like to create?</div>
        <div style="display:flex; flex-wrap:wrap; gap:6px;">
          <button class="quick-chip chart-opt-btn" data-type="doughnut">🍩 Doughnut / Pie</button>
          <button class="quick-chip chart-opt-btn" data-type="bar">📊 Bar Chart</button>
          <button class="quick-chip chart-opt-btn" data-type="column">📶 Column Chart</button>
          <button class="quick-chip chart-opt-btn" data-type="line">📈 Line Chart</button>
        </div>
      </div>
    `;

    appendInteractiveBubble(html, (bubble) => {
      const buttons = bubble.querySelectorAll(".chart-opt-btn");
      buttons.forEach((btn) => {
        btn.addEventListener("click", async () => {
          const chartType = btn.getAttribute("data-type") || "bar";
          buttons.forEach((b) => (b.disabled = true));
          btn.style.borderColor = "#0b57d0";
          btn.style.backgroundColor = "#e8f0fe";

          const prompt = `Create a high-resolution ${chartType} chart for PowerPoint visualizing the following selected data and context:
"""
${selectedText}
"""

Requirements:
1. Output a structured JSON code block with the exact chart specification:
\`\`\`json
{
  "chartType": "${chartType}",
  "title": "<Short Title (2-5 words max)>",
  "data": [
    { "label": "<Category/Metric>", "value": <NumericValue> }
  ]
}
\`\`\`
2. CRITICAL CONSTRAINT FOR TITLE:
   - The chart title MUST be short and punchy (maximum 2 to 5 words, e.g. "Coffee Output by State", "Q4 Revenue Share").
   - NEVER create long titles, never include parenthetical subtitles in the title, and never exceed 5 words!
3. Include 2 concise, executive takeaway bullet points analyzing the data under the chart.
4. CRITICAL SCOPE CONTRACT: Output content for EXACTLY ONE single slide. DO NOT regenerate, repeat, or expand upon previous presentation slides from earlier in this conversation under any circumstances.`;

          // A chart request is a single-visual action: pin the parser to one slide so
          // a stray second heading in the model output cannot split it into a deck.
          if (typeof window !== "undefined") window.__forceSingleSlide = true;
          const displayBubble = `📈 [Create Chart] Generating ${chartType} chart from selected text...`;
          await executeGeminiWorkflow(prompt, displayBubble);
        });
      });
    });
    return;
  }

  // Case 2: No text is selected OR multiple slides are selected
  // Requirement: "If no text is selected or multiple slides selected, ask the user what type of chart (you can provide some options) and also ask about what."
  const notice = slideCount > 1
    ? `${slideCount} slides are currently selected.`
    : `No specific text is currently selected.`;

  const html = `
    <div>
      <div style="font-weight:600; color:#0b57d0; margin-bottom:4px;">📈 Create Chart</div>
      <div style="font-size:11.5px; color:#444746; margin-bottom:8px;">${notice} What type of chart would you like to create, and what data or topic should it visualize?</div>
      
      <div style="font-weight:600; font-size:11.5px; margin-bottom:4px; color:#1f1f1f;">1. Select Chart Type:</div>
      <div style="display:flex; flex-wrap:wrap; gap:6px; margin-bottom:10px;">
        <button class="quick-chip chart-type-btn" data-type="doughnut">🍩 Doughnut / Pie</button>
        <button class="quick-chip chart-type-btn" data-type="bar">📊 Bar Chart</button>
        <button class="quick-chip chart-type-btn" data-type="column">📶 Column Chart</button>
        <button class="quick-chip chart-type-btn" data-type="line">📈 Line Chart</button>
      </div>

      <div style="font-weight:600; font-size:11.5px; margin-bottom:4px; color:#1f1f1f;">2. Enter or Paste Data (e.g. from Table):</div>
      <textarea class="chart-data-input" placeholder="Paste table rows, numbers, or describe metrics (e.g. 'Revenue: Q1 10M, Q2 15M, Q3 25M')..." style="width:100%; box-sizing:border-box; border:1px solid #c4c7c5; border-radius:8px; padding:8px; font-size:11.5px; font-family:inherit; resize:vertical; min-height:50px; margin-bottom:8px;"></textarea>
      <div style="margin-bottom:10px;">
        <button class="pill-btn primary chart-generate-btn">✨ Generate Chart from Data</button>
      </div>

      <div style="font-weight:600; font-size:11.5px; margin-bottom:4px; color:#1f1f1f;">Or choose a quick topic:</div>
      <div style="display:flex; flex-wrap:wrap; gap:6px; margin-bottom:4px;">
        <button class="quick-chip chart-topic-btn">🛒 Retail Market Share</button>
        <button class="quick-chip chart-topic-btn">💰 Quarterly Revenue Growth</button>
        <button class="quick-chip chart-topic-btn">📉 Expense & Cost Breakdown</button>
        <button class="quick-chip chart-topic-btn">⚡ Regional Performance</button>
      </div>
    </div>
  `;

  appendInteractiveBubble(html, (bubble) => {
    let currentType = "bar";
    const typeButtons = bubble.querySelectorAll(".chart-type-btn");
    const topicButtons = bubble.querySelectorAll(".chart-topic-btn");
    const generateBtn = bubble.querySelector(".chart-generate-btn");
    const dataInput = bubble.querySelector(".chart-data-input");

    typeButtons.forEach((btn) => {
      btn.addEventListener("click", () => {
        currentType = btn.getAttribute("data-type") || "bar";
        typeButtons.forEach((b) => {
          b.style.borderColor = "";
          b.style.fontWeight = "500";
          b.style.backgroundColor = "";
        });
        btn.style.borderColor = "#0b57d0";
        btn.style.backgroundColor = "#e8f0fe";
        btn.style.fontWeight = "600";

        const promptInput = document.getElementById("promptText");
        if (promptInput) {
          promptInput.value = `Create a ${currentType} chart of `;
          promptInput.focus();
          if (promptInput.setSelectionRange) {
            promptInput.setSelectionRange(promptInput.value.length, promptInput.value.length);
          }
        }
      });
    });

    if (generateBtn && dataInput) {
      generateBtn.addEventListener("click", async () => {
        const dataText = dataInput.value.trim();
        if (!dataText) {
          dataInput.focus();
          dataInput.style.borderColor = "#dc2626";
          return;
        }
        typeButtons.forEach((b) => (b.disabled = true));
        topicButtons.forEach((b) => (b.disabled = true));
        generateBtn.disabled = true;

        const prompt = `Create a high-resolution ${currentType} chart for PowerPoint visualizing the following data and metrics:
${dataText}

Requirements:
1. Output a structured JSON code block with the exact chart specification:
\`\`\`json
{
  "chartType": "${currentType}",
  "title": "<Short Title (2-5 words max)>",
  "data": [
    { "label": "<Category>", "value": <NumericValue> }
  ]
}
\`\`\`
2. CRITICAL CONSTRAINT FOR TITLE:
   - The chart title MUST be short and punchy (maximum 2 to 5 words).
   - NEVER create long titles, never include parenthetical subtitles in the title, and never exceed 5 words!
3. Include 2 concise, executive takeaway bullet points analyzing the data under the chart.
4. CRITICAL SCOPE CONTRACT: Output content for EXACTLY ONE single slide. DO NOT regenerate, repeat, or expand upon previous presentation slides from earlier in this conversation under any circumstances.`;

        if (typeof window !== "undefined") window.__forceSingleSlide = true;
        const displayBubble = `📈 [Create Chart] Generating ${currentType} chart from provided data...`;
        await executeGeminiWorkflow(prompt, displayBubble);
      });
    }

    topicButtons.forEach((btn) => {
      btn.addEventListener("click", async () => {
        const topic = btn.innerText.replace(/^[^\w]+/, "").trim();
        typeButtons.forEach((b) => (b.disabled = true));
        topicButtons.forEach((b) => (b.disabled = true));
        if (generateBtn) generateBtn.disabled = true;

        const prompt = `Create a high-resolution ${currentType} chart for PowerPoint visualizing: ${topic}.

Requirements:
1. Output a structured JSON code block with the exact chart specification:
\`\`\`json
{
  "chartType": "${currentType}",
  "title": "${topic}",
  "data": [
    { "label": "<Category>", "value": <NumericValue> }
  ]
}
\`\`\`
2. CRITICAL CONSTRAINT FOR TITLE:
   - The chart title MUST be short and punchy (maximum 2 to 5 words).
   - NEVER create long titles, never include parenthetical subtitles in the title, and never exceed 5 words!
3. Include 2 concise, executive takeaway bullet points analyzing the data under the chart.
4. CRITICAL SCOPE CONTRACT: Output content for EXACTLY ONE single slide. DO NOT regenerate, repeat, or expand upon previous presentation slides from earlier in this conversation under any circumstances.`;

        if (typeof window !== "undefined") window.__forceSingleSlide = true;
        const displayBubble = `📈 [Create Chart] Generating ${currentType} chart for "${topic}"...`;
        await executeGeminiWorkflow(prompt, displayBubble);
      });
    });

    const promptInput = document.getElementById("promptText");
    if (promptInput) {
      promptInput.value = `Create a ${currentType} chart of `;
      promptInput.focus();
      if (promptInput.setSelectionRange) {
        promptInput.setSelectionRange(promptInput.value.length, promptInput.value.length);
      }
    }
  });
}

// Interactive Excel Native Chart Builder Wizard before calling Gemini
async function handleExcelChartBuilderClick(chip = null) {
  clearChatAndInput();
  const loadingText = document.getElementById("loading");
  if (loadingText) {
    loadingText.innerText = "⚡ Inspecting worksheet columns & range for chart builder...";
    loadingText.style.display = "block";
  }

  let sheetName = "Active Worksheet";
  let scopeSummary = "Worksheet Used Range";
  let detectedColumns = [];

  try {
    if (typeof window !== "undefined" && window.Excel && typeof window.Excel.run === "function" && hostAdapter) {
      await window.Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const targetInfo = typeof hostAdapter.getSelectionTargetInfo === "function"
          ? await hostAdapter.getSelectionTargetInfo(context)
          : null;
        const schema = typeof hostAdapter.getWorksheetSchema === "function"
          ? await hostAdapter.getWorksheetSchema(context, sheet)
          : null;

        if (schema && schema.sheetName) {
          sheetName = schema.sheetName;
        } else if (targetInfo && targetInfo.sheetName) {
          sheetName = targetInfo.sheetName;
        }

        if (targetInfo && targetInfo.type === "table" && targetInfo.table) {
          scopeSummary = `Table "${targetInfo.table.name}" (${targetInfo.cleanAddress})`;
        } else if (targetInfo && targetInfo.type === "range" && targetInfo.cellCount > 1) {
          scopeSummary = `Selected Range ${targetInfo.cleanAddress} (${targetInfo.rowCount}×${targetInfo.columnCount})`;
        } else if (schema && schema.usedRangeAddress) {
          const cleanUsed = schema.usedRangeAddress.includes("!")
            ? schema.usedRangeAddress.split("!").pop()
            : schema.usedRangeAddress;
          scopeSummary = `Range ${cleanUsed} (${schema.rowCount || 0} rows)`;
        }

        if (schema && Array.isArray(schema.columns)) {
          const startRow = schema.startRow || 1;
          const endRow = schema.endRow || (startRow + (schema.rowCount || 1) - 1);
          detectedColumns = schema.columns.map((col) => ({
            letter: col.letter,
            header: col.header || `Column ${col.letter}`,
            range: `${col.letter}${startRow}:${col.letter}${endRow}`,
            samples: Array.isArray(col.samples) ? col.samples.slice(0, 2).join(", ") : ""
          }));
        }
      });
    }
  } catch (err) {
    console.warn("Could not pre-inspect Excel worksheet for chart builder:", err);
  } finally {
    if (loadingText) loadingText.style.display = "none";
  }

  const columnOptionsHtml = detectedColumns
    .map((col) => {
      const val = `Column ${col.letter} (${col.range}) - "${col.header}"`;
      const label = `Col ${col.letter}: ${col.header} (${col.range})`;
      return `<option value="${escapeHtml(val)}">${escapeHtml(label)}</option>`;
    })
    .join("");

  const html = `
    <form class="excel-chart-builder-form" action="#excel-chart-builder" method="POST" novalidate style="display:flex; flex-direction:column; gap:10px;">
      <div>
        <div style="font-weight:700; color:#0b57d0; font-size:13px; margin-bottom:2px;">📊 Native Chart Builder</div>
        <div style="font-size:11px; color:#0b57d0; background:#e8f0fe; border:1px solid #a8c7fa; border-radius:9999px; padding:3px 10px; display:inline-block; margin-bottom:6px; font-weight:500;">
          📄 ${escapeHtml(sheetName)} &bull; 🎯 ${escapeHtml(scopeSummary)}
        </div>
        <div style="font-size:11.5px; color:#444746;">
          Specify your chart preferences below before sending the request to Gemini:
        </div>
      </div>

      <fieldset style="border:none; margin:0; padding:0;">
        <legend style="font-weight:600; font-size:11.5px; color:#1f1f1f; margin-bottom:5px; padding:0;">1. Select Chart Type:</legend>
        <div style="display:flex; flex-wrap:wrap; gap:5px; margin-bottom:6px;">
          <button type="button" class="quick-chip excel-chart-type-pill" data-type="ColumnClustered" style="background:#e8f0fe; color:#0b57d0; border:1.5px solid #0b57d0; font-weight:600; cursor:pointer;">📊 Clustered Column</button>
          <button type="button" class="quick-chip excel-chart-type-pill" data-type="BarClustered" style="background:#f8fafd; color:#1f1f1f; border:1px solid #c4c7c5; font-weight:500; cursor:pointer;">📶 Bar Chart</button>
          <button type="button" class="quick-chip excel-chart-type-pill" data-type="LineMarkers" style="background:#f8fafd; color:#1f1f1f; border:1px solid #c4c7c5; font-weight:500; cursor:pointer;">📈 Line Chart</button>
          <button type="button" class="quick-chip excel-chart-type-pill" data-type="Doughnut" style="background:#f8fafd; color:#1f1f1f; border:1px solid #c4c7c5; font-weight:500; cursor:pointer;">🍩 Doughnut / Pie</button>
          <button type="button" class="quick-chip excel-chart-type-pill" data-type="ColumnStacked" style="background:#f8fafd; color:#1f1f1f; border:1px solid #c4c7c5; font-weight:500; cursor:pointer;">📚 Stacked Column</button>
          <button type="button" class="quick-chip excel-chart-type-pill" data-type="Area" style="background:#f8fafd; color:#1f1f1f; border:1px solid #c4c7c5; font-weight:500; cursor:pointer;">🌊 Area Chart</button>
          <button type="button" class="quick-chip excel-chart-type-pill" data-type="Waterfall" style="background:#f8fafd; color:#1f1f1f; border:1px solid #c4c7c5; font-weight:500; cursor:pointer;">📉 Waterfall</button>
          <button type="button" class="quick-chip excel-chart-type-pill" data-type="Auto" style="background:#f8fafd; color:#1f1f1f; border:1px solid #c4c7c5; font-weight:500; cursor:pointer;">✨ Auto (Best Fit)</button>
        </div>
        <label for="excelChartTypeSelect" style="display:block; font-size:11px; color:#5f6368; margin-bottom:2px;">Or choose any native Excel chart type:</label>
        <select id="excelChartTypeSelect" name="chartType" style="width:100%; box-sizing:border-box; border:1px solid #c4c7c5; border-radius:4px; padding:5px 6px; font-size:11.5px; background:#ffffff; color:#1f1f1f;">
          <option value="ColumnClustered" selected>Clustered Column (ColumnClustered)</option>
          <option value="BarClustered">Clustered Bar (BarClustered)</option>
          <option value="LineMarkers">Line with Markers (LineMarkers)</option>
          <option value="Line">Line (Line)</option>
          <option value="Doughnut">Doughnut (Doughnut)</option>
          <option value="Pie">Pie (Pie)</option>
          <option value="ColumnStacked">Stacked Column (ColumnStacked)</option>
          <option value="BarStacked">Stacked Bar (BarStacked)</option>
          <option value="Area">Area (Area)</option>
          <option value="Scatter">Scatter / XY (Scatter)</option>
          <option value="Waterfall">Waterfall (Waterfall)</option>
          <option value="Treemap">Treemap (Treemap)</option>
          <option value="Funnel">Funnel (Funnel)</option>
          <option value="Auto">Auto (Let Gemini pick optimal chart type)</option>
        </select>
      </fieldset>

      <div style="display:grid; grid-template-columns:1fr; gap:8px;">
        <div>
          <label for="excelChartCategoryCol" style="display:block; font-weight:600; font-size:11.5px; color:#1f1f1f; margin-bottom:3px;">2. Category / X-Axis Column:</label>
          <select id="excelChartCategoryCol" name="categoryCol" style="width:100%; box-sizing:border-box; border:1px solid #c4c7c5; border-radius:4px; padding:5px 6px; font-size:11.5px; background:#ffffff; color:#1f1f1f;">
            <option value="">✨ Auto-detect category / label column</option>
            ${columnOptionsHtml}
          </select>
        </div>
        <div>
          <label for="excelChartSeriesCol" style="display:block; font-weight:600; font-size:11.5px; color:#1f1f1f; margin-bottom:3px;">3. Metric / Y-Axis Value Column(s):</label>
          <select id="excelChartSeriesCol" name="seriesCol" style="width:100%; box-sizing:border-box; border:1px solid #c4c7c5; border-radius:4px; padding:5px 6px; font-size:11.5px; background:#ffffff; color:#1f1f1f;">
            <option value="">✨ Auto-detect metric / numeric column(s)</option>
            ${columnOptionsHtml}
          </select>
        </div>
      </div>

      <div style="display:grid; grid-template-columns:1fr 1fr; gap:8px;">
        <div>
          <label for="excelChartAggregation" style="display:block; font-weight:600; font-size:11.5px; color:#1f1f1f; margin-bottom:3px;">4. Aggregation:</label>
          <select id="excelChartAggregation" name="aggregation" style="width:100%; box-sizing:border-box; border:1px solid #c4c7c5; border-radius:4px; padding:5px 6px; font-size:11.5px; background:#ffffff; color:#1f1f1f;">
            <option value="auto" selected>✨ Auto (Sum if repeating)</option>
            <option value="sum">➕ Sum by Category</option>
            <option value="average">➗ Average by Category</option>
            <option value="count">🔢 Count by Category</option>
            <option value="none">📋 None (Raw rows)</option>
          </select>
        </div>
        <div>
          <label for="excelChartTitleInput" style="display:block; font-weight:600; font-size:11.5px; color:#1f1f1f; margin-bottom:3px;">5. Chart Title (Optional):</label>
          <input type="text" id="excelChartTitleInput" name="chartTitle" placeholder="Auto-generate title..." style="width:100%; box-sizing:border-box; border:1px solid #c4c7c5; border-radius:4px; padding:5px 6px; font-size:11.5px; font-family:inherit;" />
        </div>
      </div>

      <div>
        <label for="excelChartDetailsInput" style="display:block; font-weight:600; font-size:11.5px; color:#1f1f1f; margin-bottom:3px;">6. What to Visualize / Additional Instructions (Optional):</label>
        <textarea id="excelChartDetailsInput" name="chartDetails" rows="2" placeholder="e.g., Compare Revenue vs Quota by Region, plot top categories, or leave blank..." style="width:100%; box-sizing:border-box; border:1px solid #c4c7c5; border-radius:4px; padding:6px; font-size:11.5px; font-family:inherit; resize:vertical;"></textarea>
      </div>

      <div style="display:flex; align-items:center; gap:8px; margin-top:2px;">
        <button type="submit" class="pill-btn primary excel-chart-submit-btn">
          📊 Generate Native Chart
        </button>
      </div>
    </form>
  `;

  appendInteractiveBubble(html, (bubble) => {
    const form = bubble.querySelector(".excel-chart-builder-form");
    const typePills = bubble.querySelectorAll(".excel-chart-type-pill");
    const typeSelect = bubble.querySelector("#excelChartTypeSelect");
    const categorySelect = bubble.querySelector("#excelChartCategoryCol");
    const seriesSelect = bubble.querySelector("#excelChartSeriesCol");
    const aggSelect = bubble.querySelector("#excelChartAggregation");
    const titleInput = bubble.querySelector("#excelChartTitleInput");
    const detailsInput = bubble.querySelector("#excelChartDetailsInput");
    const submitBtn = bubble.querySelector(".excel-chart-submit-btn");

    const highlightSelectedPill = (selectedType) => {
      typePills.forEach((pill) => {
        const pillType = pill.getAttribute("data-type");
        if (pillType === selectedType) {
          pill.style.border = "1.5px solid #0b57d0";
          pill.style.backgroundColor = "#e8f0fe";
          pill.style.color = "#0b57d0";
          pill.style.fontWeight = "600";
        } else {
          pill.style.border = "1px solid #c4c7c5";
          pill.style.backgroundColor = "#f8fafd";
          pill.style.color = "#1f1f1f";
          pill.style.fontWeight = "500";
        }
      });
    };

    typePills.forEach((pill) => {
      pill.addEventListener("click", () => {
        const chosen = pill.getAttribute("data-type") || "ColumnClustered";
        if (typeSelect) {
          typeSelect.value = chosen;
        }
        highlightSelectedPill(chosen);
      });
    });

    if (typeSelect) {
      typeSelect.addEventListener("change", () => {
        highlightSelectedPill(typeSelect.value);
      });
    }

    if (form) {
      form.addEventListener("submit", async (e) => {
        e.preventDefault();
        if (submitBtn && submitBtn.disabled) return;

        const chosenType = (typeSelect && typeSelect.value) ? typeSelect.value : "ColumnClustered";
        const chosenCategory = (categorySelect && categorySelect.value) ? categorySelect.value.trim() : "";
        const chosenSeries = (seriesSelect && seriesSelect.value) ? seriesSelect.value.trim() : "";
        const chosenAgg = (aggSelect && aggSelect.value) ? aggSelect.value : "auto";
        const customTitle = (titleInput && titleInput.value) ? titleInput.value.trim() : "";
        const customDetails = (detailsInput && detailsInput.value) ? detailsInput.value.trim() : "";

        // Disable controls after valid submission to prevent double-posts
        if (submitBtn) {
          submitBtn.disabled = true;
          submitBtn.style.opacity = "0.75";
          submitBtn.innerHTML = "⚡ Requesting Chart from Gemini...";
        }
        typePills.forEach((b) => { b.disabled = true; });
        [typeSelect, categorySelect, seriesSelect, aggSelect, titleInput, detailsInput].forEach((el) => {
          if (el) el.disabled = true;
        });

        try {
          let docContext = "";
          if (hostAdapter && typeof hostAdapter.getSelectedText === "function") {
            try {
              docContext = await hostAdapter.getSelectedText();
            } catch (selErr) {
              console.warn("Chart builder getSelectedText fallback:", selErr);
            }
          }
          if ((!docContext || docContext.length < 20) && hostAdapter && typeof hostAdapter.getFullDocumentText === "function") {
            docContext = await hostAdapter.getFullDocumentText();
          }

          const typeDirective = chosenType === "Auto"
            ? "Automatically determine the best native Excel chart type (e.g., ColumnClustered, BarClustered, LineMarkers, Doughnut, Pie, Area, or Waterfall) based on the series dimensions."
            : `Chart Type (REQUIRED): Set "type" to "${chosenType}".`;

          const categoryDirective = chosenCategory
            ? `Category / X-Axis Column (REQUIRED): Use ${chosenCategory} for "categoriesRange".`
            : "Category / X-Axis Column: Automatically identify the appropriate category/label column for \"categoriesRange\".";

          const seriesDirective = chosenSeries
            ? `Metric / Y-Axis Series Column (REQUIRED): Use ${chosenSeries} for "seriesRange".`
            : "Metric / Y-Axis Series Column: Automatically identify the key numeric metric column(s) for \"seriesRange\".";

          const aggDirective = chosenAgg === "none"
            ? "Aggregation: Do NOT aggregate rows (plot raw rows directly; omit \"aggregate\")."
            : chosenAgg === "auto"
              ? "Aggregation: If the category column has repeating values across transaction/line-item rows, include \"aggregate\": \"sum\"; otherwise plot directly."
              : `Aggregation (REQUIRED): Set "aggregate": "${chosenAgg}" so repeating categories are aggregated.`;

          const titleDirective = customTitle
            ? `Chart Title (REQUIRED): Set "title" to "${customTitle}".`
            : "Chart Title: Generate a concise, descriptive chart \"title\".";

          const detailsDirective = customDetails
            ? `Additional User Instructions: ${customDetails}`
            : "";

          const instruction = [
            "Generate a native Excel chart specification visualizing the active worksheet data according to these user-selected parameters:",
            `- ${typeDirective}`,
            `- ${categoryDirective}`,
            `- ${seriesDirective}`,
            `- ${aggDirective}`,
            `- ${titleDirective}`,
            detailsDirective ? `- ${detailsDirective}` : "",
            "",
            "CRITICAL OUTPUT FORMAT: Return the chart specification inside an ```excel-chart JSON code block (including \"type\", \"categoriesRange\", \"seriesRange\", optional \"aggregate\", \"title\", and \"hasHeaders\": true), followed by 2-3 concise bullet points highlighting key takeaways from the charted metrics."
          ].filter(Boolean).join("\n");

          const fullPrompt = docContext && docContext.length > 20
            ? `Spreadsheet Context:\n"""\n${docContext.substring(0, 100000)}\n"""\n\nUser Instruction:\n${instruction}`
            : instruction;

          const summaryParts = [
            `Type: ${chosenType}`,
            chosenCategory ? `X: ${chosenCategory.split("-")[0].trim()}` : "X: Auto",
            chosenSeries ? `Y: ${chosenSeries.split("-")[0].trim()}` : "Y: Auto"
          ];
          if (chosenAgg && chosenAgg !== "auto") summaryParts.push(`Agg: ${chosenAgg}`);
          if (customTitle) summaryParts.push(`Title: "${customTitle}"`);
          if (customDetails) summaryParts.push(`Note: ${customDetails}`);

          const displayBubble = `📊 [Native Chart Builder] ${summaryParts.join(" | ")}`;
          await executeGeminiWorkflow(fullPrompt, displayBubble);
          if (submitBtn) {
            submitBtn.innerHTML = "✅ Chart Specification Generated";
          }
        } catch (err) {
          console.error("Excel Native Chart Builder execution error:", err);
          appendBubble("Error generating native chart: " + err.message, "system");
          if (submitBtn) {
            submitBtn.disabled = false;
            submitBtn.style.opacity = "1";
            submitBtn.innerHTML = "📊 Generate Native Chart";
          }
          typePills.forEach((b) => { b.disabled = false; });
          [typeSelect, categorySelect, seriesSelect, aggSelect, titleInput, detailsInput].forEach((el) => {
            if (el) el.disabled = false;
          });
        }
      });
    }
  });
}

/**
 * Intent-Based Formula Generator for Excel
 * 1. Analyzes the selected range in the spreadsheet (or the entire sheet if a single cell / empty range is selected).
 * 2. Recommends the top 4-5 formulas tailored to the selected data with plain-English explanations.
 * 3. Allows the user to select the right formula from the list and insert it into the selected cell in the sheet.
 */
function buildLocalFormulaRecommendations(columnsMeta) {
  const recommendations = [];
  const numericCols = columnsMeta.filter((c) => c.isNumeric);
  const textCols = columnsMeta.filter((c) => !c.isNumeric);
  const primaryNum = numericCols[0] || columnsMeta[columnsMeta.length - 1] || null;
  const secondaryNum = numericCols[1] || null;
  const primaryText = textCols[0] || columnsMeta[0] || null;

  if (primaryNum && primaryNum.isNumeric) {
    recommendations.push({
      title: `Total ${primaryNum.headerLabel}`,
      formula: `=SUM(${primaryNum.dataRangeStr})`,
      explanation: `Adds up all values in the "${primaryNum.headerLabel}" column (${primaryNum.dataRangeStr}).`
    });

    recommendations.push({
      title: `Average ${primaryNum.headerLabel}`,
      formula: `=AVERAGE(${primaryNum.dataRangeStr})`,
      explanation: `Calculates the average value across the "${primaryNum.headerLabel}" column (${primaryNum.dataRangeStr}).`
    });

    if (primaryText && primaryText.colLetter !== primaryNum.colLetter && primaryText.sampleVal) {
      const cleanSample = String(primaryText.sampleVal).replace(/"/g, '""');
      recommendations.push({
        title: `Total ${primaryNum.headerLabel} for "${primaryText.sampleVal}"`,
        formula: `=SUMIFS(${primaryNum.dataRangeStr}, ${primaryText.dataRangeStr}, "${cleanSample}")`,
        explanation: `Sums "${primaryNum.headerLabel}" (${primaryNum.dataRangeStr}) only for rows where "${primaryText.headerLabel}" (${primaryText.dataRangeStr}) is "${primaryText.sampleVal}".`
      });
    }

    recommendations.push({
      title: `Highest ${primaryNum.headerLabel}`,
      formula: `=MAX(${primaryNum.dataRangeStr})`,
      explanation: `Finds the highest value in the "${primaryNum.headerLabel}" column (${primaryNum.dataRangeStr}).`
    });

    if (secondaryNum) {
      recommendations.push({
        title: `Total ${secondaryNum.headerLabel}`,
        formula: `=SUM(${secondaryNum.dataRangeStr})`,
        explanation: `Adds up all values in the "${secondaryNum.headerLabel}" column (${secondaryNum.dataRangeStr}).`
      });
    } else if (primaryText && primaryText.sampleVal) {
      const cleanSample = String(primaryText.sampleVal).replace(/"/g, '""');
      recommendations.push({
        title: `Count of "${primaryText.sampleVal}" Rows`,
        formula: `=COUNTIF(${primaryText.dataRangeStr}, "${cleanSample}")`,
        explanation: `Counts how many rows in "${primaryText.headerLabel}" (${primaryText.dataRangeStr}) match "${primaryText.sampleVal}".`
      });
    } else {
      recommendations.push({
        title: `Lowest ${primaryNum.headerLabel}`,
        formula: `=MIN(${primaryNum.dataRangeStr})`,
        explanation: `Finds the lowest value in the "${primaryNum.headerLabel}" column (${primaryNum.dataRangeStr}).`
      });
    }
  } else if (primaryText) {
    recommendations.push({
      title: `Total Non-Empty Entries in ${primaryText.headerLabel}`,
      formula: `=COUNTA(${primaryText.dataRangeStr})`,
      explanation: `Counts all non-empty cells in the "${primaryText.headerLabel}" column (${primaryText.dataRangeStr}).`
    });
    if (primaryText.sampleVal) {
      const cleanSample = String(primaryText.sampleVal).replace(/"/g, '""');
      recommendations.push({
        title: `Count of "${primaryText.sampleVal}"`,
        formula: `=COUNTIF(${primaryText.dataRangeStr}, "${cleanSample}")`,
        explanation: `Counts how many entries in "${primaryText.headerLabel}" (${primaryText.dataRangeStr}) equal "${primaryText.sampleVal}".`
      });
    }
    recommendations.push({
      title: `Unique Values in ${primaryText.headerLabel}`,
      formula: `=UNIQUE(${primaryText.dataRangeStr})`,
      explanation: `Lists all distinct values from the "${primaryText.headerLabel}" column (${primaryText.dataRangeStr}).`
    });
    if (columnsMeta.length > 1) {
      const secondCol = columnsMeta[1];
      const lookupVal = primaryText.sampleVal ? `"${String(primaryText.sampleVal).replace(/"/g, '""')}"` : `${primaryText.colLetter}2`;
      recommendations.push({
        title: `Look Up ${secondCol.headerLabel} by ${primaryText.headerLabel}`,
        formula: `=XLOOKUP(${lookupVal}, ${primaryText.dataRangeStr}, ${secondCol.dataRangeStr}, "Not Found")`,
        explanation: `Finds ${lookupVal} in "${primaryText.headerLabel}" (${primaryText.dataRangeStr}) and returns the matching "${secondCol.headerLabel}" (${secondCol.dataRangeStr}).`
      });
    }
  } else {
    recommendations.push(
      {
        title: "Sum a Range of Numbers",
        formula: "=SUM(A2:A20)",
        explanation: "Adds all numeric values in cells A2 through A20."
      },
      {
        title: "Average of Numbers",
        formula: "=AVERAGE(A2:A20)",
        explanation: "Calculates the arithmetic mean of values in cells A2 through A20."
      },
      {
        title: "Count Non-Empty Cells",
        formula: "=COUNTA(A2:A20)",
        explanation: "Counts how many cells in A2 through A20 are not blank."
      },
      {
        title: "Maximum Value",
        formula: "=MAX(A2:A20)",
        explanation: "Finds the largest value in cells A2 through A20."
      }
    );
  }

  return recommendations.slice(0, 5);
}

function parseFormulaRecommendations(rawText, fallbackList = []) {
  const results = [];
  const seenFormulas = new Set();

  const addCandidate = (title, formula, explanation) => {
    if (!formula || typeof formula !== "string") return;
    let cleanF = formula.trim().replace(/^`+|`+$/g, "").trim();
    if (!cleanF) return;
    if (!cleanF.startsWith("=")) cleanF = "=" + cleanF;
    // Basic sanity check: starts with =FUNC(...) or arithmetic expression and has balanced parentheses
    const openParens = (cleanF.match(/\(/g) || []).length;
    const closeParens = (cleanF.match(/\)/g) || []).length;
    if (openParens === 0 || openParens !== closeParens) return;
    if (!/^=[A-Za-z_][A-Za-z0-9_.]*\s*\(/.test(cleanF)) return;

    const normKey = cleanF.toUpperCase().replace(/\s+/g, "");
    if (seenFormulas.has(normKey)) return;
    seenFormulas.add(normKey);

    const cleanTitle = (title || "").replace(/^[0-9]+[\).\s-]+/, "").replace(/^\*+|\*+$/g, "").trim() || "Recommended Formula";
    const cleanExp = (explanation || "").replace(/^\*+|\*+$/g, "").trim() || `Calculates ${cleanTitle.toLowerCase()} using ${cleanF}.`;

    results.push({
      title: cleanTitle,
      formula: cleanF,
      explanation: cleanExp
    });
  };

  const text = String(rawText || "");

  // 1. Try parsing JSON block (```excel-formulas, ```json, or raw JSON array)
  const jsonBlocks = [];
  const codeBlockRegex = /```(?:excel-formulas|json|excel)?\s*([\s\S]*?)```/gi;
  let m;
  while ((m = codeBlockRegex.exec(text)) !== null) {
    if (m[1] && m[1].includes("[")) {
      jsonBlocks.push(m[1].trim());
    }
  }
  const firstBracket = text.indexOf("[");
  const lastBracket = text.lastIndexOf("]");
  if (firstBracket !== -1 && lastBracket > firstBracket) {
    jsonBlocks.push(text.slice(firstBracket, lastBracket + 1));
  }

  for (const candidateJson of jsonBlocks) {
    try {
      const parsed = JSON.parse(candidateJson);
      const arr = Array.isArray(parsed) ? parsed : (parsed && Array.isArray(parsed.formulas) ? parsed.formulas : null);
      if (Array.isArray(arr)) {
        arr.forEach((item) => {
          if (item && typeof item === "object") {
            addCandidate(
              item.title || item.name || item.label || item.intent,
              item.formula || item.expression || item.code,
              item.explanation || item.description || item.summary
            );
          }
        });
      }
      if (results.length >= 4) break;
    } catch (_) {
      // Continue to fallback regex extraction
    }
  }

  // 2. Supplement from fallbackList if fewer than 4 recommendations were parsed
  if (results.length < 4 && Array.isArray(fallbackList)) {
    for (const fb of fallbackList) {
      if (results.length >= 5) break;
      addCandidate(fb.title, fb.formula, fb.explanation);
    }
  }

  return results.slice(0, 5);
}

async function handleExcelFormulaGeneratorClick(chip) {
  clearChatAndInput();
  const loadingText = document.getElementById("loading");
  if (loadingText) {
    loadingText.innerText = "⚡ Analyzing spreadsheet data & recommending formulas...";
    loadingText.style.display = "block";
  }

  let sheetName = "Sheet1";
  let scopeMode = "sheet"; // "selection" or "sheet"
  let analyzedAddress = "";
  let columnsMeta = [];
  let sampleTableMarkdown = "";

  try {
    if (typeof Excel !== "undefined") {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const usedRange = sheet.getUsedRangeOrNullObject();
        const selectedRange = context.workbook.getSelectedRange();

        sheet.load("name");
        usedRange.load("isNullObject, address, rowIndex, columnIndex, rowCount, columnCount");
        selectedRange.load("address, rowIndex, columnIndex, rowCount, columnCount");
        await context.sync();

        sheetName = sheet.name || "Sheet1";

        const selCells = (selectedRange.rowCount || 1) * (selectedRange.columnCount || 1);
        let useSelection = false;
        let selSampleVals = null;

        if (selCells > 1) {
          const checkRows = Math.min(selectedRange.rowCount, 20);
          const checkCols = Math.min(selectedRange.columnCount, 20);
          const selCheckRange = sheet.getRangeByIndexes(
            selectedRange.rowIndex,
            selectedRange.columnIndex,
            checkRows,
            checkCols
          );
          selCheckRange.load("values");
          await context.sync();
          selSampleVals = selCheckRange.values || [];
          let nonEmptyInSel = 0;
          for (let r = 0; r < selSampleVals.length; r++) {
            for (let c = 0; c < (selSampleVals[r] || []).length; c++) {
              const v = selSampleVals[r][c];
              if (v !== null && v !== undefined && String(v).trim() !== "") {
                nonEmptyInSel++;
              }
            }
          }
          if (nonEmptyInSel >= 2) {
            useSelection = true;
          }
        }

        const targetRangeInfo = useSelection ? selectedRange : (!usedRange.isNullObject ? usedRange : null);
        scopeMode = useSelection ? "selection" : "sheet";

        if (targetRangeInfo && targetRangeInfo.rowCount > 0 && targetRangeInfo.columnCount > 0) {
          const rawAddr = targetRangeInfo.address || "";
          analyzedAddress = rawAddr.includes("!") ? rawAddr.split("!").pop() : rawAddr;

          const sampleRowsCount = Math.min(targetRangeInfo.rowCount, 20);
          const sampleColsCount = Math.min(targetRangeInfo.columnCount, 20);

          const sampleRange = sheet.getRangeByIndexes(
            targetRangeInfo.rowIndex,
            targetRangeInfo.columnIndex,
            sampleRowsCount,
            sampleColsCount
          );
          sampleRange.load("values");

          // Also load the worksheet's top header row for these columns in case a sub-range without headers was selected
          const sheetHeaderRowIdx = !usedRange.isNullObject ? usedRange.rowIndex : 0;
          let sheetHeaderRange = null;
          if (targetRangeInfo.rowIndex > sheetHeaderRowIdx) {
            sheetHeaderRange = sheet.getRangeByIndexes(
              sheetHeaderRowIdx,
              targetRangeInfo.columnIndex,
              1,
              sampleColsCount
            );
            sheetHeaderRange.load("values");
          }

          await context.sync();

          const vals = sampleRange.values || [];
          const externalHeaders = (sheetHeaderRange && sheetHeaderRange.values && sheetHeaderRange.values[0])
            ? sheetHeaderRange.values[0]
            : null;

          // Determine if first row of targetRangeInfo is a header row
          const firstRow = vals[0] || [];
          let firstRowLooksLikeHeader = targetRangeInfo.rowIndex === sheetHeaderRowIdx && vals.length > 1;
          if (!firstRowLooksLikeHeader && vals.length > 1) {
            let firstRowStrings = 0;
            let secondRowNumbers = 0;
            for (let c = 0; c < firstRow.length; c++) {
              const v0 = firstRow[c];
              const v1 = vals[1] ? vals[1][c] : null;
              if (typeof v0 === "string" && isNaN(Number(v0.replace(/[$,%\s,]/g, "")))) {
                firstRowStrings++;
              }
              if (typeof v1 === "number" || (v1 && !isNaN(Number(String(v1).replace(/[$,%\s,]/g, ""))))) {
                secondRowNumbers++;
              }
            }
            if (firstRowStrings > 0 && secondRowNumbers > 0) {
              firstRowLooksLikeHeader = true;
            }
          }

          const startRowNum = targetRangeInfo.rowIndex + 1;
          const dataStartRowNum = firstRowLooksLikeHeader && targetRangeInfo.rowCount > 1
            ? startRowNum + 1
            : startRowNum;
          const endRowNum = targetRangeInfo.rowIndex + targetRangeInfo.rowCount;
          const dataLoopStartIdx = firstRowLooksLikeHeader && vals.length > 1 ? 1 : 0;

          const toColLetter = (colIdx) => {
            let temp = colIdx;
            let letter = "";
            while (temp >= 0) {
              letter = String.fromCharCode((temp % 26) + 65) + letter;
              temp = Math.floor(temp / 26) - 1;
            }
            return letter;
          };

          for (let c = 0; c < sampleColsCount; c++) {
            const colLetter = toColLetter(targetRangeInfo.columnIndex + c);
            let rawHeader = "";
            if (firstRowLooksLikeHeader && firstRow[c] !== null && firstRow[c] !== undefined) {
              rawHeader = String(firstRow[c]).trim();
            } else if (externalHeaders && externalHeaders[c] !== null && externalHeaders[c] !== undefined) {
              rawHeader = String(externalHeaders[c]).trim();
            }
            const headerLabel = rawHeader || `Column ${colLetter}`;

            let numericCount = 0;
            let nonEmptyCount = 0;
            let sampleVal = "";

            for (let r = dataLoopStartIdx; r < vals.length; r++) {
              const cellVal = vals[r] ? vals[r][c] : null;
              if (cellVal !== null && cellVal !== undefined && String(cellVal).trim() !== "") {
                nonEmptyCount++;
                if (!sampleVal) sampleVal = String(cellVal).trim().substring(0, 24);
                const cleanedNum = String(cellVal).replace(/[$,%\s,]/g, "");
                if (typeof cellVal === "number" || (!isNaN(Number(cleanedNum)) && cleanedNum !== "")) {
                  numericCount++;
                }
              }
            }

            if (nonEmptyCount > 0 || rawHeader) {
              const isNumeric = nonEmptyCount > 0 && numericCount >= Math.ceil(nonEmptyCount * 0.5);
              const dataRangeStr = `${colLetter}${dataStartRowNum}:${colLetter}${endRowNum}`;
              columnsMeta.push({
                colLetter,
                headerLabel,
                isNumeric,
                sampleVal,
                dataRangeStr,
                dataStartRowNum,
                endRowNum
              });
            }
          }

          // Build compact markdown table snapshot for Gemini
          if (columnsMeta.length > 0 && vals.length > 0) {
            const headerCells = columnsMeta.map((col) => `Col ${col.colLetter} (${col.headerLabel})`);
            const mdLines = [
              `| Row | ${headerCells.join(" | ")} |`,
              `| --- | ${columnsMeta.map(() => "---").join(" | ")} |`
            ];
            const maxSampleRows = Math.min(vals.length, dataLoopStartIdx + 12);
            for (let r = dataLoopStartIdx; r < maxSampleRows; r++) {
              const excelRowNum = targetRangeInfo.rowIndex + r + 1;
              const rowCells = columnsMeta.map((_, cIdx) => {
                const v = vals[r] && vals[r][cIdx] !== undefined && vals[r][cIdx] !== null
                  ? String(vals[r][cIdx]).replace(/\|/g, "/").replace(/[\r\n]+/g, " ").trim()
                  : "";
                return v.substring(0, 30);
              });
              mdLines.push(`| ${excelRowNum} | ${rowCells.join(" | ")} |`);
            }
            sampleTableMarkdown = mdLines.join("\n");
          }
        }
      });
    }
  } catch (scanErr) {
    console.warn("Formula Generator range analysis warning:", scanErr);
  }

  const scopeLabel = scopeMode === "selection" && analyzedAddress
    ? `Selected Range (${analyzedAddress}) on ${sheetName}`
    : (analyzedAddress ? `Entire Sheet (${analyzedAddress}) on ${sheetName}` : `Active Sheet (${sheetName})`);

  appendBubble(`✨ [Formula Generator] Analyzing ${scopeLabel}...`, "user");

  const localFallback = buildLocalFormulaRecommendations(columnsMeta);
  let recommendations = localFallback;

  try {
    const columnSchemaLines = columnsMeta.length > 0
      ? columnsMeta.map((c) => `- Column ${c.colLetter} ("${c.headerLabel}"): range ${c.dataRangeStr}, type=${c.isNumeric ? "Numeric" : "Text/Category"}${c.sampleVal ? `, sample="${c.sampleVal}"` : ""}`).join("\n")
      : "- General worksheet columns";

    const prompt = [
      "You are an expert Microsoft Excel assistant helping a user generate formulas based on their intent without needing to know Excel syntax.",
      `Analyze the following spreadsheet data from ${scopeLabel} and recommend the top 4 to 5 most useful Excel formulas for this data.`,
      "",
      `Sheet Name: "${sheetName}"`,
      `Analyzed Scope: ${scopeLabel}`,
      "Columns & Exact Data Ranges:",
      columnSchemaLines,
      "",
      sampleTableMarkdown ? `Sample Data Rows:\n${sampleTableMarkdown}\n` : "",
      "REQUIREMENTS:",
      "1. Provide ONLY 4 to 5 recommended Excel formulas that are most relevant and useful for the columns and values above (e.g., total sum, average, conditional sum/count for a specific category in the data, highest/lowest value, or lookup).",
      "2. Each formula MUST start with `=` and use the EXACT column letters and row numbers from the schema above (do NOT invent column letters or row ranges that are not in the schema).",
      "3. For each formula, provide:",
      "   - \"title\": Short, plain-English intent title (3 to 6 words, e.g., \"Total Sales Revenue\", \"Average Deal Size\", \"Sales in West Region\").",
      "   - \"formula\": The exact Excel formula starting with `=`.",
      "   - \"explanation\": A clear, simple 1-sentence explanation of what this formula calculates with respect to the selected sheet data.",
      "4. Return ONLY a JSON array inside an ```excel-formulas code block:",
      "```excel-formulas",
      "[",
      "  {",
      "    \"title\": \"...\",",
      "    \"formula\": \"=...\",",
      "    \"explanation\": \"...\"",
      "  }",
      "]",
      "```"
    ].filter(Boolean).join("\n");

    const aiData = await askGeminiEnterprise(prompt, [], null, false);
    if (aiData && aiData.sessionId) {
      SessionManager.registerSessionHost(aiData.sessionId, "Excel");
    }
    if (aiData && aiData.result) {
      const parsed = parseFormulaRecommendations(aiData.result, localFallback);
      if (parsed && parsed.length > 0) {
        recommendations = parsed;
      }
    }
  } catch (aiErr) {
    console.warn("Using local schema-based formula recommendations due to AI error:", aiErr);
  } finally {
    if (loadingText) loadingText.style.display = "none";
  }

  const optionsHtml = recommendations.map((rec, idx) => {
    const isSelected = idx === 0;
    const borderCol = isSelected ? "#0b57d0" : "#e2e8f0";
    const bgCol = isSelected ? "#e8f0fe" : "#ffffff";
    const radioDot = isSelected ? "●" : "○";
    const radioColor = isSelected ? "#0b57d0" : "#94a3b8";

    return `
      <div class="fg-rec-option ${isSelected ? "selected" : ""}" data-idx="${idx}" data-formula="${escapeHtml(rec.formula)}" style="display:flex; align-items:flex-start; gap:10px; padding:10px 12px; border:1.5px solid ${borderCol}; border-radius:8px; background:${bgCol}; cursor:pointer; transition:all 0.15s ease;">
        <div class="fg-rec-radio" style="font-size:15px; line-height:1.2; color:${radioColor}; font-weight:700; margin-top:1px; user-select:none;">${radioDot}</div>
        <div style="flex:1; min-width:0;">
          <div style="display:flex; align-items:center; justify-content:space-between; gap:6px; flex-wrap:wrap;">
            <span style="font-weight:600; font-size:12.5px; color:#0f172a;">${escapeHtml(rec.title)}</span>
            <code style="font-size:11px; font-family:monospace; color:#0b57d0; background:#e8f0fe; padding:2px 6px; border-radius:4px; border:1px solid #d3e3fd; word-break:break-all;">${escapeHtml(rec.formula)}</code>
          </div>
          <div style="font-size:11.5px; color:#475569; margin-top:4px; line-height:1.4;">
            ${escapeHtml(rec.explanation)}
          </div>
        </div>
      </div>
    `;
  }).join("");

  const cardHtml = `
    <div class="excel-formula-gen-wizard" style="display:flex; flex-direction:column; gap:10px;">
      <div style="border-bottom:1px solid #e2e8f0; padding-bottom:8px;">
        <div style="font-weight:700; font-size:14px; color:#0f172a;">✨ Recommended Formulas</div>
        <div style="font-size:11.5px; color:#475569; margin-top:2px;">
          Analyzed: <strong>${escapeHtml(scopeLabel)}</strong>
        </div>
      </div>

      <div style="font-size:11.5px; color:#334155;">
        Select what you would like to calculate below, click the target cell in your sheet, and click <strong>Insert Formula</strong>:
      </div>

      <div class="fg-rec-list" style="display:flex; flex-direction:column; gap:8px;">
        ${optionsHtml}
      </div>

      <div style="display:flex; flex-direction:column; gap:6px; margin-top:4px;">
        <div class="pill-btn-group">
          <button type="button" class="pill-btn primary fg-insert-btn">
            <span class="material-symbols-outlined">add</span> Insert Formula into Selected Cell
          </button>
        </div>
        <div class="fg-status-msg" style="display:none; font-size:11.5px; padding:7px 10px; border-radius:6px; line-height:1.35;"></div>
      </div>
    </div>
  `;

  appendInteractiveBubble(cardHtml, (bubble) => {
    const optionCards = bubble.querySelectorAll(".fg-rec-option");
    const insertBtn = bubble.querySelector(".fg-insert-btn");
    const statusMsg = bubble.querySelector(".fg-status-msg");

    let selectedFormula = recommendations[0] ? recommendations[0].formula : "=SUM(A1:A10)";

    optionCards.forEach((card) => {
      card.addEventListener("click", () => {
        selectedFormula = card.getAttribute("data-formula") || selectedFormula;
        optionCards.forEach((c) => {
          const isMatch = c === card;
          c.classList.toggle("selected", isMatch);
          c.style.borderColor = isMatch ? "#0b57d0" : "#e2e8f0";
          c.style.background = isMatch ? "#e8f0fe" : "#ffffff";
          const radio = c.querySelector(".fg-rec-radio");
          if (radio) {
            radio.textContent = isMatch ? "●" : "○";
            radio.style.color = isMatch ? "#0b57d0" : "#94a3b8";
          }
        });
        if (statusMsg) {
          statusMsg.style.display = "none";
        }
      });
    });

    if (insertBtn) {
      insertBtn.addEventListener("click", async () => {
        if (!selectedFormula) return;
        const showStatus = (text, isError = false) => {
          if (!statusMsg) return;
          statusMsg.style.display = "block";
          statusMsg.style.background = isError ? "#fef2f2" : "#e8f0fe";
          statusMsg.style.border = isError ? "1px solid #fecaca" : "1px solid #d3e3fd";
          statusMsg.style.color = isError ? "#991b1b" : "#0b57d0";
          statusMsg.innerHTML = text;
        };

        insertBtn.disabled = true;
        insertBtn.style.opacity = "0.7";
        insertBtn.innerHTML = `<span class="material-symbols-outlined">hourglass_top</span> Inserting Formula...`;

        try {
          if (typeof Excel !== "undefined") {
            let insertedCellAddr = "";
            let multiCellBlocked = false;
            let currentSelAddr = "";

            await Excel.run(async (context) => {
              const range = context.workbook.getSelectedRange();
              range.load("address, rowCount, columnCount, rowIndex, columnIndex");
              await context.sync();

              const rawAddr = range.address || "";
              currentSelAddr = rawAddr.includes("!") ? rawAddr.split("!").pop() : rawAddr;

              if ((range.rowCount || 1) > 1 || (range.columnCount || 1) > 1) {
                multiCellBlocked = true;
                return;
              }

              insertedCellAddr = currentSelAddr || "selected cell";
              const formulaMatrix = [[selectedFormula]];

              try {
                range.formulas = formulaMatrix;
                await context.sync();
              } catch (formErr) {
                try {
                  range.formulasLocal = formulaMatrix;
                  await context.sync();
                } catch (locErr) {
                  range.values = formulaMatrix;
                  await context.sync();
                }
              }

              try {
                range.format.font.name = "Segoe UI";
                range.format.font.size = 11;
                range.format.autofitColumns();
                await context.sync();
              } catch (_) {}
            });

            if (multiCellBlocked) {
              showStatus(
                `⚠️ Range <code>${escapeHtml(currentSelAddr)}</code> is currently highlighted. Please click a <strong>single target cell</strong> in Excel where you want the formula placed, then click <strong>Insert Formula</strong>.`,
                true
              );
            } else {
              showStatus(
                `✅ Inserted <code>${escapeHtml(selectedFormula)}</code> into cell <strong>${escapeHtml(insertedCellAddr)}</strong>!`
              );
            }
          } else if (hostAdapter && typeof hostAdapter.insertContent === "function") {
            await hostAdapter.insertContent(selectedFormula, selectedFormula, "formula");
            showStatus(`✅ Inserted <code>${escapeHtml(selectedFormula)}</code> into active cell!`);
          }
        } catch (insertErr) {
          console.error("Formula insertion error:", insertErr);
          showStatus(`⚠️ Could not insert formula: ${escapeHtml(insertErr.message || String(insertErr))}`, true);
        } finally {
          insertBtn.disabled = false;
          insertBtn.style.opacity = "1";
          insertBtn.innerHTML = `<span class="material-symbols-outlined">add</span> Insert Formula into Selected Cell`;
        }
      });
    }
  });
}

function isDocumentSpecificQuery(text) {
  if (!text) return false;
  return /\b(based on (this|the active|the current) (doc|document|presentation|deck|slide|sheet|spreadsheet|file)|according to (this|the active|the current) (doc|document|presentation|deck|slide|sheet|spreadsheet|file)|in (this|the active|the current) (doc|document|presentation|deck|slide|sheet|spreadsheet|file)|from (this|the active|the current) (doc|document|presentation|deck|slide|sheet|spreadsheet|file)|about (this|the active|the current) (doc|document|presentation|deck|slide|sheet|spreadsheet|file)|summarize (this|the active|the current) (doc|document|presentation|deck|slide|sheet|spreadsheet|file)|(what|does) (this|the active|the current) (doc|document|presentation|deck|slide|sheet|spreadsheet|file) (say|state|mention|contain)|active document only|this document only|strictly (from|based on) (this|the active) (doc|document|presentation|deck|slide|sheet|spreadsheet|file)|only based on (this|the active) (doc|document|presentation|deck|slide|sheet|spreadsheet|file))\b/i.test(text);
}

function isStrictDocumentOnlyQuery(text) {
  if (!text) return false;
  return isDocumentSpecificQuery(text);
}

function isSelectionSpecificQuery(text) {
  if (!text) return false;
  return /\b(selected (text|content|paragraph|sentence|part)|selection|highlighted (text|content))\b/i.test(text);
}

// Dedicated PowerPoint Slide Intelligence Action (Targets Highlighted / Selected Slides)
// Refine for PowerPoint responses.
//
// The generic Refine prompt says "rewrite the above response" and relies on the
// server session to know what that is. With enterprise grounding on, the model
// rewrote whatever the datastore search returned instead - an unrelated
// document - and the "natural prose, no bullet points" wording flattened a
// 10-slide deck into a single paragraph. Here the response text is sent
// verbatim, grounding is off, and a deck keeps the same slides, in order.
const PPT_REFINE_STYLE = {
  formal: "Rewrite it in an authoritative, formal executive tone.",
  casual: "Rewrite it in an approachable, clear, conversational tone.",
  shorten: "Make the wording significantly more concise and punchy.",
  elaborate: "Expand it with more in-depth operational and analytical detail.",
  bullets: "Present the content as clean, high-impact bullet points.",
  table: "Where the content is comparative or quantitative, structure it as a clean markdown table with clear column headers.",
  retry: "Rewrite it from scratch with a fresh structure, varied vocabulary and enhanced clarity."
};

/**
 * Refine asks Gemini to return the same N slides, but a rewrite pass can still
 * drop a slide's body or mangle its JSON visual block - the casual-tone refine
 * once returned slide 5 as a heading and subtitle only, which built a blank
 * slide. Compares each returned slide with the source slide and falls back to
 * the source content wherever the refined one is empty, broken or missing.
 */
function splitDeckBySlideHeading(text) {
  const src = String(text || "");
  const headingRe = /^#{1,3}\s*Slide\s+(\d+)\b.*$/gim;
  const marks = [];
  let m;
  while ((m = headingRe.exec(src)) !== null) {
    marks.push({ num: Number(m[1]), heading: m[0], start: m.index, bodyStart: m.index + m[0].length });
  }
  const slides = marks.map((mark, i) => {
    const end = i + 1 < marks.length ? marks[i + 1].start : src.length;
    const body = src.slice(mark.bodyStart, end).replace(/\n\s*-{3,}\s*$/, "").replace(/^\s*\n/, "");
    return { num: mark.num, heading: mark.heading, body: body.replace(/\s+$/, "") };
  });
  return { preamble: marks.length ? src.slice(0, marks[0].start) : src, slides };
}

function slideBodyHasContent(body) {
  return String(body || "")
    .split("\n")
    .some((line) => {
      const t = line.trim();
      if (!t || /^-{3,}$/.test(t) || /^#{3,6}\s/.test(t)) return false;
      return /[A-Za-z0-9]/.test(t.replace(/^[-*•>#\s]+/, ""));
    });
}

function slideJsonState(body) {
  const text = String(body || "");
  if (!/`{3}\s*json|^\s*\{\s*"/im.test(text)) return "none";
  if ((text.match(/`{3}/g) || []).length % 2 !== 0) return "broken";
  const blocks = text.match(/`{3}(?:json)?\s*([\s\S]*?)`{3}/gi) || [];
  for (const block of blocks) {
    const inner = block.replace(/^`{3}(?:json)?/i, "").replace(/`{3}$/, "").trim();
    if (!inner.startsWith("{")) continue;
    try {
      JSON.parse(inner);
    } catch (_) {
      return "broken";
    }
  }
  return "ok";
}

function restoreDroppedSlideContent(responseText, sourceText, expectedCount) {
  const response = splitDeckBySlideHeading(responseText);
  const source = splitDeckBySlideHeading(sourceText);
  if (response.slides.length === 0 || source.slides.length === 0) return responseText;

  const sourceByNum = new Map(source.slides.map((s) => [s.num, s]));
  const restored = [];
  const out = response.slides.map((slide) => {
    const original = sourceByNum.get(slide.num);
    if (!original || !slideBodyHasContent(original.body)) return slide;

    if (!slideBodyHasContent(slide.body)) {
      restored.push(`${slide.num} (empty)`);
      // Keep the refined subtitle when there is one; take everything else from the source.
      const refinedSubtitle = (slide.body.match(/^\s*#{3,6}\s.*$/m) || [null])[0];
      const originalBody = refinedSubtitle ? original.body.replace(/^\s*#{3,6}\s.*$/m, "").replace(/^\s*\n/, "") : original.body;
      return { ...slide, body: refinedSubtitle ? `${refinedSubtitle.trim()}\n\n${originalBody}` : originalBody };
    }

    if (slideJsonState(slide.body) === "broken" && slideJsonState(original.body) === "ok") {
      restored.push(`${slide.num} (visual)`);
      const originalBlocks = original.body.match(/`{3}(?:json)?\s*[\s\S]*?`{3}/gi) || [];
      const withoutJson = slide.body
        .replace(/`{3}(?:json)?\s*[\s\S]*?`{3}/gi, "")
        .replace(/`{3}(?:json)?[\s\S]*$/i, "")
        .replace(/\s+$/, "");
      return { ...slide, body: `${withoutJson}\n\n${originalBlocks.join("\n\n")}` };
    }
    return slide;
  });

  // The contract is "exactly N slides": put back any slide the model dropped.
  const present = new Set(out.map((s) => s.num));
  const total = Number(expectedCount) || source.slides.length;
  source.slides.forEach((s) => {
    if (s.num <= total && !present.has(s.num)) {
      restored.push(`${s.num} (missing)`);
      out.push(s);
    }
  });
  if (restored.length === 0) return responseText;

  out.sort((a, b) => a.num - b.num);
  console.warn(`[Refine] Kept the original content for slide(s) ${restored.join(", ")} - the refined version came back without it.`);
  const preamble = response.preamble.trim();
  const deck = out.map((s) => `${s.heading}\n${s.body}`).join("\n\n---\n\n");
  return preamble ? `${preamble}\n\n${deck}` : deck;
}

async function runPowerPointRefine(key, label, sourceText, info = {}) {
  const style = PPT_REFINE_STYLE[key];
  const source = String(sourceText || "").trim().substring(0, 60000);
  if (!style || !source) return;

  const slideCount = Number(info.slideCount) || 0;
  let prompt;
  if (info.isDeck && slideCount > 1) {
    const perSlide = key === "bullets" || key === "table" ? ", applying the instruction inside each slide" : "";
    prompt = `You are revising an existing PowerPoint deck. ${style}

STRICT DECK CONTRACT:
1. Output EXACTLY ${slideCount} slides: the same ${slideCount} slides, in the same order, as the deck below. Never merge, split, drop or add slides.
2. Start every slide with its own markdown heading in the form "## Slide N: Title", numbered 1 to ${slideCount}.
3. Keep each slide's topic. Change only the wording, tone, length or format as instructed${perSlide}.
4. Keep each slide's subtitle line, tables and JSON chart or visual code blocks, updating their text only where the instruction requires it.
5. Use ONLY the content of the deck below. Do not introduce outside facts, sources or topics.
6. Output only the revised deck in markdown, with no preamble or closing remarks.

DECK TO REVISE:
"""
${source}
"""`;
  } else {
    prompt = `${style}

Revise ONLY the content below. Keep its topic and facts, and do not introduce outside facts, sources or topics.
Provide exactly ONE definitive revision. Start with a short markdown heading (## Title, at most 8 words), then the revised content. No preamble or closing remarks.

CONTENT TO REVISE:
"""
${source}
"""`;
  }

  if (typeof window !== "undefined") {
    window.__isSummarizeSlidesAction = Boolean(info.isSummarize);
    window.__forceSingleSlide = false;
  }
  console.log(`[Refine] PowerPoint refine "${key}" on ${info.isDeck ? `a ${slideCount}-slide deck` : "a single response"} (grounding off).`);
  const refineOptions = info.isDeck && slideCount > 1
    ? { transformResponse: (text) => restoreDroppedSlideContent(text, source, slideCount) }
    : {};
  await executeGeminiWorkflow(prompt, label, false, refineOptions);
}

async function runPowerPointSlideAction(actionType) {
  clearChatAndInput();
  const loadingText = document.getElementById("loading");
  if (loadingText) {
    loadingText.innerText = "⚡ Reading slide context...";
    loadingText.style.display = "block";
  }

  let slides = [];
  try {
    if (hostAdapter && typeof hostAdapter.getSelectedSlidesText === "function") {
      slides = await hostAdapter.getSelectedSlidesText();
    }
  } catch (e) {
    console.warn("Could not retrieve selected slides:", e);
  }

  let slideContext = "";
  let slideLabel = "presentation deck";
  let displayBubble = "";

  if (actionType === "summarize" || actionType === "takeaways") {
    if (typeof window !== "undefined") {
      window.__isSummarizeSlidesAction = true;
      window.__forceSingleSlide = false;
    }
    if (slides && slides.length >= 2) {
      slideLabel = `${slides.length} highlighted slides`;
      slideContext = slides.map((s) => `[Highlighted Slide ${s.slideNumber}]:\n${s.text || "(No readable text)"}`).join("\n\n---\n\n");
      displayBubble = `📊 [Summarize Slides] Generating executive summary presentation (up to 5 slides) from ${slides.length} highlighted slides...`;
    } else {
      appendBubble(
        "ℹ️ **Summarize Slides**: Analyzing all slides across the entire presentation to generate a comprehensive deck summary (up to 5 slides)...",
        "assistant"
      );
      const fullDocText = await hostAdapter.getFullDocumentText();
      slideContext = fullDocText || (slides && slides.length === 1 ? `[Slide ${slides[0].slideNumber}]:\n${slides[0].text}` : "");
      slideLabel = "entire presentation deck";
      displayBubble = `📊 [Summarize Slides] Summarizing entire presentation deck (up to 5 slides)...`;
    }
  } else {
    if (typeof window !== "undefined") window.__isSummarizeSlidesAction = false;
    if (slides && slides.length > 0) {
      slideLabel = slides.length === 1 ? `Slide ${slides[0].slideNumber || 1}` : `${slides.length} highlighted slides`;
      slideContext = slides.map((s) => `[Highlighted Slide ${s.slideNumber || 1}]:\n${s.text || "(No readable text)"}`).join("\n\n---\n\n");
    } else {
      const fullDocText = await hostAdapter.getFullDocumentText();
      if (fullDocText && fullDocText.length > 10) {
        slideContext = fullDocText;
        slideLabel = "presentation deck";
      }
    }
  }

  let taskInstruction = "";

  switch (actionType) {
    case "risks":
      taskInstruction = `Based on the context from the ${slideLabel} provided below, extract and analyze all Key Strategic & Operational Risks, Blockers, Gaps, and Dependencies. Create content for a new PowerPoint slide titled "⚠️ Key Risks & Mitigations" with high-impact bullet points, severity assessments, and concrete mitigation actions.`;
      displayBubble = `⚠️ [Key Risks] Generating new risk analysis slide from ${slideLabel}...`;
      break;
    case "summarize":
    case "takeaways":
      taskInstruction = `Based strictly on the content from the ${slideLabel} provided below, create a comprehensive executive summary presentation (strictly between 3 and 5 slides maximum).
CRITICAL CLOSED-BOOK GROUNDING CONTRACT:
1. STRICT BOUNDARY: You are operating in 100% CLOSED-BOOK MODE. You must synthesize information ONLY AND EXCLUSIVELY from the text and data explicitly present in the selected slides below.
2. ZERO EXTERNAL KNOWLEDGE / ZERO HALLUCINATIONS: Do NOT introduce outside industry context, external market facts, assumptions, or topics that are not explicitly stated in the selected slides below. If a point or metric is not directly written in the provided slide text, omit it completely.
3. MULTI-SLIDE STRUCTURE & FORMAT (Strictly between 3 to 5 slides maximum):
   - You must generate AT MOST 5 slides total (strictly 3 to 5 slides). NEVER generate 6 or more slides.
   - Separate distinct topics, tables, and visual charts into their own slides (e.g. ## Slide 1: [Executive Overview / Main Metrics Table], ## Slide 2: [Category Breakdown / Visual Chart / Details Table], etc.).
   - If there are multiple tables or data sets, place each table on its own appropriate slide.
   - If a visual chart represents data, output a structured JSON code block with the exact data metrics (chartType: "doughnut" or "bar", title: "...", data: [...]) so our client presentation engine can render a crisp chart.
   - Break down the key takeaways into a dedicated single slide titled "## 📊 Executive Summary: Key Takeaways" with at most 3 to 4 concise executive bullet points (max 15 words per bullet). CRITICAL: Generate strictly at most 4 bullet points; do NOT exceed 4 bullet points or create continuation slides.
   - Format each slide with a clear markdown header (## Slide 1: [Title], ## Slide 2: [Title], etc.) so each section generates its own slide.
4. DO NOT output conversational preamble.`;
      if (!displayBubble) {
        displayBubble = `📊 [Summarize Slides] Generating executive summary presentation (up to 5 slides) from ${slideLabel}...`;
      }
      break;
    case "action_items":
      taskInstruction = `Based on the context from the ${slideLabel} provided below, extract all Action Items, Deliverables, Next Steps, and Ownership. Create content for a new PowerPoint slide titled "✅ Action Items & Next Steps" with actionable task bullets, suggested owners, and priority levels.`;
      displayBubble = `✅ [Action Items] Generating action items slide from ${slideLabel}...`;
      break;
  }

  let fullPrompt = "";
  if (slideContext && slideContext.length > 5) {
    fullPrompt = `CRITICAL SLIDE CONTEXT:\nThe following is the actual content extracted from the ${slideLabel}. You MUST base your summary strictly and exclusively on this content:\n"""\n${slideContext.substring(0, 50000)}\n"""\n\nTask: ${taskInstruction}`;
  } else {
    fullPrompt = taskInstruction;
  }

  // Apply PowerPoint prompt enhancer rules
  const { enhancePromptForPowerPoint } = await import('../adapters/ppt/promptEnhancer.js');
  fullPrompt = enhancePromptForPowerPoint(fullPrompt);

  await executeGeminiWorkflow(fullPrompt, displayBubble);
}

// Feature: Transform Doc to Deck (Upload document -> PowerPoint presentation)
function initDocToDeckFeature() {
  const fileInput = document.getElementById("docToDeckFileInput") || document.getElementById("docToDeckInput");
  if (!fileInput) return;

  fileInput.addEventListener("change", async (event) => {
    const files = event.target.files;
    if (!files || files.length === 0) return;

    if (typeof window !== "undefined") {
      window.__isSummarizeSlidesAction = false;
      // A document becomes a deck. Nothing about an earlier chart applies.
      window.__forceSingleSlide = false;
    }

    clearChatAndInput();
    // A new document means a new conversation: without a fresh Gemini Enterprise
    // session, earlier turns (e.g. a previous deck topic) leak into this deck.
    resetChatSession();
    const loadingText = document.getElementById("loading");
    if (loadingText) {
      loadingText.innerText = `⚡ Document(s) uploaded. Generating presentation summary deck...`;
      loadingText.style.display = "block";
    }

    try {
      const attachments = [];
      const fileNames = [];

      for (let i = 0; i < files.length; i++) {
        const file = files[i];
        fileNames.push(file.name);

        const base64Data = await new Promise((resolve, reject) => {
          const reader = new FileReader();
          reader.onload = () => {
            const result = reader.result;
            const base64 = typeof result === 'string' && result.includes(',') ? result.split(',')[1] : result;
            resolve(base64);
          };
          reader.onerror = (err) => reject(err);
          reader.readAsDataURL(file);
        });

        let mimeType = file.type;
        if (!mimeType) {
          const ext = file.name.split('.').pop().toLowerCase();
          if (ext === 'pdf') mimeType = 'application/pdf';
          else if (ext === 'docx') mimeType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
          else if (ext === 'doc') mimeType = 'application/msword';
          else if (ext === 'txt') mimeType = 'text/plain';
          else if (ext === 'md') mimeType = 'text/markdown';
          else mimeType = 'application/octet-stream';
        }

        attachments.push({
          fileName: file.name,
          mimeType: mimeType,
          fileContents: base64Data
        });
      }

      appendBubble(
        `📄 **Doc to Deck**: Uploaded \`${fileNames.join(', ')}\`.\n\nA comprehensive presentation summary is being created from your document...`,
        "assistant"
      );

      const displayUserBubble = `📄 [Doc to Deck] Uploaded: ${fileNames.join(', ')}`;
      const prompt = `Extract the key takeaways from the attached document(s) and create a structured executive presentation slide deck with a total of up to 6 slides (1 Title/Intro Slide + 5 Content Slides).

PowerPoint Slide Deck Requirements:
1. SLIDE DECK STRUCTURE (MAXIMUM 6 SLIDES TOTAL):
   - ## Slide 1: [Relevant Emoji] [Presentation Title (3-4 words max)]
     ### [Compelling Subtitle / Executive Orientation]
     Slide 1 MUST be a dedicated TITLE & EXECUTIVE SUMMARY SLIDE containing ONLY:
     * Presentation Title: A crisp, impactful title capturing the overarching subject.
     * Subtitle: An executive subtitle setting the strategic context.
     * Executive Summary: Exactly ONE small, high-impact executive summary paragraph (2 to 4 sentences max) explaining what the document is all about and its key strategic value.
     * STRICT RULES FOR SLIDE 1:
       - ONLY include Title, Subtitle, and the ONE small Executive Summary paragraph.
       - NEVER include a "Deck Scope & Outline" section!
       - NEVER list upcoming slides (e.g. "Slide 2: ...", "Slide 3: ...")!
       - NEVER include granular operational bullets, department trivia, metrics, or tables on Slide 1!
   - ## Slide 2 to Slide 6 (5 Content Slides):
     5 focused executive content slides breaking down the core insights, findings, data, and recommendations from the document.
2. Provide EXACTLY ONE presentation deck (Slide 1 Title & Summary Slide + 5 Content Slides, maximum 6 slides total). DO NOT output multiple alternative options.
3. DO NOT output conversational preamble or filler (e.g. "Here is...", "Sure!"). Output the slide deck content directly.
4. For each content slide (Slides 2-6), structure with rich executive visual hierarchy:
   - "## Slide <N>: <Emoji> <Punchy Slide Title (3-4 words max)>"
   - "### <Crisp Subtitle / Strategic Takeaway>"
   - Format slide content using clean, professional formats:
     * Qualitative / Strategic Slides (Vision, Market Drivers, Operational Pillars, Roadmap): 3 to 4 high-impact bullet points with bold lead-ins (• **Strategic Pillar**: Detailed description).
     * Quantitative / Comparative Slides (Financials, KPIs, Decarbonization, Operational Metrics): When presenting dense metrics, performance targets, or before/after comparisons, use a clean Markdown table (| Dimension / Metric | Baseline / Prior | Target / Strategic Impact |). Follow the table with 1 to 2 concise takeaway bullets. (Limit tables to at most 1 or 2 slides in the deck).
     * Optional Visual Breakdown Slide (Only if document has market share or volume distribution): A clean high-resolution chart JSON block (\`\`\`json { "chartType": "doughnut", "title": "<Short Title (2-5 words max)>", "data": [...] } \`\`\`) accompanied by 2 narrative bullets.
5. CRITICAL CHART TITLE RULE: Any chart title MUST be short and punchy (2 to 5 words max).
6. Do NOT output internal design metadata like "Visual Concept:", "Color:", or font sizes.
7. Do NOT output pseudocode visual labels like "Metric Grid" or "Comparison Card" in plain text. Format content cleanly as narrative bullets, structured tables, or chart JSON.
7. Separate each slide cleanly with "---".`;

      await executeGeminiWorkflow(prompt, displayUserBubble, attachments);
    } catch (err) {
      console.error("Error reading documents for Transform Doc to Deck:", err);
      appendBubble("Error reading documents: " + err.message, "system");
    } finally {
      event.target.value = "";
      if (loadingText) loadingText.style.display = "none";
    }
  });
}

async function callGeminiProxy(customPrompt = null, ignoreSelection = false, customDisplayBubble = null) {
  if (isGenerating || isProcessingInDocCommand) return;

  const promptInput = document.getElementById("promptText");
  const userText = customPrompt || (promptInput ? promptInput.value.trim() : "");

  // Intercept Slash Commands
  const trimmed = (userText || "").trim();
  const lower = trimmed.toLowerCase();

  if (hostAdapter && hostAdapter.name === "Outlook") {
    const resetInput = () => {
      if (promptInput) {
        promptInput.value = "";
        promptInput.style.height = "auto";
      }
      hideSlashPopup();
      updateSendButtonState();
    };
    if (lower === "/summarize" || lower === "/summarize-doc") {
      resetInput();
      await runOutlookChatAction("summarize");
      return;
    }
    if (lower === "/reply" || lower.startsWith("/reply ")) {
      const customIntent = lower.startsWith("/reply ") ? trimmed.slice(7).trim() : "";
      resetInput();
      await runOutlookChatAction("reply", customIntent);
      return;
    }
    if (lower === "/coach") {
      resetInput();
      await runOutlookChatAction("coach");
      return;
    }
    if (lower === "/actions" || lower === "/identify-action" || lower === "/identify-actions") {
      resetInput();
      await runOutlookChatAction("actions");
      return;
    }
  }

  if (lower === "/summarize-doc" || lower.startsWith("/summarize-doc ") || lower === "/summarize") {
    if (promptInput) {
      promptInput.value = "";
      promptInput.style.height = "auto";
    }
    hideSlashPopup();
    executeSummarizeDoc(true);
    return;
  }
  if (lower === "/identify-risks" || lower.startsWith("/identify-risks ") || lower === "/risks") {
    if (promptInput) {
      promptInput.value = "";
      promptInput.style.height = "auto";
    }
    hideSlashPopup();
    executeIdentifyRisks(true);
    return;
  }
  if (lower === "/identify-action" || lower.startsWith("/identify-action ") || lower === "/identify-actions" || lower.startsWith("/identify-actions ")) {
    if (promptInput) {
      promptInput.value = "";
      promptInput.style.height = "auto";
    }
    hideSlashPopup();
    executeIdentifyAction(true);
    return;
  }
  if (lower === "/exec-card" || lower.startsWith("/exec-card ") || lower === "/exec") {
    if (promptInput) {
      promptInput.value = "";
      promptInput.style.height = "auto";
    }
    hideSlashPopup();
    executeExecCard(true);
    return;
  }

  // Intercept generic Image or Chart requests in PowerPoint to prompt user interactively
  if (userText && hostAdapter && hostAdapter.name === "PowerPoint") {
    const trimmedLower = userText.trim().toLowerCase();
    const isGenericImage = /^(?:please\s+)?(?:generate|create|make|insert|add)\s+(?:an?\s+)?image\s*\.?$/i.test(trimmedLower);
    const isGenericChart = /^(?:please\s+)?(?:generate|create|make|insert|add)\s+(?:an?\s+)?(?:chart|graph|plot)\s*\.?$/i.test(trimmedLower);

    if (isGenericImage) {
      if (promptInput) promptInput.value = "";
      await handleCreateImageClick();
      return;
    }

    if (isGenericChart) {
      if (promptInput) promptInput.value = "";
      await handleCreateChartClick();
      return;
    }
  }

  isGenerating = true;
  const runButton = document.getElementById("run");
  if (runButton) {
    runButton.disabled = true;
    runButton.classList.remove("active");
  }

  try {
    let selectedText = "";
    if (!ignoreSelection && hostAdapter) {
      selectedText = (hostAdapter.name === "Excel")
        ? (await hostAdapter.getSelectedText())
        : (currentSelectedText || (await hostAdapter.getSelectedText()));
      if (!isSubstantiveText(selectedText)) {
        selectedText = "";
      }
    }

    if (!userText && !selectedText && !attachedFileContent) {
      appendBubble(`Please type a prompt or select content in ${hostAdapter ? hostAdapter.name : 'Office'} first.`, "system");
      return;
    }

    let fullPrompt = "";
    let displayUserBubble = "";
    let enableGrounding = true;

    const hostLabel = hostAdapter ? hostAdapter.name : "Office";
    const isPPT = hostLabel === "PowerPoint";
    const isDocumentReferenced = isDocumentSpecificQuery(userText);
    const isSelectionExplicit = isSelectionSpecificQuery(userText);
    const isFollowUp = !!currentSessionId && !isDocumentReferenced && !isSelectionExplicit;

    const intentGroundingPolicy = `Grounding & Intent Policy: First evaluate the user's intent from their query:
1. Conversational / Meta / Short Test Input: If the query is a greeting, pleasantry, capability question, or a bare word/phrase without a specific question or document task (e.g., "${hostLabel.toLowerCase()}", "hi", "test"), respond with a brief, friendly 1-2 sentence reply explaining how you can assist in Microsoft ${hostLabel}. Do NOT summarize any draft and do NOT search, cite, or attempt to open datastore files.
2. Active Document Task: If the user is asking to summarize, rewrite, edit, format, proofread, or analyze the active ${hostLabel} content or attachment, answer strictly using the provided context.
3. Enterprise Knowledge Lookup: Only when the user's query seeks specific company facts, metrics, policies, store directives, SLAs, or external enterprise records, synthesize across the Active Working Draft (if provided) and the connected Enterprise Datastores with inline citations. Never attempt to open or read datastore search results as files from local disk.`;

    if (isPPT) {
      if (selectedText && userText) {
        fullPrompt = `Selected Slide Context:\n"${selectedText}"\n\nUser Instruction: ${userText}`;
        displayUserBubble = customDisplayBubble || `📌 Context: "${selectedText.substring(0, 70)}${selectedText.length > 70 ? '...' : ''}"\n\n${userText}`;
      } else if (selectedText && !userText) {
        fullPrompt = `Please analyze, summarize, or explain the following selected text:\n"${selectedText}"`;
        displayUserBubble = `📌 Selected Text:\n"${selectedText.substring(0, 90)}${selectedText.length > 90 ? '...' : ''}"`;
        enableGrounding = false;
      } else {
        fullPrompt = userText;
        displayUserBubble = customDisplayBubble || userText;
      }
    } else if (!userText && selectedText) {
      fullPrompt = `Please analyze, summarize, or explain the following selected text:\n"${selectedText}"\n\nFormatting: Provide a direct, concise response in natural prose or bullet points. Do NOT generate a table unless explicitly requested.`;
      displayUserBubble = `📌 Selected Text:\n"${selectedText.substring(0, 90)}${selectedText.length > 90 ? '...' : ''}"`;
      enableGrounding = false;
    } else if (userText && (isSelectionExplicit || (selectedText && selectedText.length > 10))) {
      fullPrompt = `Selected Document Context:\n"${selectedText}"\n\nUser Instruction: ${userText}\n\n${intentGroundingPolicy}\nFormatting: Provide a direct, concise response in natural prose or bullet points. Do NOT generate a table unless explicitly requested.`;
      displayUserBubble = customDisplayBubble || `📌 Selection Context: "${selectedText.substring(0, 70)}${selectedText.length > 70 ? '...' : ''}"\n\n${userText}`;
      enableGrounding = true;
    } else if (isFollowUp) {
      fullPrompt = `User Follow-Up: ${userText}\n\n${intentGroundingPolicy}\nFormatting: Provide a direct, concise response in natural prose or bullet points. Do NOT generate a table unless explicitly requested.`;
      displayUserBubble = customDisplayBubble || userText;
      enableGrounding = true;
    } else {
      const fullDocText = hostAdapter ? await hostAdapter.getFullDocumentText() : "";
      if (fullDocText && fullDocText.trim().length > 20) {
        fullPrompt = `Active Working Draft:\n"""\n${fullDocText.substring(0, 100000)}\n"""\n\nUser Query: ${userText}\n\n${intentGroundingPolicy}\nFormatting: Provide a direct, concise response in natural prose or bullet points. Do NOT generate a table unless explicitly requested.`;
      } else {
        fullPrompt = `User Query: ${userText}\n\n${intentGroundingPolicy}\nFormatting: Provide a direct, concise response in natural prose or bullet points. Do NOT generate a table unless explicitly requested.`;
      }
      displayUserBubble = customDisplayBubble || userText;
      enableGrounding = true;
    }

    if (promptInput) {
      promptInput.value = "";
      promptInput.style.height = "auto";
    }
    updateSendButtonState();

    if (typeof window !== "undefined") {
      window.__lastUserPrompt = userText || fullPrompt || "";
      window.__isSummarizeSlidesAction = /^\s*(?:please\s+)?summarize\s+(?:the\s+)?(?:slides?|deck|presentation)\b/i.test(userText);
      window.__forceSingleSlide = false;
    }

    if (isPPT) {
      const { enhancePromptForPowerPoint } = await import('../adapters/ppt/promptEnhancer.js');
      fullPrompt = enhancePromptForPowerPoint(fullPrompt, { userInstruction: userText });
    }

    // If a document was uploaded via +, pass it as a structured Base64 attachment
    let requestAttachments = null;
    if (attachedFileContent && attachedFileName) {
      requestAttachments = [{
        fileName: attachedFileName,
        mimeType: attachedFileMimeType || "application/octet-stream",
        fileContents: attachedFileContent
      }];
      displayUserBubble = displayUserBubble
        ? `📎 Attached: ${attachedFileName}\n${displayUserBubble}`
        : `📎 Attached: ${attachedFileName}`;
      if (!fullPrompt || !fullPrompt.trim()) {
        fullPrompt = `Please analyze the attached document (${attachedFileName}).`;
      }

      // Clear attachment state
      attachedFileContent = null;
      attachedFileName = null;
      attachedFileMimeType = null;
      const fileUploadInput = document.getElementById("fileUploadInput");
      const attachmentContainer = document.getElementById("attachmentContainer");
      if (fileUploadInput) fileUploadInput.value = "";
      if (attachmentContainer) attachmentContainer.style.display = "none";
      updateSendButtonState();
    }

    await executeGeminiWorkflow(fullPrompt, displayUserBubble, requestAttachments, { enableGrounding });
  } finally {
    isGenerating = false;
    if (runButton) runButton.disabled = false;
    updateSendButtonState();
  }
}

async function executeGeminiWorkflow(fullPrompt, displayUserBubble, attachmentsOrGrounding = null, options = {}) {
  const pinnedToOneSlide = typeof window !== "undefined" && !!window.__forceSingleSlide;
  const enableGrounding = typeof attachmentsOrGrounding === "boolean"
    ? attachmentsOrGrounding
    : (options.enableGrounding !== undefined ? options.enableGrounding : true);
  const attachments = typeof attachmentsOrGrounding === "boolean" ? null : attachmentsOrGrounding;

  isGenerating = true;
  const hostName = hostAdapter ? hostAdapter.name : 'Word';
  const runButton = document.getElementById("run");
  const loadingText = document.getElementById("loading");
  const historyDiv = document.getElementById("chatHistory");

  if (runButton) {
    runButton.disabled = true;
    runButton.classList.remove("active");
  }
  if (loadingText) {
    loadingText.innerText = "⚡ Gemini is thinking...";
    loadingText.style.display = "flex";
  }

  try {
    appendBubble(displayUserBubble, "user");

    // Detect if this is an explicit image creation intent
    const isImageIntent = Boolean(
      (options && options.isolateSession) ||
      (displayUserBubble && /Generate image|Create image|Insert image/i.test(displayUserBubble)) ||
      (fullPrompt && /\b(create|generate|make|draw|show|produce|render|provide|insert|add)\s+(an?\s+)?(image|picture|photo|illustration|graphic|visual|artwork)\b/i.test(fullPrompt)) ||
      (fullPrompt && /\b(image|picture|photo|illustration|graphic|visual)\s+(of|for|showing|depicting)\b/i.test(fullPrompt))
    );

    if (typeof window !== "undefined") {
      window.__lastUserPrompt = fullPrompt || displayUserBubble || "";
      if (isImageIntent) {
        window.__isSummarizeSlidesAction = false;
      }
    }

    // If image generation, isolate from prior conversation session to prevent JSON formatting and previous visual contagion
    const sessionIdToUse = isImageIntent ? null : currentSessionId;
    const historyToUse = isImageIntent ? [] : chatHistoryState;

    let data = await askGeminiEnterprise(fullPrompt, historyToUse, sessionIdToUse, enableGrounding, attachments);

    // Fallback Safety Net: Check if model emitted a raw visual_request JSON block instead of calling Imagen
    let aiResponse = data.result || "No content returned.";
    if (aiResponse.includes('"visual_request"')) {
      try {
        const jsonMatch = aiResponse.match(/\{[\s\S]*?"visual_request"[\s\S]*?\}/);
        if (jsonMatch) {
          const parsed = JSON.parse(jsonMatch[0]);
          const extractedPrompt = parsed?.visual_request?.prompt;
          if (extractedPrompt && typeof extractedPrompt === "string") {
            console.log("[ImageFallback] Detected visual_request schema. Automatically invoking image generator with extracted prompt:", extractedPrompt.slice(0, 80));
            if (loadingText) {
              loadingText.innerText = "🎨 Generating image from visual specification...";
              loadingText.style.display = "flex";
            }
            const fallbackPrompt = `Create a high-quality, professional image visual illustration of: ${extractedPrompt}`;
            const fallbackData = await askGeminiEnterprise(fallbackPrompt, [], null, true);
            if (fallbackData && (fallbackData.result || (fallbackData.images && fallbackData.images.length > 0))) {
              data = fallbackData;
              aiResponse = fallbackData.result || aiResponse;
            }
          }
        }
      } catch (fbErr) {
        console.warn("[ImageFallback] Error handling visual_request fallback:", fbErr);
      }
    }

    if (!options.isolateSession) {
      currentSessionId = data.sessionId || currentSessionId || `local-${Date.now()}`;
      SessionManager.setActiveSessionId(currentSessionId, hostName);
    }
    if (Array.isArray(data.history) && !options.isolateSession) {
      chatHistoryState = data.history;
    }

    // Persist user prompt in local session tagged with active host
    if (currentSessionId && !options.isolateSession) {
      SessionManager.addMessage(currentSessionId, { role: 'user', text: displayUserBubble }, hostName);
    }

    // Ensure citations have excerpt text by checking groundingMetadata
    let enrichedCitations = Array.isArray(data.citations) ? [...data.citations] : [];
    if (data.groundingMetadata?.groundingChunks?.length) {
      const chunks = data.groundingMetadata.groundingChunks;
      if (enrichedCitations.length === 0) {
        chunks.forEach((chunk, i) => {
          const ctx = chunk.retrievedContext;
          if (ctx) {
            enrichedCitations.push({
              index: i + 1,
              title: ctx.title || (ctx.uri ? ctx.uri.split('/').pop() : `Enterprise Document ${i + 1}`),
              uri: ctx.uri || '',
              cleanUri: ctx.uri ? ctx.uri.replace('gs://', '') : '',
              excerpt: (ctx.text || '').trim()
            });
          }
        });
      } else {
        enrichedCitations.forEach((cit, i) => {
          if (!cit.excerpt) {
            const match = chunks.find(c => {
              const u = c.retrievedContext?.uri;
              const t = c.retrievedContext?.title;
              return (cit.uri && u === cit.uri) || (cit.title && t === cit.title);
            });
            if (match && match.retrievedContext?.text) {
              cit.excerpt = match.retrievedContext.text.trim();
            } else if (chunks[i]?.retrievedContext?.text) {
              cit.excerpt = chunks[i].retrievedContext.text.trim();
            }
          }
        });
      }
    }

    let cleanAiResponse = (aiResponse || "No content returned.")
      .replace(/---\s*>\s*\[!NOTE\]\s*\*\*Verified(?: Gemini)? Enterprise(?: Grounded)? Sources:\*\*[\s\S]*$/i, '')
      .replace(/>\s*\[!NOTE\]\s*\*\*Verified(?: Gemini)? Enterprise(?: Grounded)? Sources:\*\*[\s\S]*$/i, '')
      .trim();

    if (options && typeof options.transformResponse === "function") {
      try {
        cleanAiResponse = options.transformResponse(cleanAiResponse) || cleanAiResponse;
      } catch (transformErr) {
        console.warn("[Refine] Could not check the refined deck against the original:", transformErr);
      }
    }

    const isSummarizeFlow = Boolean(
      (typeof window !== "undefined" && window.__isSummarizeSlidesAction) ||
      (displayUserBubble && displayUserBubble.includes("[Summarize Slides]"))
    );

    appendAssistantBubble(cleanAiResponse, { ...data, citations: enrichedCitations }, displayUserBubble || fullPrompt, {
      isSummarizeSlides: isSummarizeFlow,
      forceSingleSlide: pinnedToOneSlide,
      outlookAction: options?.outlookAction || null
    });

    // Persist assistant response in local session tagged with active host
    if (currentSessionId && !options.isolateSession) {
      SessionManager.addMessage(currentSessionId, { role: 'assistant', text: cleanAiResponse, citations: enrichedCitations }, hostName);
      if (data.displayName) {
        SessionManager.setSessionTitle(currentSessionId, data.displayName);
      }
    }

  } catch (error) {
    console.error("Gemini execution error:", error);
    const friendlyError = formatErrorMessage(error);
    appendBubble(friendlyError, "error");
    if (error && error.message && error.message.includes("verifying session")) {
      currentSessionId = null;
      SessionManager.clearActiveSession(hostName);
    } else if (currentSessionId) {
      SessionManager.addMessage(currentSessionId, { role: 'error', text: friendlyError }, hostName);
    }
  } finally {
    if (typeof window !== "undefined") window.__forceSingleSlide = false;
    isGenerating = false;
    if (runButton) runButton.disabled = false;
    updateSendButtonState();
    if (loadingText) loadingText.style.display = "none";
    if (historyDiv) historyDiv.scrollTop = historyDiv.scrollHeight;
  }
}

function formatErrorMessage(err) {
  let msg = err && err.message ? err.message : String(err || "An unknown error occurred.");

  // Check for embedded JSON error payload
  try {
    const jsonMatch = msg.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      const parsed = JSON.parse(jsonMatch[0]);
      if (parsed.error === 'invalid_grant' || parsed.error_subtype === 'invalid_rapt' || (parsed.error_description && parsed.error_description.includes('reauth'))) {
        return `⚠️ Google Cloud Authentication Expired\n\nYour Application Default Credentials (ADC) token has expired and requires re-authentication.\n\nPlease run this command in your workstation terminal:\n  gcloud auth application-default login\n\nThen submit your prompt again.`;
      }
      if (parsed.details && typeof parsed.details === 'string') {
        return `⚠️ ${parsed.details}`;
      }
      if (parsed.error_description) {
        return `⚠️ Error: ${parsed.error_description}`;
      }
      if (parsed.error && typeof parsed.error === 'string') {
        return `⚠️ Error: ${parsed.error}`;
      }
    }
  } catch (e) {
    // ignore parse error and proceed
  }

  if (msg.includes("invalid_grant") || msg.includes("invalid_rapt") || msg.includes("reauth related error") || msg.includes("Authentication Expired")) {
    return `⚠️ Google Cloud Authentication Expired\n\nYour Application Default Credentials (ADC) token has expired and requires re-authentication.\n\nPlease run this command in your workstation terminal:\n  gcloud auth application-default login\n\nThen submit your prompt again.`;
  }

  if (msg.includes("Failed to fetch") || msg.includes("NetworkError")) {
    return `⚠️ Cannot connect to the Gemini proxy server (${getActiveProxyUrl()}).\nPlease ensure the proxy is running.`;
  }

  return `⚠️ ${msg}`;
}

const CONTENT_COPY_ICON_SVG = `<svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor" style="vertical-align: -2px; margin-right: 4px;"><path d="M16 1H4c-1.1 0-2 .9-2 2v14h2V3h12V1zm3 4H8c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h11c1.1 0 2-.9 2-2V7c0-1.1-.9-2-2-2zm0 16H8V7h11v14z"/></svg>`;

async function copyTextToClipboard(text) {
  if (!text) return false;
  if (navigator.clipboard && typeof navigator.clipboard.writeText === 'function') {
    try {
      await navigator.clipboard.writeText(text);
      return true;
    } catch (err) {
      console.warn("navigator.clipboard.writeText failed, using fallback:", err);
    }
  }

  // Fallback for Office WebViews / restricted iframes using hidden textarea + execCommand
  try {
    const textArea = document.createElement("textarea");
    textArea.value = text;
    textArea.style.position = "fixed";
    textArea.style.top = "-9999px";
    textArea.style.left = "-9999px";
    textArea.style.opacity = "0";
    textArea.setAttribute("readonly", "");
    document.body.appendChild(textArea);
    textArea.focus();
    textArea.select();
    textArea.setSelectionRange(0, text.length);
    const successful = document.execCommand("copy");
    document.body.removeChild(textArea);
    return successful;
  } catch (err) {
    console.error("execCommand fallback copy failed:", err);
    return false;
  }
}

function appendBubble(text, type) {
  const isOutlook = hostAdapter && hostAdapter.name === "Outlook";
  if (!isOutlook) {
    updateChatDisplayState(true);
  }
  const historyDiv = document.getElementById("chatHistory");
  if (!historyDiv) return;
  const bubble = document.createElement("div");
  bubble.className = `chat-bubble ${type}`;
  bubble.innerText = text;
  historyDiv.appendChild(bubble);

  if (isOutlook) {
    const hasConversation = Boolean(historyDiv.querySelector(".chat-bubble.user, .chat-bubble.assistant"));
    if (hasConversation || type === "user" || type === "assistant") {
      updateChatDisplayState(true);
    } else {
      // System/status bubble before any conversation starts in Outlook: keep starter chips (#docToolsBar),
      // hero (#centerHero), and prompt suggestions (#suggestionsContainer) visible.
      historyDiv.style.display = "flex";
      historyDiv.style.flex = "0 0 auto";
      historyDiv.style.order = "-1";
      historyDiv.style.paddingBottom = "0";
    }
  }
  historyDiv.scrollTop = historyDiv.scrollHeight;
}

function appendInteractiveBubble(htmlContent, onMountCallback = null) {
  updateChatDisplayState(true);
  const historyDiv = document.getElementById("chatHistory");
  if (!historyDiv) return null;
  const bubble = document.createElement("div");
  bubble.className = "chat-bubble assistant";
  bubble.innerHTML = htmlContent;
  historyDiv.appendChild(bubble);
  if (typeof onMountCallback === "function") {
    onMountCallback(bubble);
  }
  historyDiv.scrollTop = historyDiv.scrollHeight;
  return bubble;
}

/**
 * Opens an expanded high-resolution review modal for any generated visual image.
 * Provides explicit actions to review and decide to insert directly onto the current slide
 * or as a brand-new presentation slide.
 */
export function openImageZoomModal(imgSrc, altText = "Generated Visual") {
  if (!imgSrc) return;

  let modal = document.getElementById("imageZoomModal");
  if (!modal) {
    modal = document.createElement("div");
    modal.id = "imageZoomModal";
    modal.className = "image-zoom-overlay";
    document.body.appendChild(modal);
  }

  const isPPT = hostAdapter && hostAdapter.name === "PowerPoint";

  modal.innerHTML = `
    <div class="image-zoom-header">
      <div style="display:flex; align-items:center; gap:8px;">
        <span>🔍 Image Review</span>
        <div class="zoom-level-controls" style="display:inline-flex; gap:4px; background:rgba(255,255,255,0.15); border-radius:9999px; padding:3px;">
          <button type="button" class="zoom-ctrl-btn" id="zoomCtrlFit" title="Fit to window" style="background:none; border:none; color:#cbd5e1; font-size:11px; padding:3px 9px; border-radius:9999px; cursor:pointer;">1x Fit</button>
          <button type="button" class="zoom-ctrl-btn active" id="zoomCtrl2x" title="2x Enlarge" style="background:#0b57d0; border:none; color:#ffffff; font-weight:600; font-size:11px; padding:3px 9px; border-radius:9999px; cursor:pointer;">2x Zoom</button>
          <button type="button" class="zoom-ctrl-btn" id="zoomCtrl3x" title="3x Extra Large" style="background:none; border:none; color:#cbd5e1; font-size:11px; padding:3px 9px; border-radius:9999px; cursor:pointer;">3x Max</button>
        </div>
      </div>
      <button class="image-zoom-close" id="closeZoomModalBtn" title="Close preview (Esc)">✕</button>
    </div>
    <div class="image-zoom-body zoom-2x" id="zoomModalBody">
      <img src="${escapeHtml(imgSrc)}" alt="${escapeHtml(altText)}" />
    </div>
    <div class="image-zoom-footer">
      ${isPPT ? `
        <button class="image-zoom-btn" id="zoomInsertNewBtn">+ Insert as New Slide</button>
        <button class="image-zoom-btn secondary" id="zoomInsertCurrentBtn">Insert on Current Slide</button>
      ` : `
        <button class="image-zoom-btn" id="zoomInsertDocBtn">+ Insert Image</button>
      `}
      <button class="image-zoom-btn secondary" id="zoomCancelBtn">Close</button>
    </div>
  `;

  modal.style.display = "flex";

  const bodyEl = document.getElementById("zoomModalBody");
  const btnFit = document.getElementById("zoomCtrlFit");
  const btn2x = document.getElementById("zoomCtrl2x");
  const btn3x = document.getElementById("zoomCtrl3x");

  const setZoom = (level) => {
    if (!bodyEl) return;
    bodyEl.classList.remove("zoom-fit", "zoom-2x", "zoom-3x");
    bodyEl.classList.add(`zoom-${level}`);
    [btnFit, btn2x, btn3x].forEach(b => {
      if (b) {
        b.style.background = "none";
        b.style.color = "#cbd5e1";
        b.style.fontWeight = "normal";
      }
    });
    const activeBtn = level === "fit" ? btnFit : (level === "2x" ? btn2x : btn3x);
    if (activeBtn) {
      activeBtn.style.background = "#0b57d0";
      activeBtn.style.color = "#ffffff";
      activeBtn.style.fontWeight = "600";
    }
  };

  if (btnFit) btnFit.onclick = (e) => { e.stopPropagation(); setZoom("fit"); };
  if (btn2x) btn2x.onclick = (e) => { e.stopPropagation(); setZoom("2x"); };
  if (btn3x) btn3x.onclick = (e) => { e.stopPropagation(); setZoom("3x"); };

  // Smooth drag-to-pan in enlarged review mode
  let isDown = false;
  let startX = 0;
  let startY = 0;
  let scrollLeft = 0;
  let scrollTop = 0;
  if (bodyEl) {
    bodyEl.onmousedown = (e) => {
      if (e.target.tagName === "BUTTON") return;
      isDown = true;
      startX = e.pageX - bodyEl.offsetLeft;
      startY = e.pageY - bodyEl.offsetTop;
      scrollLeft = bodyEl.scrollLeft;
      scrollTop = bodyEl.scrollTop;
    };
    bodyEl.onmouseleave = () => { isDown = false; };
    bodyEl.onmouseup = () => { isDown = false; };
    bodyEl.onmousemove = (e) => {
      if (!isDown) return;
      e.preventDefault();
      const x = e.pageX - bodyEl.offsetLeft;
      const y = e.pageY - bodyEl.offsetTop;
      const walkX = (x - startX) * 1.5;
      const walkY = (y - startY) * 1.5;
      bodyEl.scrollLeft = scrollLeft - walkX;
      bodyEl.scrollTop = scrollTop - walkY;
    };
  }

  const closeModal = () => {
    modal.style.display = "none";
  };

  const closeBtn = document.getElementById("closeZoomModalBtn");
  if (closeBtn) closeBtn.onclick = closeModal;

  const cancelBtn = document.getElementById("zoomCancelBtn");
  if (cancelBtn) cancelBtn.onclick = closeModal;

  modal.onclick = (e) => {
    if (e.target === modal) {
      closeModal();
    }
  };

  const handleKeydown = (e) => {
    if (e.key === "Escape") {
      closeModal();
      window.removeEventListener("keydown", handleKeydown);
    }
  };
  window.addEventListener("keydown", handleKeydown);

  if (isPPT) {
    const insertCurrentBtn = document.getElementById("zoomInsertCurrentBtn");
    if (insertCurrentBtn) {
      insertCurrentBtn.onclick = async () => {
        insertCurrentBtn.innerText = "⏳ Inserting...";
        insertCurrentBtn.disabled = true;
        try {
          await hostAdapter.insertContent(`<img src="${imgSrc}" />`, `![Visual](${imgSrc})`, {
            mode: "insert_current",
            imageOnly: true
          });
          insertCurrentBtn.innerText = "✅ Inserted!";
          insertCurrentBtn.classList.add("success");
          setTimeout(() => { closeModal(); }, 1000);
        } catch (err) {
          console.error("Zoom insert error:", err);
          insertCurrentBtn.innerText = "❌ Failed to insert";
          setTimeout(() => {
            insertCurrentBtn.innerText = "📌 Insert on Current Slide";
            insertCurrentBtn.disabled = false;
          }, 2000);
        }
      };
    }

    const insertNewBtn = document.getElementById("zoomInsertNewBtn");
    if (insertNewBtn) {
      insertNewBtn.onclick = async () => {
        insertNewBtn.innerText = "⏳ Creating slide...";
        insertNewBtn.disabled = true;
        try {
          await hostAdapter.insertContent(`<img src="${imgSrc}" />`, `![Visual](${imgSrc})`, {
            mode: "insert",
            imageOnly: true
          });
          insertNewBtn.innerText = "✅ Created!";
          insertNewBtn.classList.add("success");
          setTimeout(() => { closeModal(); }, 1000);
        } catch (err) {
          console.error("Zoom insert as new slide error:", err);
          insertNewBtn.innerText = "❌ Failed";
          setTimeout(() => {
            insertNewBtn.innerText = "➕ Insert as New Slide";
            insertNewBtn.disabled = false;
          }, 2000);
        }
      };
    }
  } else {
    const insertDocBtn = document.getElementById("zoomInsertDocBtn");
    if (insertDocBtn) {
      insertDocBtn.onclick = async () => {
        insertDocBtn.innerText = "⏳ Inserting...";
        insertDocBtn.disabled = true;
        try {
          await performDocumentInsertion(`<img src="${imgSrc}" />`, `![Generated Image](${imgSrc})`, "insert_cursor");
          insertDocBtn.innerText = "✅ Inserted!";
          setTimeout(() => { closeModal(); }, 1000);
        } catch (err) {
          console.error("Zoom insert doc error:", err);
          insertDocBtn.innerText = "❌ Failed";
          setTimeout(() => {
            insertDocBtn.innerText = "📌 Insert Image";
            insertDocBtn.disabled = false;
          }, 2000);
        }
      };
    }
  }
}

function appendAssistantBubble(text, apiDataOrCitations = null, originalPrompt = "", bubbleOptions = {}) {
  updateChatDisplayState(true);
  const historyDiv = document.getElementById("chatHistory");
  if (!historyDiv) return;
  const bubble = document.createElement("div");
  bubble.className = "chat-bubble assistant";

  const apiData = Array.isArray(apiDataOrCitations) ? { citations: apiDataOrCitations } : (apiDataOrCitations || {});
  const citations = Array.isArray(apiDataOrCitations) ? apiDataOrCitations : (Array.isArray(apiData?.citations) ? apiData.citations : []);

  const hostName = hostAdapter ? hostAdapter.name : "Word";
  const isPPT = hostName === "PowerPoint";
  const isExcel = hostName === "Excel";
  const isOutlook = hostName === "Outlook";
  const isPowerPointHost = isPPT;

  const isSummarizeAction = Boolean(
    bubbleOptions?.isSummarizeSlides ||
    (typeof window !== "undefined" && window.__isSummarizeSlidesAction) ||
    (originalPrompt && (
      originalPrompt.toLowerCase().includes("[summarize slides]") ||
      (isPowerPointHost && /\bsummariz(?:e|ing)\s+(?:the\s+|all\s+|these\s+)?(?:slides?|deck|presentation)\b/i.test(originalPrompt))
    ))
  );
  bubble.dataset.isSummarizeSlides = isSummarizeAction ? "true" : "false";

  const forcesSingleSlide = bubbleOptions?.forceSingleSlide === true;
  bubble.dataset.forceSingleSlide = forcesSingleSlide ? "true" : "false";

  // Check if response contains structured chart JSON
  const hasStructuredChart = /```(?:json|chart|pie|bar|column|line|doughnut|donut)?\s*\{[\s\S]*?(?:chartType|chart_type|pie|bar|column|line|doughnut)[\s\S]*?```/i.test(text) ||
                             /\{\s*"(?:chartType|chart_type)"\s*:\s*"(?:pie|bar|line|doughnut|donut|column)"/i.test(text);

  // Check if user requested distinct non-chart image content (e.g. photo, illustration, logo)
  const promptLower = (originalPrompt || "").toLowerCase();
  const hasDistinctNonChartImageIntent = /(?:image|picture|photo|photograph|portrait|illustration|drawing|cat|dog|logo|camera|scenery)/i.test(promptLower);

  let fullText = normalizeEscapedNewlines(text || "");
  if (apiData && Array.isArray(apiData.images) && apiData.images.length > 0) {
    if (!hasStructuredChart || hasDistinctNonChartImageIntent) {
      for (const imgUrl of apiData.images) {
        if (imgUrl && !fullText.includes(imgUrl)) {
          fullText += `\n\n![Generated Image](${imgUrl})\n\n`;
        }
      }
    }
  }

  // Parse markdown into modern executive HTML
  const formattedHtml = parseMarkdown(fullText, { hasDistinctNonChartImageIntent });

  const textDiv = document.createElement("div");
  textDiv.innerHTML = isOutlook ? wrapOutlookTables(formattedHtml) : formattedHtml;
  bubble.appendChild(textDiv);

  // DOM SWEEP: Ensure EVERY <img> element in the assistant response has zoom controls and click-to-zoom
  const renderedImages = textDiv.querySelectorAll("img");
  renderedImages.forEach((img) => {
    img.classList.add("office-preview-img");
    img.style.cursor = "pointer";
    img.title = "Click to zoom into high-resolution visual";

    // Direct click on the image itself ALWAYS opens zoom modal
    img.onclick = (e) => {
      e.stopPropagation();
      openImageZoomModal(img.src, img.alt);
    };

    // Set data attributes and wire direct click handlers on any zoom buttons inside the card
    const card = img.closest(".office-visual-image-card, .rendered-chart-container, .chat-bubble");
    if (card) {
      card.querySelectorAll(".img-action-btn-zoom").forEach(btn => {
        btn.setAttribute("data-img-src", img.src);
        btn.setAttribute("data-img-alt", img.alt || "Generated Visual");
        btn.onclick = (e) => {
          e.stopPropagation();
          openImageZoomModal(img.src, img.alt);
        };
      });
    }

    // If there is no bottom "Zoom & Review" action button in the card or immediate container:
    if (card && !card.querySelector(".img-action-btn-zoom")) {
      const zoomReviewBar = document.createElement("div");
      zoomReviewBar.style.cssText = "margin:8px 0; display:flex; justify-content:center; gap:6px;";
      const zoomReviewBtn = document.createElement("button");
      zoomReviewBtn.type = "button";
      zoomReviewBtn.className = "img-action-btn-zoom";
      zoomReviewBtn.setAttribute("data-img-src", img.src);
      zoomReviewBtn.setAttribute("data-img-alt", img.alt || "Generated Visual");
      zoomReviewBtn.innerHTML = "🔍 Zoom & Review";
      zoomReviewBtn.onclick = (e) => {
        e.stopPropagation();
        openImageZoomModal(img.src, img.alt);
      };
      zoomReviewBar.appendChild(zoomReviewBtn);

      if (img.parentElement && img.parentElement.parentNode) {
        img.parentElement.parentNode.insertBefore(zoomReviewBar, img.parentElement.nextSibling);
      } else {
        card.appendChild(zoomReviewBar);
      }
    }
  });

  // Render Grounded Enterprise Citations if available
  if (Array.isArray(citations) && citations.length > 0) {
    const citationContainer = document.createElement("div");
    citationContainer.className = "citation-container";

    const citationHeader = document.createElement("div");
    citationHeader.className = "citation-header";
    citationHeader.innerHTML = `Verified Sources (${citations.length})`;
    citationContainer.appendChild(citationHeader);

    const citationList = document.createElement("div");
    citationList.className = "citation-list";

    citations.forEach((cit, idx) => {
      const citIdx = cit.index || (idx + 1);
      const card = document.createElement("div");
      card.className = "citation-card";
      card.setAttribute("data-index", citIdx);
      card.id = `citation-card-${citIdx}-${Date.now()}`;

      const cardHeader = document.createElement("div");
      cardHeader.className = "citation-card-header";
      const citTitle = cit.title || 'Enterprise Document';
      cardHeader.innerHTML = `
        <div class="citation-card-title">
          <span class="citation-card-index">${citIdx}</span>
          <span title="${escapeHtml(citTitle)}">📄 ${escapeHtml(citTitle)}</span>
        </div>
        <span class="citation-toggle-icon">▼</span>
      `;

      const cardBody = document.createElement("div");
      cardBody.className = "citation-card-body";

      const excerptText = cit.excerpt ? cit.excerpt : 'Document referenced and verified via Vertex AI Enterprise Search.';
      const cleanUriText = cit.cleanUri || (cit.uri ? cit.uri.replace('gs://', '') : '');

      cardBody.innerHTML = `
        <div class="citation-excerpt">&ldquo;${escapeHtml(excerptText)}&rdquo;</div>
        <div class="citation-meta">
          <span>${cleanUriText ? `Location: <code>${escapeHtml(cleanUriText)}</code>` : 'Enterprise Knowledge Base'}</span>
          <button class="icon-btn" title="Copy excerpt" style="width:26px; height:26px;">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor">
              <path d="M16 1H4c-1.1 0-2 .9-2 2v14h2V3h12V1zm3 4H8c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h11c1.1 0 2-.9 2-2V7c0-1.1-.9-2-2-2zm0 16H8V7h11v14z"/>
            </svg>
          </button>
        </div>
      `;

      // Copy excerpt button handler
      const copyBtn = cardBody.querySelector("button");
      if (copyBtn) {
        copyBtn.onclick = async (e) => {
          e.stopPropagation();
          const ok = await copyTextToClipboard(excerptText);
          if (ok) {
            showToast("Copied excerpt to clipboard!");
          } else {
            showToast("Copy failed");
          }
        };
      }

      // Header click toggles expand/collapse
      cardHeader.onclick = () => {
        card.classList.toggle("expanded");
      };

      card.appendChild(cardHeader);
      card.appendChild(cardBody);
      citationList.appendChild(card);
    });

    citationContainer.appendChild(citationList);
    bubble.appendChild(citationContainer);

    // Attach click handlers to inline citations in this bubble
    const inlineBadges = bubble.querySelectorAll(".inline-citation");
    inlineBadges.forEach(badge => {
      badge.onclick = (e) => {
        e.preventDefault();
        const refIdx = badge.getAttribute("data-citation-idx");
        const targetCard = bubble.querySelector(`.citation-card[data-index="${refIdx}"]`);
        if (targetCard) {
          targetCard.classList.add("expanded");
          targetCard.classList.add("highlighted");
          targetCard.scrollIntoView({ behavior: "smooth", block: "nearest" });
          setTimeout(() => {
            targetCard.classList.remove("highlighted");
          }, 1800);
        }
      };
    });
  }

  // Google Docs Style Action Toolbar
  const actionsContainer = document.createElement("div");
  actionsContainer.className = "response-actions-container";

  const primaryActions = document.createElement("div");
  primaryActions.className = "footer-primary-actions";

  const hasActiveSelection = currentSelectedText && currentSelectedText.length > 5;

  // Detect if this presentation response contains multiple slides
  const pptSlideHeaderMatches = isPPT ? (fullText.match(/(?:^|\n)##\s+(?:Slide\s*\d+|[^\n]+)/gi) || []) : [];
  const detectedSlideCount = pptSlideHeaderMatches.length;
  const isMultiSlideDeck = isPPT && detectedSlideCount > 1 && !forcesSingleSlide;

  // Check if this response was generated from an image prompt or contains rendered visuals
  const isImageRequest = hasDistinctNonChartImageIntent ||
                         /(?:generate|create|make|draw|show|render|provide|insert|add)\s+(?:an?\s+)?(?:image|picture|photo|visual|illustration|graphic|infographic)/i.test(promptLower) ||
                         /(?:image|picture|photo|visual|illustration)\s+(?:of|for|showing|depicting)/i.test(promptLower);

  const shouldInsertImageOnly = renderedImages.length > 0 && (isImageRequest || (!fullText.includes("##") && textDiv.innerText.trim().length < 60));
  const isBubbleSummarize = bubble.dataset.isSummarizeSlides === "true";

  // Helper: Retrieve the primary non-chart visual image from renderedImages
  const getTargetVisualImage = () => {
    if (renderedImages.length === 0) return null;
    const nonChartImg = Array.from(renderedImages).reverse().find(img => {
      const isInsideChart = img.closest(".rendered-chart-container, [data-chart-title]");
      const isChartAlt = (img.alt || "").toLowerCase().includes("chart");
      return !isInsideChart && !isChartAlt;
    });
    return nonChartImg || renderedImages[0];
  };

  if (isExcel) {
    try {
      const excelActions = renderExcelResponseActions({
        rawText: fullText,
        formattedHtml,
        textDiv,
        adapter: hostAdapter,
        prompt: originalPrompt,
        userPrompt: originalPrompt,
        hideCopyButton: true
      });
      if (excelActions) {
        actionsContainer.appendChild(excelActions);
      }
    } catch (excelErr) {
      console.warn("Excel response actions rendering error:", excelErr);
    }
  } else if (isPPT) {
    let replaceBtn = null;
    if (!isMultiSlideDeck) {
      replaceBtn = document.createElement("button");
      replaceBtn.className = "pill-btn outlined replace";
      replaceBtn.innerHTML = shouldInsertImageOnly ? `Replace with Image` : `Replace Draft`;
      replaceBtn.title = shouldInsertImageOnly ? "Replace selected shape or content with image" : "Replace existing text or placeholder on current slide";
      replaceBtn.onclick = async () => {
        if (shouldInsertImageOnly) {
          const img = getTargetVisualImage();
          if (img && img.src) {
            await performDocumentInsertion(`<img src="${img.src}" />`, `![Image](${img.src})`, "replace_draft", { imageOnly: true, isSummarizeSlides: false });
          }
        } else {
          await performDocumentInsertion(textDiv.innerHTML, fullText, "replace_draft", {
            isSummarizeSlides: isBubbleSummarize,
            maxSlides: bubble.dataset.forceSingleSlide === "true" ? 1 : 0
          });
        }
      };
    }

    const insertCurrentBtn = document.createElement("button");
    insertCurrentBtn.className = "pill-btn outlined insert-current";
    if (shouldInsertImageOnly) {
      insertCurrentBtn.innerHTML = `Insert Image on Slide`;
      insertCurrentBtn.title = "Insert image directly onto current slide";
    } else if (isMultiSlideDeck) {
      insertCurrentBtn.innerHTML = `Insert Slide 1 on Current`;
      insertCurrentBtn.title = `Inserts only Slide 1 onto active slide (use the primary button to build all ${detectedSlideCount} slides)`;
    } else {
      insertCurrentBtn.innerHTML = `Insert on Current Slide`;
      insertCurrentBtn.title = "Insert generated content or visual directly onto the currently active slide";
    }
    insertCurrentBtn.onclick = async () => {
      if (shouldInsertImageOnly) {
        const img = getTargetVisualImage();
        if (img && img.src) {
          await performDocumentInsertion(`<img src="${img.src}" />`, `![Image](${img.src})`, "insert_current_slide", { imageOnly: true, isSummarizeSlides: false });
        }
      } else {
        await performDocumentInsertion(textDiv.innerHTML, fullText, "insert_current_slide", {
          isSummarizeSlides: isBubbleSummarize,
          maxSlides: bubble.dataset.forceSingleSlide === "true" ? 1 : 0
        });
      }
    };

    const insertBtn = document.createElement("button");
    insertBtn.className = "pill-btn primary insert";
    if (shouldInsertImageOnly) {
      insertBtn.innerHTML = `+ Insert Image as New Slide`;
      insertBtn.title = "Insert image into presentation";
    } else {
      insertBtn.innerHTML = isMultiSlideDeck ? `+ Insert as ${detectedSlideCount} New Slides` : `+ Insert as New Slide`;
      insertBtn.title = isMultiSlideDeck ? `Build and insert all ${detectedSlideCount} slides into the presentation` : "Create new presentation slide at the end of the deck";
    }
    insertBtn.onclick = async () => {
      if (shouldInsertImageOnly) {
        const img = getTargetVisualImage();
        if (img && img.src) {
          await performDocumentInsertion(`<img src="${img.src}" />`, `![Image](${img.src})`, "insert_cursor", { imageOnly: true, isSummarizeSlides: false });
        }
      } else {
        await performDocumentInsertion(textDiv.innerHTML, fullText, "insert_cursor", {
          isSummarizeSlides: isBubbleSummarize,
          maxSlides: bubble.dataset.forceSingleSlide === "true" ? 1 : 0
        });
      }
    };

    primaryActions.appendChild(insertBtn);
    primaryActions.appendChild(insertCurrentBtn);
    if (replaceBtn) primaryActions.appendChild(replaceBtn);
  } else if (isOutlook) {
    const outlookAction = bubbleOptions?.outlookAction || "ask";
    const isCompose = hostAdapter && typeof hostAdapter.isComposeMode === "function" && hostAdapter.isComposeMode();
    const isCoachBubble = outlookAction === "coach" || /OPTION A\s*[:\-–\(\s]+CONCISE EXECUTIVE/i.test(fullText);

    if (isCoachBubble) {
      renderOutlookCoachInBubble(textDiv, fullText);
    }

    if (outlookAction === "summarize") {
      const replySummaryBtn = document.createElement("button");
      replySummaryBtn.className = "pill-btn primary";
      replySummaryBtn.innerHTML = `👥 Reply All with Summary`;
      replySummaryBtn.title = "Open a Reply All draft with this executive summary";
      replySummaryBtn.onclick = () => {
        hostAdapter.replyWithDraft(formattedHtml, true);
      };
      primaryActions.appendChild(replySummaryBtn);

      const replySingleBtn = document.createElement("button");
      replySingleBtn.className = "pill-btn outlined";
      replySingleBtn.innerHTML = `↩️ Reply`;
      replySingleBtn.title = "Open a single Reply draft with this summary";
      replySingleBtn.onclick = () => {
        hostAdapter.replyWithDraft(formattedHtml, false);
      };
      primaryActions.appendChild(replySingleBtn);
    } else if (outlookAction === "actions") {
      const calBtn = document.createElement("button");
      calBtn.className = "pill-btn primary";
      calBtn.innerHTML = `📅 Create Calendar Invite`;
      calBtn.title = "Open an Outlook calendar invitation pre-populated with these action items";
      calBtn.onclick = () => {
        hostAdapter.createCalendarInvite({
          subject: "Follow-up on Action Items & Milestones",
          body: formattedHtml,
          rawText: fullText
        });
      };
      primaryActions.appendChild(calBtn);

      const replyAllBtn = document.createElement("button");
      replyAllBtn.className = "pill-btn outlined";
      replyAllBtn.innerHTML = `👥 Reply All`;
      replyAllBtn.title = "Reply All with these notes";
      replyAllBtn.onclick = () => {
        hostAdapter.replyWithDraft(formattedHtml, true);
      };
      primaryActions.appendChild(replyAllBtn);

      if (isCompose) {
        const insertEmailBtn = document.createElement("button");
        insertEmailBtn.className = "pill-btn outlined";
        insertEmailBtn.innerHTML = `➕ Insert in Email`;
        insertEmailBtn.title = "Insert into active email draft";
        insertEmailBtn.onclick = async () => {
          await hostAdapter.insertContent(formattedHtml, fullText, "smart");
          showToast("Inserted into email draft");
        };
        primaryActions.appendChild(insertEmailBtn);
      }
    } else if (!isCoachBubble) {
      const replyAllBtn = document.createElement("button");
      replyAllBtn.className = "pill-btn primary";
      replyAllBtn.innerHTML = `👥 Open Reply All`;
      replyAllBtn.title = "Open a Reply All draft populated with this response";
      replyAllBtn.onclick = () => {
        hostAdapter.replyWithDraft(formattedHtml, true);
      };
      primaryActions.appendChild(replyAllBtn);

      const replyBtn = document.createElement("button");
      replyBtn.className = "pill-btn outlined";
      replyBtn.innerHTML = `↩️ Open Reply`;
      replyBtn.title = "Open a single Reply draft populated with this response";
      replyBtn.onclick = () => {
        hostAdapter.replyWithDraft(formattedHtml, false);
      };
      primaryActions.appendChild(replyBtn);

      if (isCompose) {
        const insertDraftBtn = document.createElement("button");
        insertDraftBtn.className = "pill-btn outlined";
        insertDraftBtn.innerHTML = `➕ Insert in Draft`;
        insertDraftBtn.title = "Insert into active email compose body";
        insertDraftBtn.onclick = async () => {
          await hostAdapter.insertContent(formattedHtml, fullText, "replace_draft");
          showToast("Inserted into draft");
        };
        primaryActions.appendChild(insertDraftBtn);
      }
    }
  } else if (hasActiveSelection) {
    // Word Selection state: Replace Selection + Insert below
    const replaceBtn = document.createElement("button");
    replaceBtn.className = "pill-btn primary";
    replaceBtn.innerHTML = `Replace Selection`;
    replaceBtn.title = "Replace selected text in document";
    replaceBtn.onclick = async () => {
      await performDocumentInsertion(formattedHtml, fullText, "replace_draft");
    };
    primaryActions.appendChild(replaceBtn);

    const insertBelowBtn = document.createElement("button");
    insertBelowBtn.className = "pill-btn outlined";
    insertBelowBtn.innerHTML = `Insert below`;
    insertBelowBtn.title = "Insert polished text below selection";
    insertBelowBtn.onclick = async () => {
      await performDocumentInsertion(formattedHtml, fullText, "insert_cursor");
    };
    primaryActions.appendChild(insertBelowBtn);
  } else {
    // Word Default state: + Insert
    const insertBtn = document.createElement("button");
    insertBtn.className = "pill-btn primary";
    insertBtn.innerHTML = shouldInsertImageOnly ? '+ Insert Image' : '+ Insert';
    insertBtn.title = "Insert at cursor location";
    insertBtn.onclick = async () => {
      if (shouldInsertImageOnly) {
        const img = getTargetVisualImage();
        if (img && img.src) {
          await performDocumentInsertion(`<img src="${img.src}" />`, `![Image](${img.src})`, "insert_cursor", { imageOnly: true });
          return;
        }
      }
      await performDocumentInsertion(formattedHtml, fullText, "insert_cursor");
    };
    primaryActions.appendChild(insertBtn);
  }

  if (!isExcel) {
    // Refine Dropdown Menu (Google Docs style)
    const refineBtn = document.createElement("button");
    refineBtn.className = "pill-btn outlined";
    refineBtn.innerHTML = `Refine ▾`;
    refineBtn.title = "Refine tone, length, or format";

    const refineMenu = document.createElement("div");
    refineMenu.className = "dropdown-menu";
    refineMenu.innerHTML = `
      <div class="dropdown-item" data-refine="formal">👔 Formal tone</div>
      <div class="dropdown-item" data-refine="casual">☕ Casual tone</div>
      <div class="dropdown-divider"></div>
      <div class="dropdown-item" data-refine="shorten">📉 Shorten</div>
      <div class="dropdown-item" data-refine="elaborate">📈 Elaborate</div>
      <div class="dropdown-divider"></div>
      <div class="dropdown-item" data-refine="bullets">🎯 Bullet points</div>
      <div class="dropdown-item" data-refine="table">📊 Table</div>
      <div class="dropdown-divider"></div>
      <div class="dropdown-item" data-refine="retry">🔄 Rewrite</div>
    `;

    refineBtn.onclick = (e) => {
      e.stopPropagation();
      document.querySelectorAll(".dropdown-menu.show").forEach(m => {
        if (m !== refineMenu) m.classList.remove("show");
      });
      refineMenu.classList.toggle("show");
    };

    const refineOptions = {
      formal: "Rewrite the above response in an authoritative, formal executive tone. Provide exactly ONE single, definitive revision in natural prose without multiple options, bullet points, or conversational preamble.",
      casual: "Rewrite the above response in an approachable, clear, casual conversational tone. Provide exactly ONE single, definitive revision in natural prose without multiple options, bullet points, or conversational preamble.",
      shorten: "Make the above response significantly more concise and punchy for executive reading. Provide exactly ONE single, definitive revision in natural prose without multiple options, bullet points, or conversational preamble.",
      elaborate: "Expand the above draft with more in-depth operational and analytical details. Provide exactly ONE single, definitive revision without multiple options or conversational preamble.",
      bullets: "Convert the key findings of the above response into clean, high-impact bullet points (* or -). Output ONLY the cohesive bulleted list without introductory filler, conversational preamble, or multiple options.",
      table: "Extract and structure the key comparative data, metrics, entities, and categories of the above response into a clean, well-organized markdown table with clear column headers. Output ONLY the table without conversational preamble.",
      retry: "Rewrite and rephrase the above response from scratch with a fresh structure, varied vocabulary, and enhanced clarity. Provide exactly ONE single, definitive revision without multiple options or preamble."
    };

    refineMenu.querySelectorAll(".dropdown-item").forEach(item => {
      item.onclick = (e) => {
        e.stopPropagation();
        refineMenu.classList.remove("show");
        if (isGenerating || isProcessingInDocCommand) return;
        const key = item.getAttribute("data-refine");
        if (refineOptions[key]) {
          if (isPPT) {
            // The deck itself goes into the prompt; see runPowerPointRefine.
            runPowerPointRefine(key, item.innerText.trim(), fullText, {
              isDeck: isMultiSlideDeck,
              slideCount: detectedSlideCount,
              isSummarize: isBubbleSummarize
            });
          } else if (isOutlook) {
            const outlookRefinePrompt = `${refineOptions[key]}\n\nResponse to Revise:\n\"\"\"\n${fullText}\n\"\"\"`;
            executeGeminiWorkflow(outlookRefinePrompt, `${item.innerText.trim()}`, false, {
              outlookAction: bubbleOptions?.outlookAction || "reply",
              isolateSession: true
            });
          } else {
            executeGeminiWorkflow(refineOptions[key], `${item.innerText.trim()}`);
          }
        }
      };
    });

    primaryActions.appendChild(refineBtn);
    primaryActions.appendChild(refineMenu);
    actionsContainer.appendChild(primaryActions);
  }

  // Icon actions: Copy + Feedback
  const iconActions = document.createElement("div");
  iconActions.className = "footer-icon-actions";

  const copyIconBtn = document.createElement("button");
  copyIconBtn.className = "icon-btn";
  copyIconBtn.title = "Copy to clipboard";
  copyIconBtn.innerHTML = `
    <svg width="18" height="18" viewBox="0 0 24 24" fill="currentColor">
      <path d="M16 1H4c-1.1 0-2 .9-2 2v14h2V3h12V1zm3 4H8c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h11c1.1 0 2-.9 2-2V7c0-1.1-.9-2-2-2zm0 16H8V7h11v14z"/>
    </svg>
  `;
  copyIconBtn.onclick = async () => {
    const ok = await copyTextToClipboard(text);
    if (ok) {
      showToast("Copied to clipboard!");
    } else {
      showToast("Copy failed");
    }
  };

  iconActions.appendChild(copyIconBtn);
  actionsContainer.appendChild(iconActions);
  bubble.appendChild(actionsContainer);
  historyDiv.appendChild(bubble);
  historyDiv.scrollTop = historyDiv.scrollHeight;

  // POWERPOINT INSTANT SLIDE DECK ENHANCEMENT:
  // Immediately parse slides and display deck outline preview cards if multiple slides are detected
  if (isPPT) {
    try {
      import('../adapters/ppt/slidePreviewUI.js').then(({ enhanceBubbleWithSlideDeck }) => {
        enhanceBubbleWithSlideDeck(bubble, textDiv.innerHTML, fullText, hostAdapter);
      }).catch(err => {
        console.warn("Direct slide deck preview enhancement error:", err);
      });
    } catch (e) {
      console.warn("Direct slide preview invocation failed:", e);
    }
  }

  // Ensure action buttons remain fully visible when images finish decoding
  const bubbleImages = bubble.querySelectorAll("img");
  bubbleImages.forEach(img => {
    if (img.complete) {
      historyDiv.scrollTop = historyDiv.scrollHeight;
    } else {
      img.addEventListener("load", () => {
        historyDiv.scrollTop = historyDiv.scrollHeight;
      });
    }
  });
}

async function performDocumentInsertion(htmlContent, rawText, mode = "smart", options = {}) {
  const runButton = document.getElementById("run");
  const loadingText = document.getElementById("loading");

  const cleanHtml = normalizeEscapedNewlines(htmlContent || "");
  const cleanRaw = normalizeEscapedNewlines(rawText || "");

  if (runButton) runButton.disabled = true;
  if (loadingText) {
    let msg = "⚡ Updating document...";
    if (hostAdapter?.name === 'PowerPoint') {
      if (options?.imageOnly) {
        msg = mode === 'insert_current_slide' ? "⚡ Inserting image onto slide..." : "⚡ Inserting image as slide...";
      } else if (mode === 'replace_draft') {
        msg = "⚡ Replacing slide content...";
      } else if (mode === 'insert_current_slide') {
        msg = "⚡ Inserting onto current slide...";
      } else {
        msg = "⚡ Creating PowerPoint slides...";
      }
    }
    loadingText.innerText = msg;
    loadingText.style.display = "flex";
  }

  try {
    const isPPT = hostAdapter?.name === "PowerPoint";
    const isExcel = hostAdapter?.name === "Excel";
    if (isPPT) {
      await hostAdapter.insertContent(cleanHtml, cleanRaw, {
        ...options,
        layoutChoice: "auto",
        mode: mode === "replace_draft" ? "replace" : (mode === "insert_current_slide" ? "insert_current" : "insert")
      });
    } else if (isExcel) {
      await hostAdapter.insertContent(cleanHtml, cleanRaw, mode);
    } else if (hostAdapter) {
      await hostAdapter.insertContent(cleanHtml, mode, options);
    }
    const debugStatus = document.getElementById("debugStatus");
    if (debugStatus) {
      const host = hostAdapter?.name || 'Office';
      let actionLabel = mode === 'replace_draft' ? 'Replaced' : (mode === 'insert_current_slide' ? 'Inserted on Slide' : 'Inserted');
      debugStatus.innerText = `Updated in ${host} (${actionLabel})`;
    }
    showToast(mode === 'replace_draft' ? "Replaced selection in document" : "Inserted into document");
  } catch (err) {
    console.error("Document insertion error:", err);
    appendBubble(`🔴 Insertion Error: ${err.message || err}`, "system");
  } finally {
    if (runButton) runButton.disabled = false;
    if (loadingText) loadingText.style.display = "none";
  }
}
// ---------------------------------------------------------------------------
// Cloud Logging Dispatcher for Client Diagnostics & Troubleshooting
// ---------------------------------------------------------------------------
let pendingDiagLogs = [];
let diagLogFlushTimeout = null;

export function sendDiagnosticLogToCloud(message, level = "INFO", category = "DIAGNOSTICS", details = {}) {
  const entry = {
    timestamp: new Date().toISOString(),
    level: (level || "INFO").toUpperCase(),
    category: category || "DIAGNOSTICS",
    message: String(message),
    details: details || {}
  };

  pendingDiagLogs.push(entry);

  if (!diagLogFlushTimeout) {
    diagLogFlushTimeout = setTimeout(flushDiagnosticLogsToCloud, 600);
  }
}

async function flushDiagnosticLogsToCloud() {
  diagLogFlushTimeout = null;
  if (pendingDiagLogs.length === 0) return;

  const logsToSend = [...pendingDiagLogs];
  pendingDiagLogs = [];

  const profile = getUserProfile();
  const baseUrl = getActiveProxyUrl().replace(/\/askGeminiEnterprise$/, '');
  const logUrl = `${baseUrl}/api/diagnostics/log`;

  const payload = {
    logs: logsToSend,
    client_context: {
      host: hostAdapter?.name || (typeof Office !== "undefined" && Office.context?.host) || "UnknownHost",
      platform: (typeof Office !== "undefined" && Office.context?.platform) || "UnknownPlatform",
      user_id: profile?.user_id || "anonymous",
      user_email: profile?.email || null,
      tenant_id: profile?.tenant_id || null,
      session_id: currentSessionId || null,
      user_agent: typeof navigator !== "undefined" ? navigator.userAgent : null,
      addin_version: "1.0.0"
    }
  };

  try {
    const res = await fetch(logUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
      keepalive: true
    });
    if (!res.ok) {
      console.warn("Cloud log ingestion HTTP error:", res.status);
    }
  } catch (err) {
    console.debug("Failed to dispatch diagnostic logs to Cloud Logging:", err);
  }
}

export function logToDiagBox(msg, isError = false, category = "DIAGNOSTICS", details = {}) {
  const box = document.getElementById("diagLogBox");
  if (box) {
    const time = new Date().toLocaleTimeString();
    box.innerText += `\n[${time}] ${msg}`;
    box.scrollTop = box.scrollHeight;
  }
  // Stream directly to GCP Cloud Logging
  const level = isError ? "ERROR" : "INFO";
  sendDiagnosticLogToCloud(msg, level, category, details);
}
window.logToDiagBox = logToDiagBox;
window.sendDiagnosticLogToCloud = sendDiagnosticLogToCloud;

async function updateDiagnosticsPanel() {
  try {
    // 1. Entra ID SSO diagnostics
    const ssoStatus = document.getElementById("diagSsoStatus");
    const ssoUser = document.getElementById("diagSsoUser");
    const ssoTenant = document.getElementById("diagSsoTenant");
    const profile = getUserProfile();
    const token = await getOfficeAuthToken().catch(() => null);
    const lastErr = getLastAuthError();

    if (token && profile.is_authenticated) {
      if (ssoStatus) {
        ssoStatus.innerText = "✅ Authenticated";
        ssoStatus.style.color = "#146c2e";
      }
      if (ssoUser) ssoUser.innerText = profile.email || profile.name || "Signed In";
      if (ssoTenant) ssoTenant.innerText = profile.tenant_id ? (profile.tenant_id.length > 20 ? profile.tenant_id.substring(0, 18) + '...' : profile.tenant_id) : "Default Tenant";
    } else {
      if (ssoStatus) {
        ssoStatus.innerText = lastErr ? `⚠️ ${lastErr.code || lastErr.message || 'Not Authenticated'}` : "Not Authenticated";
        ssoStatus.style.color = "#b3261e";
      }
      if (ssoUser) ssoUser.innerText = profile.email && profile.email !== 'user@organization.com' ? profile.email : "--";
      if (ssoTenant) ssoTenant.innerText = "--";
    }

    // 2. Google OAuth diagnostics
    const googleStatus = document.getElementById("diagGoogleStatus");
    const googleClientId = document.getElementById("diagGoogleClientId");
    const googleExpiry = document.getElementById("diagGoogleExpiry");
    const gToken = getGoogleAccessToken();
    const gClientId = await getGoogleOAuthClientId();

    if (googleClientId) {
      googleClientId.innerText = gClientId ? (gClientId.length > 24 ? gClientId.substring(0, 22) + '...' : gClientId) : 'Not Configured';
      googleClientId.title = gClientId || 'Google OAuth Client ID not configured';
    }

    if (gToken) {
      if (googleStatus) {
        googleStatus.innerText = "✅ Connected";
        googleStatus.style.color = "#146c2e";
      }
      const exp = parseInt(window.sessionStorage?.getItem('google_user_token_expiry') || '0', 10);
      const remainingMin = exp ? Math.max(0, Math.round((exp - Date.now()) / 60000)) : 0;
      if (googleExpiry) googleExpiry.innerText = `${remainingMin} min remaining`;
    } else {
      if (googleStatus) {
        googleStatus.innerText = "⚪ Not Connected";
        googleStatus.style.color = "#5f6368";
      }
      if (googleExpiry) googleExpiry.innerText = "--";
    }

    // 3. Backend & Config
    const proxyUrl = document.getElementById("diagProxyUrl");
    const authMode = document.getElementById("diagAuthMode");
    const projId = document.getElementById("diagProjectId");

    const appConfig = await fetchAppConfig();
    if (proxyUrl) {
      const activeUrl = getActiveProxyUrl();
      proxyUrl.innerText = activeUrl.replace('https://', '').split('/')[0];
      proxyUrl.title = activeUrl;
    }
    if (authMode) authMode.innerText = appConfig?.user_auth_mode || "Office SSO + Google OAuth";
    if (projId) projId.innerText = appConfig?.project_id || (typeof ENV_CONFIG !== "undefined" && ENV_CONFIG.PROJECT_ID) || "gemini";

    // PowerPoint & Office.js Engine Status
    const pptApiStatus = document.getElementById("diagPptApiStatus");
    if (pptApiStatus) {
      if (typeof Office !== "undefined" && Office.context?.requirements) {
        const v110 = Office.context.requirements.isSetSupported("PowerPointApi", "1.10");
        const v19 = Office.context.requirements.isSetSupported("PowerPointApi", "1.9");
        const v18 = Office.context.requirements.isSetSupported("PowerPointApi", "1.8");
        const v17 = Office.context.requirements.isSetSupported("PowerPointApi", "1.7");
        const v16 = Office.context.requirements.isSetSupported("PowerPointApi", "1.6");
        const v15 = Office.context.requirements.isSetSupported("PowerPointApi", "1.5");
        const v14 = Office.context.requirements.isSetSupported("PowerPointApi", "1.4");
        const v13 = Office.context.requirements.isSetSupported("PowerPointApi", "1.3");
        const v12 = Office.context.requirements.isSetSupported("PowerPointApi", "1.2");
        const v11 = Office.context.requirements.isSetSupported("PowerPointApi", "1.1");
        const highest = v110 ? "1.10" : v19 ? "1.9" : v18 ? "1.8" : v17 ? "1.7" : v16 ? "1.6" : v15 ? "1.5" : v14 ? "1.4" : v13 ? "1.3" : v12 ? "1.2" : v11 ? "1.1" : "Base";
        pptApiStatus.innerText = `PowerPointApi ${highest} Supported`;
        pptApiStatus.style.color = "#146c2e";
      } else {
        const isPPT = typeof PowerPoint !== "undefined";
        pptApiStatus.innerText = isPPT ? "PowerPoint.js Ready" : (Office?.context?.host || "Office Environment");
      }
    }

  } catch (e) {
    console.warn("Diagnostics update error:", e);
  }
}

function initTroubleshootPanel() {
  const panel = document.getElementById("troubleshootPanel");
  if (panel) {
    panel.ontoggle = () => {
      if (panel.open) {
        updateDiagnosticsPanel();
        logToDiagBox("Troubleshooting panel expanded.");
      }
    };
  }

  const btnRefreshSso = document.getElementById("diagRefreshSso");
  if (btnRefreshSso) {
    btnRefreshSso.onclick = async () => {
      logToDiagBox("Forcing Entra ID SSO token refresh...", false, "SSO");
      btnRefreshSso.innerText = "⏳ Refreshing...";
      try {
        const token = await getOfficeAuthToken(true);
        if (token) {
          const profile = getUserProfile();
          logToDiagBox("✅ SSO Token refreshed successfully.", false, "SSO", { email: profile?.email, tenant: profile?.tenant_id });
        } else {
          logToDiagBox("⚠️ Token refresh returned empty.", true, "SSO");
        }
      } catch (err) {
        logToDiagBox(`❌ SSO Refresh error: ${err.message || err}`, true, "SSO", { error: String(err) });
      } finally {
        btnRefreshSso.innerText = "🔄 Refresh SSO Token";
        await initAuthUI();
        await updateDiagnosticsPanel();
      }
    };
  }

  const btnSignInGoogle = document.getElementById("diagSignInGoogle");
  if (btnSignInGoogle) {
    btnSignInGoogle.onclick = async () => {
      logToDiagBox("Launching Google OAuth sign-in flow...", false, "GOOGLE_OAUTH");
      btnSignInGoogle.innerText = "⏳ Signing in...";
      try {
        const profile = getUserProfile();
        const res = await initiateGoogleSignIn(null, 'select_account');
        if (res.status === 'success') {
          logToDiagBox("✅ Google Sign-In succeeded.", false, "GOOGLE_OAUTH", { status: 'success' });
        } else {
          logToDiagBox(`⚠️ Google Sign-In canceled/failed: ${res.error || 'Unknown'}`, true, "GOOGLE_OAUTH", { error: res.error });
        }
      } catch (err) {
        logToDiagBox(`❌ Google Sign-In error: ${err.message || err}`, true, "GOOGLE_OAUTH", { error: String(err) });
      } finally {
        btnSignInGoogle.innerText = "🔑 Sign In with Google";
        await initAuthUI();
        await updateDiagnosticsPanel();
      }
    };
  }

  const btnClearGoogle = document.getElementById("diagClearGoogle");
  if (btnClearGoogle) {
    btnClearGoogle.onclick = async () => {
      setGoogleAccessToken(null);
      logToDiagBox("🗑️ Google access token cleared.", false, "GOOGLE_OAUTH");
      await initAuthUI();
      await updateDiagnosticsPanel();
    };
  }

  const btnPingBackend = document.getElementById("diagPingBackend");
  if (btnPingBackend) {
    btnPingBackend.onclick = async () => {
      const startTime = Date.now();
      btnPingBackend.innerText = "⏳ Pinging...";
      const baseUrl = getActiveProxyUrl().replace(/\/askGeminiEnterprise$/, '');
      const configUrl = `${baseUrl}/api/config`;
      logToDiagBox(`Testing backend connectivity -> ${configUrl}`, false, "BACKEND_PING");
      try {
        const res = await fetch(configUrl, { cache: 'no-store' });
        const latency = Date.now() - startTime;
        if (res.ok) {
          const cfg = await res.json();
          logToDiagBox(`✅ Backend reachable (${latency}ms). Project: ${cfg.project_id || 'OK'}, AuthMode: ${cfg.user_auth_mode || 'OK'}`, false, "BACKEND_PING", { latency, config: cfg });
        } else {
          logToDiagBox(`⚠️ Backend HTTP ${res.status} (${latency}ms)`, true, "BACKEND_PING", { latency, status: res.status });
        }
      } catch (err) {
        logToDiagBox(`❌ Connectivity failed: ${err.message || err}`, true, "BACKEND_PING", { error: String(err) });
      } finally {
        btnPingBackend.innerText = "📡 Test Connection";
        await updateDiagnosticsPanel();
      }
    };
  }

  // PowerPoint Diagnostic & Live Test Actions
  const btnDiagCheckApi = document.getElementById("btnDiagCheckApi");
  if (btnDiagCheckApi) {
    btnDiagCheckApi.onclick = async () => {
      logToDiagBox("Checking PowerPoint Office.js environment...", false, "POWERPOINT_API");
      try {
        const hasOffice = typeof Office !== "undefined";
        const hasPPT = typeof PowerPoint !== "undefined";
        const host = (hasOffice && Office.context?.host) || "Unknown";
        const platform = (hasOffice && Office.context?.platform) || "Unknown";

        let versions = {};
        if (hasOffice && Office.context?.requirements) {
          versions = {
            v11: Office.context.requirements.isSetSupported("PowerPointApi", "1.1"),
            v12: Office.context.requirements.isSetSupported("PowerPointApi", "1.2"),
            v13: Office.context.requirements.isSetSupported("PowerPointApi", "1.3"),
            v14: Office.context.requirements.isSetSupported("PowerPointApi", "1.4"),
            v15: Office.context.requirements.isSetSupported("PowerPointApi", "1.5")
          };
          logToDiagBox(`Host: ${host}, Platform: ${platform}, PowerPointApi: 1.1=${versions.v11}, 1.2=${versions.v12}, 1.3=${versions.v13}, 1.4=${versions.v14}, 1.5=${versions.v15}`, false, "POWERPOINT_API", { host, platform, versions });
        } else {
          logToDiagBox(`Host: ${host}, Platform: ${platform}, PPT Namespace: ${hasPPT}`, false, "POWERPOINT_API", { host, platform, hasPPT });
        }
      } catch (e) {
        logToDiagBox(`❌ API Check Error: ${e.message}`, true, "POWERPOINT_API", { error: String(e) });
      }
    };
  }

  const btnDiagTestSlide = document.getElementById("btnDiagTestSlide");
  if (btnDiagTestSlide) {
    btnDiagTestSlide.onclick = async () => {
      logToDiagBox("Calling PowerPoint.run(presentation.slides.add())...", false, "POWERPOINT_TEST");
      const startTime = Date.now();
      try {
        if (typeof PowerPoint === "undefined") {
          throw new Error("PowerPoint namespace is not available in current host.");
        }
        await PowerPoint.run(async (context) => {
          context.presentation.slides.add();
          logToDiagBox("Slide add queued. Calling context.sync()...", false, "POWERPOINT_TEST");
          await context.sync();
        });
        const elapsed = Date.now() - startTime;
        logToDiagBox(`✅ Successfully added 1 blank slide via PowerPoint.run! (${elapsed}ms)`, false, "POWERPOINT_TEST", { test: "add_slide", elapsed });
      } catch (e) {
        logToDiagBox(`❌ Slide Add Error: ${e.message} (code: ${e.code || "N/A"})`, true, "POWERPOINT_TEST", { error: String(e), code: e.code });
      }
    };
  }

  const btnDiagTestText = document.getElementById("btnDiagTestText");
  if (btnDiagTestText) {
    btnDiagTestText.onclick = async () => {
      logToDiagBox("Testing Add Slide + Title Textbox + Body Textbox via getCount()...", false, "POWERPOINT_TEST");
      const startTime = Date.now();
      try {
        if (typeof PowerPoint === "undefined") {
          throw new Error("PowerPoint namespace is not available in current host.");
        }
        await PowerPoint.run(async (context) => {
          const slides = context.presentation.slides;
          slides.add();
          await context.sync();

          const countResult = slides.getCount();
          await context.sync();

          const slideCount = countResult.value;
          logToDiagBox(`Slide count: ${slideCount}. Fetching slide at index ${slideCount - 1}...`, false, "POWERPOINT_TEST");
          const slide = slides.getItemAt(slideCount - 1);

          slide.shapes.addTextBox("🧪 Diagnostic Test Title", {
            left: 50,
            top: 40,
            width: 650,
            height: 55
          });

          slide.shapes.addTextBox("• Bullet item 1: Official Microsoft Office.js pattern\n• Bullet item 2: Direct geometry textbox creation verified", {
            left: 50,
            top: 110,
            width: 650,
            height: 300
          });

          await context.sync();
        });
        const elapsed = Date.now() - startTime;
        logToDiagBox(`✅ Successfully created Slide with Title & Body Textbox! (${elapsed}ms)`, false, "POWERPOINT_TEST", { test: "add_slide_textbox", elapsed });
      } catch (e) {
        logToDiagBox(`❌ Slide Text Error: ${e.message} (code: ${e.code || "N/A"})`, true, "POWERPOINT_TEST", { error: String(e), code: e.code });
      }
    };
  }

  const btnClearLog = document.getElementById("diagClearLog");
  if (btnClearLog) {
    btnClearLog.onclick = (e) => {
      e.stopPropagation();
      const box = document.getElementById("diagLogBox");
      if (box) box.innerText = "[Ready] Log cleared.";
    };
  }
}

/* ==========================================================================
   Outlook Specific Controller Functions (Additive Only)
   ========================================================================== */

let activeOutlookTone = "Professional";

/**
 * Helper to execute an SSE streaming workflow for Outlook tabs,
 * updating the UI incrementally, avoiding duplicate loading text,
 * and handling errors cleanly.
 */
async function streamOutlookWorkflow({ prompt, loadingEl, resultCardEl, resultBodyEl, options = {}, onChunk = null }) {
  if (loadingEl) loadingEl.style.display = "block";
  if (resultCardEl) resultCardEl.style.display = "flex";
  if (resultBodyEl) resultBodyEl.innerHTML = "";

  let accumulated = "";
  try {
    const data = await askGeminiEnterpriseStream(
      prompt,
      (chunk) => {
        accumulated = chunk.accumulatedText || "";
        if (typeof onChunk === 'function') {
          onChunk(chunk);
        } else if (resultBodyEl && accumulated) {
          resultBodyEl.innerHTML = parseMarkdown(accumulated);
          resultBodyEl.scrollTop = resultBodyEl.scrollHeight;
        }
      },
      options
    );

    const finalRaw = data.result || accumulated;
    if (resultBodyEl && !onChunk) {
      resultBodyEl.innerHTML = parseMarkdown(finalRaw || "No content returned.");
    }
    return { data, finalRaw, accumulated };
  } catch (err) {
    console.error("streamOutlookWorkflow error:", err);
    if (resultBodyEl && !accumulated) {
      resultBodyEl.innerHTML = `<div style="color:#a4262c;">Error: ${err.message}</div>`;
    }
    throw err;
  } finally {
    if (loadingEl) loadingEl.style.display = "none";
  }
}

function setupOutlookUI() {
  const isCompose = hostAdapter && typeof hostAdapter.isComposeMode === 'function' && hostAdapter.isComposeMode();

  // 1. 5-Tab Navigation Switcher
  const tabs = [
    { btnId: "tabBtnSummarize", panelId: "panelSummarize" },
    { btnId: "tabBtnReply", panelId: "panelReply" },
    { btnId: "tabBtnCoach", panelId: "panelCoach" },
    { btnId: "tabBtnActionItems", panelId: "panelActionItems" },
    { btnId: "tabBtnAsk", panelId: "panelAsk" }
  ];

  // Compose vs Read mode visibility for Coach tab (DF-5)
  const tabBtnCoach = document.getElementById("tabBtnCoach");
  if (tabBtnCoach) {
    tabBtnCoach.style.display = isCompose ? "" : "none";
  }

  // Ensure default active tab in Read mode is Summarize
  if (!isCompose) {
    const activeTab = document.querySelector(".outlook-tabs-nav .tab-btn.active");
    if (activeTab && activeTab.id === "tabBtnCoach") {
      const tabSummarize = document.getElementById("tabBtnSummarize");
      if (tabSummarize) tabSummarize.click();
    }
  }

  tabs.forEach(tab => {
    const btn = document.getElementById(tab.btnId);
    if (btn) {
      btn.onclick = () => {
        tabs.forEach(t => {
          const b = document.getElementById(t.btnId);
          const p = document.getElementById(t.panelId);
          if (b) b.classList.remove("active");
          if (p) {
            p.classList.remove("active");
            p.style.display = "none";
          }
        });
        btn.classList.add("active");
        const activePanel = document.getElementById(tab.panelId);
        if (activePanel) {
          activePanel.classList.add("active");
          activePanel.style.display = "flex";
        }
      };
    }
  });

  // 2. Tab 1: Summarize
  const btnSummarize = document.getElementById("btnSummarizeThread");
  if (btnSummarize) {
    btnSummarize.onclick = handleOutlookSummarize;
  }

  // 4. Tab 2: Reply
  const tonePills = document.querySelectorAll(".outlook-tone-bar .tone-pill");
  tonePills.forEach(pill => {
    pill.onclick = () => {
      tonePills.forEach(p => p.classList.remove("active"));
      pill.classList.add("active");
      activeOutlookTone = pill.getAttribute("data-tone") || "Professional";
    };
  });

  const btnGenerateReply = document.getElementById("btnGenerateReply");
  if (btnGenerateReply) {
    btnGenerateReply.onclick = handleOutlookReply;
  }

  const replyInput = document.getElementById("replyIntentText");
  if (replyInput) {
    replyInput.addEventListener("keydown", (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key === "Enter") {
        e.preventDefault();
        handleOutlookReply();
      }
    });
  }

  // 5. Tab 3: Coach My Email (OUT-04)
  const btnCoach = document.getElementById("btnCoachEmail");
  if (btnCoach) {
    btnCoach.onclick = handleOutlookCoachEmail;
  }

  // 6. Tab 4: Meeting Agenda & Action Items (OUT-03)
  const btnPrepareBriefing = document.getElementById("btnPrepareBriefing");
  if (btnPrepareBriefing) {
    btnPrepareBriefing.onclick = handleOutlookMeetingBriefing;
  }

  const btnExtractActions = document.getElementById("btnExtractActionItems");
  if (btnExtractActions) {
    btnExtractActions.onclick = handleOutlookActionItems;
  }

  // 7. Tab 5: Ask
  const btnAskGemini = document.getElementById("btnAskGemini");
  if (btnAskGemini) {
    btnAskGemini.onclick = handleOutlookAsk;
  }

  const askInput = document.getElementById("askQuestionText");
  if (askInput) {
    askInput.addEventListener("keydown", (e) => {
      if (e.key === "Enter" && !e.shiftKey) {
        e.preventDefault();
        handleOutlookAsk();
      }
    });
  }

  // 8. Attach Pinnable Taskpane ItemChanged handler
  try {
    if (Office.context && Office.context.mailbox && typeof Office.context.mailbox.addHandlerAsync === 'function') {
      Office.context.mailbox.addHandlerAsync(
        Office.EventType.ItemChanged,
        () => handleOutlookItemChanged()
      );
    }
  } catch (e) {
    console.warn("Could not attach Outlook ItemChanged handler:", e);
  }
}

function handleOutlookItemChanged() {
  // Option A: Save previous conversation in Recent Chats and reset to fresh empty state for the new email
  resetChatSession();

  // Refresh unified Outlook header mode badge, compact action chips, and mode-aware '↵' starter suggestions
  adaptUIForHost("Outlook");

  // Re-synchronize compose vs read mode UI for the newly active item
  const isCompose = hostAdapter && typeof hostAdapter.isComposeMode === 'function' && hostAdapter.isComposeMode();
  const tabCoach = document.getElementById("tabBtnCoach");
  const panelCoach = document.getElementById("panelCoach");
  if (tabCoach) {
    tabCoach.style.display = isCompose ? "block" : "none";
  }
  if (!isCompose && panelCoach && panelCoach.classList.contains("active")) {
    const tabSummarize = document.getElementById("tabBtnSummarize");
    if (tabSummarize) tabSummarize.click();
  }
}

function parseOutlookCoachingResponse(text) {
  if (!text) return { assessment: "", optionA: "", optionB: "" };
  const markerARegex = /(?:[\*\#=]+\s*)?OPTION A\s*[:\-–\(\s]+CONCISE EXECUTIVE[\s\*\#=\)]*/i;
  const markerBRegex = /(?:[\*\#=]+\s*)?OPTION B\s*[:\-–\(\s]+DIPLOMATIC[\s\*\#=\)]*/i;

  let assessment = "";
  let optionA = "";
  let optionB = "";

  const matchA = markerARegex.exec(text);
  const matchB = markerBRegex.exec(text);

  if (matchA) {
    assessment = text.substring(0, matchA.index).trim();
    if (matchB && matchB.index > matchA.index) {
      optionA = text.substring(matchA.index + matchA[0].length, matchB.index).trim();
      optionB = text.substring(matchB.index + matchB[0].length).trim();
    } else {
      optionA = text.substring(matchA.index + matchA[0].length).trim();
    }
  } else {
    assessment = text.trim();
  }

  return { assessment, optionA, optionB };
}

function renderOutlookCoachInBubble(containerEl, rawText) {
  if (!containerEl) return;
  const { assessment, optionA, optionB } = parseOutlookCoachingResponse(rawText);
  if (!optionA && !optionB) return;

  const cleanOptionA = optionA || "No concise executive version generated.";
  const cleanOptionB = optionB || "No diplomatic version generated.";
  const htmlA = parseMarkdown(cleanOptionA);
  const htmlB = parseMarkdown(cleanOptionB);
  const htmlAssessment = parseMarkdown(assessment || "### Tone & Clarity Assessment\nReview complete.");

  containerEl.innerHTML = `
    <div class="coach-analysis-card" style="margin-bottom: 10px;">${htmlAssessment}</div>
    <div class="coach-options-wrapper">
      <div class="coach-option-card">
        <div class="coach-option-header">
          <span class="coach-option-badge badge-executive">⚡ Option A: Concise Executive</span>
          <span class="coach-option-sub">Direct, high-impact</span>
        </div>
        <div class="coach-option-body">${htmlA}</div>
        <div class="coach-option-actions">
          <button type="button" class="outlook-action-btn primary-action btn-coach-apply-a">✨ Apply Concise Executive</button>
          <button type="button" class="outlook-action-btn btn-coach-copy-a">📋 Copy</button>
        </div>
      </div>
      <div class="coach-option-card">
        <div class="coach-option-header">
          <span class="coach-option-badge badge-diplomatic">🤝 Option B: Diplomatic</span>
          <span class="coach-option-sub">Relationship-preserving</span>
        </div>
        <div class="coach-option-body">${htmlB}</div>
        <div class="coach-option-actions">
          <button type="button" class="outlook-action-btn primary-action btn-coach-apply-b">🤝 Apply Diplomatic</button>
          <button type="button" class="outlook-action-btn btn-coach-copy-b">📋 Copy</button>
        </div>
      </div>
    </div>
  `;

  const isCompose = hostAdapter && typeof hostAdapter.isComposeMode === "function" && hostAdapter.isComposeMode();
  const btnApplyA = containerEl.querySelector(".btn-coach-apply-a");
  const btnApplyB = containerEl.querySelector(".btn-coach-apply-b");
  const btnCopyA = containerEl.querySelector(".btn-coach-copy-a");
  const btnCopyB = containerEl.querySelector(".btn-coach-copy-b");

  if (btnApplyA) {
    btnApplyA.onclick = async () => {
      if (isCompose) {
        await hostAdapter.insertContent(htmlA, cleanOptionA, "replace_all");
        showToast("Applied Concise Executive option to draft");
      } else {
        hostAdapter.replyWithDraft(htmlA, true);
      }
    };
  }
  if (btnApplyB) {
    btnApplyB.onclick = async () => {
      if (isCompose) {
        await hostAdapter.insertContent(htmlB, cleanOptionB, "replace_all");
        showToast("Applied Diplomatic option to draft");
      } else {
        hostAdapter.replyWithDraft(htmlB, true);
      }
    };
  }
  if (btnCopyA) {
    btnCopyA.onclick = () => copyOutlookTextToClipboard(cleanOptionA, btnCopyA, "📋 Copy");
  }
  if (btnCopyB) {
    btnCopyB.onclick = () => copyOutlookTextToClipboard(cleanOptionB, btnCopyB, "📋 Copy");
  }
}

/**
 * Shared prompt builder for all 5 structured Outlook actions ("summarize", "reply",
 * "coach", "briefing", "actions"), used by both the unified chat stream and legacy handlers.
 */
export function wrapOutlookTables(html = "") {
  if (!html || !html.includes("<table")) return html;
  return String(html).replace(/<table\b[\s\S]*?<\/table>/gi, (match) => `<div class="table-scroll-wrapper">${match}</div>`);
}

export function buildOutlookActionPrompt(actionType, contextStr = "", extra = {}) {
  if (actionType === "summarize") {
    return `You are an elite executive email assistant for an Account Manager. Analyze and summarize the following email thread:
${contextStr}

Please strictly format your output according to these requirements:

## 📋 Executive Briefing
Provide exactly 3 concise, high-impact bullet points summarizing the core developments, status, and business impact:
* **[Key Milestone / Topic 1]:** [Executive takeaway with numbers/commitments/decisions]
* **[Key Milestone / Topic 2]:** [Executive takeaway with numbers/commitments/decisions]
* **[Key Milestone / Topic 3]:** [Executive takeaway with numbers/commitments/decisions]

## ✅ Assigned Action Items & Follow-ups
Extract all commitments, deliverables, and open questions with explicit assignees and deadlines:
| Action Item / Deliverable | Assignee / Owner | Due Date / Timing | Status / Context |
| :--- | :--- | :--- | :--- |
(List each action item row clearly)

Format cleanly in Markdown with bold headers and standard Markdown table syntax.`;
  }

  if (actionType === "reply") {
    const senderName = extra.senderName || "Colleague";
    const currentUserName = extra.currentUserName || "User";
    const tone = extra.tone || activeOutlookTone;
    const intent = extra.intent || "";
    return `You are an elite Customer Success Representative and executive communications specialist. Draft a grounded email reply to the following customer inquiry email thread:
${contextStr}

Recipient & Identity Context:
- Incoming Sender Name (to whom you are replying): ${senderName}
- Current User / Author Name (who is sending this reply): ${currentUserName}

Reply Requirements & Contract Grounding:
- Ground your response in contract data stores, enterprise agreement records, and SLA terms.
- Cite exact SLA clauses, turnaround commitments, resolution windows, escalation policies, and verified enterprise facts from the data store when addressing contract terms, service commitments, or customer inquiries.
- Tone: ${tone}
- Greeting / Addressee: Address the greeting to "${senderName}".
- Sign-off / Signature: Sign off using ONLY the current user's name ("${currentUserName}").
- CRITICAL CONSTRAINT: NEVER sign off with the incoming sender's name ("${senderName}"). Always sign off as "${currentUserName}".
- Specific Instructions / Points to include: ${intent ? intent : "Politely acknowledge customer questions, address SLA and contract terms with verified enterprise clauses, and detail clear next steps."}
- Do not include subject lines or email headers (Date:, To:, From:, Subject:) in the body; output only the ready-to-send email body in clean Markdown.`;
  }

  if (actionType === "coach") {
    const draftText = extra.draftText || "";
    return `You are an elite executive communications coach and deal strategist. Review the following drafted email for tone, clarity, brevity, nuance, and potential miscommunication risks on critical deals:

--- THREAD / RECIPIENT CONTEXT ---
${contextStr ? contextStr.substring(0, 1500) : "New email compose"}
----------------------------------

--- USER DRAFT EMAIL TO COACH ---
${draftText}
--------------------------------

Please provide your review in the following exact format with these structural section headers:

### Tone & Clarity Assessment
* **Clarity & Brevity Score:** [High / Moderate / Needs Improvement]
* **Tone & Deal Risk:** [1-sentence concise risk synthesis]

=== OPTION A: CONCISE EXECUTIVE ===
[Draft a high-impact, direct, brevity-focused version that gets straight to decisions/metrics and eliminates unnecessary fluff. Suitable for busy C-level stakeholders.]

=== OPTION B: DIPLOMATIC ===
[Draft a relationship-preserving, warm, collaborative version that softens hard edges, fosters alignment, and handles sensitive nuances gracefully without losing clarity.]

IMPORTANT: Maintain the exact delimiter markers "=== OPTION A: CONCISE EXECUTIVE ===" and "=== OPTION B: DIPLOMATIC ===" so the UI can parse each option into interactive cards. Output ready-to-send email body text without subject lines.`;
  }

  if (actionType === "actions") {
    return `You are an executive email assistant. Extract all Action Items, Deliverables, Commitments, and Key Dates from the following email thread:
${contextStr}

Please format the response as:
1. Structured Markdown Table with columns: Task | Assignee | Due Date | Priority
2. Key Dates & Milestones List
3. Recommended Next Meeting or Follow-up Topic`;
  }

  return contextStr;
}

/**
 * Executes any of the 5 structured Outlook capabilities ("summarize", "reply", "coach",
 * "briefing", "actions") directly inside the unified #officeDocHostContainer conversational stream.
 */
export async function runOutlookChatAction(actionType, customIntent = "") {
  if (isGenerating || isProcessingInDocCommand) return;
  const loadingText = document.getElementById("loading");
  if (loadingText) {
    loadingText.innerText = "⚡ Reading email context...";
    loadingText.style.display = "flex";
  }

  try {
    if (actionType === "summarize") {
      const emailContext = await hostAdapter.getFullDocumentText();
      const prompt = buildOutlookActionPrompt("summarize", emailContext);
      await executeGeminiWorkflow(prompt, "📋 Summarize Thread", false, {
        outlookAction: "summarize",
        isolateSession: true
      });
      return;
    }

    if (actionType === "reply") {
      const emailContext = await hostAdapter.getFullDocumentText();
      const userProfile = hostAdapter.getUserProfile();
      const senderProfile = await hostAdapter.getSenderProfile();
      const prompt = buildOutlookActionPrompt("reply", emailContext, {
        senderName: senderProfile.displayName || "Colleague",
        currentUserName: userProfile.displayName || "User",
        tone: activeOutlookTone,
        intent: customIntent
      });

      const displayLabel = customIntent ? `✉️ Draft Grounded Reply: "${customIntent}"` : "✉️ Draft Grounded Reply";
      await executeGeminiWorkflow(prompt, displayLabel, true, {
        outlookAction: "reply",
        isolateSession: false
      });

      // In Compose mode, auto-populate the active draft body
      if (hostAdapter && typeof hostAdapter.isComposeMode === "function" && hostAdapter.isComposeMode()) {
        const historyDiv = document.getElementById("chatHistory");
        const lastBubble = historyDiv ? historyDiv.querySelector(".chat-bubble.assistant:last-of-type > div:first-child") : null;
        if (lastBubble && lastBubble.innerHTML) {
          try {
            await hostAdapter.insertContent(lastBubble.innerHTML, lastBubble.innerText || "", "replace_draft");
          } catch (insertErr) {
            console.warn("Auto-insertion into active composer notice:", insertErr);
          }
        }
      }
      return;
    }

    if (actionType === "coach") {
      let draftText = await hostAdapter.getDraftBody();
      if (!draftText || draftText.trim().length < 5) {
        draftText = await hostAdapter.getSelectedText();
      }
      if (!draftText || draftText.trim().length < 5) {
        updateChatDisplayState(true);
        appendBubble("⚠️ Please write or select a draft in your email composer before requesting coaching.", "system");
        return;
      }

      const emailThreadContext = await hostAdapter.getFullDocumentText();
      const prompt = buildOutlookActionPrompt("coach", emailThreadContext, { draftText });
      await executeGeminiWorkflow(prompt, "🎯 Coach My Email", false, {
        outlookAction: "coach",
        isolateSession: true
      });
      return;
    }

    if (actionType === "actions") {
      const emailContext = await hostAdapter.getFullDocumentText();
      const prompt = buildOutlookActionPrompt("actions", emailContext);
      await executeGeminiWorkflow(prompt, "✅ Extract Action Items & Dates", false, {
        outlookAction: "actions",
        isolateSession: true
      });
      return;
    }
  } finally {
    if (loadingText && !isGenerating) {
      loadingText.style.display = "none";
    }
  }
}

async function handleOutlookSummarize() {
  const btn = document.getElementById("btnSummarizeThread");
  const loading = document.getElementById("summarizeLoading");
  const resultCard = document.getElementById("summarizeResultCard");
  const resultBody = document.getElementById("summarizeResultBody");

  if (btn) btn.disabled = true;

  try {
    const emailContext = await hostAdapter.getFullDocumentText();
    const prompt = buildOutlookActionPrompt("summarize", emailContext);

    const { finalRaw } = await streamOutlookWorkflow({
      prompt,
      loadingEl: loading,
      resultCardEl: resultCard,
      resultBodyEl: resultBody,
      options: { history: [] }
    });

    const rawResult = finalRaw || "No summary generated.";
    const formattedHtml = parseMarkdown(rawResult);

    if (resultBody) resultBody.innerHTML = formattedHtml;

    // Bind action buttons: Copy Summary, Forward with Summary, Reply with Summary
    const btnCopy = document.getElementById("btnCopySummary");
    const btnInsert = document.getElementById("btnInsertSummary"); // "Forward with Summary" in HTML
    const btnReply = document.getElementById("btnReplySummary");   // "Reply with Summary" in HTML

    if (btnCopy) {
      btnCopy.innerHTML = `${CONTENT_COPY_ICON_SVG}Copy Summary`;
      btnCopy.onclick = () => copyOutlookTextToClipboard(rawResult, btnCopy, `${CONTENT_COPY_ICON_SVG}Copy Summary`);
    }

    if (btnInsert) {
      btnInsert.onclick = () => {
        hostAdapter.forwardWithDraft(formattedHtml);
      };
    }

    if (btnReply) {
      btnReply.onclick = () => {
        hostAdapter.replyWithDraft(formattedHtml, true);
      };
    }
  } catch (err) {
    if (resultBody && !resultBody.innerHTML) {
      resultBody.innerHTML = `<div style="color:#a4262c;">Error: ${err.message}</div>`;
    }
  } finally {
    if (btn) btn.disabled = false;
  }
}

/**
 * Handles User Story [OUT-04]: Email Tone & Clarity Coaching
 * Evaluates active composer draft and suggests dual versions (Concise Executive vs Diplomatic).
 */
async function handleOutlookCoachEmail() {
  const btn = document.getElementById("btnCoachEmail");
  const loading = document.getElementById("coachLoading");
  const resultContainer = document.getElementById("coachResultContainer");
  const analysisBox = document.getElementById("coachAnalysisBox");
  const optionABody = document.getElementById("coachOptionABody");
  const optionBBody = document.getElementById("coachOptionBBody");
  const btnApplyA = document.getElementById("btnApplyOptionA");
  const btnApplyB = document.getElementById("btnApplyOptionB");
  const btnCopyA = document.getElementById("btnCopyOptionA");
  const btnCopyB = document.getElementById("btnCopyOptionB");

  if (btn) btn.disabled = true;
  if (loading) loading.style.display = "block";
  if (resultContainer) resultContainer.style.display = "none";

  try {
    const draftText = await hostAdapter.getDraftBody();
    if (!draftText || draftText.trim().length < 5) {
      if (analysisBox) {
        analysisBox.innerHTML = `<div>⚠️ <strong>No draft email detected.</strong> Please write or select a draft in your email composer before requesting coaching.</div>`;
      }
      if (optionABody) optionABody.innerHTML = "<em>Draft text required for coaching options.</em>";
      if (optionBBody) optionBBody.innerHTML = "<em>Draft text required for coaching options.</em>";
      if (resultContainer) resultContainer.style.display = "flex";
      return;
    }

    const emailThreadContext = await hostAdapter.getFullDocumentText();
    const prompt = buildOutlookActionPrompt("coach", emailThreadContext, { draftText });

    let accumulated = '';
    if (resultContainer) resultContainer.style.display = "flex";
    if (analysisBox) analysisBox.innerHTML = "";
    if (optionABody) optionABody.innerHTML = "";
    if (optionBBody) optionBBody.innerHTML = "";

    const data = await askGeminiEnterpriseStream(
      prompt,
      (chunk) => {
        accumulated = chunk.accumulatedText || '';
        const parsed = parseOutlookCoachingResponse(accumulated);
        if (analysisBox && parsed.assessment) {
          analysisBox.innerHTML = parseMarkdown(parsed.assessment);
        }
        if (optionABody && parsed.optionA) {
          optionABody.innerHTML = parseMarkdown(parsed.optionA);
        }
        if (optionBBody && parsed.optionB) {
          optionBBody.innerHTML = parseMarkdown(parsed.optionB);
        }
      },
      { history: [] }
    );

    const finalRaw = data.result || accumulated;
    const { assessment, optionA, optionB } = parseOutlookCoachingResponse(finalRaw);

    const cleanOptionA = optionA || "No concise executive version generated.";
    const cleanOptionB = optionB || "No diplomatic version generated.";

    const htmlA = parseMarkdown(cleanOptionA);
    const htmlB = parseMarkdown(cleanOptionB);

    if (analysisBox) analysisBox.innerHTML = parseMarkdown(assessment || "Tone & Clarity Review Complete.");
    if (optionABody) optionABody.innerHTML = htmlA;
    if (optionBBody) optionBBody.innerHTML = htmlB;

    // Bind Apply & Copy buttons
    if (btnApplyA) {
      btnApplyA.onclick = async () => {
        btnApplyA.disabled = true;
        const originalText = btnApplyA.innerText;
        btnApplyA.innerText = "⏳ Applying...";
        try {
          await hostAdapter.insertContent(htmlA, cleanOptionA, "replace_all");
          btnApplyA.innerText = "✅ Applied!";
          setTimeout(() => { btnApplyA.innerText = originalText; btnApplyA.disabled = false; }, 2000);
        } catch (err) {
          console.error("Failed to apply Option A:", err);
          btnApplyA.innerText = "🔴 Failed";
          setTimeout(() => { btnApplyA.innerText = originalText; btnApplyA.disabled = false; }, 2000);
        }
      };
    }

    if (btnApplyB) {
      btnApplyB.onclick = async () => {
        btnApplyB.disabled = true;
        const originalText = btnApplyB.innerText;
        btnApplyB.innerText = "⏳ Applying...";
        try {
          await hostAdapter.insertContent(htmlB, cleanOptionB, "replace_all");
          btnApplyB.innerText = "✅ Applied!";
          setTimeout(() => { btnApplyB.innerText = originalText; btnApplyB.disabled = false; }, 2000);
        } catch (err) {
          console.error("Failed to apply Option B:", err);
          btnApplyB.innerText = "🔴 Failed";
          setTimeout(() => { btnApplyB.innerText = originalText; btnApplyB.disabled = false; }, 2000);
        }
      };
    }

    if (btnCopyA) {
      btnCopyA.innerHTML = `${CONTENT_COPY_ICON_SVG}Copy`;
      btnCopyA.onclick = () => copyOutlookTextToClipboard(cleanOptionA, btnCopyA, `${CONTENT_COPY_ICON_SVG}Copy`);
    }

    if (btnCopyB) {
      btnCopyB.innerHTML = `${CONTENT_COPY_ICON_SVG}Copy`;
      btnCopyB.onclick = () => copyOutlookTextToClipboard(cleanOptionB, btnCopyB, `${CONTENT_COPY_ICON_SVG}Copy`);
    }
  } catch (err) {
    console.error("handleOutlookCoachEmail error:", err);
    if (analysisBox) analysisBox.innerHTML = `<div style="color:#a4262c;">Error: ${err.message}</div>`;
    if (resultContainer) resultContainer.style.display = "flex";
  } finally {
    if (btn) btn.disabled = false;
    if (loading) loading.style.display = "none";
  }
}

/**
 * Handles User Story [OUT-02]: Grounded Smart Reply Drafter
 * Drafts professional email replies grounded in contract data stores citing exact SLA clauses.
 */
async function handleOutlookReply() {
  const btn = document.getElementById("btnGenerateReply");
  const loading = document.getElementById("replyLoading");
  const resultCard = document.getElementById("replyResultCard");
  const resultBody = document.getElementById("replyResultBody");
  const intentInput = document.getElementById("replyIntentText");

  const intent = intentInput ? intentInput.value.trim() : "";

  if (btn) btn.disabled = true;

  try {
    const emailContext = await hostAdapter.getFullDocumentText();
    const userProfile = hostAdapter.getUserProfile();
    const senderProfile = await hostAdapter.getSenderProfile();

    const prompt = buildOutlookActionPrompt("reply", emailContext, {
      senderName: senderProfile.displayName || "Colleague",
      currentUserName: userProfile.displayName || "User",
      tone: activeOutlookTone,
      intent
    });

    const { finalRaw } = await streamOutlookWorkflow({
      prompt,
      loadingEl: loading,
      resultCardEl: resultCard,
      resultBodyEl: resultBody,
      options: { history: [], enableGrounding: true }
    });

    const rawResult = finalRaw || "No reply generated.";
    const formattedHtml = parseMarkdown(rawResult);

    if (resultBody) resultBody.innerHTML = formattedHtml;

    // In Compose mode, populate text into active compose editor
    if (hostAdapter && typeof hostAdapter.isComposeMode === 'function' && hostAdapter.isComposeMode()) {
      try {
        await hostAdapter.insertContent(formattedHtml, rawResult, "replace_draft");
      } catch (insertErr) {
        console.warn("Auto-insertion into active composer notice:", insertErr);
      }
    }

    // Bind action buttons: [Open Reply All] | [Open Reply] | [Copy Draft]
    const btnOpenReplyAll = document.getElementById("btnOpenReplyAll");
    if (btnOpenReplyAll) {
      btnOpenReplyAll.onclick = () => hostAdapter.replyWithDraft(formattedHtml, true);
    }
    const btnOpenReply = document.getElementById("btnOpenReply");
    if (btnOpenReply) {
      btnOpenReply.onclick = () => hostAdapter.replyWithDraft(formattedHtml, false);
    }
    const btnCopy = document.getElementById("btnCopyReply");
    if (btnCopy) {
      btnCopy.innerHTML = `${CONTENT_COPY_ICON_SVG}Copy Draft`;
      btnCopy.onclick = () => copyOutlookTextToClipboard(rawResult, btnCopy, `${CONTENT_COPY_ICON_SVG}Copy Draft`);
    }
  } catch (err) {
    if (resultBody && !resultBody.innerHTML) {
      resultBody.innerHTML = `<div style="color:#a4262c;">Error: ${err.message}</div>`;
    }
  } finally {
    if (btn) btn.disabled = false;
  }
}

/**
 * Handles User Story [OUT-03]: Meeting Agenda & Briefing Prep
 * Generates pre-meeting briefing notes and structured bulleted agenda based on thread context and enterprise records.
 */
async function handleOutlookMeetingBriefing() {
  const btn = document.getElementById("btnPrepareBriefing");
  const loading = document.getElementById("actionItemsLoading");
  const resultCard = document.getElementById("actionItemsResultCard");
  const resultBody = document.getElementById("actionItemsResultBody");

  if (btn) btn.disabled = true;
  if (loading) loading.innerText = "⚡ Querying data store & preparing meeting briefing notes...";

  try {
    const documentContext = await hostAdapter.getFullDocumentText();
    const userProfile = hostAdapter.getUserProfile();
    const hostDisplayName = userProfile.displayName || "Meeting Host";
    const prompt = `You are an executive meeting facilitator and chief of staff. Generate comprehensive pre-meeting briefing notes and a structured agenda based on the following thread context, calendar invite details, and enterprise records:
${documentContext}

Meeting Host / Organizer Context:
- Current User / Host: ${hostDisplayName}

Please strictly format your output according to these requirements:

## 📋 Pre-Meeting Briefing & Strategic Objective
* **Meeting Purpose:** [1-2 concise sentences defining the core goal and desired outcome of this session]
* **Background & Enterprise Context:** [Bulleted summary of key background facts, contract terms, account history, or recent developments grounded in enterprise records]

## ⏱️ Structured Meeting Agenda
| Time / Duration | Agenda Topic / Discussion Item | Lead / Owner | Objective / Key Takeaway |
| :--- | :--- | :--- | :--- |
(List 3-5 structured agenda topics covering review, core decision-making, and next steps)

## 🎯 Key Decisions Required & Discussion Prompts
* [Decision Point 1 / Key Question to resolve]
* [Decision Point 2 / Key Question to resolve]

## 👥 Attendee Preparation & Action Items
* [Specific preparation materials or reviews assigned to attendees prior to joining]

Format cleanly in Markdown with bold headers and standard Markdown table syntax.`;

    const { finalRaw } = await streamOutlookWorkflow({
      prompt,
      loadingEl: loading,
      resultCardEl: resultCard,
      resultBodyEl: resultBody,
      options: { history: [], enableGrounding: true }
    });

    const rawResult = finalRaw || "No meeting briefing generated.";
    const formattedHtml = parseMarkdown(rawResult);

    if (resultBody) resultBody.innerHTML = formattedHtml;

    // Bind action buttons
    const btnCalendar = document.getElementById("btnCreateCalendarInvite");
    if (btnCalendar) {
      btnCalendar.onclick = () => {
        hostAdapter.createCalendarInvite({
          subject: "Pre-Meeting Briefing & Agenda",
          body: formattedHtml,
          rawText: rawResult
        });
      };
    }
    const btnCopy = document.getElementById("btnCopyActionItems");
    if (btnCopy) {
      btnCopy.innerHTML = `${CONTENT_COPY_ICON_SVG}Copy Output`;
      btnCopy.onclick = () => copyOutlookTextToClipboard(rawResult, btnCopy, `${CONTENT_COPY_ICON_SVG}Copy Output`);
    }
    const btnInsert = document.getElementById("btnInsertActionItems");
    if (btnInsert) {
      btnInsert.onclick = async () => {
        await hostAdapter.insertContent(formattedHtml, rawResult, "smart");
      };
    }
  } catch (err) {
    if (resultBody && !resultBody.innerHTML) {
      resultBody.innerHTML = `<div style="color:#a4262c;">Error: ${err.message}</div>`;
    }
  } finally {
    if (btn) btn.disabled = false;
  }
}

async function handleOutlookActionItems() {
  const btn = document.getElementById("btnExtractActionItems");
  const loading = document.getElementById("actionItemsLoading");
  const resultCard = document.getElementById("actionItemsResultCard");
  const resultBody = document.getElementById("actionItemsResultBody");

  if (btn) btn.disabled = true;
  if (loading) loading.innerText = "⚡ Extracting action items, deliverables, and key dates...";

  try {
    const emailContext = await hostAdapter.getFullDocumentText();
    const prompt = buildOutlookActionPrompt("actions", emailContext);

    const { finalRaw } = await streamOutlookWorkflow({
      prompt,
      loadingEl: loading,
      resultCardEl: resultCard,
      resultBodyEl: resultBody,
      options: { history: [] }
    });

    const rawResult = finalRaw || "No action items found.";
    const formattedHtml = parseMarkdown(rawResult);

    if (resultBody) resultBody.innerHTML = formattedHtml;

    // Bind action buttons
    const btnCalendar = document.getElementById("btnCreateCalendarInvite");
    if (btnCalendar) {
      btnCalendar.onclick = () => {
        hostAdapter.createCalendarInvite({
          subject: "Follow-up on Action Items & Milestones",
          body: formattedHtml,
          rawText: rawResult
        });
      };
    }
    const btnCopy = document.getElementById("btnCopyActionItems");
    if (btnCopy) {
      btnCopy.innerHTML = `${CONTENT_COPY_ICON_SVG}Copy Table`;
      btnCopy.onclick = () => copyOutlookTextToClipboard(rawResult, btnCopy, `${CONTENT_COPY_ICON_SVG}Copy Table`);
    }
    const btnInsert = document.getElementById("btnInsertActionItems");
    if (btnInsert) {
      btnInsert.onclick = async () => {
        await hostAdapter.insertContent(formattedHtml, rawResult, "smart");
      };
    }
  } catch (err) {
    if (resultBody && !resultBody.innerHTML) {
      resultBody.innerHTML = `<div style="color:#a4262c;">Error: ${err.message}</div>`;
    }
  } finally {
    if (btn) btn.disabled = false;
  }
}

async function handleOutlookAsk() {
  const btn = document.getElementById("btnAskGemini");
  const loading = document.getElementById("askLoading");
  const resultCard = document.getElementById("askResultCard");
  const resultBody = document.getElementById("askResultBody");
  const questionInput = document.getElementById("askQuestionText");

  const question = questionInput ? questionInput.value.trim() : "";
  if (!question) return;

  if (btn) btn.disabled = true;

  try {
    const emailContext = await hostAdapter.getFullDocumentText();
    const prompt = `Email Context:
${emailContext}

User Question: ${question}

Please provide a clear, accurate, and grounded answer based on the email thread context above.`;

    const { data, finalRaw } = await streamOutlookWorkflow({
      prompt,
      loadingEl: loading,
      resultCardEl: resultCard,
      resultBodyEl: resultBody,
      options: { history: askChatHistory, sessionId: askSessionId }
    });

    if (data) {
      if (data.sessionId) askSessionId = data.sessionId;
      if (Array.isArray(data.history)) askChatHistory = data.history;
    }

    const rawResult = finalRaw || "No answer generated.";
    const formattedHtml = parseMarkdown(rawResult);

    if (resultBody) resultBody.innerHTML = formattedHtml;

    // Bind action buttons
    const btnCopy = document.getElementById("btnCopyAnswer");
    if (btnCopy) {
      btnCopy.innerHTML = `${CONTENT_COPY_ICON_SVG}Copy Answer`;
      btnCopy.onclick = () => copyOutlookTextToClipboard(rawResult, btnCopy, `${CONTENT_COPY_ICON_SVG}Copy Answer`);
    }
    const btnInsert = document.getElementById("btnInsertAnswer");
    if (btnInsert) {
      btnInsert.onclick = async () => {
        await hostAdapter.insertContent(formattedHtml, rawResult, "smart");
      };
    }
    const btnReply = document.getElementById("btnReplyAnswer");
    if (btnReply) {
      btnReply.onclick = () => {
        hostAdapter.replyWithDraft(formattedHtml, false);
      };
    }
  } catch (err) {
    if (resultBody && !resultBody.innerHTML) {
      resultBody.innerHTML = `<div style="color:#a4262c;">Error: ${err.message}</div>`;
    }
  } finally {
    if (btn) btn.disabled = false;
  }
}

async function copyOutlookTextToClipboard(text, btnElement = null, originalLabel = "") {
  let success = false;

  // 1. Try modern Clipboard API
  if (navigator.clipboard && window.isSecureContext) {
    try {
      await navigator.clipboard.writeText(text);
      success = true;
    } catch (e) {
      console.warn("navigator.clipboard failed in iframe context, using execCommand fallback:", e);
    }
  }

  // 2. Reliable Fallback via hidden textarea & execCommand('copy') for Office / WebKit iframes
  if (!success) {
    try {
      const textarea = document.createElement("textarea");
      textarea.value = text;
      textarea.style.position = "fixed";
      textarea.style.left = "-9999px";
      textarea.style.top = "0";
      textarea.style.opacity = "0";
      document.body.appendChild(textarea);
      textarea.focus();
      textarea.select();
      success = document.execCommand("copy");
      document.body.removeChild(textarea);
    } catch (err) {
      console.error("Clipboard copy completely failed:", err);
    }
  }

  if (success && btnElement) {
    const prevHtml = btnElement.innerHTML;
    btnElement.innerHTML = "✅ Copied!";
    setTimeout(() => {
      btnElement.innerHTML = originalLabel || prevHtml;
    }, 2000);
  }
}



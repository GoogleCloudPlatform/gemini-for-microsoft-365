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
 * Authentication Service for Microsoft Office 365 Add-in SSO
 * 
 * Handles silent token acquisition via Office.auth.getAccessToken(),
 * token caching, fallback consent dialogs, and JWT payload parsing.
 * 
 * @author Sathya AG & Antigravity Team
 */

let cachedToken = null;
let tokenExpiry = 0;
let cachedUserProfile = null;

/**
 * Auth-proxy URL baked in at build time (GEMINI_PROXY_URL in microsoft-addin/.env).
 *
 * webpack's DefinePlugin substitutes the literal expression
 * `process.env.GEMINI_PROXY_URL` with a string, but it does not create a
 * `process` global. The previous `typeof process !== 'undefined' && ...`
 * guards were therefore always false in the browser: every deployment whose
 * hostname is not special-cased (anything other than the shared 16933400417
 * and WIF services) fell back to http://localhost:8080, so chat, slash
 * commands, history and /api/config never reached the configured auth proxy.
 */
export const BUILD_TIME_PROXY_URL = (() => {
  try {
    return (process.env.GEMINI_PROXY_URL || '').trim();
  } catch (e) {
    return ''; // Not built by webpack, so there is no substituted value.
  }
})();

/**
 * Global build-time environment parameters injected from .env via Webpack DefinePlugin.
 */
export const ENV_CONFIG = (() => {
  try {
    return {
      PROJECT_ID: (process.env.PROJECT_ID || '').trim(),
      REGION: (process.env.REGION || 'us-central1').trim(),
      SERVICE_ACCOUNT: (process.env.SERVICE_ACCOUNT || '').trim(),
      GE_GCP_PROJECT_ID: (process.env.GE_GCP_PROJECT_ID || '').trim(),
      GE_GCP_LOCATION: (process.env.GE_GCP_LOCATION || 'us').trim(),
      STREAM_ASSIST_ENDPOINT_LOCATION: (process.env.STREAM_ASSIST_ENDPOINT_LOCATION || 'us').trim(),
      GEMINI_ENTERPRISE_APP_ID: (process.env.GEMINI_ENTERPRISE_APP_ID || '').trim(),
      ENTERPRISE_COLLECTION_ID: (process.env.ENTERPRISE_COLLECTION_ID || 'default_collection').trim(),
      ENTERPRISE_ASSISTANT_ID: (process.env.ENTERPRISE_ASSISTANT_ID || 'default_assistant').trim(),
      MICROSOFT_ENTRA_APP_ID: (process.env.MICROSOFT_ENTRA_APP_ID || '').trim(),
      MICROSOFT_ENTRA_TENANT_ID: (process.env.MICROSOFT_ENTRA_TENANT_ID || '').trim(),
      GOOGLE_OAUTH_CLIENT_ID: (process.env.GOOGLE_OAUTH_CLIENT_ID || '').trim(),
    };
  } catch (e) {
    return {};
  }
})();

/**
 * Decodes the base64url payload of a JWT without verifying signature
 * (signature verification is performed securely on the auth-proxy backend).
 */
export function decodeJwtPayload(token) {
  if (!token || typeof token !== 'string') return null;
  try {
    const parts = token.split('.');
    if (parts.length !== 3) return null;
    const base64Url = parts[1];
    const base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/');
    const jsonPayload = decodeURIComponent(
      atob(base64)
        .split('')
        .map(c => '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2))
        .join('')
    );
    return JSON.parse(jsonPayload);
  } catch (err) {
    console.warn('Could not decode JWT payload:', err);
    return null;
  }
}

/**
 * Retrieves the signed-in user's profile information from Office.js or decoded JWT claims.
 */
export function getUserProfile() {
  if (cachedUserProfile) return cachedUserProfile;

  // 1. If we have a cached JWT token, extract from claims
  if (cachedToken) {
    const claims = decodeJwtPayload(cachedToken);
    if (claims) {
      cachedUserProfile = {
        name: claims.name || claims.preferred_username || claims.upn || 'Corporate User',
        email: claims.preferred_username || claims.email || claims.upn || 'user@contoso.com',
        user_id: claims.oid || claims.sub || 'entra_user',
        tenant_id: claims.tid || null,
        roles: claims.roles || [],
        is_authenticated: true
      };
      return cachedUserProfile;
    }
  }

  // 2. Fallback to Office.context user if available
  try {
    if (typeof Office !== 'undefined' && Office.context) {
      if (Office.context.user && Office.context.user.displayName) {
        return {
          name: Office.context.user.displayName,
          email: Office.context.user.email || 'user@office.com',
          user_id: Office.context.user.accountId || 'office_user',
          is_authenticated: false
        };
      }
      if (Office.context.mailbox && Office.context.mailbox.userProfile) {
        const up = Office.context.mailbox.userProfile;
        return {
          name: up.displayName || 'Corporate User',
          email: up.emailAddress || 'user@office.com',
          user_id: up.emailAddress || 'office_user',
          is_authenticated: false
        };
      }
    }
  } catch (e) {
    console.debug('Office context user inspection skipped:', e);
  }

  return {
    name: 'Office 365 User',
    email: 'user@organization.com',
    user_id: 'dev_user',
    is_authenticated: false
  };
}

let lastAuthError = null;

export function getLastAuthError() {
  return lastAuthError;
}

/**
 * Silent Office SSO normally answers within a second or two, but it can also
 * never settle (for example when a consent prompt cannot be shown, or the task
 * pane is opened outside an Office client). Chat requests await this token, so
 * a silent attempt is bounded and, after a timeout, not retried for a while.
 * Interactive attempts (forceRefresh) stay unbounded because they may be
 * waiting on the user to finish signing in.
 */
const SILENT_SSO_TIMEOUT_MS = 10000;
const SILENT_SSO_RETRY_AFTER_MS = 5 * 60 * 1000;
const SSO_TIMED_OUT = { timedOut: true };
let silentSsoInFlight = null;
let silentSsoRetryAfter = 0;

function settleWithin(promise, ms) {
  let timer = null;
  const timeout = new Promise((resolve) => {
    timer = setTimeout(() => resolve(SSO_TIMED_OUT), ms);
  });
  return Promise.race([promise, timeout]).then(
    (value) => { clearTimeout(timer); return value; },
    (error) => { clearTimeout(timer); throw error; }
  );
}

function rememberOfficeToken(token) {
  cachedToken = token;
  lastAuthError = null;
  silentSsoRetryAfter = 0;

  // Update cached user profile and compute expiry from JWT exp claim (2-minute safety buffer)
  const claims = decodeJwtPayload(token);
  const expMs = (claims && typeof claims.exp === 'number')
    ? (claims.exp * 1000) - (2 * 60 * 1000)
    : Date.now() + (50 * 60 * 1000);
  tokenExpiry = expMs;

  if (claims) {
    cachedUserProfile = {
      name: claims.name || claims.preferred_username || claims.upn || 'Corporate User',
      email: claims.preferred_username || claims.email || claims.upn || 'user@contoso.com',
      user_id: claims.oid || claims.sub || 'entra_user',
      tenant_id: claims.tid || null,
      roles: claims.roles || [],
      is_authenticated: true
    };
  }

  try {
    if (typeof window !== 'undefined' && window.localStorage) {
      window.localStorage.setItem('office_entra_access_token', token);
      window.localStorage.setItem('office_entra_token_expiry', String(tokenExpiry));
    }
    if (typeof window !== 'undefined' && window.sessionStorage) {
      window.sessionStorage.setItem('office_entra_access_token', token);
      window.sessionStorage.setItem('office_entra_token_expiry', String(tokenExpiry));
    }
  } catch (_) {}
}

async function requestOfficeToken(interactive) {
  try {
    console.log('Acquiring Microsoft Entra ID SSO token via Office.auth.getAccessToken()...');
    const tokenRequest = Office.auth.getAccessToken({
      allowSignInPrompt: interactive,
      allowConsentPrompt: interactive,
      forMSGraphAccess: false
    });
    const token = interactive ? await tokenRequest : await settleWithin(tokenRequest, SILENT_SSO_TIMEOUT_MS);

    if (token === SSO_TIMED_OUT) {
      silentSsoRetryAfter = Date.now() + SILENT_SSO_RETRY_AFTER_MS;
      lastAuthError = {
        code: 'SSO_TIMEOUT',
        message: `Office SSO did not respond within ${SILENT_SSO_TIMEOUT_MS / 1000}s`
      };
      console.warn(`Office SSO timed out; continuing without a Microsoft Entra ID token. Silent SSO retries are paused for ${SILENT_SSO_RETRY_AFTER_MS / 60000} minutes.`);
      // Keep a late answer for the next request rather than discarding it.
      tokenRequest.then((late) => { if (late) rememberOfficeToken(late); }, () => {});
      return null;
    }

    if (token) {
      const claims = decodeJwtPayload(token);
      if (!interactive && claims && typeof claims.exp === 'number' && (claims.exp * 1000 - 120000 <= Date.now())) {
        return requestOfficeToken(true);
      }
      rememberOfficeToken(token);
      console.log('Successfully acquired Entra ID SSO token for user:', cachedUserProfile ? cachedUserProfile.email : 'authenticated');
      return token;
    }
  } catch (error) {
    lastAuthError = error;
    console.error('Office SSO token acquisition failed:', error);

    // Specific Office SSO error code diagnostics
    const code = error && error.code;
    if (code === 13001) {
      console.warn('SSO Error 13001: User is not signed into Office with a Microsoft Entra ID account.');
    } else if (code === 13002) {
      console.warn('SSO Error 13002: User cancelled the consent dialog.');
    } else if (code === 13003) {
      console.warn('SSO Error 13003: User type is not supported (e.g. personal Microsoft Account).');
    } else if (code === 13007) {
      console.warn('SSO Error 13007: Invalid Application ID URI or untrusted client application (Manifest Resource domain must match hosting domain).');
    } else if (code === 13012) {
      console.warn('SSO Error 13012: Precondition failed (e.g. platform API not supported).');
    }

    // Return null so the client can attempt calling the proxy in fallback mode if allowed
    return null;
  }

  return null;
}

/**
 * Retrieves the Microsoft Entra ID JWT access token via Office.js SSO.
 * Automatically handles caching and silent acquisition.
 *
 * @param {boolean} forceRefresh - If true, bypasses in-memory cache
 * @returns {Promise<string|null>} Entra ID JWT access token or null in dev/fallback mode
 */
export async function getOfficeAuthToken(forceRefresh = false) {
  const now = Date.now();
  if (forceRefresh) {
    cachedToken = null;
    tokenExpiry = 0;
    try {
      if (typeof window !== 'undefined' && window.localStorage) {
        window.localStorage.removeItem('office_entra_access_token');
        window.localStorage.removeItem('office_entra_token_expiry');
      }
      if (typeof window !== 'undefined' && window.sessionStorage) {
        window.sessionStorage.removeItem('office_entra_access_token');
        window.sessionStorage.removeItem('office_entra_token_expiry');
      }
    } catch (_) {}
  } else {
    if (cachedToken && now < tokenExpiry) {
      return cachedToken;
    }
    try {
      const storage = (typeof window !== 'undefined' && (window.localStorage || window.sessionStorage)) || null;
      if (storage) {
        const storedToken = (window.localStorage && window.localStorage.getItem('office_entra_access_token')) ||
          (window.sessionStorage && window.sessionStorage.getItem('office_entra_access_token'));
        const storedExp = parseInt(
          (window.localStorage && window.localStorage.getItem('office_entra_token_expiry')) ||
          (window.sessionStorage && window.sessionStorage.getItem('office_entra_token_expiry')) ||
          '0',
          10
        );
        if (storedToken && now < storedExp) {
          rememberOfficeToken(storedToken);
          return cachedToken;
        }
      }
    } catch (_) {}
  }

  if (typeof Office === 'undefined' || !Office.auth || !Office.auth.getAccessToken) {
    console.warn('Office.auth.getAccessToken is not available in this environment. Running in unauthenticated / dev mode.');
    lastAuthError = { code: 'NO_OFFICE_AUTH', message: 'Office.auth API not available in this host environment.' };
    return null;
  }

  if (forceRefresh) {
    return requestOfficeToken(true);
  }

  // Office rejects overlapping getAccessToken calls, so concurrent silent
  // callers (identity bar, diagnostics panel, chat requests) share one attempt.
  if (silentSsoInFlight) {
    return silentSsoInFlight;
  }
  if (now < silentSsoRetryAfter) {
    return null;
  }
  silentSsoInFlight = requestOfficeToken(false);
  try {
    return await silentSsoInFlight;
  } finally {
    silentSsoInFlight = null;
  }
}

// ---------------------------------------------------------------------------
// Google 3-Legged OAuth Management (No DWD Required)
// ---------------------------------------------------------------------------

let cachedGoogleToken = null;
let googleTokenExpiry = 0;
let cachedAppConfig = null;
let googleAuthDialog = null;

export function getProxyBaseUrl() {
  if (typeof window !== 'undefined' && window.location && window.location.hostname) {
    const host = window.location.hostname;
    if (host === 'localhost' || host === '127.0.0.1') {
      if (window.localStorage) {
        const override = window.localStorage.getItem('gemini_proxy_url');
        if (override && (override.includes('localhost') || override.includes('127.0.0.1'))) {
          return override.replace(/\/askGeminiEnterprise$/, '');
        }
      }
      return window.location.origin || 'https://localhost:3000';
    }
    // Dynamic Cloud Run deduction: Automatically pair gemini-frontend with auth-proxy in any GCP project
    if (host.includes('gemini-frontend')) {
      return window.location.origin.replace('gemini-frontend', 'auth-proxy');
    }
  }
  if (typeof window !== 'undefined' && window.localStorage) {
    const override = window.localStorage.getItem('gemini_proxy_url');
    if (override) return override.replace(/\/askGeminiEnterprise$/, '');
  }
  if (BUILD_TIME_PROXY_URL) {
    return BUILD_TIME_PROXY_URL.replace(/\/askGeminiEnterprise$/, '');
  }
  return 'http://localhost:8080';
}

/**
 * Loads dynamic frontend configuration from the backend auth-proxy (/api/config)
 * so that Google OAuth Client IDs and settings are never hardcoded in client code.
 */
export async function fetchAppConfig() {
  if (cachedAppConfig) return cachedAppConfig;

  // Check localStorage override
  if (typeof window !== 'undefined' && window.localStorage) {
    const localClientId = window.localStorage.getItem('google_oauth_client_id');
    if (localClientId) {
      cachedAppConfig = { google_oauth_client_id: localClientId };
      return cachedAppConfig;
    }
  }

  try {
    const configUrl = `${getProxyBaseUrl()}/api/config`;
    const resp = await fetch(configUrl);
    if (resp.ok) {
      cachedAppConfig = await resp.json();
      console.log('Successfully fetched dynamic app config:', {
        google_client_id_configured: bool(cachedAppConfig?.google_oauth_client_id),
        user_auth_mode: cachedAppConfig?.user_auth_mode
      });
      return cachedAppConfig;
    }
  } catch (err) {
    console.warn('Could not fetch backend app config:', err);
  }

  return cachedAppConfig || {};
}

function bool(val) {
  return !!val;
}

/**
 * Retrieves the configured Google OAuth 2.0 Web Client ID dynamically.
 */
export async function getGoogleOAuthClientId() {
  if (typeof window !== 'undefined' && window.localStorage) {
    const stored = window.localStorage.getItem('google_oauth_client_id');
    if (stored) return stored;
  }
  const config = await fetchAppConfig();
  if (config?.google_oauth_client_id) return config.google_oauth_client_id;
  try {
    return (process.env.GOOGLE_OAUTH_CLIENT_ID || '').trim();
  } catch (e) {
    return '';
  }
}

/**
 * Stores Google OAuth access token with expiration in memory, localStorage, and sessionStorage,
 * and optionally syncs it to the backend auth proxy so other Office surfaces (Outlook, Excel, Word, PowerPoint)
 * running in isolated browser storage partitions can share the session via Microsoft Entra ID SSO.
 */
export function setGoogleAccessToken(token, expiresIn = 3600, syncToBackend = true) {
  if (!token) {
    cachedGoogleToken = null;
    googleTokenExpiry = 0;
    try {
      if (typeof window !== 'undefined' && window.localStorage) {
        window.localStorage.removeItem('google_user_access_token');
        window.localStorage.removeItem('google_user_token_expiry');
      }
      if (typeof window !== 'undefined' && window.sessionStorage) {
        window.sessionStorage.removeItem('google_user_access_token');
        window.sessionStorage.removeItem('google_user_token_expiry');
      }
    } catch (_) {}
    if (syncToBackend) {
      pushGoogleTokenToBackend(null, 0).catch(() => {});
    }
    return;
  }

  cachedGoogleToken = token;
  const now = Date.now();
  // Safe buffer: expire 5 minutes earlier
  const ttlMs = Math.max((expiresIn - 300) * 1000, 60000);
  googleTokenExpiry = now + ttlMs;

  try {
    if (typeof window !== 'undefined' && window.localStorage) {
      window.localStorage.setItem('google_user_access_token', token);
      window.localStorage.setItem('google_user_token_expiry', String(googleTokenExpiry));
    }
    if (typeof window !== 'undefined' && window.sessionStorage) {
      window.sessionStorage.setItem('google_user_access_token', token);
      window.sessionStorage.setItem('google_user_token_expiry', String(googleTokenExpiry));
    }
  } catch (_) {}
  if (syncToBackend) {
    pushGoogleTokenToBackend(token, expiresIn).catch(() => {});
  }
  console.log('Google user access token stored. Valid for ~' + Math.round(ttlMs / 60000) + ' minutes.');
}

/**
 * Pushes (or clears) the user's Google OAuth token to the backend auth proxy keyed by their
 * verified Microsoft Entra ID SSO token so all 4 Office surfaces share the Google session.
 */
async function pushGoogleTokenToBackend(googleToken, expiresIn = 3600) {
  if (typeof fetch === 'undefined') return false;
  try {
    const entraToken = await getOfficeAuthToken();
    if (!entraToken) return false;
    const baseUrl = getProxyBaseUrl();
    if (!googleToken) {
      const res = await fetch(`${baseUrl}/api/auth/google-token`, {
        method: 'DELETE',
        headers: {
          'Authorization': `Bearer ${entraToken}`
        }
      });
      return res.ok;
    }
    const res = await fetch(`${baseUrl}/api/auth/google-token`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${entraToken}`
      },
      body: JSON.stringify({
        google_token: googleToken,
        expires_in: expiresIn
      })
    });
    return res.ok;
  } catch (_) {
    return false;
  }
}

/**
 * Synchronizes the user's active Google OAuth token from the backend auth proxy if not present
 * in the current surface's localStorage partition (e.g., logging into Outlook on `outlook.office.com`
 * and then opening Excel/Word/PowerPoint on `*.sharepoint.com`, or vice versa).
 *
 * @returns {Promise<string|null>} Resolved Google access token if available
 */
export async function syncGoogleTokenFromBackend() {
  const existing = getGoogleAccessToken();
  if (existing) return existing;
  if (typeof fetch === 'undefined') return null;

  try {
    const entraToken = await getOfficeAuthToken();
    if (!entraToken) return null;

    const baseUrl = getProxyBaseUrl();
    const res = await fetch(`${baseUrl}/api/auth/me`, {
      method: 'GET',
      headers: {
        'Accept': 'application/json',
        'Authorization': `Bearer ${entraToken}`
      }
    });
    if (!res.ok) return null;

    const data = await res.json();
    if (data && data.google_token) {
      const expiresIn = Number(data.google_token_expires_in) > 0 ? Number(data.google_token_expires_in) : 3600;
      // Store locally without re-pushing to backend
      setGoogleAccessToken(data.google_token, expiresIn, false);
      return data.google_token;
    }
  } catch (_) {
    // Non-fatal fallback if backend is unreachable
  }
  return null;
}

/**
 * Retrieves the active Google user OAuth token if valid.
 */
export function getGoogleAccessToken() {
  const now = Date.now();
  if (cachedGoogleToken && now < googleTokenExpiry) {
    return cachedGoogleToken;
  }

  // Check localStorage and sessionStorage across email item switches (Desktop & Web)
  try {
    if (typeof window !== 'undefined') {
      const stored = (window.localStorage && window.localStorage.getItem('google_user_access_token')) ||
        (window.sessionStorage && window.sessionStorage.getItem('google_user_access_token'));
      const exp = parseInt(
        (window.localStorage && window.localStorage.getItem('google_user_token_expiry')) ||
        (window.sessionStorage && window.sessionStorage.getItem('google_user_token_expiry')) ||
        '0',
        10
      );
      if (stored && now < exp) {
        cachedGoogleToken = stored;
        googleTokenExpiry = exp;
        return cachedGoogleToken;
      }
      if (stored && now >= exp) {
        if (window.localStorage) {
          window.localStorage.removeItem('google_user_access_token');
          window.localStorage.removeItem('google_user_token_expiry');
        }
        if (window.sessionStorage) {
          window.sessionStorage.removeItem('google_user_access_token');
          window.sessionStorage.removeItem('google_user_token_expiry');
        }
      }
    }
  } catch (_) {}

  return null;
}

export function isGoogleTokenValid() {
  return !!getGoogleAccessToken();
}

/**
 * Initiates the 3-Legged Google OAuth Sign-In flow using Office Dialog API.
 * 
 * @param {string|null} loginHint - Email hint (e.g. scim@jeansson.demo.altostrat.com)
 * @param {string} prompt - OAuth prompt mode ('select_account' or 'none')
 * @returns {Promise<{status: string, token?: string, error?: string}>}
 */
export async function initiateGoogleSignIn(loginHint = null, prompt = 'select_account') {
  const clientId = await getGoogleOAuthClientId();
  if (!clientId) {
    const errMsg = 'Google OAuth Client ID is not configured. Please verify auth-proxy /api/config and ensure GOOGLE_OAUTH_CLIENT_ID is set.';
    console.error(errMsg);
    return { status: 'error', error: errMsg };
  }

  // Allow flexible account selection: Do not force loginHint from Office profile
  // so users can freely choose their Google Account across cross-tenant or federated domains
  // unless an explicit loginHint argument was provided by the caller.

  return new Promise((resolve) => {
    const origin = (typeof window !== 'undefined' && window.location && window.location.origin) ? window.location.origin : '';
    const authUrl = `${origin}/google-auth.html?client_id=${encodeURIComponent(clientId)}&prompt=${encodeURIComponent(prompt)}`;

    console.log('Opening Google OAuth dialog at:', authUrl);

    if (typeof Office !== 'undefined' && Office.context && Office.context.ui && Office.context.ui.displayDialogAsync) {
      Office.context.ui.displayDialogAsync(
        authUrl,
        { height: 60, width: 40, promptBeforeOpen: false },
        (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.error('Failed to open Office Google Auth dialog:', asyncResult.error);
            resolve({ status: 'error', error: asyncResult.error.message });
            return;
          }

          googleAuthDialog = asyncResult.value;

          googleAuthDialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
            try {
              const data = JSON.parse(arg.message);
              if (data.google_token) {
                setGoogleAccessToken(data.google_token, data.expires_in || 3600);
                if (googleAuthDialog) {
                  googleAuthDialog.close();
                  googleAuthDialog = null;
                }
                resolve({ status: 'success', token: data.google_token });
              } else {
                if (googleAuthDialog) {
                  googleAuthDialog.close();
                  googleAuthDialog = null;
                }
                resolve({ status: 'error', error: data.error || 'Google authentication was not completed.' });
              }
            } catch (e) {
              console.error('Error handling dialog message:', e);
              if (googleAuthDialog) {
                googleAuthDialog.close();
                googleAuthDialog = null;
              }
              resolve({ status: 'error', error: e.message });
            }
          });

          googleAuthDialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
            console.warn('Dialog event / closed by user:', arg);
            googleAuthDialog = null;
            resolve({ status: 'error', error: 'Authentication window closed.' });
          });
        }
      );
    } else {
      // Standalone browser fallback popup
      const popup = window.open(authUrl, 'google_auth_popup', 'width=500,height=650');
      const messageHandler = (event) => {
        // Enforce strict origin and window source verification
        if (origin && event.origin !== origin) {
          return;
        }
        if (popup && event.source !== popup) {
          return;
        }
        if (event.data && event.data.type === 'GOOGLE_AUTH_RESULT') {
          window.removeEventListener('message', messageHandler);
          const payload = event.data.payload || {};
          if (payload.google_token) {
            setGoogleAccessToken(payload.google_token, payload.expires_in || 3600);
            resolve({ status: 'success', token: payload.google_token });
          } else {
            resolve({ status: 'error', error: payload.error || 'Authentication failed' });
          }
        }
      };
      window.addEventListener('message', messageHandler);
    }
  });
}

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
 * Gemini API Client for Microsoft 365
 * 
 * @author Sathya AG, Principal Architect, Google
 */

import { getOfficeAuthToken, getUserProfile, getGoogleAccessToken, BUILD_TIME_PROXY_URL } from './authService.js';
import { SessionManager } from './sessionManager.js';

function detectActiveHostName() {
  try {
    if (typeof window !== 'undefined' && window.__activeOfficeHost) {
      return String(window.__activeOfficeHost);
    }
    if (typeof Office !== 'undefined' && Office.context) {
      const rawHost = String(Office.context.host || Office.context.diagnostics?.host || '').toLowerCase();
      if (rawHost.includes('excel')) return 'Excel';
      if (rawHost.includes('powerpoint')) return 'PowerPoint';
      if (rawHost.includes('outlook') || Office.context.mailbox) return 'Outlook';
      if (rawHost.includes('word')) return 'Word';
    }
    if (typeof Excel !== 'undefined' && typeof Word === 'undefined') return 'Excel';
    if (typeof PowerPoint !== 'undefined' && typeof Word === 'undefined') return 'PowerPoint';
  } catch (e) {
    // ignore
  }
  return null;
}

// Default endpoint points directly to the local dev gateway or secure auth-proxy gateway
const DEFAULT_AUTH_PROXY_URL = 'http://localhost:8080/askGeminiEnterprise';

function resolveBaseProxyUrl() {
  // 1. Check for URL query parameter (?proxyUrl=... or ?backendUrl=...)
  // Security: Restrict query overrides strictly to same-origin or local development to prevent token exfiltration
  if (typeof window !== 'undefined' && window.location && window.location.search) {
    try {
      const urlParams = new URLSearchParams(window.location.search);
      const queryUrl = urlParams.get('proxyUrl') || urlParams.get('backendUrl') || urlParams.get('endpoint');
      if (queryUrl && queryUrl.trim()) {
        const trimmed = queryUrl.trim();
        const parsed = new URL(trimmed, window.location.origin);
        const isSameOrigin = parsed.origin === window.location.origin;
        const isLocal = parsed.hostname === 'localhost' || parsed.hostname === '127.0.0.1';
        if (isSameOrigin || isLocal) {
          return trimmed;
        } else {
          console.warn('[SECURITY] Blocked untrusted proxy URL query parameter:', trimmed);
        }
      }
    } catch (e) {
      // ignore
    }
  }

  if (typeof window !== 'undefined' && window.location && window.location.hostname) {
    const host = window.location.hostname;
    if (host === 'localhost' || host === '127.0.0.1') {
      try {
        if (window.localStorage) {
          const override = window.localStorage.getItem('gemini_proxy_url');
          if (override && (override.includes('localhost') || override.includes('127.0.0.1'))) {
            return override.trim();
          }
        }
      } catch (e) {
        // ignore
      }
      const origin = window.location.origin || 'https://localhost:3000';
      return `${origin}/askGeminiEnterprise`;
    }
    if (host.includes('gemini-frontend')) {
      return `${window.location.origin.replace('gemini-frontend', 'auth-proxy')}/askGeminiEnterprise`;
    }
  }
  if (typeof window !== 'undefined') {
    try {
      if (window.localStorage) {
        const override = window.localStorage.getItem('gemini_proxy_url');
        if (override && override.trim()) return override.trim();
      }
    } catch (e) {
      // ignore
    }
  }
  // Build-time GEMINI_PROXY_URL (see BUILD_TIME_PROXY_URL in authService.js for
  // why this must not be guarded by `typeof process`).
  if (BUILD_TIME_PROXY_URL) {
    return BUILD_TIME_PROXY_URL;
  }
  return DEFAULT_AUTH_PROXY_URL;
}

export function getActiveProxyUrl(mode = 'unary') {
  const url = resolveBaseProxyUrl();
  if (mode === 'stream') {
    return url.replace(/\/askGemini(?:Enterprise)?(?=\?|$)/, '/askGeminiStream');
  }
  return url;
}

export function setProxyUrlOverride(url) {
  if (typeof window !== 'undefined') {
    try {
      if (window.localStorage) {
        if (url && url.trim()) {
          window.localStorage.setItem('gemini_proxy_url', url.trim());
        } else {
          window.localStorage.removeItem('gemini_proxy_url');
        }
      }
    } catch (e) {
      // localStorage may be blocked in cross-origin iframe
    }
  }
}

export function getProxyBaseUrl() {
  const full = getActiveProxyUrl();
  try {
    const u = new URL(full);
    return `${u.protocol}//${u.host}`;
  } catch (e) {
    return 'http://localhost:8080';
  }
}

export async function askGeminiEnterprise(prompt, history = [], sessionId = null, enableGrounding = true, attachments = null) {
  const functionUrl = getActiveProxyUrl();
  const userProfile = getUserProfile();

  // 1. Acquire Microsoft Entra ID SSO token
  let authToken = null;
  try {
    authToken = await getOfficeAuthToken();
  } catch (authErr) {
    console.warn('Proceeding without SSO token (server may reject if REQUIRE_ENTRA_AUTH=true):', authErr);
  }

  // 2. Check for 3-legged Google User OAuth token (for Google Drive grounding without DWD)
  const googleUserToken = getGoogleAccessToken();
  const activeHost = detectActiveHostName();

  const payload = { 
    prompt: prompt,
    history: history,
    enableGrounding: enableGrounding,
    userPseudoId: userProfile.email || userProfile.user_id || 'office_365_user'
  };
  if (activeHost) {
    payload.host = activeHost;
  }
  if (sessionId) {
    payload.sessionId = sessionId;
  }
  if (attachments && Array.isArray(attachments) && attachments.length > 0) {
    payload.attachments = attachments;
  }

  const headers = { 
    'Content-Type': 'application/json' 
  };
  if (authToken) {
    headers['Authorization'] = `Bearer ${authToken}`;
  }
  if (googleUserToken) {
    headers['X-End-User-Google-Token'] = googleUserToken;
    console.log('Attaching 3-legged Google User Token to request header.');
  }
  if (activeHost) {
    headers['X-Office-Host'] = activeHost;
  }

  console.log(`Sending authenticated request to proxy endpoint: ${functionUrl}`);

  const response = await fetch(functionUrl, {
    method: 'POST',
    headers: headers,
    body: JSON.stringify(payload)
  });

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}));
    throw new Error(errorData.detail || errorData.details || errorData.error || `Server returned status ${response.status}`);
  }

  const data = await response.json();
  if (data && data.sessionId && activeHost) {
    SessionManager.registerSessionHost(data.sessionId, activeHost);
  }
  return data;
}

/**
 * Streams a Gemini response in real-time via Server-Sent Events (SSE).
 * Invokes `onChunk({ delta, accumulatedText, status, citations, sessionId, done })` as tokens arrive.
 * Automatically falls back to unary `askGeminiEnterprise` if the streaming endpoint is unreachable.
 */
export async function askGeminiEnterpriseStream(prompt, onChunk, options = {}) {
  const {
    history = [],
    sessionId = null,
    enableGrounding = true,
    userPseudoId = null,
    attachments = null,
    host = null
  } = options;

  const streamUrl = getActiveProxyUrl('stream');
  const userProfile = getUserProfile();

  // 1. Acquire Microsoft Entra ID SSO token
  let authToken = null;
  try {
    authToken = await getOfficeAuthToken();
  } catch (authErr) {
    console.warn('Proceeding without SSO token for stream (server may reject if REQUIRE_ENTRA_AUTH=true):', authErr);
  }

  // 2. Check for 3-legged Google User OAuth token (for Google Drive grounding without DWD)
  const googleUserToken = getGoogleAccessToken();
  const activeHost = host || detectActiveHostName();

  const payload = {
    prompt: prompt,
    history: history,
    enableGrounding: enableGrounding,
    userPseudoId: userPseudoId || userProfile.email || userProfile.user_id || 'office_365_user'
  };
  if (activeHost) {
    payload.host = activeHost;
  }
  if (sessionId) {
    payload.sessionId = sessionId;
  }
  if (attachments && Array.isArray(attachments) && attachments.length > 0) {
    payload.attachments = attachments;
  }

  const headers = {
    'Content-Type': 'application/json',
    'Accept': 'text/event-stream, application/json'
  };
  if (authToken) {
    headers['Authorization'] = `Bearer ${authToken}`;
  }
  if (googleUserToken) {
    headers['X-End-User-Google-Token'] = googleUserToken;
  }
  if (activeHost) {
    headers['X-Office-Host'] = activeHost;
  }

  let accumulatedText = '';
  let finalCitations = [];
  let finalGroundingMetadata = null;
  let finalSessionId = sessionId;
  let finalHistory = history;

  try {
    const response = await fetch(streamUrl, {
      method: 'POST',
      headers: headers,
      body: JSON.stringify(payload)
    });

    if (!response.ok) {
      const errorData = await response.json().catch(() => ({}));
      throw new Error(errorData.detail || errorData.details || errorData.error || `Stream returned status ${response.status}`);
    }

    if (!response.body || typeof response.body.getReader !== 'function') {
      const unaryData = await askGeminiEnterprise(prompt, history, sessionId, enableGrounding, attachments);
      if (typeof onChunk === 'function') {
        onChunk({
          delta: unaryData.result || '',
          accumulatedText: unaryData.result || '',
          citations: unaryData.citations || [],
          sessionId: unaryData.sessionId || sessionId,
          done: true
        });
      }
      return unaryData;
    }

    const reader = response.body.getReader();
    const decoder = new TextDecoder('utf-8');
    let sseBuffer = '';

    while (true) {
      const { value, done } = await reader.read();
      if (done) break;

      sseBuffer += decoder.decode(value, { stream: true });
      const events = sseBuffer.split('\n\n');
      sseBuffer = events.pop() || '';

      for (const evtBlock of events) {
        const lines = evtBlock.split('\n');
        for (const line of lines) {
          const trimmed = line.trim();
          if (!trimmed.startsWith('data:')) continue;
          const jsonStr = trimmed.slice(5).trim();
          if (!jsonStr || jsonStr === '[DONE]') continue;

          try {
            const eventData = JSON.parse(jsonStr);
            if (eventData.error) {
              throw new Error(eventData.details || eventData.error);
            }

            if (eventData.sessionId) {
              finalSessionId = eventData.sessionId;
              if (activeHost) {
                SessionManager.registerSessionHost(finalSessionId, activeHost);
              }
            }
            if (Array.isArray(eventData.citations) && eventData.citations.length > 0) {
              finalCitations = eventData.citations;
            }
            if (eventData.groundingMetadata) {
              finalGroundingMetadata = eventData.groundingMetadata;
            }
            if (Array.isArray(eventData.history)) {
              finalHistory = eventData.history;
            }

            if (typeof eventData.accumulatedText === 'string') {
              accumulatedText = eventData.accumulatedText;
            } else if (typeof eventData.delta === 'string') {
              accumulatedText += eventData.delta;
            } else if (typeof eventData.result === 'string' && eventData.done) {
              accumulatedText = eventData.result;
            }

            if (typeof onChunk === 'function') {
              onChunk({
                delta: eventData.delta || eventData.text || '',
                text: eventData.delta || eventData.text || '',
                accumulatedText: accumulatedText,
                fullText: accumulatedText,
                status: eventData.status || null,
                citations: finalCitations,
                groundingMetadata: finalGroundingMetadata,
                sessionId: finalSessionId,
                done: Boolean(eventData.done),
                isDone: Boolean(eventData.done)
              });
            }
          } catch (parseErr) {
            if (parseErr.message && !parseErr.message.includes('JSON')) {
              throw parseErr;
            }
          }
        }
      }
    }

    return {
      result: accumulatedText,
      citations: finalCitations,
      groundingMetadata: finalGroundingMetadata,
      sessionId: finalSessionId,
      history: finalHistory
    };
  } catch (streamErr) {
    console.warn('SSE stream interrupted or unavailable, falling back to unary endpoint:', streamErr.message);
    if (accumulatedText.length > 0) {
      return {
        result: accumulatedText,
        citations: finalCitations,
        groundingMetadata: finalGroundingMetadata,
        sessionId: finalSessionId,
        history: finalHistory
      };
    }
    const fallbackData = await askGeminiEnterprise(prompt, history, sessionId, enableGrounding, attachments);
    if (typeof onChunk === 'function') {
      onChunk({
        delta: fallbackData.result || '',
        accumulatedText: fallbackData.result || '',
        citations: fallbackData.citations || [],
        groundingMetadata: fallbackData.groundingMetadata || null,
        sessionId: fallbackData.sessionId || sessionId,
        done: true
      });
    }
    return fallbackData;
  }
}

export async function fetchEnterpriseSessions(pageSize = 20, host = null) {
  const baseUrl = getProxyBaseUrl();
  let authToken = null;
  try {
    authToken = await getOfficeAuthToken();
  } catch (e) {
    // Proceed without Entra token if not authenticated
  }
  const googleUserToken = getGoogleAccessToken();
  const activeHost = host || detectActiveHostName();

  const headers = { 'Accept': 'application/json' };
  if (authToken) {
    headers['Authorization'] = `Bearer ${authToken}`;
  }
  if (googleUserToken) {
    headers['X-End-User-Google-Token'] = googleUserToken;
  }
  if (activeHost) {
    headers['X-Office-Host'] = activeHost;
  }

  const hostQuery = activeHost ? `&host=${encodeURIComponent(activeHost)}` : '';
  const res = await fetch(`${baseUrl}/sessions?pageSize=${pageSize}${hostQuery}`, {
    method: 'GET',
    headers: headers
  });
  if (!res.ok) {
    throw new Error(`Failed to fetch sessions from server (${res.status})`);
  }
  const data = await res.json();
  return Array.isArray(data.sessions) ? data.sessions : [];
}

export async function deleteEnterpriseSession(sessionId) {
  if (!sessionId) return false;
  const baseUrl = getProxyBaseUrl();
  let authToken = null;
  try {
    authToken = await getOfficeAuthToken();
  } catch (e) {
    // Proceed without Entra token if not authenticated
  }
  const googleUserToken = getGoogleAccessToken();

  const headers = { 'Content-Type': 'application/json' };
  if (authToken) {
    headers['Authorization'] = `Bearer ${authToken}`;
  }
  if (googleUserToken) {
    headers['X-End-User-Google-Token'] = googleUserToken;
  }

  try {
    const res = await fetch(`${baseUrl}/deleteSession`, {
      method: 'POST',
      headers: headers,
      body: JSON.stringify({ sessionId })
    });
    return res.ok;
  } catch (err) {
    console.warn('Could not delete remote session:', err);
    return false;
  }
}




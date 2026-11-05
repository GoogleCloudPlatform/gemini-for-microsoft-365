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
 * Gemini Backend Proxy for Microsoft 365 (Word, PowerPoint, Excel)
 * 
 * Exclusively powered by Google Cloud Discovery Engine / Gemini Enterprise (:streamAssist).
 * Provides grounded enterprise generative AI inference with server-side session management
 * and Office-native interactive visualizations.
 * 
 * @author Sathya AG, Principal Architect, Google
 */

import functions from '@google-cloud/functions-framework';
import { GoogleAuth, OAuth2Client } from 'google-auth-library';
import { GoogleGenAI } from '@google/genai';
import { execFileSync } from 'child_process';
import crypto from 'crypto';
import express from 'express';
import corsLib from 'cors';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// No deployment URLs are hardcoded: the Cloud Run pattern accepts every URL
// format Cloud Run issues in any project/region (<svc>-<projnum>.<region>.run.app,
// <svc>-<hash>-<rc>.a.run.app, <svc>.run.app). Extra exact origins such as a
// custom domain can be added per deployment via CORS_ALLOWED_ORIGINS (comma-separated).
const CORS_EXTRA_ORIGINS = (process.env.CORS_ALLOWED_ORIGINS || '')
  .split(',')
  .map((o) => o.trim().replace(/\/+$/, ''))
  .filter(Boolean);

const cors = corsLib({
  origin: (origin, callback) => {
    if (!origin) return callback(null, true);
    if (CORS_EXTRA_ORIGINS.includes(origin)) return callback(null, true);
    if (/^https?:\/\/(localhost|127\.0\.0\.1)(:\d+)?$/.test(origin)) return callback(null, true);
    if (/^https?:\/\/([a-zA-Z0-9-]+\.)*(office\.com|officeapps\.live\.com)$/.test(origin)) return callback(null, true);
    if (/^https:\/\/[a-zA-Z0-9-]+(\.[a-zA-Z0-9-]+)?\.run\.app$/.test(origin)) return callback(null, true);
    return callback(new Error('CORS not allowed from origin: ' + origin));
  },
  credentials: true
});

// Auto-load .env file if available (Node 20+)
if (typeof process.loadEnvFile === 'function') {
  const rootEnvPath = path.resolve(__dirname, '..', '.env');
  const localEnvPath = path.resolve(__dirname, '.env');
  if (fs.existsSync(rootEnvPath)) {
    try { process.loadEnvFile(rootEnvPath); } catch (e) {}
  }
  if (fs.existsSync(localEnvPath)) {
    try { process.loadEnvFile(localEnvPath); } catch (e) {}
  }
}

// Ensure GOOGLE_APPLICATION_CREDENTIALS points to an existing file, else clear it to allow Cloud Run metadata service auth
if (process.env.GOOGLE_APPLICATION_CREDENTIALS && !fs.existsSync(process.env.GOOGLE_APPLICATION_CREDENTIALS)) {
  delete process.env.GOOGLE_APPLICATION_CREDENTIALS;
}

// Environment Configuration (Configured via .env or GCP Cloud Run environment variables)
const PROJECT_ID = process.env.PROJECT_ID || process.env.GE_GCP_PROJECT_ID || process.env.GCP_PROJECT_ID || process.env.GOOGLE_CLOUD_PROJECT || '';
const REGION = process.env.REGION || process.env.GCP_REGION || 'us-central1';
const SERVICE_ACCOUNT = process.env.SERVICE_ACCOUNT || '';
const DATASTORE_ID = process.env.DATASTORE_ID || process.env.VERTEX_DATASTORE_ID || process.env.VERTEX_DATASTORE || '';
const IMAGE_MODEL_NAME = process.env.GEMINI_IMAGE_MODEL || 'gemini-2.5-flash-image';
const INTENT_CLASSIFIER_MODEL = process.env.GEMINI_INTENT_MODEL || 'gemini-2.5-flash';

// Gemini Enterprise StreamAssist API Configuration (STRICT: All requests route through StreamAssist)
const BACKEND_MODE = (process.env.BACKEND_MODE || 'streamassist').toLowerCase();
const GE_GCP_PROJECT_ID = process.env.GE_GCP_PROJECT_ID || process.env.GCP_PROJECT_ID || PROJECT_ID;
const GE_GCP_LOCATION = process.env.GE_GCP_LOCATION || process.env.GCP_LOCATION || 'us';
const GCP_LOCATION = GE_GCP_LOCATION;
const STREAM_ASSIST_ENDPOINT_LOCATION = process.env.STREAM_ASSIST_ENDPOINT_LOCATION || GE_GCP_LOCATION || 'us';
const DISCOVERY_ENGINE_HOST = (!STREAM_ASSIST_ENDPOINT_LOCATION || STREAM_ASSIST_ENDPOINT_LOCATION.toLowerCase() === 'global')
  ? 'discoveryengine.googleapis.com'
  : `${STREAM_ASSIST_ENDPOINT_LOCATION}-discoveryengine.googleapis.com`;
const ENTERPRISE_APP_ID = process.env.GEMINI_ENTERPRISE_APP_ID || process.env.VERTEX_DATASTORE_ID || '';
const ENTERPRISE_COLLECTION_ID = process.env.ENTERPRISE_COLLECTION_ID || process.env.GEMINI_ENTERPRISE_COLLECTION_ID || 'default_collection';
const COLLECTION_ID = ENTERPRISE_COLLECTION_ID;
const ENTERPRISE_ASSISTANT_ID = process.env.ENTERPRISE_ASSISTANT_ID || process.env.GEMINI_ENTERPRISE_ASSISTANT_ID || 'default_assistant';
const ASSISTANT_ID = ENTERPRISE_ASSISTANT_ID;
const ALLOW_SERVICE_ACCOUNT_FALLBACK = (process.env.ALLOW_SERVICE_ACCOUNT_FALLBACK || 'true').toLowerCase() === 'true';

const MICROSOFT_ENTRA_APP_ID = process.env.MICROSOFT_ENTRA_APP_ID || '';
const MICROSOFT_ENTRA_TENANT_ID = process.env.MICROSOFT_ENTRA_TENANT_ID || '';
const GOOGLE_OAUTH_CLIENT_ID = process.env.GOOGLE_OAUTH_CLIENT_ID || '';
const REQUIRE_ENTRA_AUTH = (process.env.REQUIRE_ENTRA_AUTH || 'false').toLowerCase() === 'true';
const USER_AUTH_MODE = (process.env.USER_AUTH_MODE || '').toLowerCase();
const ALLOWED_GOOGLE_CLIENT_IDS = [
  GOOGLE_OAUTH_CLIENT_ID,
  ...(process.env.GOOGLE_CLIENT_ID ? process.env.GOOGLE_CLIENT_ID.split(',') : [])
].map(c => c.trim()).filter(Boolean);

const googleAuth = new GoogleAuth({
  scopes: 'https://www.googleapis.com/auth/cloud-platform'
});

let cachedGcloudCliToken = { token: '', expiresAt: 0 };

const auth = {
  async getClient() {
    if (USER_AUTH_MODE === 'local_adc' && !process.env.K_SERVICE && process.env.NODE_ENV !== 'production' && process.env.NODE_ENV !== 'test') {
      return {
        async getAccessToken() {
          const now = Date.now();
          if (cachedGcloudCliToken.token && cachedGcloudCliToken.expiresAt > now) {
            return { token: cachedGcloudCliToken.token };
          }
          try {
            const token = execFileSync('gcloud', ['auth', 'print-access-token'], {
              encoding: 'utf8',
              timeout: 8000,
              stdio: ['ignore', 'pipe', 'ignore']
            }).trim();
            if (token) {
              cachedGcloudCliToken = { token, expiresAt: now + 45 * 60 * 1000 };
              return { token };
            }
          } catch (_) {
            // Fall back to standard GoogleAuth ADC below if gcloud CLI fails
          }
          const adcClient = await googleAuth.getClient();
          return adcClient.getAccessToken();
        }
      };
    }
    return googleAuth.getClient();
  }
};

const oauth2Client = new OAuth2Client();
let msJwksCache = { keys: [], expiresAt: 0 };

async function getMicrosoftJWKS() {
  const now = Date.now();
  if (msJwksCache.keys.length > 0 && msJwksCache.expiresAt > now) {
    return msJwksCache.keys;
  }
  try {
    const res = await fetch('https://login.microsoftonline.com/common/discovery/v2.0/keys');
    if (res.ok) {
      const data = await res.json();
      if (Array.isArray(data.keys)) {
        msJwksCache = { keys: data.keys, expiresAt: now + 3600000 };
        return msJwksCache.keys;
      }
    }
  } catch (e) {
    console.warn('Failed to fetch Microsoft JWKS:', e.message);
  }
  return msJwksCache.keys;
}

/**
 * Object-Level Authorization: Verify session ownership in Discovery Engine (CWE-639 / BOLA mitigation)
 * Fail-Closed: Access is denied unless session existence AND matching ownership are confirmed.
 */
async function verifySessionAccess(fullSessionName, requestingUserId, bearerToken) {
  if (!fullSessionName) return { authorized: false, statusCode: 400, reason: 'Empty session identifier' };
  if (!requestingUserId) return { authorized: false, statusCode: 401, reason: 'Unauthorized: Missing requesting user identity' };

  try {
    const endpointUrl = `https://${DISCOVERY_ENGINE_HOST}/v1/${fullSessionName}`;
    const headers = {
      'Authorization': `Bearer ${bearerToken}`
    };
    if (GE_GCP_PROJECT_ID) {
      headers['x-goog-user-project'] = GE_GCP_PROJECT_ID;
    }
    const res = await fetch(endpointUrl, { method: 'GET', headers });
    if (res.ok) {
      const sessionData = await res.json();
      // Fail-closed: The session MUST possess a userPseudoId matching requestingUserId
      if (!sessionData.userPseudoId || sessionData.userPseudoId !== requestingUserId) {
        console.warn(`[SECURITY_WARN] BOLA/IDOR attempt blocked: User '${requestingUserId}' attempted to access session '${fullSessionName}' belonging to '${sessionData.userPseudoId || 'unassigned'}'`);
        return {
          authorized: false,
          statusCode: 403,
          reason: 'Forbidden: You are not authorized to access sessions belonging to other users.'
        };
      }
      return { authorized: true, session: sessionData };
    } else if (res.status === 404) {
      return { authorized: false, statusCode: 404, reason: 'Session not found or expired.' };
    } else {
      // Fail closed on any other upstream HTTP response (401, 403, 500, etc.)
      console.warn(`[SECURITY_WARN] Upstream Discovery Engine returned HTTP ${res.status} when verifying session access for '${requestingUserId}'.`);
      return {
        authorized: false,
        statusCode: res.status,
        reason: `Upstream error verifying session (${res.status})`
      };
    }
  } catch (err) {
    console.warn(`[SECURITY_WARN] Session verification error (failing closed): ${err.message}`);
    return { authorized: false, statusCode: 500, reason: 'Internal error verifying session access.' };
  }
}

// Cryptographic Authentication for All Protected Routes (CWE-287, CWE-306)
async function verifyAuthentication(req) {
  const authHeader = req.headers['authorization'] || '';
  const userGoogleToken = req.headers['x-end-user-google-token'] || req.headers['x-end-user-token'] || '';

  let token = '';
  if (authHeader.startsWith('Bearer ')) {
    token = authHeader.slice(7).trim();
  } else if (userGoogleToken) {
    token = userGoogleToken.trim();
  }

  // Reject callers without any token (unless in non-production local_adc mode with REQUIRE_ENTRA_AUTH=false)
  if (!token) {
    if (
      process.env.NODE_ENV !== 'production' &&
      !process.env.K_SERVICE &&
      (!PROJECT_ID && !ENTERPRISE_APP_ID || (!REQUIRE_ENTRA_AUTH && USER_AUTH_MODE === 'local_adc'))
    ) {
      return {
        authenticated: true,
        userId: 'local_dev_user',
        email: 'dev@localhost'
      };
    }
    return {
      authenticated: false,
      statusCode: 401,
      error: 'Unauthorized: Authentication required. Bearer authorization token or Google user token must be provided.'
    };
  }

  // A. Google OAuth Access Token (ya29....)
  if (token.startsWith('ya29.')) {
    if (REQUIRE_ENTRA_AUTH) {
      return {
        authenticated: false,
        statusCode: 401,
        error: 'Unauthorized: Microsoft Entra ID authentication is required. Google OAuth tokens are not accepted when REQUIRE_ENTRA_AUTH is active.'
      };
    }

    try {
      const tokenInfo = await oauth2Client.getTokenInfo(token);
      
      // Enforce client ID / audience verification on Google OAuth token (CWE-287)
      const tokenAud = tokenInfo.aud || tokenInfo.azp || '';
      if (ALLOWED_GOOGLE_CLIENT_IDS.length > 0) {
        if (!tokenAud || !ALLOWED_GOOGLE_CLIENT_IDS.includes(tokenAud)) {
          return {
            authenticated: false,
            statusCode: 401,
            error: `Unauthorized: Google OAuth token audience '${tokenAud}' does not match configured authorized client IDs.`
          };
        }
      } else if (process.env.NODE_ENV === 'production' || PROJECT_ID) {
        return {
          authenticated: false,
          statusCode: 401,
          error: 'Unauthorized: Google OAuth access tokens require GOOGLE_OAUTH_CLIENT_ID to be configured on the server.'
        };
      }

      const verifiedUserId = tokenInfo.email || tokenInfo.sub;
      if (!verifiedUserId) {
        return {
          authenticated: false,
          statusCode: 401,
          error: 'Unauthorized: Google OAuth access token lacks valid identity.'
        };
      }

      // Domain validation for Google OAuth accounts (CWE-287)
      const userEmail = (tokenInfo.email || '').toLowerCase();
      const allowedDomains = (process.env.ALLOWED_DOMAINS || process.env.GOOGLE_WORKSPACE_DOMAIN || '')
        .split(',')
        .map(d => d.trim().toLowerCase())
        .filter(Boolean);

      if (allowedDomains.length > 0) {
        const userDomain = userEmail.split('@')[1] || '';
        if (!allowedDomains.includes(userDomain)) {
          return {
            authenticated: false,
            statusCode: 401,
            error: `Unauthorized: User email domain '${userDomain}' is not permitted.`
          };
        }
      } else if (process.env.NODE_ENV === 'production' || PROJECT_ID) {
        const userDomain = userEmail.split('@')[1] || '';
        if (userDomain === 'gmail.com' || userDomain === 'googlemail.com') {
          return {
            authenticated: false,
            statusCode: 401,
            error: 'Unauthorized: Consumer Google accounts (@gmail.com) are not authorized for Enterprise Assistant access.'
          };
        }
      }

      return {
        authenticated: true,
        userId: verifiedUserId,
        email: tokenInfo.email || verifiedUserId
      };
    } catch (err) {
      return {
        authenticated: false,
        statusCode: 401,
        error: `Unauthorized: Google OAuth access token is invalid or expired (${err.message}).`
      };
    }
  }

  // B. Cryptographic JWT Token Verification
  const parts = token.split('.');
  if (parts.length === 3) {
    let header, payload;
    try {
      header = JSON.parse(Buffer.from(parts[0], 'base64url').toString('utf8'));
      payload = JSON.parse(Buffer.from(parts[1], 'base64url').toString('utf8'));
    } catch (e) {
      return {
        authenticated: false,
        statusCode: 401,
        error: 'Unauthorized: Malformed JWT token structure.'
      };
    }

    // 1. Verify expiration
    const nowSec = Math.floor(Date.now() / 1000);
    if (payload.exp && typeof payload.exp === 'number' && payload.exp < nowSec) {
      return {
        authenticated: false,
        statusCode: 401,
        error: 'Unauthorized: Authentication token has expired.'
      };
    }

    const iss = payload.iss || '';

    // 2. Google Cloud IAM ID Token / Cloud Run S2S Token (accounts.google.com)
    if (iss.includes('accounts.google.com')) {
      try {
        const expectedAudiences = [
          GOOGLE_OAUTH_CLIENT_ID,
          process.env.CLOUD_RUN_SERVICE_URL,
          process.env.GEMINI_ENTERPRISE_PROXY_URL,
          process.env.GEMINI_PROXY_URL,
          process.env.SERVICE_URL
        ].filter(Boolean);

        const verifyOptions = { idToken: token };
        if (expectedAudiences.length > 0) {
          verifyOptions.audience = expectedAudiences;
        }

        const ticket = await oauth2Client.verifyIdToken(verifyOptions);
        const googlePayload = ticket.getPayload();
        
        // Strict identity validation: Identity must come from verified token claims (CWE-287)
        const callerEmail = (googlePayload.email || '').toLowerCase();
        const callerSub = googlePayload.sub || '';
        
        // Only trusted internal Cloud Run / GCP Service Accounts belonging to this project or explicitly configured can assert a forwarded end-user identity (CWE-287)
        const allowedServiceAccounts = [
          SERVICE_ACCOUNT,
          ...(process.env.ALLOWED_SERVICE_ACCOUNTS ? process.env.ALLOWED_SERVICE_ACCOUNTS.split(',') : [])
        ].map(s => s.trim().toLowerCase()).filter(Boolean);

        const trustedProjectSuffix = PROJECT_ID ? `@${PROJECT_ID}.iam.gserviceaccount.com` : '';
        const geProjectSuffix = GE_GCP_PROJECT_ID ? `@${GE_GCP_PROJECT_ID}.iam.gserviceaccount.com` : '';

        const isTrustedServiceAccount = Boolean(
          (callerEmail && allowedServiceAccounts.includes(callerEmail)) ||
          (trustedProjectSuffix && callerEmail.endsWith(trustedProjectSuffix.toLowerCase())) ||
          (geProjectSuffix && callerEmail.endsWith(geProjectSuffix.toLowerCase()))
        );

        // When REQUIRE_ENTRA_AUTH is active, only trusted S2S service accounts (e.g. authproxy) may call with Google tokens
        if (REQUIRE_ENTRA_AUTH && !isTrustedServiceAccount) {
          return {
            authenticated: false,
            statusCode: 401,
            error: 'Unauthorized: Microsoft Entra ID authentication is required. Direct Google user ID tokens are not permitted when REQUIRE_ENTRA_AUTH is active.'
          };
        }

        // For non-S2S callers, enforce audience and organization domain restrictions
        if (!isTrustedServiceAccount) {
          if (expectedAudiences.length > 0) {
            const tokenAud = googlePayload.aud || '';
            if (!expectedAudiences.includes(tokenAud)) {
              return {
                authenticated: false,
                statusCode: 401,
                error: `Unauthorized: Token audience '${tokenAud}' does not match authorized audiences.`
              };
            }
          } else if (process.env.NODE_ENV === 'production' || PROJECT_ID) {
            return {
              authenticated: false,
              statusCode: 401,
              error: 'Unauthorized: Google ID tokens require GOOGLE_OAUTH_CLIENT_ID or CLOUD_RUN_SERVICE_URL audience configuration.'
            };
          }

          // Enforce allowed user domain if configured
          const allowedDomains = (process.env.ALLOWED_DOMAINS || process.env.GOOGLE_WORKSPACE_DOMAIN || '')
            .split(',')
            .map(d => d.trim().toLowerCase())
            .filter(Boolean);
          if (allowedDomains.length > 0) {
            const userDomain = callerEmail.split('@')[1] || '';
            if (!allowedDomains.includes(userDomain)) {
              return {
                authenticated: false,
                statusCode: 401,
                error: `Unauthorized: User email domain '${userDomain}' is not permitted.`
              };
            }
          }
        }

        let resolvedUserId = callerEmail || callerSub;
        let resolvedEmail = callerEmail;

        if (isTrustedServiceAccount && req.headers['x-end-user-id']) {
          // Trusted downstream proxy (authproxy) forwarding authenticated end-user
          resolvedUserId = req.headers['x-end-user-id'];
          resolvedEmail = req.headers['x-end-user-email'] || resolvedUserId;
        }

        if (!resolvedUserId) {
          return {
            authenticated: false,
            statusCode: 401,
            error: 'Unauthorized: Google ID token lacks valid subject or email claim.'
          };
        }

        return {
          authenticated: true,
          userId: resolvedUserId,
          email: resolvedEmail
        };
      } catch (err) {
        return {
          authenticated: false,
          statusCode: 401,
          error: `Unauthorized: Invalid Google ID token signature (${err.message}).`
        };
      }
    }

    // 3. Microsoft Entra ID Token (login.microsoftonline.com or sts.windows.net)
    if (iss.includes('login.microsoftonline.com') || iss.includes('sts.windows.net')) {
      try {
        // Enforce audience claim presence and validation (CWE-287)
        const aud = payload.aud;
        if (!aud) {
          return {
            authenticated: false,
            statusCode: 401,
            error: 'Unauthorized: Missing audience claim in Microsoft Entra token.'
          };
        }

        // Enforce audience verification for Microsoft Entra ID tokens (CWE-287)
        if (!MICROSOFT_ENTRA_APP_ID) {
          if (process.env.NODE_ENV === 'production' || PROJECT_ID || REQUIRE_ENTRA_AUTH) {
            return {
              authenticated: false,
              statusCode: 500,
              error: 'Server configuration error: MICROSOFT_ENTRA_APP_ID must be configured to validate Entra ID token audience.'
            };
          }
        } else {
          const allowedAppIds = MICROSOFT_ENTRA_APP_ID.split(',').map(a => a.trim()).filter(Boolean);
          const audMatches = allowedAppIds.some(aid => aud === aid || aud === `api://${aid}`);
          if (!audMatches) {
            return {
              authenticated: false,
              statusCode: 401,
              error: `Unauthorized: Token audience '${aud}' does not match expected application ID.`
            };
          }
        }

        const keys = await getMicrosoftJWKS();
        const signingKey = keys.find(k => k.kid === header.kid);
        if (!signingKey) {
          return {
            authenticated: false,
            statusCode: 401,
            error: `Unauthorized: Unknown signing key 'kid' in Microsoft Entra token.`
          };
        }

        const pubKey = crypto.createPublicKey({ key: signingKey, format: 'jwk' });
        const signedData = `${parts[0]}.${parts[1]}`;
        const isSigValid = crypto.createVerify('RSA-SHA256').update(signedData).verify(pubKey, parts[2], 'base64url');

        if (!isSigValid) {
          return {
            authenticated: false,
            statusCode: 401,
            error: 'Unauthorized: Invalid Microsoft Entra ID token signature.'
          };
        }

        // Enforce strict tenant verification on Microsoft Entra token (CWE-287)
        if (!payload.tid) {
          return {
            authenticated: false,
            statusCode: 401,
            error: 'Unauthorized: Microsoft Entra token lacks tenant ID (tid) claim.'
          };
        }

        const allowedTenants = (MICROSOFT_ENTRA_TENANT_ID || '').split(',').map(t => t.trim()).filter(Boolean);
        if (allowedTenants.length > 0) {
          if (!allowedTenants.includes(payload.tid)) {
            return {
              authenticated: false,
              statusCode: 401,
              error: `Unauthorized: Token issued by untrusted tenant '${payload.tid}'.`
            };
          }
        } else if (process.env.NODE_ENV === 'production' || PROJECT_ID) {
          return {
            authenticated: false,
            statusCode: 401,
            error: 'Unauthorized: MICROSOFT_ENTRA_TENANT_ID must be configured on the server to accept Entra ID tokens.'
          };
        }

        // Derived strictly from cryptographically verified claims, rejecting spoofable client headers
        const verifiedEntraUser = payload.preferred_username || payload.email || payload.upn || payload.sub;
        if (!verifiedEntraUser) {
          return {
            authenticated: false,
            statusCode: 401,
            error: 'Unauthorized: Microsoft Entra token lacks user identity claims.'
          };
        }

        return {
          authenticated: true,
          userId: verifiedEntraUser,
          email: payload.preferred_username || payload.email || payload.upn || verifiedEntraUser
        };
      } catch (err) {
        return {
          authenticated: false,
          statusCode: 401,
          error: `Unauthorized: Microsoft Entra verification failed: ${err.message}`
        };
      }
    }

    // Untrusted issuer
    if (!PROJECT_ID && !ENTERPRISE_APP_ID && process.env.NODE_ENV !== 'production') {
      return {
        authenticated: true,
        userId: payload.preferred_username || payload.email || payload.sub || 'dev_user'
      };
    }

    return {
      authenticated: false,
      statusCode: 401,
      error: `Unauthorized: Untrusted token issuer '${iss}'.`
    };
  }

  return {
    authenticated: false,
    statusCode: 401,
    error: 'Unauthorized: Unrecognized or invalid Bearer token.'
  };
}

if (!PROJECT_ID || !ENTERPRISE_APP_ID) {
  console.warn('⚠️ WARNING: GCP_PROJECT_ID and GEMINI_ENTERPRISE_APP_ID are required for Gemini execution.');
}

// In-memory session store for multi-turn chats
const sessionStore = new Map();

// Global Google Gen AI (Vertex AI) client instance
let genAIClient;
function getGenAIClient() {
  if (!genAIClient) {
    genAIClient = new GoogleGenAI({
      vertexai: true,
      project: PROJECT_ID || '',
      location: REGION,
    });
  }
  return genAIClient;
}

async function generateVertexImage(prompt) {
  try {
    console.log(`Generating image via Vertex AI (${IMAGE_MODEL_NAME}):`, prompt.substring(0, 70) + '...');
    const res = await getGenAIClient().models.generateContent({
      model: IMAGE_MODEL_NAME,
      contents: `Create a professional corporate financial infographic chart reflecting exact data: ${prompt}. Use exact company name and exact numbers provided. DO NOT use generic placeholder names like "Global Innovations Inc.". Flat 2D vector graphic design, crisp typography, clean white background.`,
      config: {
        responseModalities: ['TEXT', 'IMAGE']
      }
    });
    const parts = res?.candidates?.[0]?.content?.parts || [];
    const imgPart = parts.find(p => p.inlineData);
    if (imgPart && imgPart.inlineData?.data) {
      return `data:${imgPart.inlineData.mimeType || 'image/png'};base64,${imgPart.inlineData.data}`;
    }
  } catch (err) {
    console.warn('Vertex image generation error:', err.message);
  }
  return null;
}

const SYSTEM_INSTRUCTION_TEXT = `You are Ask Gemini, a versatile, intelligent corporate AI assistant for Microsoft 365 (Word, PowerPoint, Excel, and Outlook).

CRITICAL FORMATTING & HOST-SPECIFIC RULES:
0. CONVERSATIONAL GREETINGS, META INPUTS & SMALL TALK:
   - When the user's query or instruction is a greeting, pleasantry, capability question, or a bare word/test input (e.g., "hi", "hello", "hey", "good morning", "thanks", "word", "excel", "ppt", "outlook", "test", "what can you do"), respond with a brief, warm, and helpful 1-2 sentence reply explaining how you can assist with their document, presentation, spreadsheet, or email.
   - NEVER summarize, analyze, or dump the contents of the Active Working Draft or Enterprise Datastores when the user only sends a greeting or short conversational/test input! Wait for an explicit question or instruction before analyzing document or datastore content.
1. DOCUMENT STRUCTURE & HEADERS:
   - NEVER output email or memo headers (e.g., "Date:", "To:", "From:", "Subject:", "Memo:", or salutations) when generating standalone Word documents or reports. Start directly with the main title (# Title).
   - When editing, rewriting, shortening, or refining existing text snippets or selected text, output ONLY the revised replacement text. Do NOT add titles, headings, greetings, or conversational commentary.
2. DYNAMIC VISUAL CHARTS & DIAGRAMS:
   - For Microsoft Excel requests: When asked for charts, graphs, or visual data plots, emit a declarative \`\`\`excel-chart JSON specification block. When asked to modify, update, re-title, switch type, rebind series, toggle legend, or format an existing or active chart, emit a declarative \`\`\`excel-chart-edit JSON specification block. DO NOT generate raster images (image:) for Excel spreadsheet data.
     * MULTI-RANGE & TARGETED FIELD SELECTION RULES (CRITICAL):
       1. Always inspect the provided Column Schema and Worksheet Grid to identify the exact columns matching the user request.
       2. DO NOT select the entire worksheet or unrelated columns when the user asks to chart specific fields!
       3. For targeted field charts (e.g., charting a specific metric by a name/category field):
          - Specify "categoriesRange" (or "categoryRange") for the category / label / display name column (e.g. "B1:B50").
          - Specify "seriesRange" (or "valuesRange") for the metric data column(s) (e.g. "E1:E50" or "D1:E50").
       4. CATEGORY TOTALS & AGGREGATIONS (CRITICAL FOR TRANSACTION/LINE-ITEM DATA):
          - When the user asks to chart totals, subtotals, sums, averages, or distributions by category/group from a line-item or transaction table where categories repeat across rows (e.g., 'distribution of total transaction amount for each category'):
            * Set "aggregate": "sum" (or "average", "count") in the \`\`\`excel-chart JSON block.
            * The Excel add-in automatically creates a dynamic, formula-linked summary table (=SUMIF) and builds the native chart from the aggregated unique categories.
       5. For single contiguous tables: Specify "dataRange" (e.g. "A1:C20").
       6. Supported 'type' / 'chartType' values: ColumnClustered, ColumnStacked, BarClustered, BarStacked, Line, LineMarkers, Pie, Doughnut, Area, Scatter, Waterfall, Treemap, Sunburst, Histogram, Funnel.
       7. Example Targeted Field Chart Block with Aggregation:
\`\`\`excel-chart
{
  "type": "BarClustered",
  "categoriesRange": "D1:D50",
  "seriesRange": "E1:E50",
  "aggregate": "sum",
  "title": "Distribution of Total Transaction Amount by Category",
  "hasHeaders": true,
  "position": { "placement": "adjacent", "offsetColumns": 2 }
}
\`\`\`
       8. Example Existing Chart Modification Block:
\`\`\`excel-chart-edit
{
  "targetChart": "active",
  "chartType": "Doughnut",
  "title": "Revenue Distribution",
  "legend": { "visible": true, "position": "Bottom" },
  "dataLabels": { "visible": true, "showPercentage": true, "showCategoryName": true }
}
\`\`\`
   - For Microsoft Word & PowerPoint requests: Whenever the user asks for charts, diagrams, or visual infographics, embed an image using:
     ![Chart Title](image: A modern corporate financial bar chart comparing Google Services ($94.5B) vs Google Cloud ($24.7B) revenue and operating income for Q1 and Q2 2026. Clean white background, modern flat 2D graphic design, crisp typography.)
     OR
     \`\`\`image
     A modern corporate financial bar chart comparing Google Services vs Google Cloud Q1 and Q2 2026 revenue and operating income ($M). Crisp flat 2D vector style, white background.
     \`\`\`
3. EXECUTIVE BULLET POINTS & TYPOGRAPHY:
   - For executive briefings, presentations, and reports, format bullet points with bold lead-ins and highlighted metrics (e.g. * **Google Services ($94.54B):** Reported strong growth driven by Search and YouTube with **$39.54B** in operating income).
   - Use clean, punchy takeaways that C-level executives can scan and absorb immediately.
4. Use standard markdown tables (| Header 1 | Header 2 | ...) for tabular data and executive callouts (> [!NOTE] ...). For markdown tables, always format the separator line properly using standard hyphens (e.g. | --- | --- |) and NEVER escape currency symbols or numbers with backslashes (always use $1428.25 or $500, NEVER \\$1428.25 or \\$500).
   - For Excel Table Insights & Executive Summary Tables: Keep the metric/category column (e.g., "Total Pipeline", "Sarah Jenkins", "Elena Rostova") in Column 1, and the numeric/currency value (e.g., "$1,672,000", "$525,000", "$537,000") clean in its own dedicated numeric Column 2 without embedding long narrative sentences into the numeric cell. Place qualitative notes or status emojis in a separate Trend Status column. This ensures Excel can chart the summary data accurately.
5. EXCEL SPREADSHEET FORMULAS & CALCULATIONS (CRITICAL FOR EXCEL):
   When the user asks for calculations, formulas, Lookups (XLOOKUP, VLOOKUP, INDEX/MATCH), Aggregations (SUMIFS, COUNTIFS, AVERAGEIFS), or Logical functions for Excel:
   - Always output the exact, ready-to-evaluate Excel formula starting with \`=\` inside a dedicated \`\`\`excel code block, for example:
     \`\`\`excel
     =SUMIFS(C:C, D:D, "Approved")
     \`\`\`
   - Always use standard uppercase function names (e.g., \`SUMIFS\`, \`XLOOKUP\`, \`VLOOKUP\`, \`COUNTIFS\`, \`AVERAGEIFS\`, \`IF\`, \`INDEX\`, \`MATCH\`, \`FILTER\`, \`UNIQUE\`, \`LET\`).
   - Use double quotes \`"\` for string literals inside Excel formulas (e.g., \`"Approved"\`, NOT single quotes).
   - Carefully map column letters (e.g. A, B, C, D) and sheet names based on the spreadsheet context provided in the prompt.
   - Follow the formula code block with a clear, concise bulleted breakdown of the formula arguments and how it works.
   - DO NOT output extraneous \`\`\`excel-range-format blocks, formatting recommendations, summary tables, or chart blocks when the user asks for a formula. Keep formula responses strictly focused on the formula and its explanation.
6. OUTLOOK EMAIL INTELLIGENCE (CRITICAL FOR OUTLOOK):
   - When generating email replies or drafts, do NOT include placeholder brackets like [Your Name] or [Insert Date]. Use the known sender/recipient information or keep it natural.
   - For 3-Bullet Thread Briefings: Output exactly 3 high-impact bullets summarizing the core issue, stakes/financial impact, and required decision.
   - For Tone Coaching: Provide explicit analysis of clarity, brevity, and tone risks, followed by distinct Option A (Concise Executive) and Option B (Diplomatic) drafts.
7. DUAL-SOURCE KNOWLEDGE & INTENT-DRIVEN ENTERPRISE GROUNDING:
   - You have access to two complementary knowledge sources when appropriate:
     1. Active Working Draft / Sheet / Slides / Email Thread / Attachment provided in context.
     2. Connected Enterprise Datastores: Accessed via Vertex AI Search Grounding across corporate repositories, financial reports, vendor SLAs, store fleet scorecards, and policies.
   - ALWAYS evaluate the user's intent first:
     * Conversational / Meta / Short Test Inputs (e.g., "hi", "hello", "word", "test", "what can you do?"): Respond warmly and concisely in 1-2 sentences explaining how you can help in their active Microsoft 365 app. Do NOT summarize any draft and do NOT search, cite, or mention datastore files.
     * Active Document Tasks (e.g., summarizing, rewriting, formatting, proofreading, translating, or analyzing the open document/slides/sheet/email): Answer strictly from the provided Active Working Draft or Selected Context.
     * Enterprise Knowledge Lookups (queries seeking company facts, store operational directives, vendor SLAs, financial figures, or policies): Synthesize information across the Active Working Draft (if provided) and the Connected Enterprise Datastores with inline citations [1], [2].
   - CRITICAL FILE ACCESS RULE: When enterprise grounding returns search results or document snippets (such as .docx, .pdf, or .xlsx files from the datastore), use the provided grounding text snippets directly. NEVER attempt to open, download, or read datastore files from the local disk or working environment, and NEVER tell the user that a datastore file "is not present on disk".
8. INLINE CITATIONS & ATTRIBUTION:
   - When citing specific facts, numbers, dates, SLA commitments, or metrics retrieved from the datastores, place clean inline reference numbers [1], [2] directly after the relevant claims or figures (e.g., "Ann Taylor reached 38.4% DTC digital penetration [1] with $3.1M expedited air freight [1]").
9. EXCEL TAXONOMY CATEGORIZATION & DATA CLEANING (CRITICAL FOR EXCEL):
   When the user highlights, clicks Taxonomy Categorizer, or asks to categorize, classify, tag, map, clean, or standardize entries in a spreadsheet or text column (e.g., Opportunity Names, Products, Raw Merchant Names -> Expense Types, Vendors -> Categories, Inquiries -> Intents, Job Roles -> Departments):
   - Identify the primary text/entity column (e.g. Opportunity Name, Merchant, Product, Description) if a whole sheet is provided, or use the selected column.
   - Analyze the raw text entries row-by-row and map each row into standardized, professional taxonomy categories.
   - Always return a declarative \`\`\`excel-taxonomy JSON specification code block.
   - Specify:
     * 'sourceColumn': The source text column letter (e.g. 'B').
     * 'targetColumn': The destination column letter to populate (e.g. 'C' or adjacent column).
     * 'insertColumn': true (inserts a new column to the right).
     * 'header': The title for the new category column header (e.g. 'Opportunity Category' or 'Expense Category').
     * 'startRow': The first data row index (e.g. 2).
     * 'categories': An array of objects with { 'row': number, 'sourceValue': string, 'category': string }.
     * 'taxonomySummary': An object summarizing the count distribution for each category.
   - Provide a clean summary markdown table showing the mapping distribution and key taxonomy insights.
   - Do NOT return generic multi-column metadata dictionaries or abstract catalogs like {"version": "1.0", "taxonomy": [...]}. Always return the row-by-row \`\`\`excel-taxonomy block.
   - Example Taxonomy Categorization Block:
\`\`\`excel-taxonomy
{
  "action": "categorize",
  "sourceColumn": "B",
  "targetColumn": "C",
  "insertColumn": true,
  "header": "Expense Category",
  "startRow": 2,
  "categories": [
    { "row": 2, "sourceValue": "Uber *Trip San Francisco", "category": "Travel / Transportation" },
    { "row": 3, "sourceValue": "AWS Cloud Services", "category": "Cloud & Infrastructure" },
    { "row": 4, "sourceValue": "Sweetgreen SF", "category": "Meals & Entertainment" }
  ],
  "taxonomySummary": {
    "Travel / Transportation": 1,
    "Cloud & Infrastructure": 1,
    "Meals & Entertainment": 1
  }
}
\`\`\``;

// In-memory LRU cache for semantic intent classifications (TTL: 10 minutes, max 500 entries)
const intentClassificationCache = new Map();
const INTENT_CACHE_TTL_MS = 10 * 60 * 1000;
const INTENT_CACHE_MAX_SIZE = 500;

/**
 * Extracts the raw user query and document/selection context flags from an Office add-in prompt.
 */
function extractUserQueryContext(rawPrompt = '', attachments = null) {
  const text = typeof rawPrompt === 'string' ? rawPrompt : '';
  const hasAttachments = Array.isArray(attachments) && attachments.length > 0;
  const hasClosedBookContract = (
    /CRITICAL (?:EXECUTION|CLOSED-BOOK GROUNDING) CONTRACT/i.test(text) ||
    /Constraint:\s*Answer strictly and solely based on the provided Active Document Context/i.test(text) ||
    /\b(?:DECK TO REVISE|CONTENT TO REVISE):/i.test(text)
  );
  const hasSelectionContext = (
    /^(?:Selected Document Context|Selected Slide Context|CRITICAL SLIDE CONTEXT):/im.test(text) ||
    /^Please analyze, summarize, or explain the following selected text:/im.test(text)
  );
  const hasActiveDraft = (
    /^(?:Active Working Draft|Active Document Context|Spreadsheet Context|Presentation Context):/im.test(text)
  );

  let rawUserQuery = text
    .replace(/\s*\[OfficeHost:\s*(?:Word|Excel|PowerPoint|Outlook)\]\s*/gi, '')
    .replace(/\s*\[IntentRouting:\s*[A-Z_]+\]\s*/gi, '')
    .trim();

  const greetingMatch = rawUserQuery.match(/^User Greeting:\s*"([^"]+)"/i);
  if (greetingMatch) {
    rawUserQuery = greetingMatch[1].trim();
  } else if (/^Please analyze, summarize, or explain the following selected text:/i.test(rawUserQuery)) {
    rawUserQuery = 'Analyze, summarize, or explain the selected text';
  } else {
    rawUserQuery = rawUserQuery
      .replace(/^Active (?:Document Context|Working Draft):[\s\S]*?(?:User (?:Instruction|Query|Follow-Up):)\s*/i, '')
      .replace(/^(?:Selected Document Context|Selected Slide Context|Presentation Context|CRITICAL SLIDE CONTEXT):[\s\S]*?(?:User (?:Instruction|Query|Follow-Up):|Task:)\s*/i, '')
      .replace(/^Spreadsheet Context:[\s\S]*?(?:User (?:Instruction|Query|Follow-Up):)\s*/i, '')
      .replace(/^User (?:Query|Instruction|Follow-Up):\s*/i, '')
      .replace(/\n+Grounding & Intent Policy(?:\s*\[Intent:\s*[A-Z_]+\])?:[\s\S]*$/i, '')
      .replace(/\n+Instruction:[\s\S]*$/i, '')
      .replace(/\n+Constraint:[\s\S]*$/i, '')
      .replace(/\n+Formatting:[\s\S]*$/i, '')
      .replace(/\n+CRITICAL INSTRUCTIONS FOR (?:POWERPOINT SLIDE|SINGLE-SLIDE|VISUAL CONTENT)[\s\S]*$/i, '')
      .trim();
  }

  return {
    rawUserQuery,
    hasActiveDraft,
    hasSelectionContext,
    hasClosedBookContract,
    hasAttachments
  };
}

/**
 * Fast, deterministic semantic intent classification used for instant conversational/closed-book
 * patterns and as a reliable fallback if the Vertex AI classifier is unreachable or times out.
 *
 * Intent Categories:
 * - CONVERSATIONAL_OR_GENERAL (enableGrounding: false)
 * - ACTIVE_DOCUMENT_TASK      (enableGrounding: false)
 * - ENTERPRISE_LOOKUP         (enableGrounding: true)
 */
function classifyIntentFallback(contextObj = {}) {
  const {
    rawUserQuery = '',
    hasActiveDraft = false,
    hasSelectionContext = false,
    hasClosedBookContract = false,
    hasAttachments = false
  } = contextObj;

  if (hasClosedBookContract) {
    return {
      intent: 'ACTIVE_DOCUMENT_TASK',
      enableGrounding: false,
      reason: 'closed_book_contract',
      rawUserQuery
    };
  }

  const q = String(rawUserQuery || '').trim();
  const normalized = q.replace(/^["'`]+|["'`]+$/g, '').trim();
  const words = normalized.split(/\s+/).filter(Boolean);

  if (!normalized) {
    return {
      intent: (hasSelectionContext || hasActiveDraft || hasAttachments)
        ? 'ACTIVE_DOCUMENT_TASK'
        : 'CONVERSATIONAL_OR_GENERAL',
      enableGrounding: false,
      reason: 'empty_query',
      rawUserQuery
    };
  }

  // 1. Bare greetings, pleasantries, capability inquiries, or single-word host/test inputs (e.g. "word", "excel", "ppt", "outlook", "test", "hi")
  const isConversationalOrMeta = /^(?:hi|hello|hey|greetings|good\s+(?:morning|afternoon|evening)|howdy|thanks|thank\s+you|yo|sup|ok|okay|yes|no|test|testing|ping|word|excel|ppt|powerpoint|outlook|office|gemini|help|who\s+are\s+you|what\s+can\s+you\s+do(?:\s+for\s+me)?|how\s+can\s+you\s+help(?:\s+me)?|how\s+does\s+this\s+work|how\s+are\s+you)(?:[\s!.,?]*)$/i.test(normalized);
  if (isConversationalOrMeta) {
    return {
      intent: 'CONVERSATIONAL_OR_GENERAL',
      enableGrounding: false,
      reason: 'conversational_or_meta_input',
      rawUserQuery
    };
  }

  // 2. Explicit references to enterprise datastores, company policies, store directives, SLAs, corporate entities, or financial metrics
  const hasExplicitEnterpriseSignals = /\b(?:datastore|data\s*store|knowledge\s+base|enterprise\s+search|company\s+(?:policy|policies|data|records|report|metrics|financials)|store\s+operation|directive|sla|service\s+level|vendor\s+(?:contract|sla|agreement)|scorecard|ann\s+taylor|loft|lane\s+bryant|knitwell|talbots|cymbal|jeansson|fy\s*\d{2,4}|q[1-4]\s+\d{4}|\bq[1-4]\b.*\b(?:revenue|margin|ebitda|sales|penetration|freight)\b|dtc\s+digital|air\s+freight|benchmark\s+against\s+our|according\s+to\s+our\s+company)\b/i.test(normalized);

  // 3. Explicit active-document / selection / spreadsheet / slide / email manipulation tasks
  const hasActiveDocumentTaskSignals = /\b(?:summarize\s+(?:this|the|my|selected|current|active|attached|above)|(?:in|from|based\s+on|according\s+to)\s+(?:this|the\s+active|the\s+current|the\s+selected|the\s+attached|the\s+open)\s+(?:doc|document|draft|paragraph|section|selection|text|slide|deck|presentation|sheet|spreadsheet|table|column|row|email|thread|message|file|attachment)|rewrite|rephrase|paraphrase|proofread|spell\s*check|fix\s+(?:the\s+)?grammar|make\s+(?:it|this|the\s+\w+)\s+(?:shorter|longer|more\s+concise|more\s+formal|professional|clearer|punchier)|translate\s+(?:this|the)|format\s+(?:this|the)|categorize|taxonomy|write\s+a\s+formula|excel\s+formula|vlookup|xlookup|sumifs|countifs|averageifs|extract\s+(?:action\s+items|key\s+risks|takeaways)\s+from\s+(?:this|the)|tone\s+coach|draft\s+a\s+reply\s+to\s+this)\b/i.test(normalized);

  if (hasActiveDocumentTaskSignals && !hasExplicitEnterpriseSignals) {
    return {
      intent: 'ACTIVE_DOCUMENT_TASK',
      enableGrounding: false,
      reason: 'explicit_active_document_task',
      rawUserQuery
    };
  }

  if ((hasSelectionContext || hasAttachments) && !hasExplicitEnterpriseSignals) {
    return {
      intent: 'ACTIVE_DOCUMENT_TASK',
      enableGrounding: false,
      reason: 'selection_or_attachment_task',
      rawUserQuery
    };
  }

  if (hasExplicitEnterpriseSignals) {
    return {
      intent: 'ENTERPRISE_LOOKUP',
      enableGrounding: true,
      reason: 'explicit_enterprise_signal',
      rawUserQuery
    };
  }

  // 4. Single-word inputs without enterprise signals (e.g., "summary", "outline", "draft", or random single words)
  if (words.length === 1) {
    if (hasActiveDraft && /^(?:summarize|summary|outline|proofread|rewrite|analyze|review|risks|actions|takeaways)$/i.test(normalized)) {
      return {
        intent: 'ACTIVE_DOCUMENT_TASK',
        enableGrounding: false,
        reason: 'single_word_document_command',
        rawUserQuery
      };
    }
    return {
      intent: 'CONVERSATIONAL_OR_GENERAL',
      enableGrounding: false,
      reason: 'single_word_non_enterprise_input',
      rawUserQuery
    };
  }

  // 5. General creative/drafting/coding/math requests that do not ask for internal company records
  const isGeneralWritingOrUtility = /^(?:write|draft|compose|create|generate)\s+(?:a|an)\s+(?:generic|sample|template|thank\s+you|apology|welcome|out\s+of\s+office|congratulations|birthday)\b/i.test(normalized);
  if (isGeneralWritingOrUtility) {
    return {
      intent: 'CONVERSATIONAL_OR_GENERAL',
      enableGrounding: false,
      reason: 'general_template_writing',
      rawUserQuery
    };
  }

  // Default for multi-word analytical/factual questions
  return {
    intent: 'ENTERPRISE_LOOKUP',
    enableGrounding: true,
    reason: 'default_enterprise_lookup',
    rawUserQuery
  };
}

/**
 * Hybrid Intent-Based Grounding Classifier (Option 3).
 * Uses a fast deterministic check for unambiguous cases (explicit closed-book, bare greetings/meta words)
 * and Vertex AI Gemini Flash (`INTENT_CLASSIFIER_MODEL`) with in-memory caching for semantic intent routing.
 */
async function classifyGroundingIntent({ prompt, enableGrounding = true, attachments = null, host = null, llmClassifierFn = null }) {
  if (enableGrounding === false) {
    const ctx = extractUserQueryContext(prompt, attachments);
    return {
      intent: 'ACTIVE_DOCUMENT_TASK',
      enableGrounding: false,
      source: 'client_explicit_false',
      reason: 'enableGrounding_false',
      rawUserQuery: ctx.rawUserQuery
    };
  }

  const ctx = extractUserQueryContext(prompt, attachments);
  const fallbackDecision = classifyIntentFallback(ctx);

  // Short-circuit unambiguous deterministic cases (closed-book contracts, empty queries, bare greetings/meta inputs like "word", or selection-only tasks)
  if (
    fallbackDecision.reason === 'closed_book_contract' ||
    fallbackDecision.reason === 'empty_query' ||
    fallbackDecision.reason === 'conversational_or_meta_input' ||
    fallbackDecision.reason === 'single_word_non_enterprise_input' ||
    fallbackDecision.reason === 'single_word_document_command'
  ) {
    return {
      ...fallbackDecision,
      source: 'deterministic_fast_path'
    };
  }

  const normalizedHost = String(host || 'Office').trim();
  const cacheKey = `${normalizedHost}|${ctx.hasActiveDraft}|${ctx.hasSelectionContext}|${ctx.hasAttachments}|${ctx.rawUserQuery.toLowerCase()}`;
  const now = Date.now();
  const cached = intentClassificationCache.get(cacheKey);
  if (cached && cached.expiresAt > now) {
    return { ...cached.decision, source: 'intent_cache' };
  }

  const canRunVertexClassifier = Boolean(
    typeof llmClassifierFn === 'function' ||
    (PROJECT_ID && process.env.NODE_ENV !== 'test' && process.env.DISABLE_LLM_INTENT_CLASSIFIER !== 'true')
  );

  if (canRunVertexClassifier) {
    try {
      const classifierPrompt = `You are an intent router for an enterprise Microsoft ${normalizedHost} AI assistant.
Classify the user's query into EXACTLY ONE of these three intents:
1. "CONVERSATIONAL_OR_GENERAL": Greetings, small talk, single-word or short test inputs (e.g., "word", "test", "hello"), questions about what the assistant can do, or general writing/brainstorming that does NOT seek internal company records, store directives, SLAs, or financial data.
2. "ACTIVE_DOCUMENT_TASK": Requests to summarize, rewrite, edit, format, proofread, translate, chart, or analyze ONLY the currently open ${normalizedHost} document/slide/spreadsheet/email or uploaded attachment, without asking for external company records.
3. "ENTERPRISE_LOOKUP": Queries asking for specific company facts, internal policies, store operation directives, vendor SLAs/contracts, financial metrics, brand performance, or information from connected Enterprise Datastores.

Context flags:
- Office Host: ${normalizedHost}
- Has Active Document Draft: ${ctx.hasActiveDraft}
- Has Highlighted Selection: ${ctx.hasSelectionContext}
- Has Uploaded Attachment: ${ctx.hasAttachments}

User Query:
"""
${ctx.rawUserQuery.substring(0, 1500)}
"""

Respond ONLY with valid JSON matching: {"intent": "CONVERSATIONAL_OR_GENERAL" | "ACTIVE_DOCUMENT_TASK" | "ENTERPRISE_LOOKUP"}`;

      const runClassifier = async () => {
        if (typeof llmClassifierFn === 'function') {
          return await llmClassifierFn(classifierPrompt, ctx);
        }
        const resp = await getGenAIClient().models.generateContent({
          model: INTENT_CLASSIFIER_MODEL,
          contents: classifierPrompt,
          config: {
            temperature: 0,
            maxOutputTokens: 64,
            responseMimeType: 'application/json',
            thinkingConfig: { thinkingBudget: 0 }
          }
        });
        const textPart = resp?.text || resp?.candidates?.[0]?.content?.parts?.[0]?.text || '';
        return JSON.parse(textPart.trim());
      };

      const timeoutPromise = new Promise((_, reject) =>
        setTimeout(() => reject(new Error('Intent classifier timeout (2000ms)')), 2000)
      );

      const parsed = await Promise.race([runClassifier(), timeoutPromise]);
      const classifiedIntent = String(parsed?.intent || '').toUpperCase().trim();
      if (
        classifiedIntent === 'CONVERSATIONAL_OR_GENERAL' ||
        classifiedIntent === 'ACTIVE_DOCUMENT_TASK' ||
        classifiedIntent === 'ENTERPRISE_LOOKUP'
      ) {
        const decision = {
          intent: classifiedIntent,
          enableGrounding: classifiedIntent === 'ENTERPRISE_LOOKUP',
          source: 'vertex_llm_classifier',
          reason: 'llm_semantic_classification',
          rawUserQuery: ctx.rawUserQuery
        };
        if (intentClassificationCache.size >= INTENT_CACHE_MAX_SIZE) {
          const oldestKey = intentClassificationCache.keys().next().value;
          if (oldestKey) intentClassificationCache.delete(oldestKey);
        }
        intentClassificationCache.set(cacheKey, { decision, expiresAt: now + INTENT_CACHE_TTL_MS });
        return decision;
      }
    } catch (err) {
      console.warn(`[INTENT_ROUTER] LLM intent classification fallback (${err.message})`);
    }
  }

  return {
    ...fallbackDecision,
    source: 'semantic_fallback'
  };
}

/**
 * Appends intent-specific response guidance to the prompt sent to Discovery Engine (:streamAssist)
 * so the model's behavior aligns with the classified intent and never attempts to read datastore files from local disk.
 */
function applyIntentGuidanceToPrompt(rawPrompt = '', groundingDecision = {}, normalizedHost = 'Office') {
  if (!rawPrompt || typeof rawPrompt !== 'string') return rawPrompt;
  const hostName = normalizedHost || 'Office';
  const intent = groundingDecision.intent || 'ENTERPRISE_LOOKUP';
  const userQuery = groundingDecision.rawUserQuery || '';

  // Avoid duplicating if already tagged
  if (/\[IntentRouting:\s*[A-Z_]+\]/i.test(rawPrompt)) {
    return rawPrompt;
  }

  let policyDirective = '';
  if (intent === 'CONVERSATIONAL_OR_GENERAL') {
    policyDirective = `\n\nGrounding & Intent Policy [Intent: CONVERSATIONAL_OR_GENERAL]: The user's input ("${userQuery.substring(0, 120)}") is conversational, a greeting, a short test/meta input, or a general capability question—NOT a request to search Enterprise Datastores or summarize a document. Respond naturally, warmly, and concisely (1-3 sentences) explaining how you can help in Microsoft ${hostName} (e.g., drafting, editing, summarizing, analyzing data, or answering enterprise questions). Do NOT summarize any draft and do NOT search, cite, or mention datastore files.\n[IntentRouting: CONVERSATIONAL_OR_GENERAL]`;
  } else if (intent === 'ACTIVE_DOCUMENT_TASK') {
    policyDirective = `\n\nGrounding & Intent Policy [Intent: ACTIVE_DOCUMENT_TASK]: Fulfill the user's request strictly using the provided Active Working Draft, Selected Context, or Attached File in this prompt. Do NOT reference external Enterprise Datastores and do NOT attempt to open or read files from local disk.\n[IntentRouting: ACTIVE_DOCUMENT_TASK]`;
  } else {
    policyDirective = `\n\nGrounding & Intent Policy [Intent: ENTERPRISE_LOOKUP]: The user is asking for enterprise information, metrics, policies, or facts. Synthesize your response using the provided Active Working Draft (if any) and the connected Enterprise Datastores with inline citations [1], [2]. CRITICAL: Use the grounded search snippets directly—NEVER attempt to open, download, or read datastore files (such as .docx, .pdf, or .xlsx) from the local disk or working environment.\n[IntentRouting: ENTERPRISE_LOOKUP]`;
  }

  // Replace any generic client-side Grounding & Intent Policy block with the resolved intent-specific policy
  const cleanedPrompt = rawPrompt.replace(/\n\nGrounding & Intent Policy:[\s\S]*?(?=\nFormatting:|\n\n\[OfficeHost:|$)/i, '');
  return `${cleanedPrompt}${policyDirective}`;
}

// Inlines images into base64 data URIs for native Office insertion
/**
 * SSRF Prevention: Validate remote image URLs against private CIDR ranges, internal hostnames, and non-HTTPS schemes (CWE-918)
 */
function isAllowedImageUrl(rawUrl) {
  try {
    const parsed = new URL(rawUrl);
    if (parsed.protocol !== 'https:') return false;
    
    // Disallow non-standard ports to prevent internal port scanning
    if (parsed.port && parsed.port !== '443') return false;

    const hostname = parsed.hostname.toLowerCase();

    // Block localhost, link-local, cloud metadata, and internal hostnames
    if (
      hostname === 'localhost' ||
      hostname.endsWith('.localhost') ||
      hostname.endsWith('.internal') ||
      hostname.endsWith('.local') ||
      hostname === 'metadata.google.internal'
    ) {
      return false;
    }

    // Block private IP literals (RFC 1918, loopback, link-local)
    if (/^(127\.|10\.|192\.168\.|172\.(1[6-9]|2[0-9]|3[0-1])\.|169\.254\.|0\.)/.test(hostname)) {
      return false;
    }

    // Block IPv6 loopback / unique local / link local
    if (hostname.startsWith('[') || hostname.includes(':')) {
      if (hostname.includes('::1') || hostname.includes('fe80:') || hostname.includes('fc00:')) {
        return false;
      }
    }

    // Allow reputable public image storage domains or explicit cloud assets
    const allowedDomainSuffixes = [
      'storage.googleapis.com',
      'googleusercontent.com',
      'gstatic.com',
      'microsoft.com',
      'windows.net',
      'office.com'
    ];

    const isAllowedDomain = allowedDomainSuffixes.some(domain => 
      hostname === domain || hostname.endsWith(`.${domain}`)
    );

    return isAllowedDomain;
  } catch (e) {
    return false;
  }
}

async function inlineImagesInContent(text) {
  if (!text) return '';
  let processed = text;

  // 1. Process ```image or ```imagen or ```visual blocks
  const codeBlockRegex = /```(?:image|imagen|visual)\s*([\s\S]*?)```/gi;
  const codeMatches = [...processed.matchAll(codeBlockRegex)];
  for (const match of codeMatches) {
    const [fullMatch, promptText] = match;
    const cleanPrompt = promptText.trim();
    if (cleanPrompt) {
      const dataUri = await generateVertexImage(cleanPrompt);
      if (dataUri) {
        const imgTag = `<div style="margin:18px 0; text-align:center;"><img src="${dataUri}" alt="Visual Chart" style="max-width:100%; border-radius:6px; border:1px solid #c7e0f4; box-shadow:0 3px 10px rgba(0,0,0,0.08);" /></div>`;
        processed = processed.replace(fullMatch, imgTag);
      }
    }
  }

  // 2. Process ![alt](image: ...) with full balanced parenthesis support
  const startMarkerRegex = /!\[([^\]]*)\]\((?:image:|image-prompt:|imagen:)\s*/gi;
  let match;
  const itemsToReplace = [];

  while ((match = startMarkerRegex.exec(processed)) !== null) {
    const fullStart = match[0];
    const alt = match[1];
    const startIndex = match.index;
    const contentStartIndex = startIndex + fullStart.length;

    let openCount = 1;
    let i = contentStartIndex;
    while (i < processed.length && openCount > 0) {
      if (processed[i] === '(') openCount++;
      else if (processed[i] === ')') openCount--;
      i++;
    }

    if (openCount === 0) {
      const prompt = processed.substring(contentStartIndex, i - 1).trim();
      const rawMatch = processed.substring(startIndex, i);
      itemsToReplace.push({ rawMatch, alt, prompt });
    }
  }

  for (const item of itemsToReplace) {
    const dataUri = await generateVertexImage(item.prompt);
    if (dataUri) {
      const imgTag = `<div style="margin:18px 0; text-align:center;"><img src="${dataUri}" alt="${escapeXml(item.alt || 'Visual')}" style="max-width:100%; border-radius:6px; border:1px solid #c7e0f4; box-shadow:0 3px 10px rgba(0,0,0,0.08);" /></div>`;
      processed = processed.replace(item.rawMatch, imgTag);
    } else {
      processed = processed.replace(item.rawMatch, '');
    }
  }

  // 3. Fetch remote image URLs (https://...) and convert to base64 data URIs for native Office insertion (SSRF mitigated)
  const remoteImgRegex = /!\[([^\]]*)\]\((https:\/\/[^\s\)]+)\)/gi;
  const remoteMatches = [...processed.matchAll(remoteImgRegex)];
  for (const m of remoteMatches) {
    const [fullMatch, alt, url] = m;
    try {
      if (!isAllowedImageUrl(url)) {
        console.warn('[SECURITY_WARN] Discarded disallowed or internal image URL (SSRF prevention):', url);
        continue;
      }

      const imgRes = await fetch(url, { signal: AbortSignal.timeout(5000) });
      const contentType = (imgRes.headers.get('content-type') || '').toLowerCase();
      if (!contentType.startsWith('image/')) {
        console.warn('[SECURITY_WARN] Discarded non-image MIME type response:', contentType, url);
        continue;
      }

      const contentLength = parseInt(imgRes.headers.get('content-length') || '0', 10);
      if (contentLength > 10 * 1024 * 1024) {
        console.warn('[SECURITY_WARN] Image exceeds 10MB limit:', contentLength, url);
        continue;
      }

      if (imgRes.ok) {
        const buffer = await imgRes.arrayBuffer();
        if (buffer.byteLength > 10 * 1024 * 1024) continue;
        const base64 = Buffer.from(buffer).toString('base64');
        const rawMime = contentType.split(';')[0].trim().toLowerCase();
        // Allowlist accepted MIME types strictly to prevent XSS via content-type header injection
        const mimeMap = {
          'image/png': 'image/png',
          'image/jpeg': 'image/jpeg',
          'image/jpg': 'image/jpeg',
          'image/webp': 'image/webp',
          'image/gif': 'image/gif'
        };
        const mime = mimeMap[rawMime];
        if (!mime) {
          console.warn('[SECURITY_WARN] Discarded image with non-allowlisted MIME type:', rawMime, url);
          continue;
        }
        // Strictly validate base64 encoding
        if (!/^[A-Za-z0-9+/=]+$/.test(base64)) {
          console.warn('[SECURITY_WARN] Discarded image with malformed base64 payload:', url);
          continue;
        }
        const dataUri = escapeXml(`data:${mime};base64,${base64}`);
        const imgTag = `<div style="margin:18px 0; text-align:center;"><img src="${dataUri}" alt="${escapeXml(alt || 'Generated Image')}" style="max-width:100%; border-radius:6px; border:1px solid #c7e0f4; box-shadow:0 3px 10px rgba(0,0,0,0.08);" /></div>`;
        processed = processed.replace(fullMatch, imgTag);
      }
    } catch (fetchErr) {
      console.warn('[INLINE_IMAGES] Could not fetch remote image URL:', url, fetchErr.message);
    }
  }

  // 4. Convert markdown data:image URIs to styled HTML images (strictly validate MIME and base64 charset, escapeXml)
  const dataImgRegex = /!\[([^\]]*)\]\((data:image\/(?:png|jpeg|jpg|webp|gif);base64,[A-Za-z0-9+/=]+)\)/gi;
  const dataMatches = [...processed.matchAll(dataImgRegex)];
  for (const m of dataMatches) {
    const [fullMatch, alt, dataUri] = m;
    const safeDataUri = escapeXml(dataUri);
    const imgTag = `<div style="margin:18px 0; text-align:center;"><img src="${safeDataUri}" alt="${escapeXml(alt || 'Visual')}" style="max-width:100%; border-radius:6px; border:1px solid #c7e0f4; box-shadow:0 3px 10px rgba(0,0,0,0.08);" /></div>`;
    processed = processed.replace(fullMatch, imgTag);
  }

  // 5. Catch any trailing prompt fragments
  processed = processed.replace(/(?:and operating income[^\n]*\.\s*)?Clean white background[^\n]*\)?/gi, '');
  processed = processed.replace(/['"],?data:,?backgroundColor:[^)]+\)\)?/gi, '');

  return processed;
}


let cachedDataStoreIds = null;
let lastDataStoreFetch = 0;
const DATASTORE_CACHE_TTL_MS = 5 * 60 * 1000; // 5 minutes cache

// Helper: Dynamically fetch all datastores attached to the Gemini Enterprise App/Engine
async function getEngineDataStoreIds(accessToken) {
  const now = Date.now();
  if (cachedDataStoreIds && (now - lastDataStoreFetch < DATASTORE_CACHE_TTL_MS)) {
    return cachedDataStoreIds;
  }

  // If user explicitly configured comma-separated IDs via environment, honor it directly
  if (process.env.DATASTORE_ID && process.env.DATASTORE_ID.includes(',')) {
    cachedDataStoreIds = process.env.DATASTORE_ID.split(',').map(s => s.trim()).filter(Boolean);
    return cachedDataStoreIds;
  }

  try {
    const engineUrl = `https://${DISCOVERY_ENGINE_HOST}/v1/projects/${GE_GCP_PROJECT_ID}/locations/${GCP_LOCATION}/collections/${COLLECTION_ID}/engines/${ENTERPRISE_APP_ID}`;
    const res = await fetch(engineUrl, {
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'x-goog-user-project': GE_GCP_PROJECT_ID
      }
    });

    if (res.ok) {
      const data = await res.json();
      if (Array.isArray(data.dataStoreIds) && data.dataStoreIds.length > 0) {
        console.log(`Auto-discovered ${data.dataStoreIds.length} attached datastore(s) for engine ${ENTERPRISE_APP_ID}:`, data.dataStoreIds);
        cachedDataStoreIds = data.dataStoreIds;
        lastDataStoreFetch = now;
        return cachedDataStoreIds;
      }
    } else {
      console.warn(`Could not auto-discover engine datastores (HTTP ${res.status}), using configured default.`);
    }
  } catch (err) {
    console.warn('Error fetching engine datastores:', err.message);
  }

  const fallback = process.env.DATASTORE_ID || process.env.VERTEX_DATASTORE_ID || '';
  return [fallback].filter(Boolean);
}

// Helper: Build dataStoreSpecs array for StreamAssist ToolsSpec with Object-Level Authorization (CWE-639 / BOLA mitigation)
async function buildDataStoreSpecs(configuredDataStores, accessToken) {
  let list = [];
  if (configuredDataStores) {
    const rawList = Array.isArray(configuredDataStores) ? configuredDataStores : String(configuredDataStores).split(',');
    list = rawList.map(s => String(s).trim()).filter(Boolean);
  } else {
    list = await getEngineDataStoreIds(accessToken);
  }

  const validIdRegex = /^[a-zA-Z0-9_-]+$/;
  const canonicalPrefix = `projects/${GE_GCP_PROJECT_ID}/locations/${GCP_LOCATION}/collections/${COLLECTION_ID}/dataStores/`;

  // Object-level authorization: Restrict datastore access strictly to datastores attached to the engine (CWE-639 / BOLA mitigation)
  const engineAllowedIds = await getEngineDataStoreIds(accessToken);
  const configuredAllowed = (DATASTORE_ID || '').split(',').map(s => s.trim()).filter(Boolean);
  const authorizedShortIds = new Set([...engineAllowedIds, ...configuredAllowed]);

  // Fail closed if no datastores are configured or authorized
  if (authorizedShortIds.size === 0) {
    console.warn('[DATASTORE_SECURITY] No authorized datastores available on the engine.');
    return [];
  }

  const validShortIds = [];
  for (const ds of list) {
    // Strictly forbid cross-project, cross-location, or cross-collection resource paths
    if (ds.startsWith('projects/') && !ds.startsWith(canonicalPrefix)) {
      console.warn(`[DATASTORE_SECURITY] Blocked cross-project/cross-collection datastore access attempt: ${ds}`);
      continue;
    }

    const shortId = ds.startsWith(canonicalPrefix) ? ds.slice(canonicalPrefix.length) : ds;
    if (!validIdRegex.test(shortId)) {
      console.warn(`[DATASTORE_SECURITY] Discarded invalid datastore format: ${ds}`);
      continue;
    }

    if (!authorizedShortIds.has(shortId)) {
      console.warn(`[DATASTORE_SECURITY] Blocked unauthorized datastore access (BOLA): ${shortId}`);
      continue;
    }

    validShortIds.push(shortId);
  }

  // Always construct canonical path strictly tied to current project/location/collection
  return validShortIds.map(shortId => ({
    dataStore: `${canonicalPrefix}${shortId}`
  }));
}

// Helper: Escape XML special characters
function escapeXml(unsafe) {
  if (typeof unsafe !== 'string') return String(unsafe || '');
  return unsafe
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

// ==========================================
// Office-Native HTML/CSS Visual Renderers
// 100% Compatible with Word, PPT, and Excel
// ==========================================

// 1. Office-Native Visual Storyboard Table
function renderOfficeStoryboard(title, subtitle, scenes) {
  const count = Math.min(scenes.length, 4);
  let cells = '';

  scenes.slice(0, 4).forEach((scene, index) => {
    const sceneLabel = escapeXml(String(scene.scene || index + 1));
    cells += `
      <td style="width:${Math.floor(100 / count)}%; vertical-align:top; padding:10px 8px; border:1px solid #c7e0f4; background-color:#ffffff;">
        <div style="background-color:#004e8c; color:#ffffff; font-weight:700; font-size:9.5pt; padding:4px 6px; border-radius:3px; text-align:center; margin-bottom:8px;">SCENE ${sceneLabel}</div>
        <div style="font-weight:700; font-size:10pt; color:#201f1e; margin-bottom:4px;">${escapeXml(scene.title || '')}</div>
        ${scene.description ? `<div style="font-size:8.5pt; color:#605e5c; margin-bottom:6px;">${escapeXml(scene.description)}</div>` : ''}
        ${scene.visualPrompt ? `<div style="font-size:8pt; color:#004e8c; background-color:#f0f6fc; padding:5px; border-left:2px solid #004e8c;"><strong>Visual:</strong> ${escapeXml(scene.visualPrompt)}</div>` : ''}
      </td>`;
  });

  return `<table style="width:100%; border-collapse:collapse; margin:16px 0; font-family:'Segoe UI',Calibri,sans-serif; background-color:#f8fafd; border:1px solid #c7e0f4;" cellpadding="0" cellspacing="0">
      <tr>
        <td colspan="${count}" style="padding:12px 14px; border-bottom:1px solid #c7e0f4;">
          <div style="font-size:13pt; font-weight:700; color:#004e8c;">🎨 ${escapeXml(title || 'Visual Storyboard')}</div>
          ${subtitle ? `<div style="font-size:10pt; color:#605e5c;">${escapeXml(subtitle)}</div>` : ''}
        </td>
      </tr>
      <tr>${cells}</tr>
    </table>`;
}

// 2. Office-Native Visual Bar Chart
function renderOfficeBarChart(title, subtitle, data) {
  const maxVal = Math.max(...data.map(d => Number(d.value) || 0), 1);
  const colors = ['#004e8c', '#0f6cbd', '#2b88d8', '#0078d4', '#107c10', '#5c2d91', '#605e5c'];

  let rows = '';
  data.forEach((item, index) => {
    const val = Number(item.value) || 0;
    const pct = Math.max(5, Math.min(100, Math.round((val / maxVal) * 100)));
    const emptyPct = 100 - pct;
    const rawColor = item.color || colors[index % colors.length];
    const safeColor = /^#[0-9a-fA-F]{3,8}$|^[a-zA-Z]{3,20}$/.test(rawColor) ? rawColor : colors[index % colors.length];
    const escapedColor = escapeXml(safeColor);
    const displayVal = item.displayValue || (item.unit ? `${val} ${item.unit}` : (val <= 100 ? `${val}%` : val.toLocaleString()));

    rows += `
      <tr style="border-bottom:1px solid #edebe9;">
        <td style="width:28%; padding:8px 10px; font-weight:600; color:#201f1e; font-size:10pt; vertical-align:middle;">${escapeXml(item.label)}</td>
        <td style="width:56%; padding:8px 10px; vertical-align:middle;">
          <table style="width:100%; border-collapse:collapse; border:none; margin:0;" cellpadding="0" cellspacing="0" border="0">
            <tr style="height:16px;">
              <td bgcolor="${escapedColor}" style="background-color:${escapedColor}; width:${pct}%; height:16px; font-size:1px; line-height:1px; border:none;">&nbsp;</td>
              ${emptyPct > 0 ? `<td bgcolor="#f3f2f1" style="background-color:#f3f2f1; width:${emptyPct}%; height:16px; font-size:1px; line-height:1px; border:none;">&nbsp;</td>` : ''}
            </tr>
          </table>
        </td>
        <td style="width:16%; padding:8px 10px; text-align:right; font-weight:700; color:#004e8c; font-size:10pt; vertical-align:middle;">${escapeXml(displayVal)}</td>
      </tr>`;
  });

  return `<table style="width:100%; border-collapse:collapse; margin:16px 0; font-family:'Segoe UI',Calibri,sans-serif; background-color:#ffffff; border:1px solid #c7e0f4;" cellpadding="0" cellspacing="0">
      <tr>
        <td colspan="3" bgcolor="#f8fafd" style="background-color:#f8fafd; padding:12px 14px; border-bottom:1px solid #c7e0f4;">
          <div style="font-size:12pt; font-weight:700; color:#004e8c;">📊 ${escapeXml(title)}</div>
          ${subtitle ? `<div style="font-size:9.5pt; color:#605e5c; margin-top:2px;">${escapeXml(subtitle)}</div>` : ''}
        </td>
      </tr>
      ${rows}
    </table>`;
}

// 3. Office-Native Visual Process Flow
function renderOfficeProcessFlow(title, subtitle, data) {
  const count = Math.min(data.length, 5);
  let cells = '';

  data.slice(0, 5).forEach((step, i) => {
    const rawColor = i === 0 ? '#004e8c' : (i === count - 1 ? '#107c10' : '#0f6cbd');
    const color = /^#[0-9a-fA-F]{3,8}$|^[a-zA-Z]{3,20}$/.test(rawColor) ? rawColor : '#004e8c';
    const escapedColor = escapeXml(color);
    cells += `
      <td bgcolor="#f8fafd" style="width:${Math.floor(100 / count)}%; padding:10px 8px; vertical-align:top; border:2px solid ${escapedColor}; background-color:#f8fafd; text-align:center;">
        <div style="font-size:9pt; font-weight:700; color:${escapedColor}; margin-bottom:4px;">STEP ${i + 1}</div>
        <div style="font-size:10pt; font-weight:600; color:#201f1e; margin-bottom:4px;">${escapeXml(step.label)}</div>
        ${step.description ? `<div style="font-size:8.5pt; color:#605e5c;">${escapeXml(step.description)}</div>` : ''}
      </td>`;
    if (i < count - 1) {
      cells += `<td style="width:16px; text-align:center; vertical-align:middle; font-size:14pt; color:#004e8c; font-weight:bold;">&rarr;</td>`;
    }
  });

  return `<table style="width:100%; border-collapse:collapse; margin:16px 0; font-family:'Segoe UI',Calibri,sans-serif; background-color:#ffffff;" cellpadding="0" cellspacing="0">
      <tr>
        <td colspan="${count * 2 - 1}" bgcolor="#f8fafd" style="background-color:#f8fafd; padding:12px 14px; border:1px solid #c7e0f4; border-bottom:none;">
          <div style="font-size:12pt; font-weight:700; color:#004e8c;">🔄 ${escapeXml(title)}</div>
          ${subtitle ? `<div style="font-size:9.5pt; color:#605e5c; margin-top:2px;">${escapeXml(subtitle)}</div>` : ''}
        </td>
      </tr>
      <tr>${cells}</tr>
    </table>`;
}

// 4. Office-Native KPI Metric Grid
function renderOfficeKpiGrid(title, subtitle, data) {
  const count = Math.min(data.length, 4);
  let cells = '';

  data.slice(0, 4).forEach(kpi => {
    const rawColor = kpi.trend === 'up' ? '#107c10' : (kpi.trend === 'down' ? '#d13438' : '#004e8c');
    const color = /^#[0-9a-fA-F]{3,8}$|^[a-zA-Z]{3,20}$/.test(rawColor) ? rawColor : '#004e8c';
    const escapedColor = escapeXml(color);
    const trendIcon = kpi.trend === 'up' ? '▲ ' : (kpi.trend === 'down' ? '▼ ' : '');

    cells += `
      <td bgcolor="#f8fafd" style="width:${Math.floor(100 / count)}%; padding:14px 10px; vertical-align:top; border:1px solid #c7e0f4; background-color:#f8fafd; text-align:center;">
        <div style="font-size:9pt; font-weight:600; color:#605e5c; margin-bottom:4px; text-transform:uppercase;">${escapeXml(kpi.label)}</div>
        <div style="font-size:18pt; font-weight:700; color:#004e8c; margin-bottom:4px;">${escapeXml(String(kpi.value))}</div>
        ${kpi.change ? `<div style="font-size:9pt; font-weight:600; color:${escapedColor};">${trendIcon}${escapeXml(kpi.change)}</div>` : ''}
        ${kpi.description ? `<div style="font-size:8pt; color:#605e5c; margin-top:4px;">${escapeXml(kpi.description)}</div>` : ''}
      </td>`;
  });

  return `<table style="width:100%; border-collapse:collapse; margin:16px 0; font-family:'Segoe UI',Calibri,sans-serif; background-color:#ffffff;" cellpadding="0" cellspacing="0">
      <tr>
        <td colspan="${count}" bgcolor="#ffffff" style="padding:10px 14px; border:1px solid #c7e0f4; border-bottom:none;">
          <div style="font-size:12pt; font-weight:700; color:#004e8c;">📈 ${escapeXml(title)}</div>
          ${subtitle ? `<div style="font-size:9.5pt; color:#605e5c; margin-top:2px;">${escapeXml(subtitle)}</div>` : ''}
        </td>
      </tr>
      <tr>${cells}</tr>
    </table>`;
}

function renderChartToImageTag(chartSpec) {
  const title = chartSpec.title || 'Executive Analysis';
  const subtitle = chartSpec.subtitle || '';
  const type = (chartSpec.type || 'bar').toLowerCase();
  const data = chartSpec.data || [];

  if (type === 'storyboard' || chartSpec.scenes) {
    return renderOfficeStoryboard(title, subtitle, chartSpec.scenes || data);
  }
  if (type === 'process' || type === 'flow') {
    return renderOfficeProcessFlow(title, subtitle, data);
  }
  if (type === 'kpi' || type === 'metric' || type === 'metrics') {
    return renderOfficeKpiGrid(title, subtitle, data);
  }
  return renderOfficeBarChart(title, subtitle, data);
}

function processChartsInContent(text) {
  if (!text) return '';
  return text.replace(/```(?:json)?\s*chart\s*([\s\S]*?)```/gi, (match, jsonString) => {
    try {
      const cleanJson = jsonString.trim();
      const chartSpec = JSON.parse(cleanJson);
      return renderChartToImageTag(chartSpec);
    } catch (e) {
      console.warn('Could not parse chart JSON block:', e);
      return '';
    }
  });
}

async function handleGeminiRequest(req, res) {
  // STRICT: Route all requests through Gemini Enterprise StreamAssist API
  return handleGeminiEnterpriseRequest(req, res);
}

function extractImagesFromObject(obj, targetList) {
  if (!obj || typeof obj !== 'object') return;

  // 1. Blob object with base64 data
  if (obj.blob && (obj.blob.data || obj.blob.bytesBase64Encoded)) {
    const mime = obj.blob.mimeType || '';
    if (mime.startsWith('image/')) {
      const b64 = obj.blob.data || obj.blob.bytesBase64Encoded;
      const uri = `data:${mime};base64,${b64}`;
      if (!targetList.includes(uri)) targetList.push(uri);
    }
  }

  // 2. inlineData with base64 data
  if (obj.inlineData && (obj.inlineData.data || obj.inlineData.bytesBase64Encoded)) {
    const mime = obj.inlineData.mimeType || '';
    if (mime.startsWith('image/')) {
      const b64 = obj.inlineData.data || obj.inlineData.bytesBase64Encoded;
      const uri = `data:${mime};base64,${b64}`;
      if (!targetList.includes(uri)) targetList.push(uri);
    }
  }

  // 3. Object with image mimeType and data
  if (obj.mimeType && typeof obj.mimeType === 'string' && obj.mimeType.startsWith('image/')) {
    const b64 = obj.data || obj.bytesBase64Encoded;
    if (b64 && typeof b64 === 'string') {
      const uri = `data:${obj.mimeType};base64,${b64}`;
      if (!targetList.includes(uri)) targetList.push(uri);
    }
  }

  // 4. Object with direct fileUri or imageUri
  if (obj.fileData && obj.fileData.fileUri) {
    if (!targetList.includes(obj.fileData.fileUri)) targetList.push(obj.fileData.fileUri);
  }
  if (obj.imageUri && typeof obj.imageUri === 'string') {
    if (!targetList.includes(obj.imageUri)) targetList.push(obj.imageUri);
  }

  // Recurse into arrays and child objects
  if (Array.isArray(obj)) {
    for (const item of obj) extractImagesFromObject(item, targetList);
  } else {
    for (const key of Object.keys(obj)) {
      if (typeof obj[key] === 'object' && obj[key] !== null) {
        extractImagesFromObject(obj[key], targetList);
      }
    }
  }
}

function extractReplies(chunk) {
  if (chunk.answer?.reply) {
    return [chunk.answer.reply];
  }
  if (Array.isArray(chunk.answer?.replies)) {
    return chunk.answer.replies;
  }
  if (Array.isArray(chunk.replies)) {
    return chunk.replies;
  }
  if (chunk.reply && typeof chunk.reply === 'object') {
    return [chunk.reply];
  }
  return [];
}

function findPendingFiles(obj, targetList, currentSession) {
  if (!obj || typeof obj !== 'object') return;
  if (obj.file && (obj.file.fileId || obj.file.id)) {
    const fId = obj.file.fileId || obj.file.id;
    const mime = obj.file.mimeType || 'image/png';
    if (!targetList.some(item => item.fileId === fId)) {
      targetList.push({ fileId: fId, mimeType: mime, session: currentSession });
    }
  }
  if (Array.isArray(obj)) {
    for (const item of obj) findPendingFiles(item, targetList, currentSession);
  } else {
    for (const key of Object.keys(obj)) {
      if (typeof obj[key] === 'object' && obj[key] !== null) {
        findPendingFiles(obj[key], targetList, currentSession);
      }
    }
  }
}

async function processStreamAssistChunks(parsedChunks, originalSessionId, sessionResetOccurred = false, headers = null) {
  let aggregatedText = '';
  let returnedSessionResource = null;
  const citations = [];
  const seenCitations = new Set();
  const extractedImages = [];
  const thoughtParts = [];
  const pendingFiles = [];

  for (const chunk of parsedChunks) {
    if (chunk.sessionInfo?.session) {
      returnedSessionResource = chunk.sessionInfo.session;
    }

    // Check if image generation tool was invoked
    const invokedTools = chunk.invocationTools || chunk.answer?.invocationTools || [];
    if (Array.isArray(invokedTools) && invokedTools.includes('image_generation')) {
      console.log('[STREAM_ASSIST] Image generation tool invoked by Gemini Enterprise');
    }

    // Comprehensive recursive image and file extraction from the entire chunk
    extractImagesFromObject(chunk, extractedImages);
    findPendingFiles(chunk, pendingFiles, chunk.sessionInfo?.session || returnedSessionResource || originalSessionId);

    // Defensive reply normalization (supports both singular reply and plural replies)
    const replies = extractReplies(chunk);

    for (const reply of replies) {
      const contentObj = reply.groundedContent?.content || reply.content;

      if (contentObj) {
        // Direct text fragment
        if (contentObj.text && !contentObj.thought) {
          aggregatedText += contentObj.text;
        } else if (contentObj.thought && contentObj.text) {
          thoughtParts.push(contentObj.text);
        }

        // Direct blob (Base64 image/media bytes)
        if (contentObj.blob && (contentObj.blob.data || contentObj.blob.bytesBase64Encoded)) {
          const mime = contentObj.blob.mimeType || 'image/png';
          if (mime.startsWith('image/')) {
            const b64 = contentObj.blob.data || contentObj.blob.bytesBase64Encoded;
            const uri = `data:${mime};base64,${b64}`;
            if (!extractedImages.includes(uri)) {
              extractedImages.push(uri);
              console.log(`[STREAM_ASSIST] Found content.blob image artifact (${b64.length} bytes)`);
            }
          }
        }

        // Direct file reference (fileId)
        if (contentObj.file && (contentObj.file.fileId || contentObj.file.id)) {
          const fId = contentObj.file.fileId || contentObj.file.id;
          const mime = contentObj.file.mimeType || 'image/png';
          if (!pendingFiles.some(f => f.fileId === fId)) {
            pendingFiles.push({ fileId: fId, mimeType: mime, session: chunk.sessionInfo?.session || returnedSessionResource || originalSessionId });
          }
        }

        // Content parts (if structured as parts array)
        if (Array.isArray(contentObj.parts) && contentObj.parts.length > 0) {
          for (const part of contentObj.parts) {
            if (part.text && !part.thought) {
              aggregatedText += part.text;
            } else if (part.thought && part.text) {
              thoughtParts.push(part.text);
            } else if (part.inlineData?.data) {
              const mime = part.inlineData.mimeType || 'image/png';
              if (mime.startsWith('image/')) {
                const uri = `data:${mime};base64,${part.inlineData.data}`;
                if (!extractedImages.includes(uri)) extractedImages.push(uri);
                console.log(`[STREAM_ASSIST] Found inlineData image artifact (${part.inlineData.data.length} bytes)`);
              }
            } else if (part.blob && (part.blob.data || part.blob.bytesBase64Encoded)) {
              const mime = part.blob.mimeType || 'image/png';
              if (mime.startsWith('image/')) {
                const b64 = part.blob.data || part.blob.bytesBase64Encoded;
                const uri = `data:${mime};base64,${b64}`;
                if (!extractedImages.includes(uri)) extractedImages.push(uri);
                console.log(`[STREAM_ASSIST] Found part.blob image artifact (${b64.length} bytes)`);
              }
            } else if (part.file && (part.file.fileId || part.file.id)) {
              const fId = part.file.fileId || part.file.id;
              const mime = part.file.mimeType || 'image/png';
              if (!pendingFiles.some(f => f.fileId === fId)) {
                pendingFiles.push({ fileId: fId, mimeType: mime, session: chunk.sessionInfo?.session || returnedSessionResource || originalSessionId });
              }
            } else if (part.fileData?.fileUri) {
              if (!extractedImages.includes(part.fileData.fileUri)) extractedImages.push(part.fileData.fileUri);
            } else if (part.image?.uri || part.image?.url) {
              const uri = part.image.uri || part.image.url;
              if (!extractedImages.includes(uri)) extractedImages.push(uri);
            }
          }
        }
      } else {
        // Fallback to top-level reply text properties
        if (reply.text) {
          aggregatedText += reply.text;
        } else if (reply.replyText) {
          aggregatedText += reply.replyText;
        }
      }

      // 2. Check reply media/images
      const replyMedia = reply.media || reply.images || reply.groundedContent?.media || reply.groundedContent?.images;
      if (replyMedia) {
        const mediaList = Array.isArray(replyMedia) ? replyMedia : [replyMedia];
        for (const m of mediaList) {
          if (typeof m === 'string') {
            if (!extractedImages.includes(m)) extractedImages.push(m);
          } else if (m.bytesBase64Encoded || m.data) {
            const uri = `data:${m.mimeType || 'image/png'};base64,${m.bytesBase64Encoded || m.data}`;
            if (!extractedImages.includes(uri)) extractedImages.push(uri);
          } else if (m.uri || m.url) {
            const uri = m.uri || m.url;
            if (!extractedImages.includes(uri)) extractedImages.push(uri);
          }
        }
      }

      // 3. Grounded citations
      const grounded = reply.groundedContent || {};

      // 3a. Text Grounding Metadata references (Gemini Enterprise StreamAssist format)
      const groundingMeta = grounded.textGroundingMetadata || chunk.answer?.groundedContent?.textGroundingMetadata || chunk.textGroundingMetadata;
      if (groundingMeta && Array.isArray(groundingMeta.references)) {
        for (const ref of groundingMeta.references) {
          const docMeta = ref.documentMetadata || {};
          let uri = docMeta.uri || '';
          let title = docMeta.title || '';
          if ((!title || title === 'Enterprise Data Source') && uri) {
            const parts = uri.split('/');
            title = decodeURIComponent(parts[parts.length - 1]);
          }
          if (!title) title = `Enterprise Document ${citations.length + 1}`;
          const cleanUri = uri ? uri.replace('gs://', '') : '';
          const rawContent = ref.content || '';
          const excerpt = rawContent.replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').trim();

          const key = uri || title;
          if (!seenCitations.has(key)) {
            seenCitations.add(key);
            citations.push({
              index: citations.length + 1,
              title,
              uri,
              cleanUri,
              excerpt
            });
          }
        }
      }

      // 3b. Search Chunk / Web citations fallback
      const searchChunk = grounded.searchChunk || grounded.search_chunk;
      if (searchChunk || grounded.web || reply.replyId || grounded.uri) {
        let uri = grounded.uri || searchChunk?.uri || '';
        let title = grounded.title || searchChunk?.title || '';
        if ((!title || title === 'Enterprise Data Source') && uri) {
          const parts = uri.split('/');
          title = decodeURIComponent(parts[parts.length - 1]);
        }
        if (!title && uri) title = `Enterprise Document ${citations.length + 1}`;
        if (uri || title) {
          const cleanUri = uri ? uri.replace('gs://', '') : '';
          const excerpt = (searchChunk?.chunkText || searchChunk?.text || grounded.text || '').trim();
          const key = uri || title;
          if (!seenCitations.has(key)) {
            seenCitations.add(key);
            citations.push({
              index: citations.length + 1,
              title: title || 'Enterprise Data Source',
              uri,
              cleanUri,
              excerpt
            });
          }
        }
      }
    }

    if (!aggregatedText) {
      const chunkText = chunk.answer?.reply?.groundedContent?.content?.text
        || chunk.answer?.reply?.content?.text
        || chunk.answer?.reply?.text
        || chunk.answer?.replyText 
        || chunk.answer?.text 
        || chunk.answer?.content?.text 
        || chunk.replyText 
        || chunk.text 
        || '';
      if (chunkText) {
        aggregatedText += chunkText;
      }
    }
  }

  // 4. Download any pending fileId resources emitted by Gemini Enterprise
  if (pendingFiles.length > 0 && headers) {
    const downloadHeaders = {
      'Authorization': headers['Authorization'] || headers['authorization'] || '',
      'Accept': '*/*'
    };
    const userProj = headers['X-Goog-User-Project'] || headers['x-goog-user-project'];
    if (userProj) {
      downloadHeaders['X-Goog-User-Project'] = userProj;
    }

    for (const file of pendingFiles) {
      const sessionPath = file.session || returnedSessionResource || originalSessionId;
      if (!sessionPath) continue;
      const regionalHost = DISCOVERY_ENGINE_HOST;
      const encodedId = encodeURIComponent(file.fileId);

      // Multiple candidate endpoints in order of priority:
      // 1. :downloadFile with alt=media (standard Google Cloud API raw media streaming)
      // 2. :downloadFile without alt=media (standard custom method)
      // 3. snake_case file_id variants
      // 4. direct session /files/{fileId} resource path
      const candidates = [
        `https://${regionalHost}/v1alpha/${sessionPath}:downloadFile?fileId=${encodedId}&alt=media`,
        `https://${regionalHost}/v1alpha/${sessionPath}:downloadFile?fileId=${encodedId}`,
        `https://${regionalHost}/v1alpha/${sessionPath}:downloadFile?file_id=${encodedId}&alt=media`,
        `https://${regionalHost}/v1alpha/${sessionPath}:downloadFile?file_id=${encodedId}`,
        `https://${regionalHost}/v1alpha/${sessionPath}/files/${encodedId}?alt=media`,
        `https://${regionalHost}/v1alpha/${sessionPath}/files/${encodedId}`
      ];

      let downloaded = false;
      for (const downloadUrl of candidates) {
        if (downloaded) break;
        try {
          console.log(`[STREAM_ASSIST] Attempting to fetch file '${file.fileId}' from ${downloadUrl}...`);
          const fileRes = await fetch(downloadUrl, {
            method: 'GET',
            headers: downloadHeaders
          });

          const contentType = fileRes.headers.get('content-type') || '';
          console.log(`[STREAM_ASSIST] Response status: ${fileRes.status} ${fileRes.statusText}, Content-Type: ${contentType}`);

          if (!fileRes.ok) {
            const errText = await fileRes.text();
            console.warn(`[STREAM_ASSIST] Candidate failed (HTTP ${fileRes.status}):`, errText.slice(0, 200));
            continue;
          }

          const arrayBuffer = await fileRes.arrayBuffer();
          const buf = Buffer.from(arrayBuffer);
          console.log(`[STREAM_ASSIST] Downloaded payload size: ${buf.length} bytes`);

          if (buf.length === 0) {
            console.log(`[STREAM_ASSIST] Empty body returned (0 bytes), trying next candidate...`);
            continue;
          }

          // Check for image magic bytes
          let detectedMime = null;
          if (buf.length >= 8 && buf[0] === 0x89 && buf[1] === 0x50 && buf[2] === 0x4E && buf[3] === 0x47) {
            detectedMime = 'image/png';
          } else if (buf.length >= 3 && buf[0] === 0xFF && buf[1] === 0xD8 && buf[2] === 0xFF) {
            detectedMime = 'image/jpeg';
          } else if (buf.length >= 6 && buf.toString('ascii', 0, 4) === 'GIF8') {
            detectedMime = 'image/gif';
          } else if (buf.length >= 12 && buf.toString('ascii', 0, 4) === 'RIFF' && buf.toString('ascii', 8, 12) === 'WEBP') {
            detectedMime = 'image/webp';
          }

          if (detectedMime || contentType.startsWith('image/')) {
            const mime = detectedMime || contentType.split(';')[0].trim() || file.mimeType || 'image/png';
            const b64 = buf.toString('base64');
            const uri = `data:${mime};base64,${b64}`;
            if (!extractedImages.includes(uri)) {
              extractedImages.push(uri);
              console.log(`[STREAM_ASSIST] Successfully extracted binary image '${file.fileId}' (${mime}, ${b64.length} base64 chars)`);
            }
            downloaded = true;
            break;
          }

          // Check if payload is JSON
          const rawText = buf.toString('utf-8').trim();
          if (rawText.startsWith('{') || rawText.startsWith('[')) {
            try {
              const fileJson = JSON.parse(rawText);
              console.log(`[STREAM_ASSIST] JSON keys returned:`, Object.keys(fileJson));
              const mime = fileJson.mimeType || file.mimeType || 'image/png';
              const b64 = fileJson.data || fileJson.bytesBase64Encoded || fileJson.fileContents || fileJson.content;
              if (b64) {
                const uri = b64.startsWith('data:') ? b64 : `data:${mime};base64,${b64}`;
                if (!extractedImages.includes(uri)) {
                  extractedImages.push(uri);
                  console.log(`[STREAM_ASSIST] Successfully extracted JSON base64 image '${file.fileId}' (${b64.length} chars)`);
                }
                downloaded = true;
                break;
              }

              // Check if a signed download URL or URI was returned
              const downloadUri = fileJson.downloadUri || fileJson.downloadUrl || fileJson.signedUrl || fileJson.uri || fileJson.url;
              if (downloadUri && typeof downloadUri === 'string') {
                console.log(`[STREAM_ASSIST] Found download URL in JSON: ${downloadUri.slice(0, 100)}... Fetching...`);
                const signedRes = await fetch(downloadUri);
                if (signedRes.ok) {
                  const sBuf = Buffer.from(await signedRes.arrayBuffer());
                  const sMime = signedRes.headers.get('content-type') || mime;
                  const uri = `data:${sMime};base64,${sBuf.toString('base64')}`;
                  if (!extractedImages.includes(uri)) {
                    extractedImages.push(uri);
                    console.log(`[STREAM_ASSIST] Successfully fetched image from signed URL (${sBuf.length} bytes)`);
                  }
                  downloaded = true;
                  break;
                }
              }
            } catch (jsonErr) {
              console.warn(`[STREAM_ASSIST] Could not parse text as JSON:`, jsonErr.message, rawText.slice(0, 200));
            }
          }
        } catch (candErr) {
          console.warn(`[STREAM_ASSIST] Exception querying candidate '${downloadUrl}':`, candErr.message);
        }
      }

      if (!downloaded) {
        console.warn(`[STREAM_ASSIST] Unable to retrieve file content for '${file.fileId}' across all candidates`);
      }
    }
  }

  // 5. Append extracted images as Markdown images if not already embedded
  if (extractedImages.length > 0) {
    if (!aggregatedText.trim()) {
      aggregatedText = "Here is the image generated based on your request:\n\n";
    }
    for (const imgUrl of extractedImages) {
      if (!aggregatedText.includes(imgUrl)) {
        aggregatedText += `\n\n![Generated Image](${imgUrl})\n\n`;
      }
    }
  }

  // Graceful fallback if no direct text was found
  if (!aggregatedText.trim()) {
    if (thoughtParts.length > 0) {
      aggregatedText = thoughtParts.join('\n\n');
    } else {
      console.warn('[STREAM_ASSIST] StreamAssist returned empty text and no images. Parsed chunk count:', parsedChunks.length, 'Chunk sample:', JSON.stringify(parsedChunks).slice(0, 1000));
      aggregatedText = "Hello! How can I help you with your document or presentation today? You can ask me to draft content, summarize key points, build slides, or analyze data.";
    }
  }

  const effectiveSessionResource = returnedSessionResource || (
    !sessionResetOccurred && typeof originalSessionId === 'string' && originalSessionId.includes('/sessions/')
      ? originalSessionId
      : null
  );
  let shortSessionId = effectiveSessionResource || (sessionResetOccurred ? null : originalSessionId);
  if (shortSessionId && shortSessionId.includes('/sessions/')) {
    shortSessionId = shortSessionId.split('/sessions/').pop();
  }

  return {
    resultText: aggregatedText,
    sessionResource: effectiveSessionResource,
    sessionId: shortSessionId,
    citations: citations,
    images: extractedImages
  };
}


async function callStreamAssistAPI({ prompt, sessionId, userId, userPseudoId, userGoogleToken, authMode, attachments, dataStoreId, enableGrounding = true, host = null }) {
  if (!PROJECT_ID) {
    throw new Error('GCP_PROJECT_ID environment variable is required for Gemini execution');
  }
  if (!ENTERPRISE_APP_ID) {
    throw new Error('GEMINI_ENTERPRISE_APP_ID environment variable is required for Gemini execution');
  }

  const activeUserId = userPseudoId || userId || 'office_365_user';
  let bearerToken = null;
  // Which identity actually signs the request. activeUserId is the Microsoft
  // profile email and says nothing about the credential; conflating the two is
  // what made a service-account denial look like an end-user licensing problem.
  let credentialLabel = `end user '${activeUserId}'`;

  if (userGoogleToken) {
    console.log(JSON.stringify({
      severity: 'INFO',
      message: `[AUTH] Using End-User Google Token for Discovery Engine StreamAssist (Mode: ${authMode || 'user_token'}, User: ${activeUserId})`,
      user_id: activeUserId,
      auth_mode: authMode || 'user_token',
      token_source: 'X-End-User-Google-Token'
    }));
    bearerToken = userGoogleToken;
    credentialLabel = `end user '${activeUserId}' via their Google OAuth token`;
  } else {
    if (ALLOW_SERVICE_ACCOUNT_FALLBACK) {
      console.warn(JSON.stringify({
        severity: 'WARNING',
        message: `[AUTH_FALLBACK] No end-user Google token provided for user '${activeUserId}'. ALLOW_SERVICE_ACCOUNT_FALLBACK is enabled. Falling back to Cloud Run Service Account ADC credentials.`,
        user_id: activeUserId,
        fallback_reason: 'MISSING_END_USER_GOOGLE_TOKEN',
        allow_service_account_fallback: true
      }));

      const client = await auth.getClient();
      const accessTokenObj = await client.getAccessToken();
      bearerToken = typeof accessTokenObj === 'string' ? accessTokenObj : accessTokenObj.token;

      // Best effort. The Cloud Run metadata credential does not always expose
      // an address, and not knowing the name must never cost us the request.
      let serviceAccountEmail = '';
      try {
        const creds = await auth.getCredentials();
        serviceAccountEmail = (creds && creds.client_email) || '';
      } catch (credErr) {
        console.warn(`[AUTH_FALLBACK] Could not read the service account address: ${credErr.message}`);
      }
      credentialLabel = serviceAccountEmail
        ? `service account '${serviceAccountEmail}' (ALLOW_SERVICE_ACCOUNT_FALLBACK)`
        : 'the Cloud Run service account (ALLOW_SERVICE_ACCOUNT_FALLBACK)';
    } else {
      console.error(JSON.stringify({
        severity: 'ERROR',
        message: `[AUTH_REJECTED] Request for user '${activeUserId}' rejected: End-user Google token is required to enforce Gemini licensing, and ALLOW_SERVICE_ACCOUNT_FALLBACK is false.`,
        user_id: activeUserId,
        allow_service_account_fallback: false
      }));

      const authErr = new Error(`End-user Google authentication token is required to access Gemini. Service account fallback is disabled.`);
      authErr.statusCode = 403;
      throw authErr;
    }
  }

  const normalizedHost = normalizeOfficeHost(host);
  const groundingDecision = await classifyGroundingIntent({
    prompt,
    enableGrounding,
    attachments,
    host: normalizedHost
  });
  const effectiveGrounding = groundingDecision.enableGrounding;
  const guidedPrompt = applyIntentGuidanceToPrompt(prompt, groundingDecision, normalizedHost || 'Office');

  console.log(JSON.stringify({
    severity: 'INFO',
    message: `[INTENT_ROUTER] StreamAssist grounding decision: intent=${groundingDecision.intent}, enableGrounding=${effectiveGrounding} (source=${groundingDecision.source}, reason=${groundingDecision.reason})`,
    intent: groundingDecision.intent,
    enable_grounding: effectiveGrounding,
    source: groundingDecision.source,
    reason: groundingDecision.reason,
    host: normalizedHost || 'Office'
  }));

  const endpointUrl = `https://${DISCOVERY_ENGINE_HOST}/v1alpha/projects/${GE_GCP_PROJECT_ID}/locations/${GCP_LOCATION}/collections/${ENTERPRISE_COLLECTION_ID}/engines/${ENTERPRISE_APP_ID}/assistants/${ENTERPRISE_ASSISTANT_ID}:streamAssist`;
  const requestBody = {
    query: {
      text: ensurePromptHasHostTag(guidedPrompt, normalizedHost)
    },
    toolsSpec: {
      vertexAiSearchSpec: {}
    }
  };

  if (effectiveGrounding === false) {
    delete requestBody.toolsSpec.vertexAiSearchSpec;
  } else if (dataStoreId) {
    const dataStoreSpecs = await buildDataStoreSpecs(dataStoreId, bearerToken);
    if (dataStoreSpecs.length > 0) {
      requestBody.toolsSpec.vertexAiSearchSpec = {
        dataStoreSpecs: dataStoreSpecs
      };
    }
  }

  // Only enable imageGenerationSpec if the prompt explicitly asks for a raster image/drawing/photo (Word/PPT) and is not an Excel/spreadsheet request
  const lowerPrompt = (prompt || '').toLowerCase();
  const isExplicitImageRequest = (
    /\b(generate|create|render|draw|produce|make|output|insert)\b.*\b(image|picture|photo|photograph|illustration|drawing|logo|visual art)\b/i.test(lowerPrompt) ||
    lowerPrompt.includes("imagegenerationspec") ||
    lowerPrompt.includes("critical instructions for image generation")
  ) && !/\b(excel|worksheet|spreadsheet|formula|cells?|range)\b/i.test(lowerPrompt);

  if (isExplicitImageRequest) {
    requestBody.toolsSpec.imageGenerationSpec = {};
  }
  const validSessionIdRegex = /^[a-zA-Z0-9_-]+$/;
  const canonicalSessionPrefix = `projects/${GE_GCP_PROJECT_ID}/locations/${GCP_LOCATION}/collections/${COLLECTION_ID}/engines/${ENTERPRISE_APP_ID}/sessions/`;
  let isValidSession = false;
  if (sessionId) {
    if (validSessionIdRegex.test(sessionId)) {
      isValidSession = true;
    } else if (sessionId.startsWith(canonicalSessionPrefix)) {
      const rest = sessionId.slice(canonicalSessionPrefix.length);
      isValidSession = validSessionIdRegex.test(rest);
    }
    if (!isValidSession) {
      console.warn(`[SECURITY_WARN] Discarded invalid or cross-engine sessionId: ${sessionId}`);
    }
  }

  let sessionToUse = null;
  if (sessionId && isValidSession) {
    let fullSessionName = sessionId.startsWith('projects/')
      ? sessionId
      : `${canonicalSessionPrefix}${sessionId}`;

    // Verify object-level authorization (BOLA / IDOR mitigation)
    const sessionAuth = await verifySessionAccess(fullSessionName, activeUserId, bearerToken);
    if (!sessionAuth.authorized) {
      if (sessionAuth.statusCode === 403) {
        const forbiddenErr = new Error(sessionAuth.reason || 'Forbidden: Unauthorized session access');
        forbiddenErr.statusCode = 403;
        throw forbiddenErr;
      }
      console.warn(`[SESSION_WARN] Session '${sessionId}' invalid or expired. Creating fresh session.`);
    } else {
      sessionToUse = fullSessionName;
    }
  }

  if (sessionToUse) {
    requestBody.session = sessionToUse;
    if (normalizedHost) {
      recordSessionHost(sessionToUse, normalizedHost);
    }
  } else {
    // Explicitly create session tagged with the userPseudoId for user history, attribution, and context files
    try {
      const sessionCreateUrl = `https://${DISCOVERY_ENGINE_HOST}/v1alpha/projects/${GE_GCP_PROJECT_ID}/locations/${GCP_LOCATION}/collections/${ENTERPRISE_COLLECTION_ID}/engines/${ENTERPRISE_APP_ID}/sessions`;
      const sessionRes = await fetch(sessionCreateUrl, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${bearerToken}`,
          'Content-Type': 'application/json',
          'X-Goog-User-Project': GE_GCP_PROJECT_ID
        },
        body: JSON.stringify({
          userPseudoId: activeUserId,
          ...(normalizedHost ? { displayName: `[${normalizedHost}] Office 365 Session` } : {})
        })
      });
      if (sessionRes.ok) {
        const sessionData = await sessionRes.json();
        if (sessionData.name) {
          requestBody.session = sessionData.name;
          if (normalizedHost) {
            recordSessionHost(sessionData.name, normalizedHost);
          }
          console.log(`[SESSION] Created new session tagged for user '${activeUserId}': ${sessionData.name}`);
        }
      }
    } catch (sessionErr) {
      console.warn(`[SESSION_WARN] Could not pre-create session with userPseudoId '${activeUserId}':`, sessionErr.message);
    }
  }

/**
 * Safely extracts raw text from document attachments (PDF, Word DOCX, Markdown, Text, CSV, JSON)
 */
async function extractDocumentText(att) {
  if (!att || !att.fileContents) return '';
  try {
    const buffer = Buffer.from(att.fileContents, 'base64');
    const mime = (att.mimeType || '').toLowerCase();
    const name = (att.fileName || '').toLowerCase();

    // Plain text / Markdown / JSON / CSV
    if (mime.includes('text') || mime.includes('markdown') || mime.includes('json') || mime.includes('csv') || name.endsWith('.txt') || name.endsWith('.md')) {
      return buffer.toString('utf-8');
    }

    // Word documents (.docx)
    if (mime.includes('word') || name.endsWith('.docx')) {
      const mammoth = await import('mammoth');
      const result = await (mammoth.default || mammoth).extractRawText({ buffer });
      return result.value || '';
    }

    // PDF documents (.pdf)
    if (mime.includes('pdf') || name.endsWith('.pdf')) {
      const pdfModule = await import('pdf-parse');
      const { PDFParse } = pdfModule.default || pdfModule;
      if (PDFParse) {
        const parser = new PDFParse({ data: buffer });
        await parser.load();
        const textResult = await parser.getText();
        return (typeof textResult === 'string' ? textResult : (textResult.text || '')).trim();
      }
    }
  } catch (err) {
    console.warn(`[EXTRACT_TEXT_WARN] Could not extract text from '${att.fileName}':`, err.message);
  }
  return '';
}

  // Handle document attachments:
  // 1. Extract text from documents and inject into query.text so StreamAssist ALWAYS has full grounded context
  // 2. Try Discovery Engine addContextFile across multiple API versions and locations
  const uploadedFileIds = [];
  const extractedDocs = [];

  if (attachments && Array.isArray(attachments) && attachments.length > 0) {
    for (const att of attachments) {
      // 1. Extract text first
      try {
        const text = await extractDocumentText(att);
        if (text && text.trim().length > 0) {
          extractedDocs.push({ fileName: att.fileName, text: text.trim() });
          console.log(`[STREAM_ASSIST] Successfully extracted text from '${att.fileName}' (${text.trim().length} chars)`);
        }
      } catch (extractErr) {
        console.warn(`[STREAM_ASSIST_WARN] Could not extract text from '${att.fileName}':`, extractErr.message);
      }

      // 2. Attempt addContextFile via Discovery Engine if session exists
      if (requestBody.session) {
        const regionalHost = STREAM_ASSIST_ENDPOINT_LOCATION && STREAM_ASSIST_ENDPOINT_LOCATION !== 'global'
          ? `${STREAM_ASSIST_ENDPOINT_LOCATION}-discoveryengine.googleapis.com`
          : 'discoveryengine.googleapis.com';

        const candidateUrls = [
          `https://${regionalHost}/v1alpha/${requestBody.session}:addContextFile`,
          `https://${regionalHost}/v1alpha/projects/${GE_GCP_PROJECT_ID}/locations/${GCP_LOCATION}/collections/${ENTERPRISE_COLLECTION_ID}/engines/${ENTERPRISE_APP_ID}/sessions/-:addContextFile`
        ];

        let fileAdded = false;
        for (const addFileUrl of candidateUrls) {
          if (fileAdded) break;
          try {
            const addRes = await fetch(addFileUrl, {
              method: 'POST',
              headers: {
                'Authorization': `Bearer ${bearerToken}`,
                'Content-Type': 'application/json',
                'X-Goog-User-Project': GE_GCP_PROJECT_ID
              },
              body: JSON.stringify({
                fileName: att.fileName,
                mimeType: att.mimeType || 'application/pdf',
                fileContents: att.fileContents
              })
            });

            if (addRes.ok) {
              const addData = await addRes.json();
              if (addData.fileId) {
                uploadedFileIds.push(addData.fileId);
                if (addData.session) {
                  requestBody.session = addData.session;
                }
                fileAdded = true;
                console.log(`[STREAM_ASSIST] Successfully registered context file '${att.fileName}' -> fileId: ${addData.fileId}, tokenCount: ${addData.tokenCount || 'N/A'}`);
              }
            } else if (addRes.status === 404) {
              // Endpoint does not support addContextFile ("Method not found"). Break early and use reliable text grounding.
              if (process.env.VERBOSE_LOGGING === 'true') {
                console.log(`[STREAM_ASSIST] addContextFile not supported on this engine (404). Proceeding with text grounding.`);
              }
              break;
            } else {
              if (process.env.VERBOSE_LOGGING === 'true') {
                const addErrText = await addRes.text();
                console.log(`[STREAM_ASSIST] addContextFile status ${addRes.status}: ${addErrText.slice(0, 100)}`);
              }
            }
          } catch (attErr) {
            if (process.env.VERBOSE_LOGGING === 'true') {
              console.log(`[STREAM_ASSIST] addContextFile network note: ${attErr.message}`);
            }
          }
        }
      }
    }

    if (uploadedFileIds.length > 0) {
      requestBody.fileIds = uploadedFileIds;
      console.log(`[STREAM_ASSIST] Successfully bound ${uploadedFileIds.length} context fileId(s) to streamAssist request:`, uploadedFileIds);
    } else if (extractedDocs.length > 0) {
      // Fallback: inject extracted text only if addContextFile could not bind fileIds
      const docContextBlocks = extractedDocs.map(d => `--- BEGIN DOCUMENT: ${d.fileName} ---\n${d.text}\n--- END DOCUMENT: ${d.fileName} ---`).join('\n\n');
      requestBody.query.text = `${requestBody.query.text}\n\n[ATTACHED DOCUMENTS CONTEXT]\n${docContextBlocks}\n[END ATTACHED DOCUMENTS CONTEXT]`;
      console.log(`[STREAM_ASSIST] Fallback: Grounded query.text with ${extractedDocs.length} attached document(s) (${docContextBlocks.length} chars total)`);
    }
  }

  const headers = {
    'Authorization': `Bearer ${bearerToken}`,
    'Content-Type': 'application/json',
    'X-Goog-User-Project': GE_GCP_PROJECT_ID
  };

  console.log(`Calling StreamAssist API (${endpointUrl})... Session: ${requestBody.session || 'NEW'} Grounding: ${enableGrounding !== false}`);

  let sessionResetOccurred = false;

  let apiRes = await fetch(endpointUrl, {
    method: 'POST',
    headers: headers,
    body: JSON.stringify(requestBody)
  });

  if (!apiRes.ok) {
    let errText = await apiRes.text();
    
    // 1. Auto-recover if session overflowed token limit (PROMPT_TOO_LARGE) or session ownership conflict occurs
    if (requestBody.session && (
      (apiRes.status === 400 && (errText.includes('PROMPT_TOO_LARGE') || errText.includes('INVALID_ARGUMENT'))) ||
      (apiRes.status === 403 && errText.includes('Session is not owned'))
    )) {
      sessionResetOccurred = true;
      console.warn(JSON.stringify({
        severity: 'WARNING',
        message: `[SESSION_RECOVERY] Discovery Engine session '${requestBody.session}' failed (${errText.includes('PROMPT_TOO_LARGE') ? 'PROMPT_TOO_LARGE token overflow' : 'session error'}). Retrying automatically with a fresh clean session...`,
        user_id: activeUserId
      }));
      delete requestBody.session;
      apiRes = await fetch(endpointUrl, {
        method: 'POST',
        headers: headers,
        body: JSON.stringify(requestBody)
      });
      if (!apiRes.ok) {
        errText = await apiRes.text();
      }
    }

    // 2. Auto-recover if toolsSpec (e.g. vertexAiSearchSpec / imageGenerationSpec) causes downstream timeout/error
    if (!apiRes.ok && (apiRes.status === 499 || (apiRes.status >= 400 && requestBody.toolsSpec))) {
      console.warn(JSON.stringify({
        severity: 'WARNING',
        message: `[TOOLS_SPEC_FALLBACK] StreamAssist failed with HTTP ${apiRes.status} (details: ${errText.substring(0, 100)}). Retrying with default assistant tools configuration...`,
        user_id: activeUserId
      }));
      delete requestBody.toolsSpec;
      apiRes = await fetch(endpointUrl, {
        method: 'POST',
        headers: headers,
        body: JSON.stringify(requestBody)
      });
      if (!apiRes.ok) {
        errText = await apiRes.text();
      }
    }

    // 3. Auto-recover if end-user token lacked Gemini Enterprise license (LICENSE_WITHOUT_SUBSCRIPTION_TIER)
    if (!apiRes.ok && ALLOW_SERVICE_ACCOUNT_FALLBACK && errText.includes('LICENSE_WITHOUT_SUBSCRIPTION_TIER')) {
      console.warn(JSON.stringify({
        severity: 'WARNING',
        message: `[LICENSE_FALLBACK] User token lacks Gemini Enterprise license. Falling back to local Service Account / ADC credentials...`,
        user_id: activeUserId
      }));
      const client = await auth.getClient();
      const accessTokenObj = await client.getAccessToken();
      const adcBearerToken = typeof accessTokenObj === 'string' ? accessTokenObj : accessTokenObj.token;
      headers['Authorization'] = `Bearer ${adcBearerToken}`;
      apiRes = await fetch(endpointUrl, {
        method: 'POST',
        headers: headers,
        body: JSON.stringify(requestBody)
      });
      if (!apiRes.ok) {
        errText = await apiRes.text();
      }
    }

    if (!apiRes.ok) {
      console.error(JSON.stringify({
        severity: 'ERROR',
        message: `StreamAssist API call failed with HTTP ${apiRes.status}`,
        status_code: apiRes.status,
        user_id: activeUserId,
        credential: credentialLabel,
        target_project: PROJECT_ID,
        error_detail: errText
      }));

      if (apiRes.status === 403) {
        const forbiddenErr = new Error(
          `Google Cloud Discovery Engine rejected the request (HTTP 403). ` +
          `The call was signed by ${credentialLabel}, so that is the principal that was ` +
          `denied - not necessarily '${activeUserId}'. It needs ` +
          `'discoveryengine.assistants.assist' (role roles/discoveryengine.user) on engine ` +
          `'${ENTERPRISE_APP_ID}' in project '${PROJECT_ID}', plus an active Gemini ` +
          `Enterprise license if it is a human user. Details: ${errText}`
        );
        forbiddenErr.statusCode = 403;
        throw forbiddenErr;
      }

      const streamErr = new Error(`StreamAssist API returned HTTP ${apiRes.status}: ${errText}`);
      streamErr.statusCode = apiRes.status;
      throw streamErr;
    }
  }

  const rawResponseBody = await apiRes.text();
  let parsedChunks = [];

  try {
    const data = JSON.parse(rawResponseBody);
    parsedChunks = Array.isArray(data) ? data : [data];
  } catch (e) {
    const lines = rawResponseBody.split('\n').filter(l => l.trim().length > 0);
    for (const line of lines) {
      try {
        parsedChunks.push(JSON.parse(line));
      } catch (err) {
        console.warn('Could not parse streaming line chunk:', line.substring(0, 80));
      }
    }
  }

  const processedResult = await processStreamAssistChunks(parsedChunks, requestBody.session || sessionId, sessionResetOccurred, headers);
  if (normalizedHost && processedResult.sessionId) {
    recordSessionHost(processedResult.sessionId, normalizedHost);
  }
  return processedResult;
}

async function handleGeminiEnterpriseRequest(req, res) {
  cors(req, res, async () => {
    try {
      if (req.method !== 'POST') {
        return res.status(405).json({ error: 'Method Not Allowed' });
      }

      // Enforce cryptographic authentication (CWE-306)
      const auth = await verifyAuthentication(req);
      if (!auth.authenticated) {
        return res.status(auth.statusCode || 401).json({ error: auth.error });
      }

      const { prompt, sessionId, attachments, dataStoreId, enableGrounding, host } = req.body;
      const effectiveHost = normalizeOfficeHost(host || req.headers['x-office-host']);
      // Object-Level Authorization: User identity must strictly come from cryptographically verified auth context (CWE-639 / BOLA mitigation)
      const userId = auth.userId;
      const userPseudoId = auth.userId;
      const endUserName = req.body.authenticatedUser?.name || req.headers['x-end-user-name'] || '';
      const rawUserGoogleToken = req.headers['x-end-user-google-token'] || req.headers['x-end-user-token'] || '';
      if (rawUserGoogleToken) {
        cacheOauth3GoogleToken(auth, rawUserGoogleToken, 3600);
      }
      const userGoogleToken = rawUserGoogleToken || getCachedOauth3GoogleToken(auth).token || '';
      const authMode = req.headers['x-user-auth-mode'] || '';

      if (!prompt) {
        return res.status(400).json({ error: 'Prompt is required' });
      }

      console.log(`Processing Gemini request (Mode: ${BACKEND_MODE})... Authenticated User: ${userPseudoId} (${endUserName || 'Corporate User'}, AuthMode: ${authMode || 'default'}, Grounding: ${enableGrounding !== false}, Attachments: ${attachments ? attachments.length : 0})`);

      if (BACKEND_MODE === 'streamassist' && ENTERPRISE_APP_ID) {
        try {
          const streamAssistResult = await callStreamAssistAPI({
            prompt,
            sessionId,
            userId,
            userPseudoId,
            userGoogleToken,
            authMode,
            attachments,
            dataStoreId,
            enableGrounding,
            host: effectiveHost
          });
          const rawText = streamAssistResult.resultText || '';
          const chartProcessedText = processChartsInContent(rawText);
          const processedText = await inlineImagesInContent(chartProcessedText);

          return res.status(200).json({
            result: processedText,
            sessionId: streamAssistResult.sessionId,
            sessionResource: streamAssistResult.sessionResource,
            host: effectiveHost,
            citations: streamAssistResult.citations,
            images: streamAssistResult.images || [],
            backendMode: 'streamassist',
            backendEngine: 'gemini-enterprise-discovery-engine',
            turn: {
              userPrompt: prompt,
              assistantText: processedText
            }
          });
        } catch (streamAssistErr) {
          console.error('StreamAssist execution failed:', streamAssistErr.message);
          return res.status(streamAssistErr.statusCode || 500).json({
            error: 'Gemini StreamAssist failed',
            details: streamAssistErr.message,
            statusCode: streamAssistErr.statusCode || 500
          });
        }
      }

      return res.status(400).json({
        error: 'Invalid Configuration',
        details: 'BACKEND_MODE must be streamassist and GEMINI_ENTERPRISE_APP_ID must be configured.'
      });
    } catch (error) {
      console.error('Error in Gemini handler:', error);
      let details = error.message;
      if (typeof details === 'string' && (details.includes('invalid_grant') || details.includes('invalid_rapt') || details.includes('reauth related error'))) {
        details = 'Google Cloud Authentication Expired (invalid_grant/invalid_rapt). Run "gcloud auth application-default login" to re-authenticate.';
      }
      return res.status(500).json({
        error: 'Failed to process Gemini request',
        details: details,
      });
    }
  });
}

// Cross-Surface OAuth3 User Google Token Cache (keyed by verified Entra ID userId / email)
const oauth3UserTokenCache = new Map();

// Per-Session Office Surface Attribution Cache (survives Discovery Engine displayName overwrites)
const SESSION_HOST_CACHE_FILE = '/tmp/gemini_office_session_hosts.json';
const sessionHostCache = new Map();

function normalizeOfficeHost(rawHost) {
  if (!rawHost || typeof rawHost !== 'string') return null;
  const lower = rawHost.trim().toLowerCase();
  if (lower === 'word') return 'Word';
  if (lower === 'excel') return 'Excel';
  if (lower === 'powerpoint' || lower === 'ppt') return 'PowerPoint';
  if (lower === 'outlook') return 'Outlook';
  return null;
}

try {
  if (fs.existsSync(SESSION_HOST_CACHE_FILE)) {
    const raw = JSON.parse(fs.readFileSync(SESSION_HOST_CACHE_FILE, 'utf8'));
    if (raw && typeof raw === 'object') {
      for (const [k, v] of Object.entries(raw)) {
        const norm = normalizeOfficeHost(v);
        if (norm) sessionHostCache.set(k, norm);
      }
    }
  }
} catch (_) {
  // Ignore cache load errors
}

function recordSessionHost(sessionIdOrName, rawHost) {
  const normHost = normalizeOfficeHost(rawHost);
  if (!sessionIdOrName || !normHost) return;
  const shortId = String(sessionIdOrName).includes('/sessions/')
    ? String(sessionIdOrName).split('/sessions/').pop()
    : String(sessionIdOrName).trim();
  if (!shortId) return;
  if (sessionHostCache.get(shortId) === normHost) return;
  sessionHostCache.set(shortId, normHost);
  try {
    const serialized = Object.fromEntries(sessionHostCache.entries());
    fs.writeFileSync(SESSION_HOST_CACHE_FILE, JSON.stringify(serialized), 'utf8');
  } catch (_) {
    // Non-fatal in read-only environments
  }
}

function getRecordedSessionHost(sessionIdOrName) {
  if (!sessionIdOrName) return null;
  const shortId = String(sessionIdOrName).includes('/sessions/')
    ? String(sessionIdOrName).split('/sessions/').pop()
    : String(sessionIdOrName).trim();
  return sessionHostCache.get(shortId) || null;
}

function ensurePromptHasHostTag(promptText, rawHost) {
  const normHost = normalizeOfficeHost(rawHost);
  if (!normHost || !promptText || typeof promptText !== 'string') return promptText;
  if (/\[OfficeHost:\s*(?:Word|Excel|PowerPoint|Outlook)\]/i.test(promptText)) {
    return promptText;
  }
  return `${promptText}\n\n[OfficeHost: ${normHost}]`;
}

function getUserCacheKeys(authObj) {
  const keys = [];
  if (authObj?.userId) keys.push(String(authObj.userId).trim().toLowerCase());
  if (authObj?.email) {
    const emailKey = String(authObj.email).trim().toLowerCase();
    if (emailKey && !keys.includes(emailKey)) keys.push(emailKey);
  }
  return keys;
}

function cacheOauth3GoogleToken(authObj, googleToken, expiresIn = 3600) {
  const cleanToken = String(googleToken || '').trim();
  const keys = getUserCacheKeys(authObj);
  if (keys.length === 0) return;

  if (!cleanToken) {
    for (const k of keys) oauth3UserTokenCache.delete(k);
    return;
  }

  const nowSec = Math.floor(Date.now() / 1000);
  const ttlSec = Math.max(Math.min(parseInt(expiresIn, 10) || 3600, 7200) - 180, 60);
  const entry = { token: cleanToken, expiresAt: nowSec + ttlSec };
  for (const k of keys) {
    oauth3UserTokenCache.set(k, entry);
  }
}

function getCachedOauth3GoogleToken(authObj) {
  const nowSec = Math.floor(Date.now() / 1000);
  for (const k of getUserCacheKeys(authObj)) {
    const entry = oauth3UserTokenCache.get(k);
    if (entry && entry.token && entry.expiresAt > nowSec) {
      return { token: entry.token, remainingSec: Math.max(entry.expiresAt - nowSec, 60) };
    }
    if (entry && entry.expiresAt <= nowSec) {
      oauth3UserTokenCache.delete(k);
    }
  }
  return { token: '', remainingSec: 0 };
}

// Session Management Handlers
async function resolveSessionBearerToken(userGoogleToken = '', activeUserId = 'office_365_user') {
  if (userGoogleToken) {
    return userGoogleToken;
  }
  const cached = getCachedOauth3GoogleToken({ userId: activeUserId, email: activeUserId });
  if (cached.token) {
    return cached.token;
  }
  if (ALLOW_SERVICE_ACCOUNT_FALLBACK) {
    const client = await auth.getClient();
    const accessTokenObj = await client.getAccessToken();
    return typeof accessTokenObj === 'string' ? accessTokenObj : accessTokenObj.token;
  }
  const authErr = new Error(`End-user Google authentication token is required to access Gemini sessions. Service account fallback is disabled.`);
  authErr.statusCode = 403;
  throw authErr;
}

function cleanStoredPromptText(rawText = '') {
  if (!rawText || typeof rawText !== 'string') return '';
  return rawText
    .replace(/\s*\[OfficeHost:\s*(?:Word|Excel|PowerPoint|Outlook)\]\s*/gi, '')
    .replace(/\s*\[IntentRouting:\s*[A-Z_]+\]\s*/gi, '')
    .replace(/^User Greeting:\s*"([^"]+)"[\s\S]*$/i, '$1')
    .replace(/^Active (?:Document Context|Working Draft):[\s\S]*?(?:User (?:Instruction|Query|Follow-Up):)\s*/i, '')
    .replace(/^(?:Selected Document Context|Selected Slide Context|Presentation Context|CRITICAL SLIDE CONTEXT):[\s\S]*?(?:User (?:Instruction|Query|Follow-Up):|Task:)\s*/i, '')
    .replace(/^Spreadsheet Context:[\s\S]*?(?:User (?:Instruction|Query|Follow-Up):)\s*/i, '')
    .replace(/^User (?:Query|Instruction|Follow-Up):\s*/i, '')
    .replace(/\n+Grounding & Intent Policy(?:\s*\[Intent:\s*[A-Z_]+\])?:[\s\S]*$/i, '')
    .replace(/\n+Instruction:[\s\S]*$/i, '')
    .replace(/\n+Constraint:[\s\S]*$/i, '')
    .replace(/\n+Formatting:[\s\S]*$/i, '')
    .trim();
}

function inferSessionHost(sessionObj = {}) {
  const shortId = sessionObj.name
    ? String(sessionObj.name).split('/sessions/').pop()
    : (sessionObj.sessionId || '');
  const recordedHost = getRecordedSessionHost(shortId);
  if (recordedHost) {
    return recordedHost;
  }

  const rawDisplayName = String(sessionObj.displayName || '');
  const prefixMatch = rawDisplayName.match(/^\[(Word|Excel|PowerPoint|Outlook)\]/i);
  if (prefixMatch) {
    const matched = normalizeOfficeHost(prefixMatch[1]);
    if (matched) {
      if (shortId) recordSessionHost(shortId, matched);
      return matched;
    }
  }

  const turns = Array.isArray(sessionObj.turns) ? sessionObj.turns : [];
  const rawTurnTexts = [];
  for (const turn of turns) {
    if (turn?.query?.text) rawTurnTexts.push(turn.query.text);
    const rawReply =
      turn?.assistAnswer?.replies?.[0]?.groundedContent?.content?.text ||
      turn?.reply?.reply ||
      turn?.reply?.summary?.summaryText ||
      turn?.answer ||
      '';
    if (typeof rawReply === 'string' && rawReply) rawTurnTexts.push(rawReply);
  }

  const combinedRaw = [rawDisplayName, ...rawTurnTexts].join('\n');
  if (!combinedRaw.trim()) return null;

  const explicitHostTagMatch = combinedRaw.match(/\[OfficeHost:\s*(Word|Excel|PowerPoint|Outlook)\]/i);
  if (explicitHostTagMatch) {
    const tagged = normalizeOfficeHost(explicitHostTagMatch[1]);
    if (tagged) {
      if (shortId) recordSessionHost(shortId, tagged);
      return tagged;
    }
  }

  const greetingHostMatch = combinedRaw.match(/assist with their (Word|Excel|PowerPoint|Outlook) document today/i);
  if (greetingHostMatch) {
    const gHost = normalizeOfficeHost(greetingHostMatch[1]);
    if (gHost) {
      if (shortId) recordSessionHost(shortId, gHost);
      return gHost;
    }
  }

  if (
    /```excel-(?:chart|formulas|taxonomy|range-format)\b/i.test(combinedRaw) ||
    /\b(?:Selected Spreadsheet Data:|Spreadsheet Context:|Active Worksheet:|Excel Formula Architect|Native Chart Builder|Formula Generator|Formula Builder|Table Insights|spreadsheet data|spreadsheet schema|from this spreadsheet|from this sheet)\b/i.test(combinedRaw) ||
    /(?:^|\n)\s*(?:Worksheet:\s*"|Analyzed Range:\s*|Column Schema:\s*)/i.test(combinedRaw) ||
    /\b(?:XLOOKUP|VLOOKUP|SUMIFS|COUNTIFS|AVERAGEIFS|INDEX\/MATCH)\b/.test(combinedRaw)
  ) {
    if (shortId) recordSessionHost(shortId, 'Excel');
    return 'Excel';
  }

  if (
    /\b(?:Selected Slide Context:|Presentation Context:|CRITICAL SLIDE CONTEXT:|CRITICAL POWERPOINT|CRITICAL STRUCTURE CONTRACT FOR SLIDE GENERATION|CRITICAL INSTRUCTIONS FOR (?:EXECUTIVE VISUAL|SLIDE EDITING|CHART|IMAGE|SUMMARIZE SLIDES|EXECUTIVE SUMMARY) GENERATION|PowerPoint Slide Deck Requirements:|STRICT DECK CONTRACT:|revising an existing PowerPoint deck|Executive Slide Summary|for PowerPoint visualizing|for this presentation topic|Summarize key takeaways from these slides|presentation slides|presentation deck|Slide Takeaway Callout|pitch deck|visual slide ideas|Make this slide more concise)\b/i.test(combinedRaw) ||
    /\[(?:Summarize Slides|Doc to Deck|Create Chart|Create Image|Build Presentation|Slide Insights)\]/i.test(combinedRaw) ||
    /\[(?:Highlighted )?Slide \d+\]:/i.test(combinedRaw) ||
    /(?:^|\n)\s*##+\s*Slide\s+\d+\b/i.test(combinedRaw) ||
    /"visualType"\s*:\s*"(?:metric_grid_3col|three_column_cards|process_timeline|before_after)"/i.test(combinedRaw) ||
    /"chartType"\s*:\s*"(?:pie|doughnut|bar|column|line)"/i.test(combinedRaw)
  ) {
    if (shortId) recordSessionHost(shortId, 'PowerPoint');
    return 'PowerPoint';
  }

  if (
    /\b(?:Email Thread Context:|Tone & Clarity Assessment|OPTION A\s*[:\-–\(\s]+CONCISE EXECUTIVE|OPTION B\s*[:\-–\(\s]+DIPLOMATIC|Draft a grounded email reply|elite executive email assistant|Meeting Host\s*\/\s*Organizer Context)\b/i.test(combinedRaw) ||
    /\[(?:Summarize Thread|Draft Reply|Email Coach|Action Items)\]/i.test(combinedRaw)
  ) {
    if (shortId) recordSessionHost(shortId, 'Outlook');
    return 'Outlook';
  }

  if (
    /\b(?:Active Document Context:|Active Working Draft:|Selected Document Text:|Selected Document Context:|Summarize this document thoroughly|Executive Summary Box for this document|from this document\. Present them in a structured table)\b/i.test(combinedRaw)
  ) {
    if (shortId) recordSessionHost(shortId, 'Word');
    return 'Word';
  }

  return null;
}

function extractMessagesFromTurns(turns = []) {
  if (!Array.isArray(turns)) return [];
  const messages = [];
  for (const turn of turns) {
    const rawUserText = turn?.query?.text || '';
    const cleanUserText = cleanStoredPromptText(rawUserText);
    if (cleanUserText) {
      messages.push({
        role: 'user',
        text: cleanUserText,
        citations: []
      });
    }
    const rawReply =
      turn?.assistAnswer?.replies?.[0]?.groundedContent?.content?.text ||
      turn?.reply?.reply ||
      turn?.reply?.summary?.summaryText ||
      turn?.answer ||
      '';
    if (typeof rawReply === 'string' && rawReply.trim()) {
      messages.push({
        role: 'assistant',
        text: rawReply.trim(),
        citations: []
      });
    }
  }
  return messages;
}

async function listSessionsAPI(pageSize = 20, filterUserId = '', userGoogleToken = '') {
  const accessToken = await resolveSessionBearerToken(userGoogleToken, filterUserId);

  const endpointUrl = `https://${DISCOVERY_ENGINE_HOST}/v1alpha/projects/${GE_GCP_PROJECT_ID}/locations/${GCP_LOCATION}/collections/${ENTERPRISE_COLLECTION_ID}/engines/${ENTERPRISE_APP_ID}/sessions?pageSize=${pageSize}`;

  const apiHeaders = {
    'Authorization': `Bearer ${accessToken}`,
    'Content-Type': 'application/json'
  };
  if (PROJECT_ID) {
    apiHeaders['x-goog-user-project'] = GE_GCP_PROJECT_ID;
  }

  const res = await fetch(endpointUrl, {
    method: 'GET',
    headers: apiHeaders
  });

  if (!res.ok) {
    const errText = await res.text();
    throw new Error(`Failed to list sessions (HTTP ${res.status}): ${errText}`);
  }

  const data = await res.json();
  const rawSessions = Array.isArray(data.sessions) ? data.sessions : [];

  // Enforce strict object-level authorization: filter sessions by authenticated user (CWE-639)
  if (!filterUserId && !userGoogleToken) {
    return [];
  }
  const normalizedFilterId = (filterUserId || '').toLowerCase();
  const userFilteredSessions = rawSessions.filter(s => {
    if (s.userPseudoId && normalizedFilterId) {
      return s.userPseudoId.toLowerCase() === normalizedFilterId;
    }
    // When authenticated directly with the user's own Google OAuth token, Discovery Engine scopes sessions to that user
    return Boolean(userGoogleToken);
  });

  return userFilteredSessions.map(s => {
    const shortId = s.name ? s.name.split('/sessions/').pop() : '';
    const inferredHost = inferSessionHost(s);
    const messages = extractMessagesFromTurns(s.turns);
    const firstUserMsg = messages.find(m => m.role === 'user');
    const derivedTitle = firstUserMsg
      ? (firstUserMsg.text.length > 36 ? firstUserMsg.text.substring(0, 36) + '...' : firstUserMsg.text)
      : 'Untitled Conversation';
    const strippedDisplayName = (s.displayName || '').replace(/^\[(Word|Excel|PowerPoint|Outlook)\]\s*/i, '').trim();
    const rawDisplayName = (strippedDisplayName && strippedDisplayName !== 'Office 365 Session' && strippedDisplayName !== 'New Conversation')
      ? strippedDisplayName
      : derivedTitle;
    return {
      sessionId: shortId,
      name: s.name,
      displayName: rawDisplayName,
      host: inferredHost,
      startTime: s.startTime,
      endTime: s.endTime,
      turnsCount: Array.isArray(s.turns) ? s.turns.length : 0,
      messages,
      state: s.state,
      userPseudoId: s.userPseudoId
    };
  });
}

async function deleteSessionAPI(sessionId, requestingUserId = '', userGoogleToken = '') {
  if (!sessionId) throw new Error('sessionId is required');
  const validSessionIdRegex = /^[a-zA-Z0-9_-]+$/;
  const canonicalSessionPrefix = `projects/${GE_GCP_PROJECT_ID}/locations/${GCP_LOCATION}/collections/${ENTERPRISE_COLLECTION_ID}/engines/${ENTERPRISE_APP_ID}/sessions/`;
  let cleanId = sessionId;
  if (sessionId.startsWith(canonicalSessionPrefix)) {
    cleanId = sessionId.slice(canonicalSessionPrefix.length);
  }
  if (!validSessionIdRegex.test(cleanId) || cleanId.includes('/') || cleanId.includes('..')) {
    const err = new Error('Invalid sessionId format: must be alphanumeric (no path traversal allowed)');
    err.statusCode = 400;
    throw err;
  }

  const accessToken = await resolveSessionBearerToken(userGoogleToken, requestingUserId);

  const fullSessionName = `${canonicalSessionPrefix}${cleanId}`;
  const endpointUrl = `https://${DISCOVERY_ENGINE_HOST}/v1alpha/${fullSessionName}`;

  const apiHeaders = {
    'Authorization': `Bearer ${accessToken}`
  };
  if (PROJECT_ID) {
    apiHeaders['x-goog-user-project'] = GE_GCP_PROJECT_ID;
  }

  // Object-level authorization check: Verify session ownership before deletion (CWE-639, fail-closed)
  const sessionAuth = await verifySessionAccess(fullSessionName, requestingUserId, accessToken);
  if (!sessionAuth.authorized) {
    const err = new Error(sessionAuth.reason || 'Forbidden: Unauthorized session access');
    err.statusCode = sessionAuth.statusCode || 403;
    throw err;
  }

  const res = await fetch(endpointUrl, {
    method: 'DELETE',
    headers: apiHeaders
  });

  if (!res.ok) {
    const errText = await res.text();
    throw new Error(`Failed to delete session (HTTP ${res.status}): ${errText}`);
  }

  return { success: true, sessionId };
}

/**
 * Helper to extract balanced JSON objects from an incoming stream buffer
 * (Discovery Engine :streamAssist returns a JSON array `[ {...}, {...} ]` or NDJSON).
 */
function extractStreamJsonObjects(buffer) {
  const objects = [];
  let depth = 0;
  let objStart = -1;
  let inString = false;
  let escaped = false;
  let lastConsumedIndex = 0;

  for (let i = 0; i < buffer.length; i++) {
    const ch = buffer[i];
    if (escaped) {
      escaped = false;
      continue;
    }
    if (ch === '\\' && inString) {
      escaped = true;
      continue;
    }
    if (ch === '"') {
      inString = !inString;
      continue;
    }
    if (inString) continue;

    if (ch === '{') {
      if (depth === 0) objStart = i;
      depth++;
    } else if (ch === '}') {
      depth--;
      if (depth === 0 && objStart !== -1) {
        const rawObj = buffer.slice(objStart, i + 1);
        try {
          objects.push(JSON.parse(rawObj));
          lastConsumedIndex = i + 1;
        } catch (e) {
          // Wait for more bytes if malformed
        }
        objStart = -1;
      }
    }
  }

  return {
    objects,
    remaining: lastConsumedIndex > 0 ? buffer.slice(lastConsumedIndex) : buffer
  };
}

/**
 * Real-time Server-Sent Events (SSE) Streaming Endpoint (`/askGeminiStream`).
 * Streams incremental tokens and citations directly from Discovery Engine `:streamAssist`
 * to the Office Add-in client as `text/event-stream` events.
 */
async function handleGeminiStreamRequest(req, res) {
  cors(req, res, async () => {
    if (req.method !== 'POST' && req.method !== 'GET') {
      return res.status(405).json({ error: 'Method Not Allowed' });
    }

    // Enforce cryptographic authentication (CWE-306)
    const auth = await verifyAuthentication(req);
    if (!auth.authenticated) {
      return res.status(auth.statusCode || 401).json({ error: auth.error });
    }

    const body = req.method === 'POST' ? (req.body || {}) : (req.query || {});
    const { prompt, history, sessionId, attachments, dataStoreId, enableGrounding = true, host = null } = body;
    const normalizedHost = normalizeOfficeHost(host || req.headers['x-office-host']);
    // Object-Level Authorization: User identity must strictly come from cryptographically verified auth context (CWE-639 / BOLA mitigation)
    const userId = auth.userId;
    const userPseudoId = auth.userId;
    const rawUserGoogleToken = req.headers['x-end-user-google-token'] || req.headers['x-end-user-token'] || '';
    if (rawUserGoogleToken) {
      cacheOauth3GoogleToken(auth, rawUserGoogleToken, 3600);
    }
    const userGoogleToken = rawUserGoogleToken || getCachedOauth3GoogleToken(auth).token || '';
    const authMode = req.headers['x-user-auth-mode'] || '';

    if (!prompt) {
      return res.status(400).json({ error: 'Prompt is required' });
    }

    // Configure Server-Sent Events (SSE) headers
    res.setHeader('Content-Type', 'text/event-stream; charset=utf-8');
    res.setHeader('Cache-Control', 'no-cache, no-transform');
    res.setHeader('Connection', 'keep-alive');
    res.setHeader('X-Accel-Buffering', 'no');
    if (typeof res.flushHeaders === 'function') {
      res.flushHeaders();
    }

    const sendEvent = (data) => {
      res.write(`data: ${JSON.stringify(data)}\n\n`);
    };

    try {
      if (BACKEND_MODE !== 'streamassist' || !ENTERPRISE_APP_ID || !PROJECT_ID) {
        sendEvent({
          error: 'Invalid Configuration',
          details: 'BACKEND_MODE must be streamassist and GCP_PROJECT_ID / GEMINI_ENTERPRISE_APP_ID must be configured.',
          done: true
        });
        res.write('data: [DONE]\n\n');
        return res.end();
      }

      const activeUserId = userPseudoId || userId || 'office_365_user';
      let bearerToken = null;

      if (userGoogleToken) {
        console.log(JSON.stringify({
          severity: 'INFO',
          message: `[AUTH_STREAM] Using End-User Google Token for Discovery Engine StreamAssist (Mode: ${authMode || 'user_token'}, User: ${activeUserId})`,
          user_id: activeUserId,
          auth_mode: authMode || 'user_token',
          token_source: 'X-End-User-Google-Token'
        }));
        bearerToken = userGoogleToken;
      } else if (ALLOW_SERVICE_ACCOUNT_FALLBACK) {
        console.warn(JSON.stringify({
          severity: 'WARNING',
          message: `[AUTH_STREAM_FALLBACK] No end-user Google token provided for user '${activeUserId}'. ALLOW_SERVICE_ACCOUNT_FALLBACK is enabled. Falling back to Cloud Run Service Account ADC credentials.`,
          user_id: activeUserId,
          allow_service_account_fallback: true
        }));
        const client = await auth.getClient();
        const accessTokenObj = await client.getAccessToken();
        bearerToken = typeof accessTokenObj === 'string' ? accessTokenObj : accessTokenObj.token;
      } else {
        sendEvent({
          error: 'End-user Google authentication token is required to access Gemini. Service account fallback is disabled.',
          statusCode: 403,
          done: true
        });
        res.write('data: [DONE]\n\n');
        return res.end();
      }

      const groundingDecision = await classifyGroundingIntent({
        prompt,
        enableGrounding,
        attachments,
        host: normalizedHost
      });
      const effectiveGrounding = groundingDecision.enableGrounding;
      const guidedPrompt = applyIntentGuidanceToPrompt(prompt, groundingDecision, normalizedHost || 'Office');

      console.log(JSON.stringify({
        severity: 'INFO',
        message: `[INTENT_ROUTER] SSE StreamAssist grounding decision: intent=${groundingDecision.intent}, enableGrounding=${effectiveGrounding} (source=${groundingDecision.source}, reason=${groundingDecision.reason})`,
        intent: groundingDecision.intent,
        enable_grounding: effectiveGrounding,
        source: groundingDecision.source,
        reason: groundingDecision.reason,
        host: normalizedHost || 'Office'
      }));

      const endpointUrl = `https://${DISCOVERY_ENGINE_HOST}/v1alpha/projects/${GE_GCP_PROJECT_ID}/locations/${GCP_LOCATION}/collections/${ENTERPRISE_COLLECTION_ID}/engines/${ENTERPRISE_APP_ID}/assistants/${ENTERPRISE_ASSISTANT_ID}:streamAssist`;

      const requestBody = {
        query: { text: ensurePromptHasHostTag(guidedPrompt, normalizedHost) },
        toolsSpec: {
          vertexAiSearchSpec: {}
        }
      };

      if (effectiveGrounding === false) {
        delete requestBody.toolsSpec.vertexAiSearchSpec;
      } else if (dataStoreId) {
        const dataStoreSpecs = await buildDataStoreSpecs(dataStoreId, bearerToken);
        if (dataStoreSpecs.length > 0) {
          requestBody.toolsSpec.vertexAiSearchSpec = { dataStoreSpecs };
        }
      }

      const validSessionIdRegex = /^[a-zA-Z0-9_-]+$/;
      const canonicalSessionPrefix = `projects/${GE_GCP_PROJECT_ID}/locations/${GCP_LOCATION}/collections/${COLLECTION_ID}/engines/${ENTERPRISE_APP_ID}/sessions/`;
      let isValidSession = false;
      if (sessionId) {
        if (validSessionIdRegex.test(sessionId)) {
          isValidSession = true;
        } else if (sessionId.startsWith(canonicalSessionPrefix)) {
          const rest = sessionId.slice(canonicalSessionPrefix.length);
          isValidSession = validSessionIdRegex.test(rest);
        }
        if (!isValidSession) {
          console.warn(`[SECURITY_WARN] Discarded invalid or cross-engine sessionId: ${sessionId}`);
        }
      }

      let sessionToUse = null;
      if (sessionId && isValidSession) {
        let fullSessionName = sessionId.startsWith('projects/')
          ? sessionId
          : `${canonicalSessionPrefix}${sessionId}`;

        // Verify object-level authorization (BOLA / IDOR mitigation)
        const sessionAuth = await verifySessionAccess(fullSessionName, activeUserId, bearerToken);
        if (!sessionAuth.authorized) {
          if (sessionAuth.statusCode === 403) {
            const forbiddenErr = new Error(sessionAuth.reason || 'Forbidden: Unauthorized session access');
            forbiddenErr.statusCode = 403;
            throw forbiddenErr;
          }
          console.warn(`[SESSION_WARN] Session '${sessionId}' invalid or expired. Creating fresh session.`);
        } else {
          sessionToUse = fullSessionName;
        }
      }

      if (sessionToUse) {
        requestBody.session = sessionToUse;
        if (normalizedHost) {
          recordSessionHost(sessionToUse, normalizedHost);
        }
      } else {
        try {
          const sessionCreateUrl = `https://${DISCOVERY_ENGINE_HOST}/v1alpha/projects/${GE_GCP_PROJECT_ID}/locations/${GCP_LOCATION}/collections/${ENTERPRISE_COLLECTION_ID}/engines/${ENTERPRISE_APP_ID}/sessions`;
          const sessionRes = await fetch(sessionCreateUrl, {
            method: 'POST',
            headers: {
              'Authorization': `Bearer ${bearerToken}`,
              'Content-Type': 'application/json',
              'X-Goog-User-Project': GE_GCP_PROJECT_ID
            },
            body: JSON.stringify({
              userPseudoId: activeUserId,
              ...(normalizedHost ? { displayName: `[${normalizedHost}] Office 365 Session` } : {})
            })
          });
          if (sessionRes.ok) {
            const sessionData = await sessionRes.json();
            if (sessionData.name) {
              requestBody.session = sessionData.name;
              if (normalizedHost) {
                recordSessionHost(sessionData.name, normalizedHost);
              }
            }
          }
        } catch (sessionErr) {
          console.warn(`[SESSION_STREAM_WARN] Could not pre-create session for '${activeUserId}':`, sessionErr.message);
        }
      }

      sendEvent({ status: 'Connecting to Gemini...', backendMode: 'streamassist' });

      const headers = {
        'Authorization': `Bearer ${bearerToken}`,
        'Content-Type': 'application/json',
        'X-Goog-User-Project': GE_GCP_PROJECT_ID
      };

      let apiRes = await fetch(endpointUrl, {
        method: 'POST',
        headers,
        body: JSON.stringify(requestBody)
      });

      if (!apiRes.ok) {
        let errText = await apiRes.text();
        if (requestBody.session && (
          (apiRes.status === 400 && (errText.includes('PROMPT_TOO_LARGE') || errText.includes('INVALID_ARGUMENT'))) ||
          (apiRes.status === 403 && errText.includes('Session is not owned'))
        )) {
          delete requestBody.session;
          apiRes = await fetch(endpointUrl, {
            method: 'POST',
            headers,
            body: JSON.stringify(requestBody)
          });
          if (!apiRes.ok) errText = await apiRes.text();
        }

        if (!apiRes.ok && (apiRes.status === 499 || (apiRes.status >= 400 && requestBody.toolsSpec))) {
          delete requestBody.toolsSpec;
          apiRes = await fetch(endpointUrl, {
            method: 'POST',
            headers,
            body: JSON.stringify(requestBody)
          });
          if (!apiRes.ok) errText = await apiRes.text();
        }

        if (!apiRes.ok && ALLOW_SERVICE_ACCOUNT_FALLBACK && errText.includes('LICENSE_WITHOUT_SUBSCRIPTION_TIER')) {
          const client = await auth.getClient();
          const accessTokenObj = await client.getAccessToken();
          const adcBearerToken = typeof accessTokenObj === 'string' ? accessTokenObj : accessTokenObj.token;
          headers['Authorization'] = `Bearer ${adcBearerToken}`;
          apiRes = await fetch(endpointUrl, {
            method: 'POST',
            headers,
            body: JSON.stringify(requestBody)
          });
          if (!apiRes.ok) errText = await apiRes.text();
        }

        if (!apiRes.ok) {
          throw new Error(`StreamAssist API returned HTTP ${apiRes.status}: ${errText}`);
        }
      }

      let accumulatedText = '';
      let returnedSessionId = sessionId || null;
      const citations = [];
      const seenCitations = new Set();
      const thoughtParts = [];
      let streamBuffer = '';

      const processParsedChunk = (chunk) => {
        if (chunk.sessionInfo?.session) {
          returnedSessionId = chunk.sessionInfo.session.includes('/sessions/')
            ? chunk.sessionInfo.session.split('/sessions/').pop()
            : chunk.sessionInfo.session;
          if (normalizedHost && returnedSessionId) {
            recordSessionHost(returnedSessionId, normalizedHost);
          }
        }
        const replies = chunk.answer?.replies || [];
        let chunkDelta = '';
        for (const reply of replies) {
          const content = reply.groundedContent?.content || reply.content || {};
          const textPart = content.text || reply.text || '';
          const isThought = Boolean(content.thought || reply.thought);
          if (textPart) {
            if (isThought) {
              thoughtParts.push(textPart);
            } else {
              chunkDelta += textPart;
            }
          }
          const grounded = reply.groundedContent || {};
          const groundingMeta = grounded.textGroundingMetadata || chunk.answer?.groundedContent?.textGroundingMetadata || chunk.textGroundingMetadata;
          if (groundingMeta && Array.isArray(groundingMeta.references)) {
            for (const ref of groundingMeta.references) {
              const docMeta = ref.documentMetadata || {};
              let uri = docMeta.uri || '';
              let title = docMeta.title || '';
              if ((!title || title === 'Enterprise Data Source') && uri) {
                const parts = uri.split('/');
                title = decodeURIComponent(parts[parts.length - 1]);
              }
              if (!title) title = `Enterprise Document ${citations.length + 1}`;
              const cleanUri = uri ? uri.replace('gs://', '') : '';
              const rawContent = ref.content || '';
              const excerpt = rawContent.replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').trim();
              const key = uri || title;
              if (!seenCitations.has(key)) {
                seenCitations.add(key);
                citations.push({ index: citations.length + 1, title, uri, cleanUri, excerpt });
              }
            }
          }
          const searchChunk = grounded.searchChunk || grounded.search_chunk;
          if (searchChunk || grounded.web || grounded.uri) {
            let uri = grounded.uri || searchChunk?.uri || '';
            let title = grounded.title || searchChunk?.title || '';
            if ((!title || title === 'Enterprise Data Source') && uri) {
              const parts = uri.split('/');
              title = decodeURIComponent(parts[parts.length - 1]);
            }
            if (!title && uri) title = `Enterprise Document ${citations.length + 1}`;
            if (uri || title) {
              const cleanUri = uri ? uri.replace('gs://', '') : '';
              const excerpt = (searchChunk?.chunkText || searchChunk?.text || grounded.text || '').trim();
              const key = uri || title;
              if (!seenCitations.has(key)) {
                seenCitations.add(key);
                citations.push({
                  index: citations.length + 1,
                  title: title || 'Enterprise Data Source',
                  uri,
                  cleanUri,
                  excerpt
                });
              }
            }
          }
        }
        if (chunkDelta) {
          accumulatedText += chunkDelta;
          sendEvent({
            delta: chunkDelta,
            text: chunkDelta,
            accumulatedText,
            sessionId: returnedSessionId,
            citations,
            done: false
          });
        }
      };

      if (apiRes.body && typeof apiRes.body.getReader === 'function') {
        const reader = apiRes.body.getReader();
        const decoder = new TextDecoder('utf-8');
        while (true) {
          const { value, done } = await reader.read();
          if (done) break;
          streamBuffer += decoder.decode(value, { stream: true });
          const extracted = extractStreamJsonObjects(streamBuffer);
          streamBuffer = extracted.remaining;
          for (const chunkObj of extracted.objects) {
            processParsedChunk(chunkObj);
          }
        }
      } else {
        const rawText = await apiRes.text();
        const extracted = extractStreamJsonObjects(rawText);
        for (const chunkObj of extracted.objects) {
          processParsedChunk(chunkObj);
        }
      }

      if (!accumulatedText.trim() && thoughtParts.length > 0) {
        accumulatedText = thoughtParts.join('\n\n');
      }

      const chartProcessedText = processChartsInContent(accumulatedText);
      const finalProcessedText = await inlineImagesInContent(chartProcessedText);

      const updatedHistory = [
        ...(Array.isArray(history) ? history : []),
        { role: 'user', parts: [{ text: prompt }] },
        { role: 'model', parts: [{ text: finalProcessedText }] }
      ];

      sendEvent({
        delta: '',
        text: '',
        accumulatedText: finalProcessedText,
        fullText: finalProcessedText,
        result: finalProcessedText,
        sessionId: returnedSessionId,
        citations,
        history: updatedHistory,
        backendMode: 'streamassist',
        backendEngine: 'gemini-enterprise-discovery-engine',
        done: true
      });
      res.write('data: [DONE]\n\n');
      return res.end();
    } catch (error) {
      console.error('Error in SSE stream handler:', error);
      sendEvent({
        error: 'Failed to stream Gemini response',
        details: error.message,
        done: true
      });
      res.write('data: [DONE]\n\n');
      return res.end();
    }
  });
}

// Google Cloud Functions Framework Export
functions.http('askGeminiEnterprise', handleGeminiEnterpriseRequest);
functions.http('askGemini', handleGeminiEnterpriseRequest); // Backward-compatibility route
functions.http('askGeminiStream', handleGeminiStreamRequest);

// Standalone Express Server (for Local Dev, Cloud Run, and Docker)
const app = express();
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));
app.use(cors);

app.post('/askGeminiEnterprise', handleGeminiEnterpriseRequest);
app.post('/askGemini', handleGeminiEnterpriseRequest);
app.post('/askGeminiStream', handleGeminiStreamRequest);
app.get('/askGeminiStream', handleGeminiStreamRequest);
app.post('/', handleGeminiEnterpriseRequest);
app.get('/health', (req, res) => res.status(200).json({ status: 'healthy', service: 'geminiproxy', engine: 'Gemini Discovery Engine' }));
app.get('/api/config', async (req, res) => {
  try {
    const auth = await verifyAuthentication(req);
    if (!auth.authenticated) {
      return res.status(auth.statusCode || 401).json({ error: auth.error });
    }
    return res.status(200).json({
      google_oauth_client_id: process.env.GOOGLE_OAUTH_CLIENT_ID || '',
      microsoft_entra_app_id: process.env.MICROSOFT_ENTRA_APP_ID || '',
      user_auth_mode: process.env.USER_AUTH_MODE || 'local_adc',
      require_entra_auth: (process.env.REQUIRE_ENTRA_AUTH || 'false').toLowerCase() === 'true'
    });
  } catch (error) {
    return res.status(500).json({ error: error.message });
  }
});

// Cross-Surface Auth & Google OAuth3 Session Endpoints
app.get('/api/auth/me', async (req, res) => {
  try {
    const auth = await verifyAuthentication(req);
    if (!auth.authenticated) {
      return res.status(auth.statusCode || 401).json({ error: auth.error });
    }
    const cached = getCachedOauth3GoogleToken(auth);
    return res.status(200).json({
      authenticated: true,
      user: {
        email: auth.email || auth.userId,
        sub: auth.userId
      },
      user_auth_mode: process.env.USER_AUTH_MODE || 'local_adc',
      google_token: cached.token || null,
      google_token_expires_in: cached.remainingSec || 0
    });
  } catch (error) {
    return res.status(500).json({ error: error.message });
  }
});

app.post('/api/auth/google-token', async (req, res) => {
  try {
    const auth = await verifyAuthentication(req);
    if (!auth.authenticated) {
      return res.status(auth.statusCode || 401).json({ error: auth.error });
    }
    const { google_token, expires_in = 3600 } = req.body || {};
    cacheOauth3GoogleToken(auth, google_token, expires_in);
    const cached = getCachedOauth3GoogleToken(auth);
    return res.status(200).json({
      status: 'ok',
      has_google_token: Boolean(cached.token),
      expires_in: cached.remainingSec
    });
  } catch (error) {
    return res.status(500).json({ error: error.message });
  }
});

app.delete('/api/auth/google-token', async (req, res) => {
  try {
    const auth = await verifyAuthentication(req);
    if (!auth.authenticated) {
      return res.status(auth.statusCode || 401).json({ error: auth.error });
    }
    cacheOauth3GoogleToken(auth, '', 0);
    return res.status(200).json({
      status: 'ok',
      has_google_token: false
    });
  } catch (error) {
    return res.status(500).json({ error: error.message });
  }
});

// Sessions REST Endpoints (CWE-287, CWE-639)
app.get('/sessions', async (req, res) => {
  try {
    const auth = await verifyAuthentication(req);
    if (!auth.authenticated) {
      return res.status(auth.statusCode || 401).json({ error: auth.error });
    }
    const endUserId = auth.userId || '';
    const rawUserGoogleToken = req.headers['x-end-user-google-token'] || req.headers['x-end-user-token'] || '';
    if (rawUserGoogleToken) {
      cacheOauth3GoogleToken(auth, rawUserGoogleToken, 3600);
    }
    const userGoogleToken = rawUserGoogleToken || getCachedOauth3GoogleToken(auth).token || '';
    const pageSize = Math.min(parseInt(req.query.pageSize, 10) || 20, 100);
    const sessions = await listSessionsAPI(pageSize, endUserId, userGoogleToken);
    return res.status(200).json({ sessions });
  } catch (error) {
    console.error('Error fetching sessions:', error);
    return res.status(error.statusCode || 500).json({ error: error.message });
  }
});

app.delete('/sessions/:id', async (req, res) => {
  try {
    const auth = await verifyAuthentication(req);
    if (!auth.authenticated) {
      return res.status(auth.statusCode || 401).json({ error: auth.error });
    }
    const requestingUserId = auth.userId || '';
    const rawUserGoogleToken = req.headers['x-end-user-google-token'] || req.headers['x-end-user-token'] || '';
    if (rawUserGoogleToken) {
      cacheOauth3GoogleToken(auth, rawUserGoogleToken, 3600);
    }
    const userGoogleToken = rawUserGoogleToken || getCachedOauth3GoogleToken(auth).token || '';
    const result = await deleteSessionAPI(req.params.id, requestingUserId, userGoogleToken);
    return res.status(200).json(result);
  } catch (error) {
    console.error('Error deleting session:', error);
    return res.status(error.statusCode || 500).json({ error: error.message });
  }
});

app.post('/deleteSession', async (req, res) => {
  try {
    const auth = await verifyAuthentication(req);
    if (!auth.authenticated) {
      return res.status(auth.statusCode || 401).json({ error: auth.error });
    }
    const requestingUserId = auth.userId || '';
    const rawUserGoogleToken = req.headers['x-end-user-google-token'] || req.headers['x-end-user-token'] || '';
    if (rawUserGoogleToken) {
      cacheOauth3GoogleToken(auth, rawUserGoogleToken, 3600);
    }
    const userGoogleToken = rawUserGoogleToken || getCachedOauth3GoogleToken(auth).token || '';
    const { sessionId } = req.body;
    const result = await deleteSessionAPI(sessionId, requestingUserId, userGoogleToken);
    return res.status(200).json(result);
  } catch (error) {
    console.error('Error deleting session:', error);
    return res.status(error.statusCode || 500).json({ error: error.message });
  }
});

app.get('/', (req, res) => res.status(200).json({ status: 'healthy', service: 'geminiproxy', engine: 'Gemini Discovery Engine' }));

const PORT = process.env.PORT || 8080;
if (process.env.NODE_ENV !== 'test') {
  app.listen(PORT, () => {
    console.log(`Gemini Proxy HTTP Server running on port ${PORT}`);
  });
}

export {
  extractUserQueryContext,
  classifyIntentFallback,
  classifyGroundingIntent,
  applyIntentGuidanceToPrompt,
  cleanStoredPromptText
};


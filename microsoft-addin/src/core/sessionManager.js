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
 * Session Manager for Gemini for Microsoft 365
 * 
 * Manages local conversation persistence, message histories, active session tracking,
 * and seamless switching across chat threads with surface-isolation support
 * (Word sessions are separated from Excel and PowerPoint sessions).
 * 
 * @author Sathya AG, Principal Architect, Google
 */

const STORAGE_KEY = 'gemini_o365_chat_sessions_v1';
const ACTIVE_SESSION_KEY = 'gemini_o365_active_session_id'; // Legacy key, to be cleaned up
const WINDOW_STATE_KEY = 'gemini_o365_window_state_v1';

let memoryStorage = {
  activeSessionId: null,
  activeSessionIdByHost: {},
  sessions: {}
};

function isLocalStorageAvailable() {
  try {
    if (typeof window === 'undefined' || !window.localStorage) return false;
    const testKey = '__storage_test_ls__';
    window.localStorage.setItem(testKey, testKey);
    window.localStorage.removeItem(testKey);
    return true;
  } catch (e) {
    return false;
  }
}

function isSessionStorageAvailable() {
  try {
    if (typeof window === 'undefined' || !window.sessionStorage) return false;
    const testKey = '__storage_test_ss__';
    window.sessionStorage.setItem(testKey, testKey);
    window.sessionStorage.removeItem(testKey);
    return true;
  } catch (e) {
    return false;
  }
}

function normalizeHostName(rawHost) {
  if (!rawHost || typeof rawHost !== 'string') return null;
  const lower = rawHost.trim().toLowerCase();
  if (lower === 'word') return 'Word';
  if (lower === 'excel') return 'Excel';
  if (lower === 'powerpoint' || lower === 'ppt') return 'PowerPoint';
  if (lower === 'outlook') return 'Outlook';
  return null;
}

function stripHostPrefixFromTitle(title = '') {
  if (!title || typeof title !== 'string') return '';
  return title
    .replace(/^\[(Word|Excel|PowerPoint|Outlook)\]\s*/i, '')
    .replace(/\s*\[OfficeHost:\s*(?:Word|Excel|PowerPoint|Outlook)\]\s*/gi, '')
    .trim();
}

export function detectSessionHostFromContent(sessionOrData) {
  if (!sessionOrData || typeof sessionOrData !== 'object') return null;

  const rawTitle = String(sessionOrData.displayName || sessionOrData.title || '');
  const prefixMatch = rawTitle.match(/^\[(Word|Excel|PowerPoint|Outlook)\]/i);
  if (prefixMatch) {
    const matched = normalizeHostName(prefixMatch[1]);
    if (matched) return matched;
  }

  const messages = Array.isArray(sessionOrData.messages) ? sessionOrData.messages : [];
  const turns = Array.isArray(sessionOrData.turns) ? sessionOrData.turns : [];
  const turnTexts = [];
  for (const t of turns) {
    if (t?.query?.text) turnTexts.push(t.query.text);
    const reply =
      t?.assistAnswer?.replies?.[0]?.groundedContent?.content?.text ||
      t?.reply?.reply ||
      t?.reply?.summary?.summaryText ||
      t?.answer ||
      '';
    if (typeof reply === 'string' && reply) turnTexts.push(reply);
  }

  const combinedText = [
    rawTitle,
    ...messages.map(m => (m && typeof m.text === 'string' ? m.text : '')),
    ...turnTexts
  ].join('\n');

  if (!combinedText.trim()) return null;

  // Check explicit embedded [OfficeHost: <Host>] tag
  const explicitTagMatch = combinedText.match(/\[OfficeHost:\s*(Word|Excel|PowerPoint|Outlook)\]/i);
  if (explicitTagMatch) {
    const tagged = normalizeHostName(explicitTagMatch[1]);
    if (tagged) return tagged;
  }

  // Check explicit greeting host instruction ("assist with their <Host> document today")
  const greetingHostMatch = combinedText.match(/assist with their (Word|Excel|PowerPoint|Outlook) document today/i);
  if (greetingHostMatch) {
    const gHost = normalizeHostName(greetingHostMatch[1]);
    if (gHost) return gHost;
  }

  // 1. Excel signatures
  if (
    /📊\s*\[(?:Table Insights|Native Chart Builder)\]|✨\s*\[Formula (?:Generator|Builder)\]|🏷️\s*\[Taxonomy|🎨\s*\[Highlight Data|🧹\s*\[Format & Clean/i.test(combinedText) ||
    /```excel-(?:chart|formulas|taxonomy|range-format)\b/i.test(combinedText) ||
    /\b(?:Selected Spreadsheet Data:|Spreadsheet Context:|Active Worksheet:|Excel Formula Architect|Native Chart Builder|Formula Generator|Formula Builder|Table Insights|spreadsheet data|spreadsheet schema|from this spreadsheet|from this sheet|in this spreadsheet|in this worksheet)\b/i.test(combinedText) ||
    /(?:^|\n)\s*(?:Worksheet:\s*"|Analyzed Range:\s*|Column Schema:\s*)/i.test(combinedText) ||
    /\b(?:XLOOKUP|VLOOKUP|SUMIFS|COUNTIFS|AVERAGEIFS|INDEX\/MATCH)\b/.test(combinedText) ||
    /(?:^|[\s`"=])=(?:SUM|AVERAGE|XLOOKUP|VLOOKUP|INDEX|MATCH|COUNTIF|COUNTIFS|SUMIF|SUMIFS|AVERAGEIF|IFERROR)\s*\(/im.test(combinedText)
  ) {
    return 'Excel';
  }

  // 2. PowerPoint signatures (checked before Word so slide prompts with legacy Document Context headers resolve to PowerPoint)
  if (
    /📈\s*\[Create Chart\]|🎨\s*\[Create Image\]|✨\s*\[Build Presentation\]|📑\s*\[(?:Slide Insights|Summarize Slides)\]|\[(?:Summarize Slides|Doc to Deck|Create Chart|Create Image|Build Presentation|Slide Insights)\]/i.test(combinedText) ||
    /\b(?:Selected Slide Context:|Presentation Context:|CRITICAL SLIDE CONTEXT:|CRITICAL POWERPOINT|CRITICAL STRUCTURE CONTRACT FOR SLIDE GENERATION|CRITICAL INSTRUCTIONS FOR (?:EXECUTIVE VISUAL|SLIDE EDITING|CHART|IMAGE|SUMMARIZE SLIDES|EXECUTIVE SUMMARY) GENERATION|PowerPoint Slide Deck Requirements:|STRICT DECK CONTRACT:|revising an existing PowerPoint deck|Executive Slide Summary|for PowerPoint visualizing|for this presentation topic|Summarize key takeaways from these slides|presentation slides|presentation deck|Slide Takeaway Callout|pitch deck|visual slide ideas|Make this slide more concise)\b/i.test(combinedText) ||
    /\[(?:Highlighted )?Slide \d+\]:/i.test(combinedText) ||
    /(?:^|\n)\s*##+\s*Slide\s+\d+\b/i.test(combinedText) ||
    /"visualType"\s*:\s*"(?:metric_grid_3col|three_column_cards|process_timeline|before_after)"/i.test(combinedText) ||
    /"chartType"\s*:\s*"(?:pie|doughnut|bar|column|line)"/i.test(combinedText)
  ) {
    return 'PowerPoint';
  }

  // 3. Outlook signatures
  if (
    /\b(?:Email Thread Context:|Tone & Clarity Assessment|OPTION A\s*[:\-–\(\s]+CONCISE EXECUTIVE|OPTION B\s*[:\-–\(\s]+DIPLOMATIC|Draft a grounded email reply|elite executive email assistant|Meeting Host\s*\/\s*Organizer Context)\b/i.test(combinedText) ||
    /\[(?:Summarize Thread|Draft Reply|Email Coach|Action Items)\]/i.test(combinedText)
  ) {
    return 'Outlook';
  }

  // 4. Word signatures
  if (
    /\b(?:Active Document Context:|Active Working Draft:|Selected Document Text:|Selected Document Context:|Summarize this document thoroughly|Executive Summary Box for this document|from this document\. Present them in a structured table|in this document)\b/i.test(combinedText)
  ) {
    return 'Word';
  }

  return null;
}

function sanitizeSessionsHostAttribution(sessions, sessionHosts = {}) {
  if (!sessions || typeof sessions !== 'object') return false;
  let modified = false;

  for (const [id, session] of Object.entries(sessions)) {
    if (!session || typeof session !== 'object') continue;

    if (session.title) {
      const cleanedTitle = stripHostPrefixFromTitle(session.title);
      if (cleanedTitle && cleanedTitle !== session.title) {
        session.title = cleanedTitle;
        modified = true;
      }
    }

    // Prioritize content-based host detection over stale session.host / sessionHosts[id]
    // so sessions previously misattributed to 'Word' are automatically repaired.
    const contentHost = detectSessionHostFromContent(session);
    const mappedHost = normalizeHostName(sessionHosts[id]);
    const storedHost = normalizeHostName(session.host);
    const detectedHost = contentHost || mappedHost || storedHost || null;

    if (detectedHost && session.host !== detectedHost) {
      session.host = detectedHost;
      modified = true;
    }
    if (detectedHost && sessionHosts[id] !== detectedHost) {
      sessionHosts[id] = detectedHost;
      modified = true;
    }
  }

  return modified;
}

function loadState() {
  let state = {
    activeSessionId: memoryStorage.activeSessionId,
    activeSessionIdByHost: { ...(memoryStorage.activeSessionIdByHost || {}) },
    sessions: { ...(memoryStorage.sessions || {}) },
    sessionHosts: { ...(memoryStorage.sessionHosts || {}) }
  };

  if (isLocalStorageAvailable()) {
    try {
      // Clean up legacy global active session key to prevent cross-document bleeding
      window.localStorage.removeItem(ACTIVE_SESSION_KEY);
      
      const raw = window.localStorage.getItem(STORAGE_KEY);
      if (raw) {
        const parsed = JSON.parse(raw);
        state.sessions = parsed.sessions || {};
        state.sessionHosts = parsed.sessionHosts || {};

        const repairedHosts = sanitizeSessionsHostAttribution(state.sessions, state.sessionHosts);

        // Strip any legacy activeSessionId / activeSessionIdByHost from localStorage immediately
        // or persist repaired host attributions
        if (repairedHosts || parsed.activeSessionId !== undefined || parsed.activeSessionIdByHost !== undefined) {
          window.localStorage.setItem(STORAGE_KEY, JSON.stringify({
            sessions: state.sessions,
            sessionHosts: state.sessionHosts
          }));
        }
      }
    } catch (e) {
      console.warn('Could not read sessions from localStorage:', e);
    }
  } else {
    sanitizeSessionsHostAttribution(state.sessions, state.sessionHosts);
  }
  
  if (isSessionStorageAvailable()) {
    try {
      const rawSession = window.sessionStorage.getItem(WINDOW_STATE_KEY);
      if (rawSession) {
        const parsedSession = JSON.parse(rawSession);
        state.activeSessionId = parsedSession.activeSessionId || null;
        state.activeSessionIdByHost = parsedSession.activeSessionIdByHost || {};
      } else {
        state.activeSessionId = null;
        state.activeSessionIdByHost = {};
      }
    } catch (e) {
      console.warn('Could not read active session from sessionStorage:', e);
    }
  }

  return state;
}

function saveState(state) {
  memoryStorage = state;
  if (isLocalStorageAvailable()) {
    try {
      // Save ONLY the sessions dictionary and host registry globally
      window.localStorage.setItem(STORAGE_KEY, JSON.stringify({
        sessions: state.sessions,
        sessionHosts: state.sessionHosts || {}
      }));
    } catch (e) {
      console.warn('Could not write sessions to localStorage:', e);
    }
  }
  
  if (isSessionStorageAvailable()) {
    try {
      // Save active session tracking locally to this specific document window/frame
      window.sessionStorage.setItem(WINDOW_STATE_KEY, JSON.stringify({
        activeSessionId: state.activeSessionId,
        activeSessionIdByHost: state.activeSessionIdByHost
      }));
    } catch (e) {
      console.warn('Could not write active session to sessionStorage:', e);
    }
  }
}

function getCurrentDocumentUrl() {
  try {
    if (typeof Office !== 'undefined' && Office.context && Office.context.document && Office.context.document.url) {
      return String(Office.context.document.url).trim();
    }
  } catch (e) {
    // ignore
  }
  return '';
}

function getHostScopeKey(host = null) {
  const baseHost = host || 'default';
  const docUrl = getCurrentDocumentUrl();
  return docUrl ? `${baseHost}::${docUrl}` : baseHost;
}

export const SessionManager = {
  registerSessionHost(sessionId, host) {
    if (!sessionId || !host) return;
    const state = loadState();
    if (!state.sessionHosts) state.sessionHosts = {};
    let changed = false;
    if (state.sessionHosts[sessionId] !== host) {
      state.sessionHosts[sessionId] = host;
      changed = true;
    }
    if (state.sessions[sessionId] && state.sessions[sessionId].host !== host) {
      state.sessions[sessionId].host = host;
      changed = true;
    }
    if (changed) {
      saveState(state);
    }
  },

  getActiveSessionId(host = null) {
    const state = loadState();
    const docUrl = getCurrentDocumentUrl();
    if (host && state.activeSessionIdByHost) {
      const scopedKey = getHostScopeKey(host);
      const candidateId = state.activeSessionIdByHost[scopedKey] || (!docUrl ? state.activeSessionIdByHost[host] : null);
      const session = candidateId ? state.sessions[candidateId] : null;
      if (session && Array.isArray(session.messages) && session.messages.length > 0) {
        if (docUrl && session.documentUrl && session.documentUrl !== docUrl) {
          return null;
        }
        const sessionHost = detectSessionHostFromContent(session) || (state.sessionHosts && state.sessionHosts[candidateId]) || session.host || null;
        if (sessionHost && sessionHost.toLowerCase() !== host.toLowerCase()) {
          return null;
        }
        return candidateId;
      }
      return null; // Return null if host is specified but no session exists for it in this window
    }
    return state.activeSessionId;
  },

  setActiveSessionId(sessionId, host = null) {
    const state = loadState();
    const docUrl = getCurrentDocumentUrl();
    const normalizedHost = normalizeHostName(host) || host;
    state.activeSessionId = sessionId;
    if (normalizedHost) {
      if (!state.activeSessionIdByHost) state.activeSessionIdByHost = {};
      const scopedKey = getHostScopeKey(normalizedHost);
      state.activeSessionIdByHost[scopedKey] = sessionId;
      if (!docUrl) {
        state.activeSessionIdByHost[normalizedHost] = sessionId;
      }
      if (sessionId) {
        if (!state.sessionHosts) state.sessionHosts = {};
        if (!state.sessionHosts[sessionId]) {
          state.sessionHosts[sessionId] = normalizedHost;
        }
      }
    }
    if (sessionId && !state.sessions[sessionId]) {
      state.sessions[sessionId] = {
        id: sessionId,
        title: 'New Conversation',
        host: normalizedHost || (state.sessionHosts && state.sessionHosts[sessionId]) || null,
        documentUrl: docUrl || '',
        createdAt: Date.now(),
        updatedAt: Date.now(),
        messages: []
      };
    } else if (sessionId && state.sessions[sessionId]) {
      if (normalizedHost && !state.sessions[sessionId].host) {
        state.sessions[sessionId].host = normalizedHost;
      }
      if (docUrl && !state.sessions[sessionId].documentUrl) {
        state.sessions[sessionId].documentUrl = docUrl;
      }
    }
    saveState(state);
    return sessionId;
  },

  clearActiveSession(host = null) {
    const state = loadState();
    state.activeSessionId = null;
    if (host && state.activeSessionIdByHost) {
      const scopedKey = getHostScopeKey(host);
      state.activeSessionIdByHost[scopedKey] = null;
      state.activeSessionIdByHost[host] = null;
    }
    saveState(state);
  },

  getSession(sessionId) {
    if (!sessionId) return null;
    const state = loadState();
    return state.sessions[sessionId] || null;
  },

  getActiveSession(host = null) {
    const activeId = this.getActiveSessionId(host);
    return this.getSession(activeId);
  },

  getAllSessions(hostFilter = null) {
    const state = loadState();
    let list = Object.values(state.sessions).filter(s => s && Array.isArray(s.messages) && s.messages.length > 0);
    if (hostFilter) {
      const targetHost = normalizeHostName(hostFilter) || hostFilter;
      // Strict per-surface isolation: only return sessions that belong to targetHost.
      // Never default unknown/untagged sessions to 'Word'.
      list = list.filter(s => {
        const contentHost = detectSessionHostFromContent(s);
        const mappedHost = (state.sessionHosts && normalizeHostName(state.sessionHosts[s.id])) || null;
        const storedHost = normalizeHostName(s.host) || s.host || null;
        const sessionHost = contentHost || mappedHost || storedHost || null;
        if (!sessionHost) return false;
        return sessionHost.toLowerCase() === targetHost.toLowerCase();
      });
    }
    return list.sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0));
  },

  syncRemoteSession(remoteSession, host = null) {
    if (!remoteSession || !remoteSession.sessionId) return null;
    const state = loadState();
    const id = remoteSession.sessionId;
    const existing = state.sessions[id];
    const remoteMessages = Array.isArray(remoteSession.messages) ? remoteSession.messages : [];
    const remoteUpdatedAt = remoteSession.endTime
      ? Date.parse(remoteSession.endTime) || Date.now()
      : (remoteSession.startTime ? Date.parse(remoteSession.startTime) || Date.now() : Date.now());

    const contentHost = detectSessionHostFromContent(remoteSession);
    const remoteHost = normalizeHostName(remoteSession.host);
    const mappedHost = (state.sessionHosts && normalizeHostName(state.sessionHosts[id])) || null;
    const existingHost = existing ? (normalizeHostName(existing.host) || existing.host || null) : null;
    const inferredHost = contentHost || remoteHost || mappedHost || existingHost || null;
    const cleanedDisplayName = stripHostPrefixFromTitle(remoteSession.displayName || remoteSession.title || '');

    if (inferredHost) {
      if (!state.sessionHosts) state.sessionHosts = {};
      state.sessionHosts[id] = inferredHost;
    }

    if (!existing) {
      if (remoteMessages.length === 0) return null;
      // Never blindly attribute untagged remote sessions to 'Word' or to the currently open surface.
      // Only assign a host when inferred from session metadata, server attribution, or content.
      const resolvedHost = inferredHost || null;
      state.sessions[id] = {
        id,
        title: (cleanedDisplayName && cleanedDisplayName !== 'Office 365 Session')
          ? cleanedDisplayName
          : (generateTitleFromPrompt(remoteMessages[0]?.text) || 'Conversation'),
        host: resolvedHost,
        documentUrl: '',
        createdAt: remoteSession.startTime ? Date.parse(remoteSession.startTime) || remoteUpdatedAt : remoteUpdatedAt,
        updatedAt: remoteUpdatedAt,
        messages: remoteMessages.map(m => ({
          role: m.role || 'assistant',
          text: String(m.text || '')
            .replace(/\s*\[OfficeHost:\s*(?:Word|Excel|PowerPoint|Outlook)\]\s*/gi, '')
            .replace(/\s*\[IntentRouting:\s*[A-Z_]+\]\s*/gi, '')
            .trim(),
          citations: Array.isArray(m.citations) ? m.citations : [],
          timestamp: remoteUpdatedAt
        }))
      };
      saveState(state);
      return state.sessions[id];
    }

    let changed = false;
    if (inferredHost && existing.host !== inferredHost) {
      existing.host = inferredHost;
      changed = true;
    }
    if ((!existing.messages || existing.messages.length === 0) && remoteMessages.length > 0) {
      existing.messages = remoteMessages.map(m => ({
        role: m.role || 'assistant',
        text: String(m.text || '').replace(/\s*\[OfficeHost:\s*(?:Word|Excel|PowerPoint|Outlook)\]\s*/gi, '').trim(),
        citations: Array.isArray(m.citations) ? m.citations : [],
        timestamp: remoteUpdatedAt
      }));
      changed = true;
    }
    if (
      cleanedDisplayName &&
      cleanedDisplayName !== 'Office 365 Session' &&
      (!existing.title || existing.title === 'New Conversation' || existing.title === 'Conversation' || existing.title === 'Untitled Conversation')
    ) {
      existing.title = cleanedDisplayName;
      changed = true;
    }
    if (changed) {
      saveState(state);
    }
    return existing;
  },

  addMessage(sessionId, message, host = null) {
    if (!sessionId) return;
    const state = loadState();
    const docUrl = getCurrentDocumentUrl();
    const normalizedHost = normalizeHostName(host) || host;
    if (normalizedHost) {
      if (!state.sessionHosts) state.sessionHosts = {};
      state.sessionHosts[sessionId] = normalizedHost;
    }
    if (!state.sessions[sessionId]) {
      state.sessions[sessionId] = {
        id: sessionId,
        title: message.role === 'user' ? generateTitleFromPrompt(message.text) : 'Conversation',
        host: normalizedHost || (state.sessionHosts && state.sessionHosts[sessionId]) || null,
        documentUrl: docUrl || '',
        createdAt: Date.now(),
        updatedAt: Date.now(),
        messages: []
      };
    }

    const session = state.sessions[sessionId];
    session.updatedAt = Date.now();
    if (normalizedHost) {
      session.host = normalizedHost;
    } else if (!session.host && state.sessionHosts && state.sessionHosts[sessionId]) {
      session.host = state.sessionHosts[sessionId];
    }
    if (docUrl && !session.documentUrl) {
      session.documentUrl = docUrl;
    }

    // If this is the first user message and title is generic, derive a title
    if (message.role === 'user' && (!session.title || session.title === 'New Conversation' || session.title === 'Conversation')) {
      session.title = generateTitleFromPrompt(message.text);
    }

    session.messages.push({
      role: message.role, // 'user' | 'assistant' | 'system'
      text: message.text || '',
      citations: Array.isArray(message.citations) ? message.citations : [],
      timestamp: message.timestamp || Date.now()
    });

    // Re-verify content-based host attribution if message contains explicit host markers
    const contentHost = detectSessionHostFromContent(session);
    if (contentHost && !normalizedHost) {
      session.host = contentHost;
      if (!state.sessionHosts) state.sessionHosts = {};
      state.sessionHosts[sessionId] = contentHost;
    }

    saveState(state);
    return session;
  },

  setSessionTitle(sessionId, title) {
    if (!sessionId || !title) return;
    const state = loadState();
    if (state.sessions[sessionId]) {
      state.sessions[sessionId].title = stripHostPrefixFromTitle(title);
      saveState(state);
    }
  },

  deleteSession(sessionId) {
    if (!sessionId) return;
    const state = loadState();
    const session = state.sessions[sessionId];
    const host = session ? session.host : null;
    delete state.sessions[sessionId];
    if (state.sessionHosts) {
      delete state.sessionHosts[sessionId];
    }
    if (state.activeSessionId === sessionId) {
      state.activeSessionId = null;
    }
    if (host && state.activeSessionIdByHost && state.activeSessionIdByHost[host] === sessionId) {
      state.activeSessionIdByHost[host] = null;
    }
    saveState(state);
  },

  clearAllSessions(hostFilter = null) {
    const state = loadState();
    if (!hostFilter) {
      state.sessions = {};
      state.sessionHosts = {};
      state.activeSessionId = null;
      state.activeSessionIdByHost = {};
    } else {
      const targetHost = normalizeHostName(hostFilter) || hostFilter;
      Object.keys(state.sessions).forEach(id => {
        const s = state.sessions[id];
        const sHost = detectSessionHostFromContent(s) || (state.sessionHosts && state.sessionHosts[id]) || s.host || null;
        if (sHost && sHost.toLowerCase() === targetHost.toLowerCase()) {
          delete state.sessions[id];
          if (state.sessionHosts) {
            delete state.sessionHosts[id];
          }
        }
      });
      if (state.activeSessionIdByHost && state.activeSessionIdByHost[targetHost]) {
        state.activeSessionIdByHost[targetHost] = null;
      }
      state.activeSessionId = null;
    }
    saveState(state);
  }
};

function generateTitleFromPrompt(prompt) {
  if (!prompt) return 'New Conversation';
  let clean = stripHostPrefixFromTitle(prompt)
    .replace(/\s*\[IntentRouting:\s*[A-Z_]+\]\s*/gi, '')
    .replace(/^User Greeting:\s*"([^"]+)"[\s\S]*$/i, '$1')
    .replace(/^Active (?:Document Context|Working Draft):[\s\S]*?(?:User (?:Instruction|Query|Follow-Up):)/i, '')
    .replace(/^(?:Selected Document Context|Selected Slide Context|Presentation Context|CRITICAL SLIDE CONTEXT):[\s\S]*?(?:User (?:Instruction|Query|Follow-Up):|Task:)/i, '')
    .replace(/^Spreadsheet Context:[\s\S]*?(?:User (?:Instruction|Query|Follow-Up):)/i, '')
    .replace(/^User (?:Query|Instruction|Follow-Up):\s*/i, '')
    .replace(/Grounding & Intent Policy(?:\s*\[Intent:\s*[A-Z_]+\])?:[\s\S]*$/i, '')
    .replace(/Instruction:[\s\S]*$/i, '')
    .replace(/Constraint:[\s\S]*$/i, '')
    .replace(/Formatting:[\s\S]*$/i, '')
    .trim();

  // Strip leading punctuation/quotes
  clean = clean.replace(/^["'`\s]+/, '').replace(/["'`\s]+$/, '');
  if (!clean) return 'New Conversation';
  return clean.length > 36 ? clean.substring(0, 36) + '...' : clean;
}

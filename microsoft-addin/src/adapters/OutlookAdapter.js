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
 * Host Adapter Implementation for Microsoft Outlook (Office.context.mailbox)
 * 
 * Supports MessageRead and MessageCompose modes, full email thread envelope ingestion,
 * smart replies, action item extraction, calendar appointment creation, and in-body insertion.
 * 
 * @author Sathya AG, Principal Architect, Google
 */

export class OutlookAdapter {
  constructor() {
    this.name = "Outlook";
  }

  /**
   * Helper to wrap Office.js callback-based async methods into native Promises.
   */
  _promisify(fn, ...args) {
    return new Promise((resolve, reject) => {
      fn(...args, (asyncResult) => {
        const succeeded = asyncResult && (
          asyncResult.status === 'succeeded' ||
          (typeof Office !== 'undefined' && Office.AsyncResultStatus && asyncResult.status === Office.AsyncResultStatus.Succeeded)
        );
        if (succeeded) {
          resolve(asyncResult.value);
        } else {
          reject(new Error(asyncResult && asyncResult.error ? asyncResult.error.message : "Office.js Async Operation Failed"));
        }
      });
    });
  }

  /**
   * Returns true if the user is actively composing an email (New, Reply, Forward).
   */
  isComposeMode() {
    if (typeof Office === 'undefined' || !Office.context || !Office.context.mailbox) {
      return false;
    }
    const item = Office.context.mailbox.item;
    return Boolean(item && item.body && typeof item.body.setSelectedDataAsync === 'function');
  }

  /**
   * Read currently selected text from the active email composer (Compose mode).
   */
  async getSelectedText() {
    let selectedText = "";
    try {
      if (typeof Office !== 'undefined' && Office.context && Office.context.mailbox) {
        const item = Office.context.mailbox.item;
        if (item && typeof item.getSelectedDataAsync === 'function') {
          const textCoercion = (typeof Office !== 'undefined' && Office.CoercionType?.Text) || 'text';
          selectedText = await this._promisify(
            item.getSelectedDataAsync.bind(item),
            textCoercion
          );
          if (typeof selectedText === 'object' && selectedText.data) {
            selectedText = selectedText.data;
          }
        }
      }
    } catch (err) {
      console.warn("Outlook selection read notice (may be in Read mode):", err.message);
    }
    return (typeof selectedText === 'string') ? selectedText.trim() : "";
  }

  /**
   * Reads complete email or calendar appointment metadata and body context for summarization, Q&A, and meeting briefings.
   * Handles both MessageRead/AppointmentRead (sync properties) and MessageCompose/AppointmentCompose (async getters).
   */
  async getFullDocumentText() {
    try {
      if (typeof Office === 'undefined' || !Office.context || !Office.context.mailbox) {
        return "";
      }
      const item = Office.context.mailbox.item;
      if (!item) return "";

      const isAppointment = Boolean(
        item.itemType === 'appointment' ||
        (typeof Office !== 'undefined' && Office.MailboxEnums && item.itemType === Office.MailboxEnums.ItemType.Appointment) ||
        (item.itemType !== 'message' && (item.organizer || (item.start && item.end)))
      );

      // 1. Extract Subject
      let subject = "(No Subject)";
      if (typeof item.subject === 'string') {
        subject = item.subject || "(No Subject)";
      } else if (item.subject && typeof item.subject.getAsync === 'function') {
        subject = (await this._promisify(item.subject.getAsync.bind(item.subject))) || "(No Subject)";
      }

      // 2. Extract Sender / Organizer (From)
      let from = isAppointment ? "Meeting Organizer" : "Unknown Sender";
      const formatRecipient = (r) => {
        if (!r) return "";
        if (typeof r === 'string') return r.trim();
        const name = (r.displayName || r.name || "").trim();
        const address = (r.emailAddress || r.address || "").trim();
        if (name && address && name.toLowerCase() !== address.toLowerCase()) {
          return `${name} <${address}>`;
        }
        return address || name || "";
      };

      const fromSource = item.from || item.organizer;
      if (fromSource) {
        if (typeof fromSource.getAsync === 'function') {
          try {
            const fromResult = await this._promisify(fromSource.getAsync.bind(fromSource));
            from = formatRecipient(fromResult) || from;
          } catch (e) {
            console.warn("from/organizer getAsync error:", e);
          }
        } else {
          from = formatRecipient(fromSource) || from;
        }
      }
      if (this.isComposeMode() && (from === "Unknown Sender" || from === "Meeting Organizer") && Office.context.mailbox.userProfile) {
        const profile = Office.context.mailbox.userProfile;
        from = profile.displayName ? (profile.emailAddress ? `${profile.displayName} <${profile.emailAddress}>` : profile.displayName) : (profile.emailAddress || "Current User");
      }

      // 3. Extract To / Required Attendees & Cc / Optional Attendees
      let toRecipients = "";
      const toSource = item.requiredAttendees || item.to;
      if (Array.isArray(toSource)) {
        toRecipients = toSource.map(formatRecipient).filter(Boolean).join(", ");
      } else if (toSource && typeof toSource.getAsync === 'function') {
        try {
          const toList = await this._promisify(toSource.getAsync.bind(toSource));
          if (Array.isArray(toList)) {
            toRecipients = toList.map(formatRecipient).filter(Boolean).join(", ");
          } else if (toList) {
            toRecipients = formatRecipient(toList);
          }
        } catch (e) {
          console.warn("to/requiredAttendees getAsync error:", e);
        }
      }

      let ccRecipients = "";
      const ccSource = item.optionalAttendees || item.cc;
      if (Array.isArray(ccSource)) {
        ccRecipients = ccSource.map(formatRecipient).filter(Boolean).join(", ");
      } else if (ccSource && typeof ccSource.getAsync === 'function') {
        try {
          const ccList = await this._promisify(ccSource.getAsync.bind(ccSource));
          if (Array.isArray(ccList)) {
            ccRecipients = ccList.map(formatRecipient).filter(Boolean).join(", ");
          } else if (ccList) {
            ccRecipients = formatRecipient(ccList);
          }
        } catch (e) {
          console.warn("cc/optionalAttendees getAsync error:", e);
        }
      }

      // 4. Extract Appointment-specific Fields (Start, End, Location)
      let startStr = "";
      let endStr = "";
      let locationStr = "";

      if (item.start) {
        try {
          if (item.start instanceof Date) {
            startStr = item.start.toLocaleString();
          } else if (typeof item.start.getAsync === 'function') {
            const s = await this._promisify(item.start.getAsync.bind(item.start));
            startStr = s instanceof Date ? s.toLocaleString() : (typeof s === 'string' && !isNaN(new Date(s).getTime()) ? new Date(s).toLocaleString() : String(s || ""));
          } else if (typeof item.start === 'string' && !isNaN(new Date(item.start).getTime())) {
            startStr = new Date(item.start).toLocaleString();
          } else {
            startStr = String(item.start);
          }
        } catch (e) {}
      }

      if (item.end) {
        try {
          if (item.end instanceof Date) {
            endStr = item.end.toLocaleString();
          } else if (typeof item.end.getAsync === 'function') {
            const en = await this._promisify(item.end.getAsync.bind(item.end));
            endStr = en instanceof Date ? en.toLocaleString() : (typeof en === 'string' && !isNaN(new Date(en).getTime()) ? new Date(en).toLocaleString() : String(en || ""));
          } else if (typeof item.end === 'string' && !isNaN(new Date(item.end).getTime())) {
            endStr = new Date(item.end).toLocaleString();
          } else {
            endStr = String(item.end);
          }
        } catch (e) {}
      }

      if (item.location) {
        try {
          if (typeof item.location === 'string') {
            locationStr = item.location;
          } else if (typeof item.location.getAsync === 'function') {
            const loc = await this._promisify(item.location.getAsync.bind(item.location));
            locationStr = typeof loc === 'string' ? loc : (loc?.displayName || "");
          } else if (item.location.displayName) {
            locationStr = item.location.displayName;
          }
        } catch (e) {}
      }

      // 5. Extract Date and Conversation ID
      const dateTime = item.dateTimeCreated ? new Date(item.dateTimeCreated).toLocaleString() : "";
      const conversationId = item.conversationId || "";

      // 6. Read full email / appointment body (coerced to clean plain text)
      let bodyText = "";
      if (item.body && typeof item.body.getAsync === 'function') {
        const textCoercion = (typeof Office !== 'undefined' && Office.CoercionType && Office.CoercionType.Text) || 'text';
        bodyText = await this._promisify(
          item.body.getAsync.bind(item.body),
          textCoercion
        );
      } else if (typeof item.body === 'string') {
        bodyText = item.body;
      }

      // 7. Segment and clean 15+ message thread turns
      const structuredBody = this.segmentAndCleanEmailThread(bodyText);

      // 8. Assemble structured envelope
      if (isAppointment) {
        let fullContext = `--- CALENDAR APPOINTMENT / MEETING INVITATION METADATA ---\n`;
        fullContext += `Subject: ${subject}\n`;
        fullContext += `Organizer: ${from}\n`;
        if (toRecipients) fullContext += `Required Attendees: ${toRecipients}\n`;
        if (ccRecipients) fullContext += `Optional Attendees: ${ccRecipients}\n`;
        if (startStr) fullContext += `Start Time: ${startStr}\n`;
        if (endStr) fullContext += `End Time: ${endStr}\n`;
        if (locationStr) fullContext += `Location: ${locationStr}\n`;
        if (dateTime) fullContext += `Date Created: ${dateTime}\n`;
        if (conversationId) fullContext += `ConversationID: ${conversationId}\n`;
        fullContext += `----------------------------------------------------------\n\n`;
        fullContext += structuredBody ? structuredBody : "(Empty Meeting Description / Agenda Context)";
        return fullContext;
      }

      let fullContext = `--- EMAIL ENVELOPE METADATA ---\n`;
      fullContext += `Subject: ${subject}\n`;
      fullContext += `From: ${from}\n`;
      if (toRecipients) fullContext += `To: ${toRecipients}\n`;
      if (ccRecipients) fullContext += `Cc: ${ccRecipients}\n`;
      if (dateTime) fullContext += `Date: ${dateTime}\n`;
      if (conversationId) fullContext += `ConversationID: ${conversationId}\n`;
      fullContext += `-------------------------------\n\n`;
      fullContext += structuredBody ? structuredBody : "(Empty Email Body)";

      return fullContext;
    } catch (err) {
      console.error("Outlook full email context read error:", err);
      return "";
    }
  }

  /**
   * Cleans, segments, and labels chronological message turns in multi-message threads (e.g. 15-turn threads).
   * Delimits turns into structured [Message X of N - ...] blocks for optimal LLM context.
   */
  segmentAndCleanEmailThread(rawBody) {
    if (!rawBody || typeof rawBody !== 'string') return "";

    // Split on common Outlook/Exchange, MIME, and webmail thread delimiters
    const delimiterRegex = /(?:-{3,}\s*(?:Original Message|Forwarded message|Message d'origine)\s*-{3,}|_{8,}|(?:\r?\n|^)From:\s+[^\n]+\r?\n(?:Sent|Date):\s+[^\n]+|(?:\r?\n|^)On\s+[^\n]+?\s+wrote:?)/gi;

    // Find all delimiter matches
    const delimiters = [];
    let match;
    while ((match = delimiterRegex.exec(rawBody)) !== null) {
      delimiters.push({
        index: match.index,
        length: match[0].length,
        header: match[0].trim()
      });
    }

    if (delimiters.length === 0) {
      return rawBody.trim();
    }

    const turns = [];
    // First message turn (the latest message at top)
    const firstTurn = rawBody.substring(0, delimiters[0].index).trim();
    if (firstTurn) {
      turns.push({ header: "Latest Message / Active Reply", content: firstTurn });
    }

    for (let i = 0; i < delimiters.length; i++) {
      const startIndex = delimiters[i].index + delimiters[i].length;
      const endIndex = i + 1 < delimiters.length ? delimiters[i + 1].index : rawBody.length;
      const turnRaw = rawBody.substring(startIndex, endIndex).trim();
      
      const cleanedTurn = turnRaw
        .replace(/^>+\s?/gm, '') // Remove nested quote markers
        .trim();

      if (cleanedTurn) {
        turns.push({
          header: delimiters[i].header.replace(/[\r\n]+/g, ' | ').replace(/[-_]{3,}/g, '').trim() || "Prior Thread Message",
          content: cleanedTurn
        });
      }
    }

    if (turns.length === 0) {
      return rawBody.trim();
    }

    const totalMessages = turns.length;
    let formattedThread = `=== EMAIL THREAD TURN BREAKDOWN (${totalMessages} Messages Identified) ===\n\n`;

    turns.forEach((turn, idx) => {
      formattedThread += `--- [Message ${idx + 1} of ${totalMessages} - ${turn.header}] ---\n`;
      formattedThread += `${turn.content}\n\n`;
    });

    formattedThread += `=== END OF EMAIL THREAD ===`;
    return formattedThread;
  }

  /**
   * Reads draft text from the active composer (or selection if available).
   */
  async getDraftBody() {
    try {
      if (typeof Office !== 'undefined' && Office.context && Office.context.mailbox) {
        const item = Office.context.mailbox.item;
        if (item) {
          const selectedText = await this.getSelectedText();
          if (selectedText && selectedText.length > 5) {
            return selectedText;
          }
          if (item.body && typeof item.body.getAsync === 'function') {
            const textCoercion = (typeof Office !== 'undefined' && Office.CoercionType?.Text) || 'text';
            const fullBody = await this._promisify(
              item.body.getAsync.bind(item.body),
              textCoercion
            );
            return (fullBody || "").trim();
          }
        }
      }
    } catch (err) {
      console.warn("Outlook getDraftBody error:", err);
    }
    return "";
  }

  /**
   * Replaces the entire draft body in the active email composer.
   */
  async replaceDraftBody(htmlContent, rawText = "") {
    return this.insertContent(htmlContent, rawText, "replace_all");
  }

  /**
   * Water-tight content insertion handling both Compose and Read modes.
   */
  async insertContent(htmlContent, rawText = "", mode = "smart") {
    if (!htmlContent && !rawText) return;

    if (typeof Office === 'undefined' || !Office.context || !Office.context.mailbox) {
      throw new Error("Office.js Mailbox environment is not available.");
    }

    const item = Office.context.mailbox.item;
    if (!item) throw new Error("No active email item found in Outlook context.");

    const sanitizedHtml = `<div style="font-family:'Segoe UI',Calibri,Arial,sans-serif; font-size:11pt; color:#201f1e; line-height:1.5;">${htmlContent}</div>`;
    const coercionHtml = (typeof Office !== 'undefined' && Office.CoercionType?.Html) || 'html';

    if (this.isComposeMode()) {
      // COMPOSE MODE: Directly inject into active email composer
      if (mode === "prepend") {
        await this._promisify(
          item.body.prependAsync.bind(item.body),
          sanitizedHtml,
          { coercionType: coercionHtml }
        );
      } else if (mode === "append") {
        await this._promisify(
          item.body.appendAsync.bind(item.body),
          sanitizedHtml,
          { coercionType: coercionHtml }
        );
      } else if (mode === "replace_all" || mode === "replace_draft") {
        // Replace entire body if setAsync is supported
        if (typeof item.body.setAsync === 'function') {
          await this._promisify(
            item.body.setAsync.bind(item.body),
            sanitizedHtml,
            { coercionType: coercionHtml }
          );
        } else {
          await this._promisify(
            item.body.setSelectedDataAsync.bind(item.body),
            sanitizedHtml,
            { coercionType: coercionHtml }
          );
        }
      } else {
        // Default: Insert at cursor or replace selection
        await this._promisify(
          item.body.setSelectedDataAsync.bind(item.body),
          sanitizedHtml,
          { coercionType: coercionHtml }
        );
      }
    } else {
      // READ MODE: Launch a Smart Reply draft window with the AI generated response
      if (mode === "reply_all") {
        if (typeof item.displayReplyAllForm === 'function') {
          item.displayReplyAllForm({ htmlBody: sanitizedHtml });
        }
      } else {
        if (typeof item.displayReplyForm === 'function') {
          item.displayReplyForm({ htmlBody: sanitizedHtml });
        }
      }
    }
  }

  /**
   * Reads current user profile (displayName and email) with fallback name parsing from email prefix.
   */
  getUserProfile() {
    let displayName = "";
    let email = "";
    try {
      if (typeof Office !== 'undefined' && Office.context && Office.context.mailbox && Office.context.mailbox.userProfile) {
        const profile = Office.context.mailbox.userProfile;
        displayName = (profile.displayName || "").trim();
        email = (profile.emailAddress || profile.email || "").trim();
      }
    } catch (e) {
      console.warn("Could not read user profile:", e);
    }

    if (!displayName && email) {
      const prefix = email.split('@')[0];
      const parts = prefix.split(/[._-]+/).filter(Boolean);
      if (parts.length > 0) {
        displayName = parts.map(p => p.charAt(0).toUpperCase() + p.slice(1)).join(' ');
      }
    }

    return {
      displayName: displayName || "User",
      email: email || ""
    };
  }

  /**
   * Reads incoming sender profile (displayName and email) from item.from with fallback name parsing.
   * Handles Read mode (item.from), Compose mode (item.to / thread context), and raw string formats.
   */
  async getSenderProfile() {
    let displayName = "";
    let email = "";
    try {
      if (typeof Office !== 'undefined' && Office.context && Office.context.mailbox) {
        const item = Office.context.mailbox.item;
        if (item) {
          let fromObj = null;
          if (item.from) {
            if (typeof item.from.getAsync === 'function') {
              try {
                fromObj = await this._promisify(item.from.getAsync.bind(item.from));
              } catch (e) {
                console.warn("item.from.getAsync error:", e);
              }
            } else {
              fromObj = item.from;
            }
          } else if (item.organizer) {
            if (typeof item.organizer.getAsync === 'function') {
              try {
                fromObj = await this._promisify(item.organizer.getAsync.bind(item.organizer));
              } catch (e) {
                console.warn("item.organizer.getAsync error:", e);
              }
            } else {
              fromObj = item.organizer;
            }
          }

          if (fromObj) {
            if (typeof fromObj === 'string') {
              const str = fromObj.trim();
              const angleMatch = str.match(/^(.*?)\s*<([^>]+)>/);
              if (angleMatch) {
                displayName = angleMatch[1].replace(/["']/g, '').trim();
                email = angleMatch[2].trim();
              } else {
                email = str;
              }
            } else {
              displayName = (fromObj.displayName || fromObj.name || "").trim();
              email = (fromObj.emailAddress || fromObj.address || "").trim();
            }
          }

          // Check organizer if sender still empty
          if (!displayName && !email && item.organizer) {
            let orgObj = item.organizer;
            if (typeof orgObj.getAsync === 'function') {
              try {
                orgObj = await this._promisify(orgObj.getAsync.bind(orgObj));
              } catch (e) {}
            }
            if (typeof orgObj === 'string') {
              const str = orgObj.trim();
              const angleMatch = str.match(/^(.*?)\s*<([^>]+)>/);
              if (angleMatch) {
                displayName = angleMatch[1].replace(/["']/g, '').trim();
                email = angleMatch[2].trim();
              } else {
                email = str;
              }
            } else if (orgObj) {
              displayName = (orgObj.displayName || orgObj.name || "").trim();
              email = (orgObj.emailAddress || orgObj.address || "").trim();
            }
          }

          // In Compose mode or if sender not found from item.from, check item.to
          if ((!displayName && !email) || (this.isComposeMode() && item.to)) {
            let toList = null;
            if (typeof item.to.getAsync === 'function') {
              try {
                toList = await this._promisify(item.to.getAsync.bind(item.to));
              } catch (e) {
                console.warn("item.to.getAsync error:", e);
              }
            } else if (Array.isArray(item.to) && item.to.length > 0) {
              toList = item.to;
            }

            if (Array.isArray(toList) && toList.length > 0) {
              const firstTo = toList[0];
              if (typeof firstTo === 'string') {
                const angleMatch = firstTo.match(/^(.*?)\s*<([^>]+)>/);
                if (angleMatch) {
                  displayName = displayName || angleMatch[1].replace(/["']/g, '').trim();
                  email = email || angleMatch[2].trim();
                } else {
                  email = email || firstTo.trim();
                }
              } else if (firstTo) {
                displayName = displayName || (firstTo.displayName || firstTo.name || "").trim();
                email = email || (firstTo.emailAddress || firstTo.address || "").trim();
              }
            }
          }

          // Fallback: If still not resolved, try extracting from thread history in body
          if (!displayName && !email && item.body && typeof item.body.getAsync === 'function') {
            try {
              const textCoercion = (typeof Office !== 'undefined' && Office.CoercionType?.Text) || 'text';
              const bodyText = await this._promisify(item.body.getAsync.bind(item.body), textCoercion);
              if (bodyText) {
                const fromMatch = bodyText.match(/(?:\r?\n|^)From:\s*([^<\r\n]+)(?:<([^>\r\n]+)>)?/i);
                if (fromMatch) {
                  displayName = (fromMatch[1] || "").replace(/["']/g, '').trim();
                  email = (fromMatch[2] || "").trim();
                }
              }
            } catch (e) {}
          }
        }
      }
    } catch (e) {
      console.warn("Could not read sender profile:", e);
    }

    if (!displayName && email) {
      const prefix = email.split('@')[0];
      const parts = prefix.split(/[._-]+/).filter(Boolean);
      if (parts.length > 0) {
        displayName = parts.map(p => p.charAt(0).toUpperCase() + p.slice(1)).join(' ');
      }
    }

    return {
      displayName: displayName || "Colleague",
      email: email || ""
    };
  }

  /**
   * Direct Helper: Create a Reply or Reply All Draft with pre-filled AI content.
   * Defaults to replyAll = true to preserve all original CC recipients.
   */
  replyWithDraft(htmlContent, replyAll = true) {
    if (typeof Office === 'undefined' || !Office.context || !Office.context.mailbox) return;
    const item = Office.context.mailbox.item;
    if (!item) return;

    const sanitizedHtml = `<div style="font-family:'Segoe UI',Calibri,Arial,sans-serif; font-size:11pt; color:#201f1e; line-height:1.5;">${htmlContent}</div>`;

    if (replyAll && typeof item.displayReplyAllForm === 'function') {
      item.displayReplyAllForm({ htmlBody: sanitizedHtml });
    } else if (typeof item.displayReplyForm === 'function') {
      item.displayReplyForm({ htmlBody: sanitizedHtml });
    } else if (this.isComposeMode()) {
      this.insertContent(htmlContent, '', 'replace_draft');
    }
  }

  /**
   * Direct Helper: Create a Forward Draft with pre-filled AI content.
   * Explicitly sets toRecipients: [] to ensure the To: field is empty.
   */
  forwardWithDraft(htmlContent) {
    if (typeof Office === 'undefined' || !Office.context || !Office.context.mailbox) return;
    const item = Office.context.mailbox.item;
    if (!item) return;

    const sanitizedHtml = `<div style="font-family:'Segoe UI',Calibri,Arial,sans-serif; font-size:11pt; color:#201f1e; line-height:1.5;">${htmlContent}</div>`;
    if (typeof item.displayForwardForm === 'function') {
      item.displayForwardForm({ toRecipients: [], htmlBody: sanitizedHtml });
    } else if (this.isComposeMode()) {
      this.insertContent(htmlContent, '', 'smart');
    }
  }

  /**
   * Launch Outlook New Appointment form pre-populated with extracted action items.
   */
  createCalendarInvite(appointmentData = {}) {
    if (typeof Office === 'undefined' || !Office.context || !Office.context.mailbox) return;
    if (typeof Office.context.mailbox.displayNewAppointmentForm === 'function') {
      const item = Office.context.mailbox.item;
      let itemSubject = "";
      if (item && typeof item.subject === 'string') {
        itemSubject = item.subject;
      }
      const subject = (appointmentData.subject || (itemSubject ? `Briefing: ${itemSubject}` : "Pre-Meeting Briefing & Agenda")).trim();
      
      // Clean body formatting: support rich HTML with plain text fallback
      let rawBody = appointmentData.body || appointmentData.rawText || "";
      let htmlBody = "";
      if (rawBody) {
        if (rawBody.trim().startsWith("<div") || rawBody.trim().startsWith("<table") || rawBody.trim().startsWith("<p")) {
          htmlBody = `<div style="font-family:'Segoe UI',Calibri,Arial,sans-serif; font-size:11pt; color:#201f1e; line-height:1.5;">${rawBody}</div>`;
        } else {
          htmlBody = `<div style="font-family:'Segoe UI',Calibri,Arial,sans-serif; font-size:11pt; color:#201f1e; line-height:1.5;">${rawBody.replace(/\n/g, '<br/>')}</div>`;
        }
      }
      const plainTextBody = appointmentData.rawText || rawBody.replace(/<[^>]*>/g, "").trim();

      // Gracefully validate Start and End Date objects
      let start = appointmentData.start instanceof Date ? appointmentData.start : (appointmentData.start ? new Date(appointmentData.start) : null);
      if (!start || isNaN(start.getTime())) {
        if (item && item.start) {
          start = item.start instanceof Date ? item.start : new Date(item.start);
        }
        if (!start || isNaN(start.getTime())) {
          start = new Date(Date.now() + 24 * 60 * 60 * 1000); // Tomorrow
        }
      }

      let end = appointmentData.end instanceof Date ? appointmentData.end : (appointmentData.end ? new Date(appointmentData.end) : null);
      if (!end || isNaN(end.getTime()) || end <= start) {
        if (item && item.end) {
          end = item.end instanceof Date ? item.end : new Date(item.end);
        }
        if (!end || isNaN(end.getTime()) || end <= start) {
          end = new Date(start.getTime() + 30 * 60 * 1000); // +30 mins
        }
      }

      // Helper to normalize attendee list
      const normalizeAttendees = (attendeeList) => {
        if (!attendeeList) return [];
        if (typeof attendeeList === 'string') {
          return attendeeList.split(/[,;]+/).map(s => s.trim()).filter(Boolean);
        }
        if (Array.isArray(attendeeList)) {
          return attendeeList.map(a => {
            if (!a) return "";
            if (typeof a === 'string') return a.trim();
            return (a.emailAddress || a.address || a.displayName || a.name || "").trim();
          }).filter(Boolean);
        }
        return [];
      };

      // Normalize attendees gracefully to string[] of email addresses
      const rawRequired = appointmentData.requiredAttendees || appointmentData.attendees;
      let requiredAttendees = normalizeAttendees(rawRequired);
      let optionalAttendees = normalizeAttendees(appointmentData.optionalAttendees);

      // Fallback: If no attendees specified, infer from active email thread or appointment
      if (requiredAttendees.length === 0 && item) {
        if (item.requiredAttendees && Array.isArray(item.requiredAttendees)) {
          requiredAttendees = normalizeAttendees(item.requiredAttendees);
        }
        if (item.from) {
          const fromAddr = item.from.emailAddress || item.from.address || (typeof item.from === 'string' ? item.from : '');
          if (fromAddr) requiredAttendees.push(fromAddr.trim());
        }
        if (item.organizer) {
          const orgAddr = item.organizer.emailAddress || item.organizer.address || (typeof item.organizer === 'string' ? item.organizer : '');
          if (orgAddr) requiredAttendees.push(orgAddr.trim());
        }
        if (Array.isArray(item.to)) {
          item.to.forEach(r => {
            const addr = r.emailAddress || r.address || (typeof r === 'string' ? r : '');
            if (addr) requiredAttendees.push(addr.trim());
          });
        }
        requiredAttendees = [...new Set(requiredAttendees)].filter(Boolean);
      }

      if (optionalAttendees.length === 0 && item && item.optionalAttendees && Array.isArray(item.optionalAttendees)) {
        optionalAttendees = normalizeAttendees(item.optionalAttendees);
      }
      optionalAttendees = [...new Set(optionalAttendees)].filter(email => email && !requiredAttendees.includes(email));

      const location = (appointmentData.location || (item && typeof item.location === 'string' ? item.location : "") || "Microsoft Teams Meeting").trim();

      const formParams = {
        subject,
        body: htmlBody || plainTextBody,
        requiredAttendees,
        location,
        start,
        end
      };
      if (optionalAttendees.length > 0) {
        formParams.optionalAttendees = optionalAttendees;
      }

      try {
        Office.context.mailbox.displayNewAppointmentForm(formParams);
      } catch (err) {
        console.warn("displayNewAppointmentForm with HTML failed, retrying with plain text body:", err);
        try {
          formParams.body = plainTextBody;
          Office.context.mailbox.displayNewAppointmentForm(formParams);
        } catch (retryErr) {
          console.error("displayNewAppointmentForm failed:", retryErr);
        }
      }
    } else {
      console.warn("displayNewAppointmentForm is not supported in this Outlook host environment.");
    }
  }

  /**
   * Scan for in-email @gemini commands (Compose Mode only).
   * Prevents accidental cursor displacement and race conditions via immediate placeholder
   * replacement and unique marker anchoring.
   */
  async checkInDocumentCommands(forceRun = false, callbacks = {}) {
    if (!this.isComposeMode()) {
      if (callbacks.onStatus) callbacks.onStatus("");
      return;
    }

    if (this._isProcessingCommand) return;

    try {
      const selectedText = await this.getSelectedText();
      if (selectedText && selectedText.toLowerCase().includes("@gemini")) {
        const geminiIdx = selectedText.toLowerCase().indexOf("@gemini");
        const userPrompt = selectedText.substring(geminiIdx + 7).replace(/^[:\s]+/, "").trim();

        if (userPrompt.length >= 2 && callbacks.executePrompt) {
          this._isProcessingCommand = true;
          const markerId = "GEMINI_MARKER_" + Date.now();
          const placeholderText = `⚡ [Gemini is generating response... ${markerId}]`;

          if (callbacks.onStatus) callbacks.onStatus(`Found: "@gemini ${userPrompt.substring(0, 20)}..."`);

          // Immediately replace selection with anchored marker placeholder to prevent cursor drift
          const item = Office.context.mailbox.item;
          const textCoercion = (typeof Office !== 'undefined' && Office.CoercionType?.Text) || 'text';
          const htmlCoercion = (typeof Office !== 'undefined' && Office.CoercionType?.Html) || 'html';

          if (item && item.body && typeof item.body.setSelectedDataAsync === 'function') {
            try {
              await this._promisify(
                item.body.setSelectedDataAsync.bind(item.body),
                placeholderText,
                { coercionType: textCoercion }
              );
            } catch (selErr) {
              console.warn("Failed to set placeholder for @gemini selection:", selErr);
            }
          }

          let formattedHtml = null;
          try {
            formattedHtml = await callbacks.executePrompt(userPrompt);
          } catch (execErr) {
            console.error("Execute prompt error:", execErr);
          }

          if (formattedHtml) {
            let replacedInBody = false;
            // Attempt to replace the marker in place if full body setAsync is supported
            if (item && item.body && typeof item.body.getAsync === 'function' && typeof item.body.setAsync === 'function') {
              try {
                const currentBody = await this._promisify(
                  item.body.getAsync.bind(item.body),
                  htmlCoercion
                );
                if (currentBody && currentBody.includes(markerId)) {
                  const sanitizedHtml = `<div style="font-family:'Segoe UI',Calibri,Arial,sans-serif; font-size:11pt; color:#201f1e; line-height:1.5;">${formattedHtml}</div>`;
                  const escapedMarker = markerId.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
                  const markerRegex = new RegExp(`(?:⚡|&#9889;)?\\s*\\[?Gemini is generating response\\.\\.\\.\\s*${escapedMarker}\\]?|${escapedMarker}`, 'g');
                  const updatedBody = currentBody.replace(markerRegex, sanitizedHtml);
                  await this._promisify(
                    item.body.setAsync.bind(item.body),
                    updatedBody,
                    { coercionType: htmlCoercion }
                  );
                  replacedInBody = true;
                }
              } catch (bodyErr) {
                console.warn("Marker body in-place replacement warning:", bodyErr);
              }
            }

            // Fallback if marker replacement could not be performed
            if (!replacedInBody) {
              await this.insertContent(formattedHtml, "", "replace");
            }
          } else {
            // If generation failed, clean up the marker and restore original command text
            if (item && item.body && typeof item.body.getAsync === 'function' && typeof item.body.setAsync === 'function') {
              try {
                const currentBody = await this._promisify(
                  item.body.getAsync.bind(item.body),
                  htmlCoercion
                );
                if (currentBody && currentBody.includes(markerId)) {
                  const escapedMarker = markerId.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
                  const markerRegex = new RegExp(`(?:⚡|&#9889;)?\\s*\\[?Gemini is generating response\\.\\.\\.\\s*${escapedMarker}\\]?|${escapedMarker}`, 'g');
                  const restoredBody = currentBody.replace(markerRegex, `@gemini ${userPrompt}`);
                  await this._promisify(
                    item.body.setAsync.bind(item.body),
                    restoredBody,
                    { coercionType: htmlCoercion }
                  );
                }
              } catch (cleanupErr) {
                console.warn("Marker cleanup warning:", cleanupErr);
              }
            }
          }
        }
      } else {
        if (callbacks.onStatus) callbacks.onStatus("");
      }
    } catch (err) {
      console.warn("Outlook in-doc command check error:", err);
    } finally {
      this._isProcessingCommand = false;
    }
  }
}

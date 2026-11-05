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
 * Host Adapter Factory for Microsoft 365 (Word, PowerPoint, Excel)
 * 
 * @author Sathya AG, Principal Architect, Google
 */

import { WordAdapter } from './WordAdapter.js';
import { PPTAdapter } from './PPTAdapter.js';
import { ExcelAdapter } from './ExcelAdapter.js';
import { OutlookAdapter } from './OutlookAdapter.js';

export class HostAdapterFactory {
  static getAdapter(info = null) {
    // 1. Direct host from Office.onReady(info)
    let host = info?.host;

    const officeRef = typeof Office !== 'undefined' ? Office : null;

    // 2. Diagnostics host
    try {
      if (!host && officeRef?.context?.diagnostics?.host) {
        host = officeRef.context.diagnostics.host;
      }
    } catch (_) {}

    // 3. Fallback context host
    try {
      if (!host && officeRef?.context?.host) {
        host = officeRef.context.host;
      }
    } catch (_) {}

    // 4. Runtime namespace inspection
    try {
      if (!host) {
        if (typeof PowerPoint !== 'undefined') {
          host = officeRef?.HostType?.PowerPoint || "PowerPoint";
        } else if (typeof Excel !== 'undefined') {
          host = officeRef?.HostType?.Excel || "Excel";
        } else if (typeof Word !== 'undefined') {
          host = officeRef?.HostType?.Word || "Word";
        } else if (officeRef?.context?.mailbox) {
          host = officeRef?.HostType?.Outlook || "Outlook";
        }
      }
    } catch (_) {}

    // 5. Requirements set inspection
    try {
      if (!host && officeRef?.context?.requirements) {
        if (officeRef.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
          host = officeRef?.HostType?.PowerPoint || "PowerPoint";
        } else if (officeRef.context.requirements.isSetSupported('ExcelApi', '1.1')) {
          host = officeRef?.HostType?.Excel || "Excel";
        } else if (officeRef.context.requirements.isSetSupported('WordApi', '1.1')) {
          host = officeRef?.HostType?.Word || "Word";
        } else if (officeRef.context.requirements.isSetSupported('Mailbox', '1.1')) {
          host = officeRef?.HostType?.Outlook || "Outlook";
        }
      }
    } catch (_) {}

    const hostStr = typeof host === 'string' ? host.toLowerCase() : '';

    if (host && (host === officeRef?.HostType?.PowerPoint || host === "PowerPoint" || hostStr.includes("powerpoint"))) {
      return new PPTAdapter();
    }
    if (host && (host === officeRef?.HostType?.Excel || host === "Excel" || hostStr.includes("excel"))) {
      return new ExcelAdapter();
    }
    if (host && (host === officeRef?.HostType?.Word || host === "Word" || hostStr.includes("word"))) {
      return new WordAdapter();
    }
    let hasMailbox = false;
    try {
      hasMailbox = Boolean(officeRef?.context?.mailbox);
    } catch (_) {}
    if ((host && (host === officeRef?.HostType?.Outlook || host === "Outlook" || hostStr.includes("outlook"))) || hasMailbox) {
      return new OutlookAdapter();
    }

    // 6. URL query parameter fallback (?host=outlook)
    try {
      if (typeof window !== 'undefined' && window.location?.search) {
        const params = new URLSearchParams(window.location.search);
        const urlHost = (params.get('host') || '').toLowerCase();
        if (urlHost === 'outlook') return new OutlookAdapter();
      }
    } catch (_) {}

    return new WordAdapter();
  }
}

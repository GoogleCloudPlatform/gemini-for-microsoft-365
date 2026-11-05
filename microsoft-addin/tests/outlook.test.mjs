/**
 * Comprehensive Automated Test Suite for Outlook Add-in & Gemini Proxy
 * Tests all 14 quality and UX enhancements outlined in logs/patch-improve-demo.md
 */

import { test, describe, beforeEach } from 'node:test';
import assert from 'node:assert/strict';
import { OutlookAdapter } from '../src/adapters/OutlookAdapter.js';
import { parseMarkdown } from '../src/core/markdownParser.js';

describe('Outlook Add-in Quality & UX Enhancements', () => {

  describe('1. Unit Tests (UT-1 to UT-5)', () => {

    test('UT-1: OutlookAdapter.getUserProfile & getSenderProfile name extraction with fallback', async () => {
      const adapter = new OutlookAdapter();

      // Case A: Full Office context with displayName
      global.Office = {
        AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
        context: {
          mailbox: {
            userProfile: {
              displayName: 'Jane Doe',
              emailAddress: 'jane.doe@company.com'
            },
            item: {
              from: {
                displayName: 'John Smith',
                emailAddress: 'john.smith@client.com'
              }
            }
          }
        }
      };

      const user1 = adapter.getUserProfile();
      assert.equal(user1.displayName, 'Jane Doe');
      assert.equal(user1.email, 'jane.doe@company.com');

      const sender1 = await adapter.getSenderProfile();
      assert.equal(sender1.displayName, 'John Smith');
      assert.equal(sender1.email, 'john.smith@client.com');

      // Case B: Null/empty displayName with email fallback
      global.Office.context.mailbox.userProfile = {
        displayName: '',
        emailAddress: 'alex.rivera@enterprise.org'
      };
      global.Office.context.mailbox.item.from = {
        displayName: null,
        emailAddress: 'adele.vance@customer.com'
      };

      const user2 = adapter.getUserProfile();
      assert.equal(user2.displayName, 'Alex Rivera');
      assert.equal(user2.email, 'alex.rivera@enterprise.org');

      const sender2 = await adapter.getSenderProfile();
      assert.equal(sender2.displayName, 'Adele Vance');
      assert.equal(sender2.email, 'adele.vance@customer.com');

      // Case C: String from header with angle brackets
      global.Office.context.mailbox.item.from = "Adele Vance <adele.vance@customer.com>";
      const senderAngle = await adapter.getSenderProfile();
      assert.equal(senderAngle.displayName, 'Adele Vance');
      assert.equal(senderAngle.email, 'adele.vance@customer.com');

      // Case D: Compose mode (extracting incoming sender/recipient from item.to)
      global.Office.context.mailbox.item.from = null;
      global.Office.context.mailbox.item.body = { setSelectedDataAsync: () => {} };
      global.Office.context.mailbox.item.to = [{
        displayName: 'Adele Vance',
        emailAddress: 'adele.vance@customer.com'
      }];
      const senderCompose = await adapter.getSenderProfile();
      assert.equal(senderCompose.displayName, 'Adele Vance');
      assert.equal(senderCompose.email, 'adele.vance@customer.com');

      // Case E: Missing item.from and item.to objects
      global.Office.context.mailbox.item.from = null;
      global.Office.context.mailbox.item.to = [];
      global.Office.context.mailbox.item.body = {};
      const sender3 = await adapter.getSenderProfile();
      assert.equal(sender3.displayName, 'Colleague');
    });

    test('UT-2: OutlookAdapter.forwardWithDraft passes empty toRecipients array', () => {
      const adapter = new OutlookAdapter();
      let forwardArgs = null;

      global.Office = {
        context: {
          mailbox: {
            item: {
              body: {},
              displayForwardForm: (args) => {
                forwardArgs = args;
              }
            }
          }
        }
      };

      adapter.forwardWithDraft('<p>Summary text</p>');
      assert.ok(forwardArgs, 'displayForwardForm should be called');
      assert.deepEqual(forwardArgs.toRecipients, [], 'toRecipients must be an empty array');
      assert.match(forwardArgs.htmlBody, /Summary text/, 'htmlBody should contain the summary text');
    });

    test('UT-3: OutlookAdapter.replyWithDraft defaults replyAll to true to preserve CCs', () => {
      const adapter = new OutlookAdapter();
      let replyAllCalled = false;
      let replySingleCalled = false;
      let passedHtml = '';

      global.Office = {
        context: {
          mailbox: {
            item: {
              body: {},
              displayReplyAllForm: (args) => {
                replyAllCalled = true;
                passedHtml = args.htmlBody;
              },
              displayReplyForm: (args) => {
                replySingleCalled = true;
              }
            }
          }
        }
      };

      // Default call should invoke displayReplyAllForm
      adapter.replyWithDraft('<p>Draft content</p>');
      assert.equal(replyAllCalled, true, 'Default replyWithDraft must invoke displayReplyAllForm');
      assert.equal(replySingleCalled, false);
      assert.match(passedHtml, /Draft content/);

      // Explicit single reply
      replyAllCalled = false;
      adapter.replyWithDraft('<p>Single reply</p>', false);
      assert.equal(replySingleCalled, true, 'replyWithDraft with replyAll=false must invoke displayReplyForm');
    });

    test('UT-4: OutlookAdapter.segmentAndCleanEmailThread splits 15-turn threads correctly', () => {
      const adapter = new OutlookAdapter();

      const multiTurnThread = `Hi Team, here is the latest update on the deal.
------------------- Original Message -------------------
From: Adele Vance <adele@contoso.com>
Sent: Tuesday, August 25, 2026 10:00 AM
To: Alex Rivera <alex@company.com>
Subject: Re: Pricing Discussion

Alex, let us finalize the terms today.
________________________________
From: Alex Rivera <alex@company.com>
Sent: Monday, August 24, 2026 4:00 PM
Subject: Pricing Discussion

Here are the initial quotes for review.`;

      const segmented = adapter.segmentAndCleanEmailThread(multiTurnThread);
      assert.match(segmented, /EMAIL THREAD TURN BREAKDOWN/);
      assert.match(segmented, /Message 1 of 3/);
      assert.match(segmented, /Message 2 of 3/);
      assert.match(segmented, /Message 3 of 3/);
      assert.match(segmented, /Adele Vance/);
      assert.match(segmented, /initial quotes/);
    });

    test('UT-5: Coach response parser extracts 2-line assessment and Option A/B', () => {
      const sampleCoachResponse = `### Tone & Clarity Assessment
* **Clarity & Brevity Score:** Moderate
* **Tone & Deal Risk:** The draft is overly defensive regarding delivery timelines, which may trigger unnecessary customer escalations.

=== OPTION A: CONCISE EXECUTIVE ===
Hi Adele,

We are on track for Q3 delivery. The final sign-off is scheduled for Friday.

Best,
Alex

=== OPTION B: DIPLOMATIC ===
Dear Adele,

Thank you for following up. We understand the importance of this timeline and are working diligently to finalize the sign-off by Friday.

Warmly,
Alex`;

      const markerARegex = /(?:[\*\#=]+\s*)?OPTION A\s*[:\-–\(\s]+CONCISE EXECUTIVE[\s\*\#=\)]*/i;
      const markerBRegex = /(?:[\*\#=]+\s*)?OPTION B\s*[:\-–\(\s]+DIPLOMATIC[\s\*\#=\)]*/i;

      const matchA = markerARegex.exec(sampleCoachResponse);
      const matchB = markerBRegex.exec(sampleCoachResponse);

      assert.ok(matchA, 'Option A marker should match');
      assert.ok(matchB, 'Option B marker should match');

      const assessment = sampleCoachResponse.substring(0, matchA.index).trim();
      const optionA = sampleCoachResponse.substring(matchA.index + matchA[0].length, matchB.index).trim();
      const optionB = sampleCoachResponse.substring(matchB.index + matchB[0].length).trim();

      assert.match(assessment, /Clarity & Brevity Score:\*\* Moderate/);
      assert.match(assessment, /Tone & Deal Risk:\*\*/);
      assert.match(optionA, /We are on track for Q3 delivery/);
      assert.match(optionB, /Thank you for following up/);

      // Decorated delimiters check (**=== OPTION A: CONCISE EXECUTIVE ===**)
      const decoratedResponse = `### Tone & Clarity Assessment
* **Clarity & Brevity Score:** High
* **Tone & Deal Risk:** Clean and direct.

**=== OPTION A: CONCISE EXECUTIVE ===**
Short and direct text.

**=== OPTION B: DIPLOMATIC ===**
Polite and diplomatic text.`;

      const matchDecA = markerARegex.exec(decoratedResponse);
      const matchDecB = markerBRegex.exec(decoratedResponse);
      assert.ok(matchDecA, 'Decorated Option A should match');
      assert.ok(matchDecB, 'Decorated Option B should match');
      assert.equal(decoratedResponse.substring(matchDecA.index + matchDecA[0].length, matchDecB.index).trim(), 'Short and direct text.');
      assert.equal(decoratedResponse.substring(matchDecB.index + matchDecB[0].length).trim(), 'Polite and diplomatic text.');
    });

    test('UT-7: OutlookAdapter.getFullDocumentText extracts calendar appointment and meeting metadata', async () => {
      const adapter = new OutlookAdapter();

      global.Office = {
        AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
        context: {
          mailbox: {
            item: {
              itemType: 'appointment',
              subject: 'Q3 Enterprise Architecture & SLA Review',
              organizer: {
                displayName: 'Sarah Connor',
                emailAddress: 'sarah.connor@cyberdyne.com'
              },
              requiredAttendees: [
                { displayName: 'John Connor', emailAddress: 'john@resistance.org' },
                { displayName: 'Kyle Reese', emailAddress: 'kyle@resistance.org' }
              ],
              optionalAttendees: [
                { displayName: 'Miles Dyson', emailAddress: 'miles@cyberdyne.com' }
              ],
              start: new Date('2026-09-15T14:00:00Z'),
              end: new Date('2026-09-15T15:00:00Z'),
              location: 'Microsoft Teams Meeting Room 4B',
              dateTimeCreated: '2026-08-30T10:00:00Z',
              conversationId: 'APPT_CONV_998877',
              body: {
                getAsync: (coercion, cb) => {
                  cb({
                    status: 'succeeded',
                    value: 'Please review attached SLA contract terms and architecture migration roadmap prior to the meeting.'
                  });
                }
              }
            }
          }
        }
      };

      const fullText = await adapter.getFullDocumentText();
      assert.match(fullText, /CALENDAR APPOINTMENT \/ MEETING INVITATION METADATA/);
      assert.match(fullText, /Subject: Q3 Enterprise Architecture & SLA Review/);
      assert.match(fullText, /Organizer: Sarah Connor <sarah\.connor@cyberdyne\.com>/);
      assert.match(fullText, /Required Attendees: John Connor <john@resistance\.org>, Kyle Reese <kyle@resistance\.org>/);
      assert.match(fullText, /Optional Attendees: Miles Dyson <miles@cyberdyne\.com>/);
      assert.match(fullText, /Location: Microsoft Teams Meeting Room 4B/);
      assert.match(fullText, /Start Time:/);
      assert.match(fullText, /End Time:/);
      assert.match(fullText, /Please review attached SLA contract terms/);
    });

    test('UT-8: OutlookAdapter.getSenderProfile extracts organizer from calendar appointment when item.from is missing', async () => {
      const adapter = new OutlookAdapter();

      global.Office = {
        AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
        context: {
          mailbox: {
            item: {
              itemType: 'appointment',
              from: null,
              organizer: {
                displayName: 'Dr. Emmett Brown',
                emailAddress: 'doc.brown@delorean.org'
              },
              body: {}
            }
          }
        }
      };

      const sender = await adapter.getSenderProfile();
      assert.equal(sender.displayName, 'Dr. Emmett Brown');
      assert.equal(sender.email, 'doc.brown@delorean.org');
    });

    test('UT-9: OutlookAdapter.getFullDocumentText in Read mode preserves Unknown Sender without attributing to userProfile', async () => {
      const adapter = new OutlookAdapter();

      global.Office = {
        AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
        context: {
          mailbox: {
            userProfile: {
              displayName: 'Jane Doe',
              emailAddress: 'jane.doe@company.com'
            },
            item: {
              itemType: 'message',
              subject: 'Customer SLA Inquiry',
              from: null,
              to: [{ displayName: 'Jane Doe', emailAddress: 'jane.doe@company.com' }],
              body: {
                getAsync: (coercion, cb) => {
                  cb({ status: 'succeeded', value: 'What are the response time SLAs for P1 incidents?' });
                }
              }
            }
          }
        }
      };

      const fullText = await adapter.getFullDocumentText();
      assert.match(fullText, /EMAIL ENVELOPE METADATA/);
      assert.match(fullText, /From: Unknown Sender/);
      assert.equal(fullText.includes('From: Jane Doe'), false, 'Read mode must not set From to user profile');
    });

    test('UT-10: OutlookAdapter.createCalendarInvite inherits subject, times, and attendees from active appointment context', () => {
      const adapter = new OutlookAdapter();
      let capturedForm = null;

      global.Office = {
        context: {
          mailbox: {
            displayNewAppointmentForm: (form) => {
              capturedForm = form;
            },
            item: {
              itemType: 'appointment',
              subject: 'Contract Renewal Discussion',
              start: new Date('2026-10-01T09:00:00Z'),
              end: new Date('2026-10-01T10:00:00Z'),
              location: 'Conference Room 10A',
              organizer: { displayName: 'Alice', emailAddress: 'alice@corp.com' },
              requiredAttendees: [{ displayName: 'Bob', emailAddress: 'bob@corp.com' }],
              optionalAttendees: [{ displayName: 'Charlie', emailAddress: 'charlie@corp.com' }]
            }
          }
        }
      };

      adapter.createCalendarInvite({ body: '<p>Briefing Notes</p>' });
      assert.ok(capturedForm, 'displayNewAppointmentForm must be called');
      assert.equal(capturedForm.subject, 'Briefing: Contract Renewal Discussion');
      assert.equal(capturedForm.location, 'Conference Room 10A');
      assert.equal(capturedForm.start.toISOString(), '2026-10-01T09:00:00.000Z');
      assert.equal(capturedForm.end.toISOString(), '2026-10-01T10:00:00.000Z');
      assert.ok(capturedForm.requiredAttendees.includes('bob@corp.com'));
      assert.ok(capturedForm.requiredAttendees.includes('alice@corp.com'));
      assert.ok(capturedForm.optionalAttendees.includes('charlie@corp.com'));
    });

    test('UT-11: OutlookAdapter.getFullDocumentText supports Appointment Compose-mode async metadata getters (.getAsync)', async () => {
      const adapter = new OutlookAdapter();

      global.Office = {
        AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
        MailboxEnums: { ItemType: { Appointment: 'appointment' } },
        context: {
          mailbox: {
            item: {
              itemType: 'appointment',
              subject: {
                getAsync: (cb) => cb({ status: 'succeeded', value: 'Async Compose Appointment' })
              },
              organizer: {
                getAsync: (cb) => cb({ status: 'succeeded', value: { displayName: 'Async Organizer', emailAddress: 'org@company.com' } })
              },
              requiredAttendees: {
                getAsync: (cb) => cb({ status: 'succeeded', value: [{ displayName: 'Req User', emailAddress: 'req@company.com' }] })
              },
              optionalAttendees: {
                getAsync: (cb) => cb({ status: 'succeeded', value: [{ displayName: 'Opt User', emailAddress: 'opt@company.com' }] })
              },
              start: {
                getAsync: (cb) => cb({ status: 'succeeded', value: new Date('2026-11-10T10:00:00Z') })
              },
              end: {
                getAsync: (cb) => cb({ status: 'succeeded', value: new Date('2026-11-10T11:00:00Z') })
              },
              location: {
                getAsync: (cb) => cb({ status: 'succeeded', value: 'Boardroom East' })
              },
              body: {
                setSelectedDataAsync: () => {},
                getAsync: (_coercion, cb) => cb({ status: 'succeeded', value: 'Compose appointment agenda text.' })
              }
            }
          }
        }
      };

      const fullText = await adapter.getFullDocumentText();
      assert.match(fullText, /CALENDAR APPOINTMENT \/ MEETING INVITATION METADATA/);
      assert.match(fullText, /Subject: Async Compose Appointment/);
      assert.match(fullText, /Organizer: Async Organizer <org@company\.com>/);
      assert.match(fullText, /Required Attendees: Req User <req@company\.com>/);
      assert.match(fullText, /Optional Attendees: Opt User <opt@company\.com>/);
      assert.match(fullText, /Location: Boardroom East/);
      assert.match(fullText, /Compose appointment agenda text\./);
    });

    test('UT-12: OutlookAdapter.checkInDocumentCommands handles Read mode, re-entrancy lock, marker replacement, and rollback on failure', async () => {
      const adapter = new OutlookAdapter();

      // 1. Read mode guard
      let statusMsg = 'initial';
      global.Office = {
        context: {
          mailbox: {
            item: { body: {} } // Read mode
          }
        }
      };
      await adapter.checkInDocumentCommands(false, {
        onStatus: (msg) => { statusMsg = msg; }
      });
      assert.equal(statusMsg, '');

      // 2. Re-entrancy lock guard
      adapter._isProcessingCommand = true;
      let executeCalled = false;
      await adapter.checkInDocumentCommands(false, {
        executePrompt: async () => { executeCalled = true; return '<p>Result</p>'; }
      });
      assert.equal(executeCalled, false, 'Re-entrancy lock must prevent concurrent execution');
      adapter._isProcessingCommand = false;

      // 3. Compose mode: successful marker injection and in-place body replacement
      let currentBodyHtml = '<p>Hello team,</p><p>@gemini draft a polite follow-up</p>';
      global.Office = {
        AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
        CoercionType: { Text: 'text', Html: 'html' },
        context: {
          mailbox: {
            item: {
              getSelectedDataAsync: (_coercion, cb) => {
                cb({ status: 'succeeded', value: '@gemini draft a polite follow-up' });
              },
              body: {
                setSelectedDataAsync: (placeholder, _opts, cb) => {
                  currentBodyHtml = currentBodyHtml.replace('@gemini draft a polite follow-up', placeholder);
                  cb({ status: 'succeeded' });
                },
                getAsync: (_coercion, cb) => {
                  cb({ status: 'succeeded', value: currentBodyHtml });
                },
                setAsync: (updatedHtml, _opts, cb) => {
                  currentBodyHtml = updatedHtml;
                  cb({ status: 'succeeded' });
                }
              }
            }
          }
        }
      };

      await adapter.checkInDocumentCommands(false, {
        executePrompt: async (prompt) => {
          assert.equal(prompt, 'draft a polite follow-up');
          return '<strong>Polite follow-up text</strong>';
        }
      });

      assert.match(currentBodyHtml, /Polite follow-up text/, 'Marker must be replaced with generated HTML');
      assert.equal(currentBodyHtml.includes('GEMINI_MARKER_'), false, 'Temporary marker must be removed');
      assert.equal(adapter._isProcessingCommand, false, 'Lock must be released in finally block');

      // 4. Rollback on null/failed generation restores original @gemini command text
      currentBodyHtml = '<p>@gemini summarize key risks</p>';
      global.Office.context.mailbox.item.getSelectedDataAsync = (_coercion, cb) => {
        cb({ status: 'succeeded', value: '@gemini summarize key risks' });
      };
      global.Office.context.mailbox.item.body.setSelectedDataAsync = (placeholder, _opts, cb) => {
        currentBodyHtml = currentBodyHtml.replace('@gemini summarize key risks', placeholder);
        cb({ status: 'succeeded' });
      };

      await adapter.checkInDocumentCommands(false, {
        executePrompt: async () => null
      });

      assert.match(currentBodyHtml, /@gemini summarize key risks/, 'Failed generation must roll back marker to original @gemini command');
      assert.equal(currentBodyHtml.includes('GEMINI_MARKER_'), false);
    });

    test('UT-13: OutlookAdapter.createCalendarInvite recovers from invalid/reversed dates and retries with plain text when HTML form fails', () => {
      const adapter = new OutlookAdapter();
      const callHistory = [];

      global.Office = {
        context: {
          mailbox: {
            displayNewAppointmentForm: (formParams) => {
              callHistory.push({ ...formParams });
              if (callHistory.length === 1) {
                throw new Error('HTML body not supported by legacy client');
              }
            },
            item: {
              subject: 'Quarterly Review'
            }
          }
        }
      };

      // Pass end date earlier than start date and verify automatic +30 min recovery + HTML-to-plaintext retry
      const validStart = new Date('2026-12-01T15:00:00Z');
      const invalidEarlierEnd = new Date('2026-12-01T14:00:00Z');

      adapter.createCalendarInvite({
        subject: 'Action Items Sync',
        body: '<p>Review Q4 milestones</p>',
        start: validStart,
        end: invalidEarlierEnd,
        requiredAttendees: 'alice@corp.com; bob@corp.com'
      });

      assert.equal(callHistory.length, 2, 'Must retry with plain text when first displayNewAppointmentForm call throws');
      assert.match(callHistory[0].body, /<div/, 'First attempt uses styled HTML wrapper');
      assert.equal(callHistory[1].body, 'Review Q4 milestones', 'Second retry uses stripped plain text body');
      assert.equal(
        callHistory[1].end.getTime() - callHistory[1].start.getTime(),
        30 * 60 * 1000,
        'End time <= start time must be automatically adjusted to start + 30 minutes'
      );
      assert.deepEqual(callHistory[1].requiredAttendees, ['alice@corp.com', 'bob@corp.com']);
    });

  });

  describe('2. Backend Proxy Limits (UT-6)', () => {

    test('UT-6: Express proxy index.js configures 10mb/50mb body parser limits', async () => {
      const fs = await import('node:fs');
      const proxyCode = fs.readFileSync(new URL('../../geminiproxy/index.js', import.meta.url), 'utf-8');

      assert.match(proxyCode, /app\.use\(express\.json\(\{\s*limit:\s*['"](?:10|50)mb['"]\s*\}\)\)/, 'express.json must have 10mb or 50mb limit');
      assert.match(proxyCode, /app\.use\(express\.urlencoded\(\{\s*(?:limit:\s*['"](?:10|50)mb['"],\s*extended:\s*true|extended:\s*true,\s*limit:\s*['"](?:10|50)mb['"])\s*\}\)\)/, 'express.urlencoded must have 10mb or 50mb limit');
    });

  });

  describe('3. Integration & HTML/CSS Conformance (IT-1 to IT-5)', () => {

    test('IT-1: Summarize tab HTML has correct action buttons', async () => {
      const fs = await import('node:fs');
      const html = fs.readFileSync(new URL('../src/taskpane/taskpane.html', import.meta.url), 'utf-8');

      assert.match(html, /<button id="btnCopySummary"[^>]*>(?:📋\s*|<svg[\s\S]*?<\/svg>\s*)Copy Summary<\/button>/);
      assert.equal(html.includes('id="btnInsertSummary"'), false, 'btnInsertSummary must be removed from HTML');
      assert.match(html, /<button id="btnReplySummary"[^>]*>👥 Reply All with Summary<\/button>/);
    });

    test('IT-2: Reply tab HTML action buttons are reordered and Replace button removed', async () => {
      const fs = await import('node:fs');
      const html = fs.readFileSync(new URL('../src/taskpane/taskpane.html', import.meta.url), 'utf-8');

      // Verify btnReplaceInEmail is removed
      assert.equal(html.includes('id="btnReplaceInEmail"'), false, 'btnReplaceInEmail must be removed from HTML');

      // Verify button order in Reply tab
      const replySection = html.substring(html.indexOf('id="panelReply"'), html.indexOf('id="panelCoach"'));
      const idxReplyAll = replySection.indexOf('id="btnOpenReplyAll"');
      const idxReply = replySection.indexOf('id="btnOpenReply"');
      const idxCopy = replySection.indexOf('id="btnCopyReply"');

      assert.ok(idxReplyAll !== -1, 'btnOpenReplyAll must exist');
      assert.ok(idxReply !== -1, 'btnOpenReply must exist');
      assert.ok(idxCopy !== -1, 'btnCopyReply must exist');
      assert.ok(idxReplyAll < idxReply, 'btnOpenReplyAll must come before btnOpenReply');
      assert.ok(idxReply < idxCopy, 'btnOpenReply must come before btnCopyReply');
    });

    test('IT-3: taskpane.css removes rigid max-height from .coach-option-body', async () => {
      const fs = await import('node:fs');
      const css = fs.readFileSync(new URL('../src/taskpane/taskpane.css', import.meta.url), 'utf-8');

      const bodyClassMatch = css.match(/\.coach-option-body\s*\{([^}]+)\}/);
      assert.ok(bodyClassMatch, '.coach-option-body CSS rule must exist');
      assert.equal(bodyClassMatch[1].includes('max-height: 160px'), false, '.coach-option-body must not lock max-height to 160px');
    });

    test('IT-4: taskpane.js isolates Ask chat history and history=[] on single-shot tabs', async () => {
      const fs = await import('node:fs');
      const js = fs.readFileSync(new URL('../src/taskpane/taskpane.js', import.meta.url), 'utf-8');

      // Check streamOutlookWorkflow usage
      assert.match(js, /async function streamOutlookWorkflow/, 'streamOutlookWorkflow helper must be defined');

      // Check isolated history in single-shot tabs
      assert.match(js, /handleOutlookSummarize[\s\S]*?options:\s*\{\s*history:\s*\[\]\s*\}/, 'Summarize tab must pass empty history []');
      assert.match(js, /handleOutlookReply[\s\S]*?options:\s*\{\s*history:\s*\[\]\s*\}/, 'Reply tab must pass empty history []');
      assert.match(js, /handleOutlookCoachEmail[\s\S]*?\{\s*history:\s*\[\]\s*\}/, 'Coach tab must pass empty history []');
      assert.match(js, /handleOutlookActionItems[\s\S]*?options:\s*\{\s*history:\s*\[\]\s*\}/, 'Actions tab must pass empty history []');

      // Check Ask tab maintains its own history
      assert.match(js, /let askChatHistory\s*=\s*\[\];/, 'askChatHistory must be declared');
      assert.match(js, /handleOutlookAsk[\s\S]*?history:\s*askChatHistory,\s*sessionId:\s*askSessionId/, 'Ask tab must pass askChatHistory');
    });

    test('IT-5: taskpane.js grounds sender vs author identity in reply prompt', async () => {
      const fs = await import('node:fs');
      const js = fs.readFileSync(new URL('../src/taskpane/taskpane.js', import.meta.url), 'utf-8');

      assert.match(js, /getUserProfile\(\)/, 'Must call getUserProfile');
      assert.match(js, /getSenderProfile\(\)/, 'Must call getSenderProfile');
      assert.match(js, /Incoming Sender Name/, 'Prompt must include Incoming Sender Name');
      assert.match(js, /Current User \/ Author Name/, 'Prompt must include Current User / Author Name');
      assert.match(js, /NEVER sign off with the incoming sender's name/i, 'Prompt must include negative sign-off constraint');
    });

    test('IT-6: [OUT-02] Grounded Smart Reply Drafter prompt, HTML, and editor population conformance', async () => {
      const fs = await import('node:fs');
      const html = fs.readFileSync(new URL('../src/taskpane/taskpane.html', import.meta.url), 'utf-8');
      const js = fs.readFileSync(new URL('../src/taskpane/taskpane.js', import.meta.url), 'utf-8');

      // 1. Verify HTML elements for Draft Grounded Reply
      assert.match(html, /<button id="btnGenerateReply"[^>]*>⚡ Draft Grounded Reply<\/button>/, 'Button must be labeled Draft Grounded Reply');
      assert.match(html, /id="panelReply"/, 'Reply panel must exist');
      assert.match(html, /id="replyIntentText"/, 'replyIntentText input must exist');

      // 2. Verify prompt grounding in contract data stores and SLA clauses
      assert.match(js, /Customer Success Representative/i, 'Prompt must configure Customer Success persona');
      assert.match(js, /contract data stores/i, 'Prompt must instruct grounding in contract data stores');
      assert.match(js, /exact SLA clauses/i, 'Prompt must instruct citing exact SLA clauses');
      assert.match(js, /enableGrounding:\s*true/, 'handleOutlookReply must enable grounding');

      // 3. Verify compose mode editor population
      assert.match(js, /insertContent\(formattedHtml,\s*rawResult,\s*["']replace_draft["']\)/, 'Must populate active compose body in compose mode');
    });

    test('IT-7: [OUT-03] Meeting Agenda & Briefing Prep prompt, HTML, and calendar invite conformance', async () => {
      const fs = await import('node:fs');
      const html = fs.readFileSync(new URL('../src/taskpane/taskpane.html', import.meta.url), 'utf-8');
      const js = fs.readFileSync(new URL('../src/taskpane/taskpane.js', import.meta.url), 'utf-8');

      // 1. Verify HTML buttons in panelActionItems
      assert.match(html, /<button id="btnPrepareBriefing"[^>]*>📅 Prepare Meeting Briefing<\/button>/, 'Prepare Meeting Briefing button must exist');
      assert.match(html, /<button id="btnExtractActionItems"/, 'Extract Action Items button must exist');

      // 2. Verify controller handler binding
      assert.match(js, /btnPrepareBriefing\.onclick\s*=\s*handleOutlookMeetingBriefing/, 'btnPrepareBriefing must be bound to handleOutlookMeetingBriefing');

      // 3. Verify briefing prompt structure
      assert.match(js, /Pre-Meeting Briefing & Strategic Objective/i, 'Briefing prompt must require Strategic Objective');
      assert.match(js, /Structured Meeting Agenda/i, 'Briefing prompt must require Structured Meeting Agenda table');
      assert.match(js, /Key Decisions Required/i, 'Briefing prompt must require Key Decisions Required');
      assert.match(js, /Attendee Preparation/i, 'Briefing prompt must require Attendee Preparation notes');

      // 4. Verify grounding enabled and calendar invite creation
      assert.match(js, /handleOutlookMeetingBriefing[\s\S]*?enableGrounding:\s*true/, 'Meeting briefing must enable grounding');
      assert.match(js, /createCalendarInvite\(\{\s*subject:\s*["']Pre-Meeting Briefing & Agenda["']/, 'Must wire createCalendarInvite with briefing subject');
    });

    test('IT-8: Outlook manifests declare Message and Appointment command surfaces', async () => {
      const { execSync } = await import('node:child_process');
      const fs = await import('node:fs');
      const os = await import('node:os');
      const path = await import('node:path');

      const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'outlook-manifest-'));
      const devXmlPath = path.join(tempDir, 'manifest.dev.xml');
      const prodXmlPath = path.join(tempDir, 'manifest.prod.xml');
      const scriptPath = path.resolve('../scripts/generate_manifest.py');

      try {
        execSync(`python3 "${scriptPath}" --dev --host outlook --output "${devXmlPath}"`);
        execSync(`python3 "${scriptPath}" --host outlook --frontend-url https://gemini-frontend-test.run.app --client-id 11111111-2222-3333-4444-555555555555 --output "${prodXmlPath}"`);

        const devXml = fs.readFileSync(devXmlPath, 'utf-8');
        const prodXml = fs.readFileSync(prodXmlPath, 'utf-8');

        for (const xml of [prodXml, devXml]) {
          assert.match(xml, /<ExtensionPoint xsi:type="MessageReadCommandSurface">/, 'Must declare MessageReadCommandSurface');
          assert.match(xml, /<ExtensionPoint xsi:type="MessageComposeCommandSurface">/, 'Must declare MessageComposeCommandSurface');
          assert.match(xml, /<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">/, 'Must declare AppointmentAttendeeCommandSurface');
          assert.match(xml, /<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">/, 'Must declare AppointmentOrganizerCommandSurface');
        }
      } finally {
        fs.rmSync(tempDir, { recursive: true, force: true });
      }
    });

    test('IT-9: Unified Outlook UI uses shared #officeDocHostContainer with compact action chips, slash commands, and ItemChanged reset', async () => {
      const fs = await import('node:fs');
      const html = fs.readFileSync(new URL('../src/taskpane/taskpane.html', import.meta.url), 'utf-8');
      const js = fs.readFileSync(new URL('../src/taskpane/taskpane.js', import.meta.url), 'utf-8');

      // 0. Verify taskpane.html does not force #officeDocHostContainer to display:none!important
      assert.equal(html.includes('#officeDocHostContainer{display:none!important;}'), false, 'taskpane.html must not hide #officeDocHostContainer with !important');

      // 1. Verify initTaskpane displays #officeDocHostContainer for Outlook without early return
      assert.match(js, /if\s*\(docContainer\)\s*docContainer\.style\.display\s*=\s*["']flex["']/, 'Outlook must display #officeDocHostContainer');
      assert.match(js, /if\s*\(outlookContainer\)\s*outlookContainer\.style\.display\s*=\s*["']none["']/, 'Outlook must hide legacy #outlookHostContainer');

      // 2. Verify Option B compact quick-action chips in renderAdaptiveActionChips (Briefing removed per team feedback)
      assert.match(js, /id="chipOutlookSummarize"/, 'Must render Summarize quick chip');
      assert.match(js, /id="chipOutlookReply"/, 'Must render Draft Reply quick chip');
      assert.match(js, /id="chipOutlookCoach"/, 'Must render Coach quick chip');
      assert.equal(js.includes('id="chipOutlookBriefing"'), false, 'Must NOT render Briefing quick chip');
      assert.match(js, /id="chipOutlookActions"/, 'Must render Action Items quick chip');

      // 3. Verify OUTLOOK_SLASH_COMMANDS, buildOutlookActionPrompt, and runOutlookChatAction (without /briefing)
      assert.match(js, /export const OUTLOOK_SLASH_COMMANDS/, 'Must export OUTLOOK_SLASH_COMMANDS');
      assert.equal(/OUTLOOK_SLASH_COMMANDS\s*=\s*\[[\s\S]*?\/briefing[\s\S]*?\]/.test(js), false, 'OUTLOOK_SLASH_COMMANDS must not include /briefing');
      assert.match(js, /export function buildOutlookActionPrompt/, 'Must export shared buildOutlookActionPrompt helper');
      assert.match(js, /export async function runOutlookChatAction/, 'Must export runOutlookChatAction');
      assert.equal(js.includes('local_${Date.now()}'), false, 'executeGeminiWorkflow must not mint local_ session IDs');

      // 4. Verify Option A email switch (ItemChanged) resets chat session and re-adapts UI
      assert.match(js, /function handleOutlookItemChanged\(\)\s*\{[\s\S]*?resetChatSession\(\);[\s\S]*?adaptUIForHost\(["']Outlook["']\);/, 'handleOutlookItemChanged must reset chat session and re-adapt Outlook UI');

      // 5. Verify Read mode starter action buttons (#docToolsBar) hide after first conversation and restore on reset
      assert.match(js, /if\s*\(isOutlookReadMode\s*&&\s*docToolsBar\)\s*docToolsBar\.style\.display\s*=\s*["']none["']/, 'updateChatDisplayState must hide #docToolsBar in Outlook Read mode once hasMessages is true');
      assert.match(js, /if\s*\(isOutlook\s*&&\s*docToolsBar\)\s*docToolsBar\.style\.display\s*=\s*["']flex["']/, 'updateChatDisplayState must restore #docToolsBar when hasMessages is false');

      // 6. Verify table-scoped horizontal scrollbar wrapper in wrapOutlookTables and CSS (strictly Outlook-scoped)
      assert.match(js, /export function wrapOutlookTables/, 'Must export wrapOutlookTables helper for Outlook table scrolling');
      assert.match(html, /body\.host-outlook\s+\.table-scroll-wrapper\s*\{[\s\S]*?overflow-x:\s*auto/, 'taskpane.html must scope .table-scroll-wrapper to body.host-outlook');

      // 7. Verify compact flex-wrap post-response action button sizing (strictly Outlook-scoped)
      assert.match(html, /body\.host-outlook\s+\.footer-primary-actions\s*\{[\s\S]*?flex-wrap:\s*wrap/, 'footer-primary-actions flex-wrap must be scoped to body.host-outlook');
      assert.match(html, /body\.host-outlook\s+\.pill-btn\s*\{[\s\S]*?padding:\s*5px 10px[\s\S]*?font-size:\s*11\.5px/, 'pill-btn compact sizing must be scoped to body.host-outlook');

      // 8. Verify Outlook Read/Compose mode status text is removed from #debugStatus and OutlookAdapter
      const adapterJs = fs.readFileSync(new URL('../src/adapters/OutlookAdapter.js', import.meta.url), 'utf-8');
      assert.equal(adapterJs.includes('Outlook Read Mode Ready'), false, 'OutlookAdapter must not display Outlook Read Mode Ready');
      assert.equal(adapterJs.includes('Outlook Compose Mode Ready'), false, 'OutlookAdapter must not display Outlook Compose Mode Ready');
      assert.equal(js.includes('isCompose ? "Compose Mode" : "Read Mode"'), false, 'taskpane.js must not display Read Mode / Compose Mode in #debugStatus');

      // 9. Verify pre-conversation system bubbles (e.g. Google login successful) preserve starter chips and prompt suggestions
      assert.match(js, /const hasConversation\s*=\s*Boolean\(historyDiv\.querySelector\(["']\.chat-bubble\.user,\s*\.chat-bubble\.assistant["']\)\)/, 'appendBubble must check for user/assistant conversation before hiding starter UI');
    });

  });

  describe('4. Mutation Fault Injection Checks (MUT-1 to MUT-4)', () => {

    test('MUT-1: Inverting isComposeMode correctly impacts Coach tab display', () => {
      const adapter = new OutlookAdapter();

      // Read mode (isComposeMode = false)
      global.Office = {
        context: {
          mailbox: {
            item: {
              body: {} // No setSelectedDataAsync
            }
          }
        }
      };
      assert.equal(adapter.isComposeMode(), false, 'Read mode must return false');

      // Compose mode (isComposeMode = true)
      global.Office.context.mailbox.item.body.setSelectedDataAsync = () => {};
      assert.equal(adapter.isComposeMode(), true, 'Compose mode must return true');
    });

    test('MUT-2: replyWithDraft default parameter keeps replyAll=true', () => {
      const adapter = new OutlookAdapter();
      let usedReplyAll = null;

      global.Office = {
        context: {
          mailbox: {
            item: {
              body: {},
              displayReplyAllForm: () => { usedReplyAll = true; },
              displayReplyForm: () => { usedReplyAll = false; }
            }
          }
        }
      };

      adapter.replyWithDraft('hello');
      assert.equal(usedReplyAll, true, 'Default parameter must be replyAll=true');
    });

    test('MUT-3: forwardWithDraft preserves empty toRecipients array', () => {
      const adapter = new OutlookAdapter();
      let recipients = null;

      global.Office = {
        context: {
          mailbox: {
            item: {
              body: {},
              displayForwardForm: (opts) => { recipients = opts.toRecipients; }
            }
          }
        }
      };

      adapter.forwardWithDraft('content');
      assert.deepEqual(recipients, []);
    });

    test('MUT-4: Markdown parser handles tables and callout boxes reliably', () => {
      const mdTable = `| Name | Role |\n| :--- | :--- |\n| Alex | Engineer |`;
      const htmlTable = parseMarkdown(mdTable);
      assert.match(htmlTable, /<table/);
      assert.match(htmlTable, /Alex/);

      const mdCallout = `> [!NOTE]\n> Critical finding`;
      const htmlCallout = parseMarkdown(mdCallout);
      assert.match(htmlCallout, /Executive Callout|Critical finding/);
    });

    test('MUT-5: OutlookAdapter differentiates between email message envelopes and calendar appointment envelopes', async () => {
      const adapter = new OutlookAdapter();

      // Case A: Email Message Envelope
      global.Office = {
        AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
        context: {
          mailbox: {
            item: {
              itemType: 'message',
              subject: 'Project Status Update',
              from: { displayName: 'Alice', emailAddress: 'alice@corp.com' },
              to: [{ displayName: 'Bob', emailAddress: 'bob@corp.com' }],
              body: { getAsync: (c, cb) => cb({ status: 'succeeded', value: 'Email message content.' }) }
            }
          }
        }
      };
      const emailDoc = await adapter.getFullDocumentText();
      assert.match(emailDoc, /EMAIL ENVELOPE METADATA/);
      assert.equal(emailDoc.includes('CALENDAR APPOINTMENT'), false);

      // Case B: Calendar Appointment Envelope
      global.Office.context.mailbox.item = {
        itemType: 'appointment',
        subject: 'Executive Briefing',
        organizer: { displayName: 'Alice', emailAddress: 'alice@corp.com' },
        requiredAttendees: [{ displayName: 'Bob', emailAddress: 'bob@corp.com' }],
        body: { getAsync: (c, cb) => cb({ status: 'succeeded', value: 'Appointment agenda notes.' }) }
      };
      const apptDoc = await adapter.getFullDocumentText();
      assert.match(apptDoc, /CALENDAR APPOINTMENT \/ MEETING INVITATION METADATA/);
      assert.equal(apptDoc.includes('EMAIL ENVELOPE METADATA'), false);
    });

  });

  describe('5. Regression Prevention Tests (REG-1 to REG-4)', () => {

    test('REG-1: Markdown parser handles empty strings and edge cases without throwing', () => {
      assert.equal(parseMarkdown(''), '');
      assert.equal(parseMarkdown(null), '');
      assert.equal(parseMarkdown(undefined), '');
      assert.match(parseMarkdown('**Bold Text**'), /<strong[^>]*>Bold Text<\/strong>/);
    });

    test('REG-2: OutlookAdapter handles missing Office.context safely without crashing', async () => {
      const adapter = new OutlookAdapter();
      global.Office = undefined;

      assert.equal(adapter.isComposeMode(), false);
      assert.equal(await adapter.getSelectedText(), '');
      assert.equal(await adapter.getFullDocumentText(), '');
      assert.equal(await adapter.getDraftBody(), '');

      const user = adapter.getUserProfile();
      assert.equal(user.displayName, 'User');

      const sender = await adapter.getSenderProfile();
      assert.equal(sender.displayName, 'Colleague');

      // Safe invocation of reply/forward when Office is undefined
      assert.doesNotThrow(() => adapter.replyWithDraft('text'));
      assert.doesNotThrow(() => adapter.forwardWithDraft('text'));
      assert.doesNotThrow(() => adapter.createCalendarInvite({}));
    });

    test('REG-3: OutlookAdapter handles calendar appointments with null or malformed dates gracefully', async () => {
      const adapter = new OutlookAdapter();

      global.Office = {
        AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
        context: {
          mailbox: {
            item: {
              itemType: 'appointment',
              subject: null,
              organizer: null,
              requiredAttendees: null,
              optionalAttendees: null,
              start: null,
              end: null,
              location: null,
              body: { getAsync: (c, cb) => cb({ status: 'succeeded', value: '' }) }
            }
          }
        }
      };

      const result = await adapter.getFullDocumentText();
      assert.match(result, /CALENDAR APPOINTMENT \/ MEETING INVITATION METADATA/);
      assert.match(result, /Subject: \(No Subject\)/);
      assert.match(result, /Empty Meeting Description/);
    });

    test('REG-4: OutlookAdapter._promisify resolves with string succeeded status when Office.AsyncResultStatus is undefined', async () => {
      const adapter = new OutlookAdapter();
      global.Office = { context: { mailbox: {} } }; // No AsyncResultStatus property

      const testFn = (cb) => cb({ status: 'succeeded', value: 'async success value' });
      const result = await adapter._promisify(testFn);
      assert.equal(result, 'async success value');
    });

    test('REG-5: OutlookAdapter handles missing Office.CoercionType without throwing', async () => {
      const adapter = new OutlookAdapter();
      let capturedCoercion = null;

      global.Office = {
        context: {
          mailbox: {
            item: {
              body: {
                setSelectedDataAsync: (data, opts, cb) => {
                  capturedCoercion = opts ? opts.coercionType : null;
                  if (typeof cb === 'function') cb({ status: 'succeeded' });
                },
                getAsync: (coercion, cb) => {
                  cb({ status: 'succeeded', value: 'body content' });
                }
              }
            }
          }
        }
      }; // No Office.CoercionType

      await assert.doesNotReject(async () => {
        await adapter.insertContent('<p>html</p>', 'html', 'replace_draft');
      });
      assert.equal(capturedCoercion, 'html');
    });

    test('REG-6: Email message with null location and start does not falsely classify as appointment', async () => {
      const adapter = new OutlookAdapter();

      global.Office = {
        AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
        context: {
          mailbox: {
            item: {
              itemType: 'message',
              subject: 'Status',
              location: null,
              start: null,
              end: null,
              from: { displayName: 'Sender', emailAddress: 'sender@corp.com' },
              to: [{ displayName: 'Receiver', emailAddress: 'receiver@corp.com' }],
              body: { getAsync: (c, cb) => cb({ status: 'succeeded', value: 'Hello' }) }
            }
          }
        }
      };

      const doc = await adapter.getFullDocumentText();
      assert.match(doc, /EMAIL ENVELOPE METADATA/);
      assert.equal(doc.includes('CALENDAR APPOINTMENT'), false);
    });

    test('REG-7: OutlookAdapter.createCalendarInvite deduplicates optionalAttendees already present in requiredAttendees', () => {
      const adapter = new OutlookAdapter();
      let capturedForm = null;

      global.Office = {
        context: {
          mailbox: {
            displayNewAppointmentForm: (form) => {
              capturedForm = form;
            },
            item: {
              subject: 'Architecture Sync'
            }
          }
        }
      };

      adapter.createCalendarInvite({
        requiredAttendees: ['alice@corp.com', 'bob@corp.com'],
        optionalAttendees: ['bob@corp.com', 'charlie@corp.com', 'alice@corp.com']
      });

      assert.ok(capturedForm);
      assert.deepEqual(capturedForm.requiredAttendees, ['alice@corp.com', 'bob@corp.com']);
      assert.deepEqual(capturedForm.optionalAttendees, ['charlie@corp.com'], 'Optional attendees must not contain required attendees');
    });

    test('REG-8: OutlookAdapter.getFullDocumentText parses plain string body in mock environments', async () => {
      const adapter = new OutlookAdapter();

      global.Office = {
        AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
        context: {
          mailbox: {
            item: {
              itemType: 'appointment',
              subject: 'Customer QBR Meeting',
              organizer: 'Lead Exec <exec@enterprise.com>',
              body: 'Agenda: 1. Q3 Metrics 2. SLA Compliance 3. Next Steps'
            }
          }
        }
      };

      const doc = await adapter.getFullDocumentText();
      assert.match(doc, /CALENDAR APPOINTMENT/);
      assert.match(doc, /Customer QBR Meeting/);
      assert.match(doc, /Agenda: 1\. Q3 Metrics 2\. SLA Compliance/);
    });

    test('REG-9: OutlookAdapter.getFullDocumentText parses ISO string dates for appointment start and end', async () => {
      const adapter = new OutlookAdapter();

      global.Office = {
        AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
        context: {
          mailbox: {
            item: {
              itemType: 'appointment',
              subject: 'ISO Date Appointment',
              start: '2026-09-20T15:30:00.000Z',
              end: '2026-09-20T16:30:00.000Z',
              organizer: 'Host <host@corp.com>',
              body: 'Description'
            }
          }
        }
      };

      const doc = await adapter.getFullDocumentText();
      assert.match(doc, /Start Time:/);
      assert.match(doc, /End Time:/);
      assert.match(doc, /ISO Date Appointment/);
    });

    test('REG-10: Gemini proxy index.js exposes /askGeminiStream on Discovery Engine :streamAssist', async () => {
      const { readFileSync } = await import('node:fs');
      const { resolve } = await import('node:path');
      const proxyIndexPath = resolve(process.cwd(), '../geminiproxy/index.js');
      const proxyCode = readFileSync(proxyIndexPath, 'utf-8');

      assert.match(proxyCode, /\/askGeminiStream/, 'Proxy must expose /askGeminiStream SSE route');
      assert.match(proxyCode, /handleGeminiStreamRequest/, 'Proxy must define handleGeminiStreamRequest');
      assert.match(proxyCode, /:streamAssist/, 'Proxy must invoke Discovery Engine :streamAssist');
    });

  });

});


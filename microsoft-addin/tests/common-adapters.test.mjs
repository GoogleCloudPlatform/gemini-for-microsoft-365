/**
 * Cross-Adapter Shared Unit Test Suite
 *
 * Tests the unified host adapter contract and behavior across all 4 adapters:
 *   - HostAdapterFactory (5-tier host resolution cascade + default fallback)
 *   - Shared interface & lifecycle across WordAdapter, PPTAdapter, ExcelAdapter, OutlookAdapter:
 *       1. Identity (`adapter.name` & method signatures)
 *       2. `getSelectedText()` (happy path, whitespace trimming, empty selection, missing host global, API error resilience)
 *       3. `getFullDocumentText()` (happy path, empty document, missing host global, API error resilience)
 *       4. `insertContent()` (null/empty early return, missing host global safety, happy path)
 *       5. `checkInDocumentCommands()` (case-insensitive `@gemini` trigger, `< 2` char guard, placeholder injection, callbacks, error recovery)
 *
 * Run:
 *   node tests/common-adapters.test.mjs
 */

import { test, describe, beforeEach, afterEach } from 'node:test';
import assert from 'node:assert/strict';

import { HostAdapterFactory } from '../src/adapters/HostAdapterFactory.js';
import { WordAdapter } from '../src/adapters/WordAdapter.js';
import { PPTAdapter } from '../src/adapters/PPTAdapter.js';
import { ExcelAdapter } from '../src/adapters/ExcelAdapter.js';
import { OutlookAdapter } from '../src/adapters/OutlookAdapter.js';
import {
  setGoogleAccessToken,
  getGoogleAccessToken,
  isGoogleTokenValid,
  syncGoogleTokenFromBackend
} from '../src/core/authService.js';

function cleanupHostGlobals() {
  delete globalThis.Office;
  delete globalThis.Word;
  delete globalThis.PowerPoint;
  delete globalThis.Excel;
  delete globalThis.window;
  delete globalThis.document;
}

describe('Cross-Adapter Shared Test Suite', () => {
  beforeEach(() => {
    cleanupHostGlobals();
  });

  afterEach(() => {
    cleanupHostGlobals();
  });

  // =========================================================================
  // 1. HostAdapterFactory — 5-Tier Host Resolution Cascade
  // =========================================================================
  describe('1. HostAdapterFactory 5-Tier Host Resolution Cascade', () => {
    test('UT-CA1: Tier 1 — resolves adapter from Office.onReady info.host (case-insensitive)', () => {
      assert.ok(HostAdapterFactory.getAdapter({ host: 'Word' }) instanceof WordAdapter);
      assert.ok(HostAdapterFactory.getAdapter({ host: 'word' }) instanceof WordAdapter);
      assert.ok(HostAdapterFactory.getAdapter({ host: 'PowerPoint' }) instanceof PPTAdapter);
      assert.ok(HostAdapterFactory.getAdapter({ host: 'powerpoint' }) instanceof PPTAdapter);
      assert.ok(HostAdapterFactory.getAdapter({ host: 'Excel' }) instanceof ExcelAdapter);
      assert.ok(HostAdapterFactory.getAdapter({ host: 'EXCEL' }) instanceof ExcelAdapter);
      assert.ok(HostAdapterFactory.getAdapter({ host: 'Outlook' }) instanceof OutlookAdapter);
      assert.ok(HostAdapterFactory.getAdapter({ host: 'outlook' }) instanceof OutlookAdapter);
    });

    test('UT-CA2: Tier 2 — falls back to Office.context.diagnostics.host when info.host is absent', () => {
      globalThis.Office = { context: { diagnostics: { host: 'PowerPoint' } } };
      assert.ok(HostAdapterFactory.getAdapter(null) instanceof PPTAdapter);

      globalThis.Office = { context: { diagnostics: { host: 'Excel' } } };
      assert.ok(HostAdapterFactory.getAdapter({}) instanceof ExcelAdapter);

      globalThis.Office = { context: { diagnostics: { host: 'Outlook' } } };
      assert.ok(HostAdapterFactory.getAdapter() instanceof OutlookAdapter);

      globalThis.Office = { context: { diagnostics: { host: 'Microsoft Word' } } };
      assert.ok(HostAdapterFactory.getAdapter() instanceof WordAdapter);
    });

    test('UT-CA3: Tier 3 — falls back to Office.context.host when diagnostics is absent', () => {
      globalThis.Office = { context: { host: 'Excel' } };
      assert.ok(HostAdapterFactory.getAdapter() instanceof ExcelAdapter);

      globalThis.Office = { context: { host: 'PowerPoint' } };
      assert.ok(HostAdapterFactory.getAdapter() instanceof PPTAdapter);

      globalThis.Office = { context: { host: 'Outlook' } };
      assert.ok(HostAdapterFactory.getAdapter() instanceof OutlookAdapter);
    });

    test('UT-CA4: Tier 4 — falls back to runtime global namespace inspection (PowerPoint, Excel, Word, mailbox)', () => {
      globalThis.PowerPoint = { run: async () => {} };
      assert.ok(HostAdapterFactory.getAdapter() instanceof PPTAdapter);
      delete globalThis.PowerPoint;

      globalThis.Excel = { run: async () => {} };
      assert.ok(HostAdapterFactory.getAdapter() instanceof ExcelAdapter);
      delete globalThis.Excel;

      globalThis.Word = { run: async () => {} };
      assert.ok(HostAdapterFactory.getAdapter() instanceof WordAdapter);
      delete globalThis.Word;

      globalThis.Office = { context: { mailbox: { item: {} } } };
      assert.ok(HostAdapterFactory.getAdapter() instanceof OutlookAdapter);
    });

    test('UT-CA5: Tier 5 — falls back to Office.context.requirements.isSetSupported', () => {
      globalThis.Office = {
        context: {
          requirements: {
            isSetSupported: (api) => api === 'PowerPointApi'
          }
        }
      };
      assert.ok(HostAdapterFactory.getAdapter() instanceof PPTAdapter);

      globalThis.Office.context.requirements.isSetSupported = (api) => api === 'ExcelApi';
      assert.ok(HostAdapterFactory.getAdapter() instanceof ExcelAdapter);

      globalThis.Office.context.requirements.isSetSupported = (api) => api === 'Mailbox';
      assert.ok(HostAdapterFactory.getAdapter() instanceof OutlookAdapter);

      globalThis.Office.context.requirements.isSetSupported = (api) => api === 'WordApi';
      assert.ok(HostAdapterFactory.getAdapter() instanceof WordAdapter);
    });

    test('UT-CA6: Default Fallback — returns WordAdapter when all detection tiers fail or throw', () => {
      assert.ok(HostAdapterFactory.getAdapter() instanceof WordAdapter);

      globalThis.Office = {
        get context() {
          throw new Error('Office context inaccessible');
        }
      };
      assert.ok(HostAdapterFactory.getAdapter() instanceof WordAdapter);
    });
  });

  // =========================================================================
  // 2. Unified Interface & Identity Contract Across All 4 Adapters
  // =========================================================================
  describe('2. Unified Adapter Interface Contract (All 4 Adapters)', () => {
    const adapterSpecs = [
      { Class: WordAdapter, expectedName: 'Word' },
      { Class: PPTAdapter, expectedName: 'PowerPoint' },
      { Class: ExcelAdapter, expectedName: 'Excel' },
      { Class: OutlookAdapter, expectedName: 'Outlook' }
    ];

    for (const { Class, expectedName } of adapterSpecs) {
      test(`UT-CA7 [${expectedName}]: exposes .name="${expectedName}" and required async interface methods`, () => {
        const adapter = new Class();
        assert.equal(adapter.name, expectedName);
        assert.equal(typeof adapter.getSelectedText, 'function');
        assert.equal(typeof adapter.getFullDocumentText, 'function');
        assert.equal(typeof adapter.insertContent, 'function');
        assert.equal(typeof adapter.checkInDocumentCommands, 'function');
      });

      test(`UT-CA8 [${expectedName}]: safe no-op and empty string return when host globals are undefined`, async () => {
        const adapter = new Class();
        const selected = await adapter.getSelectedText();
        assert.equal(selected, '');

        const fullDoc = await adapter.getFullDocumentText();
        assert.equal(fullDoc, '');

        // Empty content must return early across all 4 adapters without throwing
        await assert.doesNotReject(async () => {
          await adapter.insertContent('', '', 'smart');
          await adapter.insertContent(null, '', 'smart');
        });

        await assert.doesNotReject(async () => {
          await adapter.checkInDocumentCommands(false, {
            onStatus: () => {},
            executePrompt: async () => '<p>Result</p>'
          });
        });
      });

      test(`UT-CA9 [${expectedName}]: graceful error recovery when host APIs reject during getSelectedText / getFullDocumentText`, async () => {
        const adapter = new Class();

        globalThis.Word = {
          run: async () => { throw new Error('Simulated Word API failure'); }
        };
        globalThis.PowerPoint = {
          run: async () => { throw new Error('Simulated PowerPoint API failure'); }
        };
        globalThis.Excel = {
          run: async () => { throw new Error('Simulated Excel API failure'); }
        };
        globalThis.Office = {
          AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
          CoercionType: { Text: 'text', Html: 'html' },
          context: {
            document: {
              getSelectedDataAsync: (_type, cb) => cb({ status: 'failed', error: { message: 'Doc error' } })
            },
            mailbox: {
              item: {
                getSelectedDataAsync: (_type, cb) => cb({ status: 'failed', error: { message: 'Mailbox error' } }),
                body: {
                  getAsync: (_type, cb) => cb({ status: 'failed', error: { message: 'Body error' } })
                }
              }
            }
          }
        };

        const selected = await adapter.getSelectedText();
        assert.equal(selected, '');

        const fullDoc = await adapter.getFullDocumentText();
        assert.equal(fullDoc, '');
      });
    }
  });

  // =========================================================================
  // 3. Cross-Adapter Selection & Full Document Trimming / Empty Handling
  // =========================================================================
  describe('3. Cross-Adapter Selection Trimming & Empty State Contract', () => {
    test('UT-CA10: All 4 adapters trim surrounding whitespace in getSelectedText()', async () => {
      // 1. Word
      globalThis.Word = {
        run: async (cb) => cb({
          document: { getSelection: () => ({ load: () => {}, text: '   Word Selection   ' }) },
          sync: async () => {}
        })
      };
      assert.equal(await new WordAdapter().getSelectedText(), 'Word Selection');

      // 2. PowerPoint
      globalThis.PowerPoint = {
        run: async (cb) => cb({
          presentation: {
            getSelectedTextRangeOrNullObject: () => ({ isNullObject: false, text: '  PPT Selection  ', load: () => {} })
          },
          sync: async () => {}
        })
      };
      assert.equal(await new PPTAdapter().getSelectedText(), 'PPT Selection');

      // 3. Excel
      globalThis.Excel = {
        run: async (cb) => cb({
          workbook: {
            worksheets: {
              getActiveWorksheet: () => ({
                name: 'Sheet1',
                load: () => {},
                charts: { getItemOrNullObject: () => ({ isNullObject: true }), load: () => {}, items: [] },
                getUsedRangeOrNullObject: () => ({ isNullObject: true, load: () => {} })
              })
            },
            getSelectedRange: () => ({
              load: () => {},
              address: 'Sheet1!A1',
              rowCount: 1,
              columnCount: 1,
              rowIndex: 0,
              columnIndex: 0,
              values: [['   Cell A1   ']],
              formulas: [['']],
              getTables: () => ({ load: () => {}, items: [] })
            }),
            getActiveChartOrNullObject: () => ({ isNullObject: true, load: () => {} })
          },
          sync: async () => {}
        })
      };
      const excelSel = await new ExcelAdapter().getSelectedText();
      assert.ok(excelSel.includes('Cell Content: "Cell A1"'), 'ExcelAdapter must trim cell text in getSelectedText');

      // 4. Outlook
      globalThis.Office = {
        AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
        CoercionType: { Text: 'text' },
        context: {
          mailbox: {
            item: {
              getSelectedDataAsync: (_type, cb) => cb({ status: 'succeeded', value: '  Outlook Selection  ' })
            }
          }
        }
      };
      assert.equal(await new OutlookAdapter().getSelectedText(), 'Outlook Selection');
    });

    test('UT-CA11: All 4 adapters ignore @gemini commands when the prompt is shorter than 2 characters', async () => {
      let executeCalled = false;
      const callbacks = {
        onStatus: () => {},
        executePrompt: async () => {
          executeCalled = true;
          return '<p>Should not run</p>';
        }
      };

      // 1. Word with bare "@gemini " and "@gemini a"
      globalThis.Word = {
        run: async (cb) => cb({
          document: {
            getSelection: () => ({
              paragraphs: { load: () => {}, items: [{ text: 'Other line' }] }
            }),
            body: {
              paragraphs: {
                load: () => {},
                items: [{ text: '@gemini ' }, { text: '@gemini a' }]
              }
            }
          },
          sync: async () => {}
        })
      };
      await new WordAdapter().checkInDocumentCommands(true, callbacks);
      assert.equal(executeCalled, false, 'WordAdapter must ignore short @gemini prompt');

      // 2. PowerPoint with bare "@gemini"
      globalThis.PowerPoint = {
        run: async (cb) => cb({
          presentation: {
            slides: {
              load: () => {},
              items: [{
                shapes: {
                  load: () => {},
                  items: [{ textFrame: { textRange: { load: () => {}, text: '@gemini x' } } }]
                }
              }]
            }
          },
          sync: async () => {}
        })
      };
      await new PPTAdapter().checkInDocumentCommands(true, callbacks);
      assert.equal(executeCalled, false, 'PPTAdapter must ignore short @gemini prompt');

      // 3. Excel with bare "@gemini:"
      globalThis.Excel = {
        run: async (cb) => cb({
          workbook: {
            worksheets: { getActiveWorksheet: () => ({}) },
            getSelectedRange: () => ({
              load: () => {},
              address: 'Sheet1!A1',
              values: [['@gemini: a']],
              formulas: [['']]
            })
          },
          sync: async () => {}
        })
      };
      await new ExcelAdapter().checkInDocumentCommands(false, callbacks);
      assert.equal(executeCalled, false, 'ExcelAdapter must ignore short @gemini prompt');

      // 4. Outlook with bare "@gemini "
      globalThis.Office = {
        AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
        CoercionType: { Text: 'text', Html: 'html' },
        context: {
          mailbox: {
            item: {
              body: {
                setSelectedDataAsync: (_d, _o, cb) => cb({ status: 'succeeded' }),
                getAsync: (_t, cb) => cb({ status: 'succeeded', value: '@gemini !' })
              }
            }
          }
        }
      };
      await new OutlookAdapter().checkInDocumentCommands(false, callbacks);
      assert.equal(executeCalled, false, 'OutlookAdapter must ignore short @gemini prompt');
    });
  });

  // =========================================================================
  // 4. Cross-Surface Google OAuth Token Sync (Outlook <-> Excel/Word/PPT)
  // =========================================================================
  describe('4. Cross-Surface Google OAuth Token Sync via Entra ID SSO', () => {
    test('UT-CA12: Logging in on one surface pushes Google token to backend and hydrates other isolated storage partitions', async () => {
      const originalFetch = globalThis.fetch;
      let backendTokenStore = null;

      try {
        // Fake Entra ID JWT token for Office.auth.getAccessToken()
        const header = Buffer.from(JSON.stringify({ alg: 'RS256', typ: 'JWT' })).toString('base64url');
        const payload = Buffer.from(JSON.stringify({
          preferred_username: 'alice@contoso.com',
          name: 'Alice',
          oid: 'user-oid-123',
          exp: Math.floor(Date.now() / 1000) + 3600
        })).toString('base64url');
        const fakeEntraJwt = `${header}.${payload}.signature`;

        globalThis.Office = {
          auth: {
            getAccessToken: async () => fakeEntraJwt
          }
        };

        const createPartitionStorage = () => {
          const map = new Map();
          return {
            getItem: (k) => (map.has(k) ? map.get(k) : null),
            setItem: (k, v) => map.set(k, String(v)),
            removeItem: (k) => map.delete(k)
          };
        };

        globalThis.fetch = async (url, options = {}) => {
          const u = String(url);
          if (u.endsWith('/api/auth/google-token') && options.method === 'POST') {
            const body = JSON.parse(options.body || '{}');
            backendTokenStore = {
              google_token: body.google_token,
              google_token_expires_in: body.expires_in || 3600
            };
            return { ok: true, status: 200, json: async () => ({ status: 'ok', has_google_token: true }) };
          }
          if (u.endsWith('/api/auth/google-token') && options.method === 'DELETE') {
            backendTokenStore = null;
            return { ok: true, status: 200, json: async () => ({ status: 'ok', has_google_token: false }) };
          }
          if (u.endsWith('/api/auth/me')) {
            return {
              ok: true,
              status: 200,
              json: async () => ({
                authenticated: true,
                google_token: backendTokenStore ? backendTokenStore.google_token : null,
                google_token_expires_in: backendTokenStore ? backendTokenStore.google_token_expires_in : 0
              })
            };
          }
          return { ok: false, status: 404, json: async () => ({}) };
        };

        // Surface A: Outlook (outlook.office.com storage partition)
        const outlookLocalStorage = createPartitionStorage();
        globalThis.window = {
          location: { hostname: 'localhost', origin: 'http://localhost:8080', search: '' },
          localStorage: outlookLocalStorage,
          sessionStorage: createPartitionStorage()
        };

        setGoogleAccessToken('mock_oauth_cross_surface_google_token', 3600, true);
        // Wait microtask for async pushGoogleTokenToBackend
        await new Promise((r) => setTimeout(r, 20));
        assert.equal(backendTokenStore?.google_token, 'mock_oauth_cross_surface_google_token');
        assert.equal(isGoogleTokenValid(), true);

        // Simulate switching to Surface B: Excel (*.sharepoint.com isolated storage partition & fresh memory)
        setGoogleAccessToken(null, 0, false);
        const excelLocalStorage = createPartitionStorage();
        globalThis.window = {
          location: { hostname: 'localhost', origin: 'http://localhost:8080', search: '' },
          localStorage: excelLocalStorage,
          sessionStorage: createPartitionStorage()
        };
        assert.equal(getGoogleAccessToken(), null, 'Excel partition initially has no local token');

        // Hydrate from backend using shared Entra ID SSO
        const hydratedToken = await syncGoogleTokenFromBackend();
        assert.equal(hydratedToken, 'mock_oauth_cross_surface_google_token');
        assert.equal(getGoogleAccessToken(), 'mock_oauth_cross_surface_google_token');
        assert.equal(excelLocalStorage.getItem('google_user_access_token'), 'mock_oauth_cross_surface_google_token');

        // Clean up in-memory token state
        setGoogleAccessToken(null, 0, false);
      } finally {
        globalThis.fetch = originalFetch;
      }
    });
  });

  // =========================================================================
  // 5. Hybrid Intent-Based Grounding & Prompt Framing (Option 3)
  // =========================================================================
  describe('5. Hybrid Intent-Based Grounding & Prompt Framing (Option 3)', () => {
    test('UT-CA13: Bare input "word" or greeting in Word resolves to CONVERSATIONAL_OR_GENERAL with enableGrounding=false', async () => {
      process.env.NODE_ENV = 'test';
      const {
        extractUserQueryContext,
        classifyGroundingIntent,
        applyIntentGuidanceToPrompt,
        cleanStoredPromptText
      } = await import('../../geminiproxy/index.js');

      const framedWordPrompt = `Active Working Draft:\n"""\nQuarterly Planning Notes for Store Fleet.\n"""\n\nUser Query: word\n\nGrounding & Intent Policy: First evaluate the user's intent from their query...\nFormatting: Provide a direct, concise response in natural prose or bullet points. Do NOT generate a table unless explicitly requested.`;

      const ctx = extractUserQueryContext(framedWordPrompt);
      assert.equal(ctx.rawUserQuery, 'word');
      assert.equal(ctx.hasActiveDraft, true);

      const decision = await classifyGroundingIntent({
        prompt: framedWordPrompt,
        enableGrounding: true,
        host: 'Word'
      });
      assert.equal(decision.intent, 'CONVERSATIONAL_OR_GENERAL');
      assert.equal(decision.enableGrounding, false, 'Must disable enterprise datastore search for bare prompt "word"');

      const guided = applyIntentGuidanceToPrompt(framedWordPrompt, decision, 'Word');
      assert.ok(guided.includes('[IntentRouting: CONVERSATIONAL_OR_GENERAL]'));
      assert.ok(guided.includes('Do NOT summarize any draft and do NOT search, cite, or mention datastore files'));

      // Verify cleanStoredPromptText strips wrappers and recovers "word"
      assert.equal(cleanStoredPromptText(guided), 'word');
    });

    test('UT-CA14: Active document tasks disable enterprise grounding while enterprise lookups enable grounding', async () => {
      process.env.NODE_ENV = 'test';
      const {
        classifyGroundingIntent,
        applyIntentGuidanceToPrompt
      } = await import('../../geminiproxy/index.js');

      // 1. Active Document Task
      const docTaskPrompt = `Active Working Draft:\n"""\nDraft memo on team onboarding.\n"""\n\nUser Query: Summarize this document in 3 bullet points\n\nFormatting: Provide a direct, concise response.`;
      const docDecision = await classifyGroundingIntent({
        prompt: docTaskPrompt,
        enableGrounding: true,
        host: 'Word'
      });
      assert.equal(docDecision.intent, 'ACTIVE_DOCUMENT_TASK');
      assert.equal(docDecision.enableGrounding, false);

      // 2. Enterprise Knowledge Lookup
      const enterprisePrompt = `Active Working Draft:\n"""\nExecutive Retail Brief.\n"""\n\nUser Query: What are the closing procedures in our Store Operation Directive and Ann Taylor Q2 DTC digital penetration?\n\nFormatting: Provide a direct, concise response.`;
      const entDecision = await classifyGroundingIntent({
        prompt: enterprisePrompt,
        enableGrounding: true,
        host: 'Word'
      });
      assert.equal(entDecision.intent, 'ENTERPRISE_LOOKUP');
      assert.equal(entDecision.enableGrounding, true);

      const guidedEnt = applyIntentGuidanceToPrompt(enterprisePrompt, entDecision, 'Word');
      assert.ok(guidedEnt.includes('[IntentRouting: ENTERPRISE_LOOKUP]'));
      assert.ok(guidedEnt.includes('NEVER attempt to open, download, or read datastore files'));

      // 3. Vertex LLM Semantic Classifier Path & Caching
      let llmCalls = 0;
      const customQueryPrompt = `User Query: Could you help me brainstorm three creative icebreaker questions for a team offsite?`;
      const llmDecision1 = await classifyGroundingIntent({
        prompt: customQueryPrompt,
        enableGrounding: true,
        host: 'Word',
        llmClassifierFn: async () => {
          llmCalls += 1;
          return { intent: 'CONVERSATIONAL_OR_GENERAL' };
        }
      });
      assert.equal(llmDecision1.intent, 'CONVERSATIONAL_OR_GENERAL');
      assert.equal(llmDecision1.enableGrounding, false);
      assert.equal(llmDecision1.source, 'vertex_llm_classifier');
      assert.equal(llmCalls, 1);

      // Second call hits cache without invoking LLM again
      const llmDecision2 = await classifyGroundingIntent({
        prompt: customQueryPrompt,
        enableGrounding: true,
        host: 'Word',
        llmClassifierFn: async () => {
          llmCalls += 1;
          return { intent: 'CONVERSATIONAL_OR_GENERAL' };
        }
      });
      assert.equal(llmDecision2.intent, 'CONVERSATIONAL_OR_GENERAL');
      assert.equal(llmDecision2.enableGrounding, false);
      assert.equal(llmDecision2.source, 'intent_cache');
      assert.equal(llmCalls, 1);
    });
  });
});

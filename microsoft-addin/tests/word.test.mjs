/**
 * Automated Test Suite for Microsoft Word Adapter (WordAdapter.js)
 */
import { test, describe, beforeEach } from 'node:test';
import assert from 'node:assert/strict';
import { WordAdapter } from '../src/adapters/WordAdapter.js';
import { SessionManager } from '../src/core/sessionManager.js';

describe('Word Adapter Test Suite', () => {
  let adapter;

  beforeEach(() => {
    adapter = new WordAdapter();
  });

  test('UT-W1: WordAdapter initialization and name contract', () => {
    assert.equal(adapter.name, 'Word');
  });

  test('UT-W2: getSelectedText reads selection via Word.run and returns trimmed text', async () => {
    let syncCalled = false;
    global.Word = {
      run: async (callback) => {
        const mockContext = {
          document: {
            getSelection: () => ({
              load: () => {},
              text: '   Sample selected word paragraph   '
            })
          },
          sync: async () => {
            syncCalled = true;
          }
        };
        return await callback(mockContext);
      }
    };

    const text = await adapter.getSelectedText();
    assert.equal(syncCalled, true);
    assert.equal(text, 'Sample selected word paragraph');
  });

  test('UT-W3: getSelectedText handles Word.run exceptions safely and returns empty string', async () => {
    global.Word = {
      run: async () => {
        throw new Error('Selection context detached');
      }
    };

    const text = await adapter.getSelectedText();
    assert.equal(text, '');
  });

  test('UT-W4: getFullDocumentText extracts complete body text cleanly', async () => {
    let syncCalled = false;
    global.Word = {
      run: async (callback) => {
        const mockContext = {
          document: {
            body: {
              load: () => {},
              text: 'Executive Summary: Project Gemini is progressing on schedule.'
            }
          },
          sync: async () => {
            syncCalled = true;
          }
        };
        return await callback(mockContext);
      }
    };

    const fullText = await adapter.getFullDocumentText();
    assert.equal(syncCalled, true);
    assert.equal(fullText, 'Executive Summary: Project Gemini is progressing on schedule.');
  });

  test('UT-W5: getFullDocumentText returns empty string when body is empty or throws', async () => {
    global.Word = {
      run: async () => {
        throw new Error('Document body locked');
      }
    };

    const fullText = await adapter.getFullDocumentText();
    assert.equal(fullText, '');
  });

  test('UT-W6: insertContent in selection mode replaces selection and tags ContentControl', async () => {
    let insertedHtml = '';
    let insertLocation = '';
    let taggedTag = '';
    let taggedTitle = '';
    let syncCount = 0;

    global.document = {
      createElement: () => ({
        set innerHTML(val) { this._html = val; },
        get innerHTML() { return this._html || ''; }
      })
    };

    global.Word = {
      InsertLocation: { replace: 'Replace' },
      run: async (callback) => {
        const mockContext = {
          document: {
            getSelection: () => ({
              load: () => {},
              text: 'Active highlighted text',
              insertHtml: (html, loc) => {
                insertedHtml = html;
                insertLocation = loc;
                return {
                  insertContentControl: () => ({
                    set tag(t) { taggedTag = t; },
                    set title(title) { taggedTitle = title; }
                  })
                };
              }
            }),
            contentControls: {
              getByTag: () => ({
                load: () => {},
                items: []
              })
            }
          },
          sync: async () => {
            syncCount++;
          }
        };
        return await callback(mockContext);
      }
    };

    await adapter.insertContent('<p>New Gemini Content</p>', 'smart');
    assert.equal(insertedHtml, '<p>New Gemini Content</p>');
    assert.equal(insertLocation, 'Replace');
    assert.equal(taggedTag, 'GEMINI_DRAFT');
    assert.equal(taggedTitle, 'Gemini Enterprise Draft');
    assert.ok(syncCount >= 2);
  });

  test('UT-W7: insertContent in replace_draft mode replaces existing GEMINI_DRAFT in place', async () => {
    let draftDeleted = false;
    let replacedHtml = '';
    let retaggedTag = '';

    global.Word = {
      InsertLocation: { replace: 'Replace' },
      run: async (callback) => {
        const mockContext = {
          document: {
            getSelection: () => ({
              load: () => {},
              text: ''
            }),
            contentControls: {
              getByTag: (tag) => ({
                load: () => {},
                items: [
                  {
                    tag: 'GEMINI_DRAFT',
                    delete: (keepContent) => {
                      draftDeleted = true;
                    },
                    getRange: () => ({
                      insertHtml: (html, loc) => {
                        replacedHtml = html;
                        return {
                          insertContentControl: () => ({
                            set tag(t) { retaggedTag = t; }
                          })
                        };
                      }
                    })
                  }
                ]
              })
            }
          },
          sync: async () => {}
        };
        return await callback(mockContext);
      }
    };

    await adapter.insertContent('<div>Refined Draft</div>', 'replace_draft');
    assert.equal(draftDeleted, true);
    assert.equal(replacedHtml, '<div>Refined Draft</div>');
    assert.equal(retaggedTag, 'GEMINI_DRAFT');
  });

  test('UT-W8: insertContent returns early on empty or null content', async () => {
    let runCalled = false;
    global.Word = {
      run: async () => { runCalled = true; }
    };

    await adapter.insertContent(null);
    await adapter.insertContent('');
    assert.equal(runCalled, false);
  });

  test('UT-W9: checkInDocumentCommands detects @gemini and doc slash commands', async () => {
    let statusLog = [];
    let promptExecuted = '';

    global.Word = {
      InsertLocation: { replace: 'Replace' },
      run: async (callback) => {
        const mockContext = {
          document: {
            getSelection: () => ({
              paragraphs: {
                load: () => {},
                items: [{ text: 'cursor line' }]
              }
            }),
            body: {
              paragraphs: {
                load: () => {},
                items: [
                  {
                    text: '@gemini summarize the key quarterly earnings highlights',
                    insertText: () => {}
                  }
                ]
              },
              search: () => ({
                load: () => {},
                items: []
              })
            }
          },
          sync: async () => {}
        };
        return await callback(mockContext);
      }
    };

    await adapter.checkInDocumentCommands(true, {
      onStatus: (msg) => statusLog.push(msg),
      executePrompt: async (prompt) => {
        promptExecuted = prompt;
        return '<p>Summary result</p>';
      }
    });

    assert.equal(promptExecuted, 'summarize the key quarterly earnings highlights');
    assert.ok(statusLog.some(s => s.includes('Found:')));
  });

  test('UT-W10: insertContent in insert_cursor mode bypasses Priority 1 & Priority 2 and inserts at cursor', async () => {
    let draftTouched = false;
    let selectionInserted = '';
    let taggedWith = '';

    global.Word = {
      InsertLocation: { replace: 'Replace' },
      run: async (callback) => {
        const mockContext = {
          document: {
            getSelection: () => ({
              load: () => {},
              text: 'Selected Text That Should Not Trigger Priority 1',
              insertHtml: (html) => {
                selectionInserted = html;
                return {
                  insertContentControl: () => ({
                    set tag(t) { taggedWith = t; },
                    set title(_v) {}
                  })
                };
              }
            }),
            contentControls: {
              getByTag: () => ({
                load: () => {},
                items: [
                  {
                    tag: 'GEMINI_DRAFT',
                    delete: () => { draftTouched = true; },
                    getRange: () => ({ insertHtml: () => ({ insertContentControl: () => ({}) }) })
                  }
                ]
              })
            }
          },
          sync: async () => {}
        };
        return await callback(mockContext);
      }
    };

    await adapter.insertContent('<p>Cursor Insert</p>', 'insert_cursor');
    assert.equal(draftTouched, false, 'Must not delete or replace existing GEMINI_DRAFT control in insert_cursor mode');
    assert.equal(selectionInserted, '<p>Cursor Insert</p>');
    assert.equal(taggedWith, 'GEMINI_DRAFT');
  });

  test('UT-W11: insertContent Priority 2 updates only the last GEMINI_DRAFT when multiple controls exist', async () => {
    let firstDeleted = false;
    let lastDeleted = false;
    let lastReplacedHtml = '';

    global.Word = {
      InsertLocation: { replace: 'Replace' },
      run: async (callback) => {
        const mockContext = {
          document: {
            getSelection: () => ({ load: () => {}, text: '' }),
            contentControls: {
              getByTag: () => ({
                load: () => {},
                items: [
                  {
                    tag: 'GEMINI_DRAFT',
                    delete: () => { firstDeleted = true; },
                    getRange: () => ({ insertHtml: () => ({ insertContentControl: () => ({}) }) })
                  },
                  {
                    tag: 'GEMINI_DRAFT',
                    delete: () => { lastDeleted = true; },
                    getRange: () => ({
                      insertHtml: (html) => {
                        lastReplacedHtml = html;
                        return { insertContentControl: () => ({}) };
                      }
                    })
                  }
                ]
              })
            }
          },
          sync: async () => {}
        };
        return await callback(mockContext);
      }
    };

    await adapter.insertContent('<p>Updated Second Draft</p>', 'smart');
    assert.equal(firstDeleted, false, 'First GEMINI_DRAFT control must not be touched');
    assert.equal(lastDeleted, true, 'Last GEMINI_DRAFT control must be updated');
    assert.equal(lastReplacedHtml, '<p>Updated Second Draft</p>');
  });

  test('UT-W12: insertContent Priority 2 exception falls back to Priority 3 cursor insertion and tolerates ContentControl tag errors', async () => {
    let fallbackInsertedHtml = '';

    global.Word = {
      InsertLocation: { replace: 'Replace' },
      run: async (callback) => {
        const mockContext = {
          document: {
            getSelection: () => ({
              load: () => {},
              text: '',
              insertHtml: (html) => {
                fallbackInsertedHtml = html;
                return {
                  insertContentControl: () => {
                    throw new Error('ContentControl tagging unsupported on this Word build');
                  }
                };
              }
            }),
            contentControls: {
              getByTag: () => ({
                load: () => {},
                items: [
                  {
                    tag: 'GEMINI_DRAFT',
                    getRange: () => {
                      throw new Error('Detached ContentControl range');
                    }
                  }
                ]
              })
            }
          },
          sync: async () => {}
        };
        return await callback(mockContext);
      }
    };

    await assert.doesNotReject(async () => {
      await adapter.insertContent('<p>Resilient Fallback Content</p>', 'smart');
    });
    assert.equal(fallbackInsertedHtml, '<p>Resilient Fallback Content</p>');
  });

  test('UT-W13: checkInDocumentCommands detects all 4 native doc slash commands', async () => {
    const slashCommands = [
      '/summarize-doc',
      '/identify-risks',
      '/identify-action',
      '/exec-card'
    ];

    for (const cmd of slashCommands) {
      let executed = '';
      let statusMessages = [];

      global.Word = {
        InsertLocation: { replace: 'Replace' },
        run: async (callback) => {
          const mockContext = {
            document: {
              getSelection: () => ({
                paragraphs: { load: () => {}, items: [{ text: 'other paragraph' }] }
              }),
              body: {
                paragraphs: {
                  load: () => {},
                  items: [{ text: `${cmd} for Q3 board review`, insertText: () => {} }]
                },
                search: () => ({ load: () => {}, items: [] })
              }
            },
            sync: async () => {}
          };
          return await callback(mockContext);
        }
      };

      await adapter.checkInDocumentCommands(false, {
        onStatus: (msg) => statusMessages.push(msg),
        executePrompt: async (prompt) => {
          executed = prompt;
          return '<p>Done</p>';
        }
      });

      assert.equal(executed, `${cmd} for Q3 board review`);
      assert.ok(statusMessages.some(s => s.includes(`Found: "${cmd}`)));
    }
  });

  test('UT-W14: checkInDocumentCommands pauses with Typing status when cursor is in active paragraph and forceRun=false', async () => {
    let executed = false;
    let statusMessages = [];

    global.Word = {
      InsertLocation: { replace: 'Replace' },
      run: async (callback) => {
        const mockContext = {
          document: {
            getSelection: () => ({
              paragraphs: {
                load: () => {},
                items: [{ text: '@gemini draft an executive intro' }]
              }
            }),
            body: {
              paragraphs: {
                load: () => {},
                items: [{ text: '@gemini draft an executive intro', insertText: () => {} }]
              }
            }
          },
          sync: async () => {}
        };
        return await callback(mockContext);
      }
    };

    await adapter.checkInDocumentCommands(false, {
      onStatus: (msg) => statusMessages.push(msg),
      executePrompt: async () => {
        executed = true;
        return '<p>Intro</p>';
      }
    });

    assert.equal(executed, false, 'Must not execute while user cursor is actively in the paragraph and forceRun=false');
    assert.ok(statusMessages.some(s => s.includes('Typing:') && s.includes('(press ENTER)')));
  });

  test('UT-W15: checkInDocumentCommands 2nd-pass body.search locates marker, replaces paragraph with HTML, and tags GEMINI_DRAFT', async () => {
    let insertedMarkerText = '';
    let searchedMarkerId = '';
    let finalHtmlReplaced = '';
    let finalCcTag = '';
    let runCallIndex = 0;

    global.Word = {
      InsertLocation: { replace: 'Replace' },
      run: async (callback) => {
        runCallIndex++;
        const mockContext = {
          document: {
            getSelection: () => ({
              paragraphs: { load: () => {}, items: [{ text: 'different line' }] }
            }),
            body: {
              paragraphs: {
                load: () => {},
                items: [
                  {
                    text: '@gemini write a risk table',
                    insertText: (txt) => { insertedMarkerText = txt; }
                  }
                ]
              },
              search: (query) => {
                searchedMarkerId = query;
                return {
                  load: () => {},
                  items: [
                    {
                      paragraphs: {
                        load: () => {},
                        items: [
                          {
                            insertHtml: (html) => {
                              finalHtmlReplaced = html;
                              return {
                                insertContentControl: () => ({
                                  set tag(t) { finalCcTag = t; },
                                  set title(_v) {}
                                })
                              };
                            }
                          }
                        ]
                      }
                    }
                  ]
                };
              }
            }
          },
          sync: async () => {}
        };
        return await callback(mockContext);
      }
    };

    await adapter.checkInDocumentCommands(false, {
      onStatus: () => {},
      executePrompt: async () => '<table><tr><td>Risk 1</td></tr></table>'
    });

    assert.equal(runCallIndex, 2, 'Must run two Word.run passes (scan/mark + search/replace)');
    assert.ok(searchedMarkerId.startsWith('GEMINI_MARKER_'));
    assert.ok(insertedMarkerText.includes(searchedMarkerId));
    assert.equal(finalHtmlReplaced, '<table><tr><td>Risk 1</td></tr></table>');
    assert.equal(finalCcTag, 'GEMINI_DRAFT');
  });

  test('UT-W16: SessionManager strictly isolates Word chat history from Excel, PowerPoint, and Outlook sessions', () => {
    const store = new Map();
    const sessStore = new Map();
    const mockLocalStorage = {
      getItem: (k) => (store.has(k) ? store.get(k) : null),
      setItem: (k, v) => store.set(k, String(v)),
      removeItem: (k) => store.delete(k)
    };
    const mockSessionStorage = {
      getItem: (k) => (sessStore.has(k) ? sessStore.get(k) : null),
      setItem: (k, v) => sessStore.set(k, String(v)),
      removeItem: (k) => sessStore.delete(k)
    };
    global.window = {
      localStorage: mockLocalStorage,
      sessionStorage: mockSessionStorage
    };

    // Simulate a legacy state where an Excel formula session was mistakenly saved with host: 'Word'
    store.set('gemini_o365_chat_sessions_v1', JSON.stringify({
      sessions: {
        'excel-legacy-1': {
          id: 'excel-legacy-1',
          title: ' Analyze the provided Excel spreadsheet data and return EXACTLY 3 recommended Excel formulas...',
          host: 'Word',
          updatedAt: 1000,
          messages: [
            { role: 'user', text: 'Analyze the provided Excel spreadsheet data and return EXACTLY 3 recommended Excel formulas...' },
            { role: 'ai', text: '```json\n[{"title":"Sum","formula":"=SUM(A2:A10)"}]\n```' }
          ]
        },
        'word-real-1': {
          id: 'word-real-1',
          title: '[Word] Summarize Contract Clauses',
          host: 'Word',
          updatedAt: 2000,
          messages: [
            { role: 'user', text: 'Summarize the indemnification clause in this document' },
            { role: 'ai', text: 'The indemnification clause requires mutual coverage.' }
          ]
        }
      },
      sessionHosts: {}
    }));

    // Sync a remote session from Discovery Engine that belongs to Excel while inside Word
    SessionManager.syncRemoteSession({
      sessionId: 'excel-remote-2',
      title: '[Excel] Office 365 Session',
      updatedAt: 3000,
      messages: [
        { role: 'user', text: 'Build a bar chart for Q3 revenue' },
        { role: 'ai', text: '```excel-chart\n{"type":"ColumnClustered"}\n```' }
      ]
    }, 'Word');

    // Sync a remote session that was registered in sessionHosts as PowerPoint
    SessionManager.registerSessionHost('ppt-remote-3', 'PowerPoint');
    SessionManager.syncRemoteSession({
      sessionId: 'ppt-remote-3',
      title: 'Executive Deck Update',
      updatedAt: 4000,
      messages: [
        { role: 'user', text: 'Make this slide more concise' },
        { role: 'ai', text: 'Here is a tighter version.' }
      ]
    }, 'Word');

    const wordSessions = SessionManager.getAllSessions('Word');
    assert.equal(wordSessions.length, 1, 'Word history must only return genuine Word sessions');
    assert.equal(wordSessions[0].id, 'word-real-1');
    assert.equal(wordSessions[0].title, 'Summarize Contract Clauses');

    const excelSessions = SessionManager.getAllSessions('Excel');
    assert.equal(excelSessions.length, 2, 'Excel history must contain both repaired and synced Excel sessions');

    const pptSessions = SessionManager.getAllSessions('PowerPoint');
    assert.equal(pptSessions.length, 1, 'PowerPoint history must contain only PowerPoint sessions');
    assert.equal(pptSessions[0].id, 'ppt-remote-3');
  });

  test('UT-W17: SessionManager repairs poisoned PowerPoint localStorage entries and isolates all 4 surfaces (Word, PowerPoint, Excel, Outlook)', () => {
    const store = new Map();
    const sessStore = new Map();
    global.window = {
      localStorage: {
        getItem: (k) => (store.has(k) ? store.get(k) : null),
        setItem: (k, v) => store.set(k, String(v)),
        removeItem: (k) => store.delete(k)
      },
      sessionStorage: {
        getItem: (k) => (sessStore.has(k) ? sessStore.get(k) : null),
        setItem: (k, v) => sessStore.set(k, String(v)),
        removeItem: (k) => sessStore.delete(k)
      }
    };

    // Simulate a poisoned localStorage where a PowerPoint session was previously synced inside Word
    // and saved with BOTH host: 'Word' and sessionHosts['ppt-poisoned-1'] = 'Word'
    store.set('gemini_o365_chat_sessions_v1', JSON.stringify({
      sessions: {
        'ppt-poisoned-1': {
          id: 'ppt-poisoned-1',
          title: 'Q4 Executive Roadmap Presentation',
          host: 'Word',
          updatedAt: 5000,
          messages: [
            { role: 'user', text: 'Summarize these presentation slides thoroughly.' },
            { role: 'assistant', text: '## Slide 1: Q4 Roadmap\n* Key milestone achieved' }
          ]
        },
        'outlook-session-1': {
          id: 'outlook-session-1',
          title: 'Customer Escalation Follow-up',
          host: 'Word',
          updatedAt: 4500,
          messages: [
            { role: 'user', text: 'Draft a grounded email reply to the following customer inquiry email thread' },
            { role: 'assistant', text: 'Hi Alex,\n\nThank you for reaching out regarding our SLA terms.' }
          ]
        },
        'word-session-1': {
          id: 'word-session-1',
          title: '[OfficeHost: Word] Contract Review',
          host: 'Word',
          updatedAt: 4000,
          messages: [
            { role: 'user', text: '[OfficeHost: Word]\nActive Document Context:\n"""\nMaster Services Agreement\n"""\nSummarize this document.' },
            { role: 'assistant', text: 'This agreement establishes the master terms.' }
          ]
        }
      },
      sessionHosts: {
        'ppt-poisoned-1': 'Word',
        'outlook-session-1': 'Word',
        'word-session-1': 'Word'
      }
    }));

    // Also sync an untagged remote session with no surface markers while in Word
    SessionManager.syncRemoteSession({
      sessionId: 'untagged-remote-99',
      displayName: 'Generic Session',
      updatedAt: 6000,
      turns: [
        { query: { text: 'What is 2 + 2?' }, answer: '4' }
      ]
    }, 'Word');

    const wordHistory = SessionManager.getAllSessions('Word');
    const pptHistory = SessionManager.getAllSessions('PowerPoint');
    const outlookHistory = SessionManager.getAllSessions('Outlook');
    const excelHistory = SessionManager.getAllSessions('Excel');

    assert.equal(wordHistory.length, 1, 'Word must only see Word sessions (no poisoned PPT, Outlook, or untagged sessions)');
    assert.equal(wordHistory[0].id, 'word-session-1');

    assert.equal(pptHistory.length, 1, 'PowerPoint must recover and see its own session');
    assert.equal(pptHistory[0].id, 'ppt-poisoned-1');

    assert.equal(outlookHistory.length, 1, 'Outlook must recover and see its own session');
    assert.equal(outlookHistory[0].id, 'outlook-session-1');

    assert.equal(excelHistory.length, 0, 'Excel must see 0 sessions when none belong to Excel');
  });
});

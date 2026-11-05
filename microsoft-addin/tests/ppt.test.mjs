/**
 * Automated Test Suite for Microsoft PowerPoint Adapter & Slide Parser
 */
import { test, describe, beforeEach } from 'node:test';
import assert from 'node:assert/strict';
import {
  parseSlides,
  extractVisualPayload,
  normalizeVisualPayload,
  hydrateOrDropVisualPayload,
  formatVisualPayloadBody,
  extractStructuredBullets,
  filterDuplicateTableBullets,
  parseMarkdownTable,
  cleanSlideTitle,
  isConversationalPreamble,
  extractCleanBulletPoints
} from '../src/adapters/ppt/slideParser.js';
import {
  PPTAdapter,
  acquireSlideBuildLease,
  touchSlideBuildLease,
  releaseSlideBuildLease
} from '../src/adapters/PPTAdapter.js';
import {
  cardLayoutBudget,
  isCardLayoutSlide,
  applyCardLayoutBudget,
  demoteCardSlideForSlowHost
} from '../src/adapters/ppt/cardBudget.js';
import {
  normalizeHex,
  contrastRatio,
  ensureContrast,
  readableTextOn,
  themeFingerprint,
  encodePalette,
  decodePalette,
  finalizeSample
} from '../src/adapters/ppt/deckTheme.js';
import {
  fitFontSize,
  isSubstantiveSlideBody,
  __test as slideBuilderTest
} from '../src/adapters/ppt/slideBuilder.js';
import {
  extractRequestedSlideCount,
  enhancePromptForPowerPoint
} from '../src/adapters/ppt/promptEnhancer.js';

describe('PowerPoint Adapter & Slide Parser Test Suite', () => {

  beforeEach(() => {
    // Provide lightweight DOM mock for Node.js test environment
    const createMockDiv = () => ({
      _html: '',
      set innerHTML(val) { this._html = val; },
      get innerHTML() { return this._html || ''; },
      set textContent(val) { this._html = val; },
      get textContent() { return this._html ? this._html.replace(/<[^>]+>/g, ' ') : ''; },
      get innerText() { return this.textContent; },
      querySelectorAll: () => [],
      querySelector: () => null,
      cloneNode: function() {
        const cloned = createMockDiv();
        cloned.innerHTML = this._html;
        return cloned;
      }
    });

    global.document = {
      createElement: () => createMockDiv(),
      getElementById: () => null,
      head: { appendChild: () => {} },
      body: { appendChild: () => {} }
    };
    global.window = {
      __slideBuildLease: null,
      __isGeneratingSlides: false
    };
  });

  describe('1. Slide Structure Parsing', () => {
    test('UT-P1: parseSlides extracts multi-slide markdown headers', () => {
      const markdown = [
        '## Slide 1: Executive Overview',
        'Subtitle: Q3 Cloud Infrastructure Performance',
        '• Record adoption across enterprise sectors',
        '• 99.99% system availability maintained throughout peak load',
        'Takeaway: Infrastructure scalability is driving enterprise customer retention',
        '',
        '## Slide 2: Strategic Pillars',
        'Subtitle: Four Key Growth Drivers',
        '• Automated Security Compliance',
        '• Multi-region Disaster Recovery',
        '• Zero-trust Credential Delegation',
        'Takeaway: Security automation reduced manual audit effort by 80%'
      ].join('\n');

      const slides = parseSlides(markdown);
      assert.ok(Array.isArray(slides));
      assert.equal(slides.length, 2);

      assert.equal(slides[0].title, 'Executive Overview');
      assert.equal(slides[0].subtitle, 'Q3 Cloud Infrastructure Performance');
      assert.ok(slides[0].body.includes('Record adoption across enterprise sectors'));
      assert.equal(slides[0].takeaway, 'Infrastructure scalability is driving enterprise customer retention');

      assert.equal(slides[1].title, 'Strategic Pillars');
      assert.equal(slides[1].subtitle, 'Four Key Growth Drivers');
      assert.ok(slides[1].body.includes('Automated Security Compliance'));
      assert.equal(slides[1].takeaway, 'Security automation reduced manual audit effort by 80%');
    });

    test('UT-P2: parseSlides gracefully handles single-slide response without markdown headers', () => {
      const text = [
        'Here is the summary of the meeting:',
        '- Action items assigned to engineering',
        '- Timeline moved to next Tuesday',
        '- Follow-up call with clients on Thursday'
      ].join('\n');

      const slides = parseSlides(text);
      assert.ok(Array.isArray(slides));
      assert.equal(slides.length, 1);
      assert.ok(slides[0].body.includes('Action items assigned to engineering'));
    });

    test('UT-P3: extractCleanBulletPoints cleans leading markers and formats bullets', () => {
      const rawText = [
        '• First key observation',
        '- Second strategic objective',
        '• Third operational metric'
      ].join('\n');
      const bulletString = extractCleanBulletPoints(rawText);

      assert.ok(typeof bulletString === 'string');
      assert.ok(bulletString.includes('• First key observation'));
      assert.ok(bulletString.includes('• Second strategic objective'));
      assert.ok(bulletString.includes('• Third operational metric'));
    });
  });

  describe('2. Visual Archetypes & Payload Extraction', () => {
    test('UT-P4: normalizeVisualPayload normalizes metric_grid_3col structures', () => {
      const raw = {
        visualType: 'metric_grid_3col',
        title: 'Q3 Financial Performance',
        metrics: [
          { stat: '$4.2B', headline: 'Revenue', detail: 'Up 24% YoY' },
          { stat: '85%', headline: 'Gross Margin', detail: 'Exceeded guidance' },
          { stat: '1.2M', headline: 'Active Users', detail: 'Global enterprise' }
        ]
      };

      const normalized = normalizeVisualPayload(raw);
      assert.equal(normalized.visualType, 'metric_grid_3col');
      assert.equal(normalized.title, 'Q3 Financial Performance');
      assert.equal(normalized.cards.length, 3);
      assert.equal(normalized.cards[0].metric, '$4.2B');
      assert.equal(normalized.cards[0].title, 'Revenue');
    });

    test('UT-P5: normalizeVisualPayload normalizes before_after comparison layout', () => {
      const raw = {
        visualType: 'before_after',
        title: 'Modernization Impact',
        before: { title: 'Legacy Monolith', bullets: ['Slow release cycles', 'Manual tests'] },
        after: { title: 'Cloud Microservices', bullets: ['Continuous deployment', 'Automated CI/CD'] }
      };

      const normalized = normalizeVisualPayload(raw);
      assert.equal(normalized.visualType, 'before_after');
      assert.equal(normalized.before.title, 'Legacy Monolith');
      assert.equal(normalized.after.title, 'Cloud Microservices');
      assert.equal(normalized.before.bullets.length, 2);
    });

    test('UT-P6: normalizeVisualPayload normalizes process_timeline roadmap layout', () => {
      const raw = {
        visualType: 'process_timeline',
        title: 'Implementation Roadmap',
        steps: [
          { number: '01', headline: 'Assessment & Audit', text: 'Analyze existing systems' },
          { number: '02', headline: 'Migration', text: 'Transfer database and services' },
          { number: '03', headline: 'Validation', text: 'Run integration test suites' }
        ]
      };

      const normalized = normalizeVisualPayload(raw);
      assert.equal(normalized.visualType, 'process_timeline');
      assert.equal(normalized.steps.length, 3);
      assert.equal(normalized.steps[0].step, '01');
      assert.equal(normalized.steps[0].title, 'Assessment & Audit');
    });

    test('UT-P7: extractVisualPayload finds fenced JSON code blocks with layout types', () => {
      const markdown = [
        '## Slide 3: Growth Roadmap',
        '```json',
        '{',
        '  "visualType": "three_column_cards",',
        '  "title": "Three Strategic Pillars",',
        '  "cards": [',
        '    { "badge": "PILLAR 1", "headline": "Security", "body": "End-to-end encryption" },',
        '    { "badge": "PILLAR 2", "headline": "Scalability", "body": "Global multi-cloud reach" },',
        '    { "badge": "PILLAR 3", "headline": "Usability", "body": "Intuitive Office add-in" }',
        '  ]',
        '}',
        '```'
      ].join('\n');

      const payload = extractVisualPayload(markdown);
      assert.ok(payload);
      assert.equal(payload.visualType, 'three_column_cards');
      assert.equal(payload.cards.length, 3);
      assert.equal(payload.cards[0].pill, 'PILLAR 1');
      assert.equal(payload.cards[0].title, 'Security');
    });

    test('UT-P8: formatVisualPayloadBody generates readable markdown fallback', () => {
      const payload = {
        visualType: 'three_column_cards',
        cards: [
          { title: 'Security', body: 'End-to-end encryption' },
          { title: 'Scalability', body: 'Global multi-cloud reach' }
        ]
      };

      const formatted = formatVisualPayloadBody(payload);
      assert.ok(formatted.includes('• **Security**: End-to-end encryption'));
      assert.ok(formatted.includes('• **Scalability**: Global multi-cloud reach'));
    });
  });

  describe('3. Slide-Build Concurrency Lease (acquire / touch / release)', () => {
    test('UT-P9: acquireSlideBuildLease grants lease, touchSlideBuildLease refreshes timestamp, and releaseSlideBuildLease clears it', async () => {
      const token = await acquireSlideBuildLease();
      assert.ok(token.startsWith('build-'));
      assert.equal(global.window.__isGeneratingSlides, true);
      assert.equal(global.window.__slideBuildLease.token, token);

      const originalTick = global.window.__slideBuildLease.tickedAt;
      global.window.__slideBuildLease.tickedAt = originalTick - 5000;
      touchSlideBuildLease();
      assert.ok(global.window.__slideBuildLease.tickedAt >= originalTick);

      // Releasing with wrong token is ignored (owner token protection)
      releaseSlideBuildLease('wrong-token');
      assert.equal(global.window.__isGeneratingSlides, true);
      assert.equal(global.window.__slideBuildLease.token, token);

      // Releasing with matching token clears the lease
      releaseSlideBuildLease(token);
      assert.equal(global.window.__isGeneratingSlides, false);
      assert.equal(global.window.__slideBuildLease, null);
    });

    test('UT-P10: acquireSlideBuildLease reclaims an abandoned lease older than 20s without blocking', async () => {
      let diagMsg = '';
      global.window.logToDiagBox = (msg) => { diagMsg = msg; };
      global.window.__slideBuildLease = {
        token: 'build-stale-1',
        tickedAt: Date.now() - 25000 // 25s ago (> 20s threshold)
      };

      let waited = false;
      const newToken = await acquireSlideBuildLease(() => { waited = true; });
      assert.equal(waited, false);
      assert.notEqual(newToken, 'build-stale-1');
      assert.equal(global.window.__slideBuildLease.token, newToken);
      assert.match(diagMsg, /reclaiming it and continuing/i);

      releaseSlideBuildLease(newToken);
    });
  });

  describe('4. PPTAdapter Slide & Shape Text Extraction', () => {
    test('UT-P11: extractSlideText performs batch sync across TextBox and Table shapes and falls back to per-shape sync on batch error', async () => {
      const adapter = new PPTAdapter();

      // Case 1: Batch sync succeeds for TextBox + Table
      const slide1 = {
        id: 's1',
        shapes: {
          load: () => {},
          items: [
            {
              type: 'TextBox',
              textFrame: { textRange: { text: 'Executive Summary Heading', load: () => {} } }
            },
            {
              type: 'Table',
              getTable: () => ({
                values: [['Region', 'Revenue'], ['NA', '$120M']],
                load: () => {}
              })
            }
          ]
        }
      };
      const ctx1 = { sync: async () => {} };
      const res1 = await adapter.extractSlideText(slide1, 1, ctx1);
      assert.equal(res1.slideNumber, 1);
      assert.equal(res1.id, 's1');
      assert.ok(res1.text.includes('Executive Summary Heading'));
      assert.ok(res1.text.includes('| Region | Revenue |'));
      assert.ok(res1.text.includes('| NA | $120M |'));

      // Case 2: Batch sync throws on 2nd sync, triggering per-shape fallback
      let syncCount = 0;
      const ctx2 = {
        sync: async () => {
          syncCount++;
          if (syncCount === 2) {
            throw new Error('Batch sync failed on grouped shape');
          }
        }
      };
      const res2 = await adapter.extractSlideText(slide1, 2, ctx2);
      assert.equal(res2.slideNumber, 2);
      assert.ok(res2.text.includes('Executive Summary Heading'));
      assert.ok(res2.text.includes('| Region | Revenue |'));
    });

    test('UT-P12: getSelectedShapeText prioritizes active text range over shape selection and formats selected tables as Markdown', async () => {
      const adapter = new PPTAdapter();

      // Priority 1: Active highlighted text range
      global.PowerPoint = {
        run: async (cb) => cb({
          sync: async () => {},
          presentation: {
            getSelectedTextRangeOrNullObject: () => ({
              isNullObject: false,
              text: '  Highlighted sentence inside shape  ',
              load: () => {}
            })
          }
        })
      };
      const rangeText = await adapter.getSelectedShapeText();
      assert.equal(rangeText, 'Highlighted sentence inside shape');

      // Priority 2: Selected Table shape when text range is null object
      global.PowerPoint = {
        run: async (cb) => cb({
          sync: async () => {},
          presentation: {
            getSelectedTextRangeOrNullObject: () => ({ isNullObject: true, text: '', load: () => {} }),
            getSelectedShapes: () => ({
              load: () => {},
              items: [
                {
                  type: 'Table',
                  getTable: () => ({
                    values: [['Metric', 'Q1', 'Q2'], ['ARR', '$10M', '$14M']],
                    load: () => {}
                  })
                }
              ]
            })
          }
        })
      };
      const tableMarkdown = await adapter.getSelectedShapeText();
      assert.ok(tableMarkdown.includes('| Metric | Q1 | Q2 |'));
      assert.ok(tableMarkdown.includes('| --- | --- | --- |'));
      assert.ok(tableMarkdown.includes('| ARR | $10M | $14M |'));
    });

    test('UT-P13: getSelectedSlidesText maps true 1-based presentation slide numbers and getFullDocumentText joins slides with ---', async () => {
      const adapter = new PPTAdapter();
      global.PowerPoint = {
        run: async (cb) => cb({
          sync: async () => {},
          presentation: {
            slides: {
              load: () => {},
              items: [{ id: 'slide-a' }, { id: 'slide-b' }, { id: 'slide-c' }],
              getCount: () => ({ value: 2 }),
              getItemAt: (idx) => ({
                id: idx === 0 ? 'slide-a' : 'slide-b',
                shapes: {
                  load: () => {},
                  items: [{ type: 'TextBox', textFrame: { textRange: { text: `Deck Slide ${idx + 1} Body`, load: () => {} } } }]
                }
              })
            },
            getSelectedSlides: () => ({
              load: () => {},
              items: [{
                id: 'slide-c',
                shapes: {
                  load: () => {},
                  items: [{ type: 'TextBox', textFrame: { textRange: { text: 'Third Slide Selected', load: () => {} } } }]
                }
              }]
            })
          }
        })
      };

      const selected = await adapter.getSelectedSlidesText();
      assert.equal(selected.length, 1);
      assert.equal(selected[0].slideNumber, 3);
      assert.equal(selected[0].text, 'Third Slide Selected');

      const fullDoc = await adapter.getFullDocumentText();
      assert.ok(fullDoc.includes('[Slide 1]:\nDeck Slide 1 Body'));
      assert.ok(fullDoc.includes('---'));
      assert.ok(fullDoc.includes('[Slide 2]:\nDeck Slide 2 Body'));
    });
  });

  describe('5. PPTAdapter In-Place Shape Replacement & Commands', () => {
    test('UT-P14: insertContent in replace mode updates selected shape text in-place and bolds lead-in phrases before colons', async () => {
      const adapter = new PPTAdapter();
      const boldedSubstrings = [];
      const mockTextRange = {
        text: '',
        font: { size: 18 },
        getSubstring: (start, len) => {
          const sub = { start, len, font: { bold: false } };
          boldedSubstrings.push(sub);
          return sub;
        }
      };

      global.PowerPoint = {
        run: async (cb) => cb({
          sync: async () => {},
          presentation: {
            getSelectedShapes: () => ({
              load: () => {},
              items: [{ textFrame: { textRange: mockTextRange } }]
            })
          }
        })
      };

      const raw = '• Cloud Margin: Expanded by 340 bps YoY\n• Net Retention — Exceeded 120% in Q4';
      const result = await adapter.insertContent('<p>Cloud Margin</p>', raw, { mode: 'replace' });

      assert.equal(result.length, 1);
      assert.equal(result[0].title, 'Updated Shape');
      assert.equal(mockTextRange.font.size, 14);
      assert.ok(mockTextRange.text.includes('• Cloud Margin: Expanded by 340 bps YoY'));
      assert.equal(boldedSubstrings.length, 2);
      assert.equal(boldedSubstrings[0].font.bold, true);
      assert.equal(boldedSubstrings[1].font.bold, true);
    });

    test('UT-P15: checkInDocumentCommands invokes callbacks.onStatus with ready state', async () => {
      const adapter = new PPTAdapter();
      let statusText = '';
      await adapter.checkInDocumentCommands(true, {
        onStatus: (msg) => { statusText = msg; }
      });
      assert.match(statusText, /PowerPoint Adapter Ready/i);
    });
  });

  describe('6. Advanced Slide Parser (Hydration, Structured Bullets, Table Dedup, Title Cleaning)', () => {
    test('UT-P16: hydrateOrDropVisualPayload hydrates empty metric_grid_3col and three_column_cards shells from structured bullets or drops when empty', () => {
      const bodyWithMetrics = [
        '• **$4.8B Cloud Revenue**: Up 31% YoY across enterprise accounts',
        '• **99.99% Uptime SLA**: Zero P0 incidents across 42 regions',
        '• **42% Margin Expansion**: Driven by custom silicon efficiency'
      ].join('\n');

      const hydratedMetrics = hydrateOrDropVisualPayload(
        'metric_grid_3col',
        { visualType: 'metric_grid_3col', cards: [] },
        bodyWithMetrics
      );
      assert.equal(hydratedMetrics.visualType, 'metric_grid_3col');
      assert.equal(hydratedMetrics.visualData.cards.length, 3);
      assert.equal(hydratedMetrics.visualData.cards[0].metric, '$4.8B');
      assert.equal(hydratedMetrics.visualData.cards[0].title, 'Cloud Revenue');

      // Drops empty visual shell when body has no structured bullets
      const dropped = hydrateOrDropVisualPayload(
        'three_column_cards',
        { visualType: 'three_column_cards', cards: [] },
        'Plain single line without structured bullets.'
      );
      assert.equal(dropped.visualType, null);
      assert.equal(dropped.visualData, null);
    });

    test('UT-P17: extractStructuredBullets detects 2-4 bold lead-in bullets and rejects 1 or 5+ bullets', () => {
      const threeBullets = [
        '• **Zero Trust Security**: Hardware-attested identity across all endpoints',
        '• **Unified Data Fabric**: Real-time federation without ETL duplication',
        '• **Agentic Workflows**: Autonomous multi-step task execution'
      ].join('\n');

      const cards = extractStructuredBullets(threeBullets);
      assert.ok(Array.isArray(cards));
      assert.equal(cards.length, 3);
      assert.equal(cards[0].pill, 'PILLAR 01');
      assert.equal(cards[0].title, 'Zero Trust Security');
      assert.equal(cards[0].body, 'Hardware-attested identity across all endpoints');

      assert.equal(extractStructuredBullets('• **Only One**: Single bullet is not a card grid'), null);
    });

    test('UT-P18: filterDuplicateTableBullets removes pipe-delimited table remnants and verbatim cell matches while keeping genuine insights', () => {
      const tableData = {
        headers: ['Region', 'Q4 Revenue', 'YoY Growth'],
        rows: [
          ['North America', '$450M', '+22%'],
          ['EMEA', '$280M', '+18%']
        ]
      };
      const bullets = [
        '| North America | $450M | +22% |',
        'North America',
        '• North America led global expansion with strong enterprise renewals',
        '---'
      ];

      const filtered = filterDuplicateTableBullets(bullets, tableData);
      assert.deepEqual(filtered, ['• North America led global expansion with strong enterprise renewals']);
    });

    test('UT-P19: parseMarkdownTable parses pipe tables and cleanSlideTitle strips conversational preambles and caps length', () => {
      const mdTable = [
        '| Pillar | Owner | Status |',
        '| :--- | :---: | ---: |',
        '| **Security** | SecOps | Complete |',
        '| **Compute** | Infra | In Progress |'
      ].join('\n');

      const parsed = parseMarkdownTable(mdTable);
      assert.ok(parsed);
      assert.deepEqual(parsed.headers, ['Pillar', 'Owner', 'Status']);
      assert.equal(parsed.dataRows.length, 2);
      assert.deepEqual(parsed.dataRows[0], ['Security', 'SecOps', 'Complete']);

      assert.equal(cleanSlideTitle('## Slide 4: **Global Supply Chain Resilience**', 4), 'Global Supply Chain Resilience');
      assert.equal(cleanSlideTitle('Sure, here is a table of Regional Revenue Breakdown', 2), 'Regional Revenue Breakdown');
      assert.equal(isConversationalPreamble('Here is a concise and punchy version optimized for your slide:'), true);
      assert.equal(isConversationalPreamble('Enterprise cloud adoption grew 34% in Q4.'), false);
    });
  });

  describe('7. Card Layout Budget & Slow-Host Demotion', () => {
    test('UT-P20: cardLayoutBudget computes 40% ceiling using exact integer math', () => {
      assert.equal(cardLayoutBudget(0), 0);
      assert.equal(cardLayoutBudget(1), 1);
      assert.equal(cardLayoutBudget(2), 1);
      assert.equal(cardLayoutBudget(5), 2);
      assert.equal(cardLayoutBudget(10), 4);
    });

    test('UT-P21: applyCardLayoutBudget preserves first N card slides and demotes excess card slides to bold lead-in bullets', () => {
      const makeCardSlide = (num) => ({
        title: `Slide ${num}`,
        visualType: 'three_column_cards',
        visualData: {
          cards: [
            { pill: 'P1', title: `Pillar A${num}`, body: 'Detail A' },
            { pill: 'P2', title: `Pillar B${num}`, body: 'Detail B' }
          ]
        },
        body: ''
      });

      // 5-slide deck has a card budget of 2
      const slides = [makeCardSlide(1), makeCardSlide(2), makeCardSlide(3)];
      // Pad with 2 non-card table slides so total = 5 (budget = 2)
      slides.push({ title: 'Slide 4', tableData: { rows: [['a', 'b']] }, body: '' });
      slides.push({ title: 'Slide 5', tableData: { rows: [['c', 'd']] }, body: '' });

      applyCardLayoutBudget(slides);
      assert.equal(slides[0].visualType, 'three_column_cards');
      assert.equal(slides[1].visualType, 'three_column_cards');
      // 3rd card slide exceeds budget of 2 -> demoted to bullets
      assert.equal(slides[2].visualType, null);
      assert.ok(slides[2].cardBudgetDemoted);
      assert.ok(slides[2].body.includes('• **Pillar A3**: Detail A'));

      // Slow-host demotion on remaining card slide
      const demoted = demoteCardSlideForSlowHost(slides[1]);
      assert.equal(demoted, true);
      assert.equal(slides[1].visualType, null);
      assert.deepEqual(slides[1].slowHostDemoted, { layout: 'three_column_cards' });
      assert.equal(isCardLayoutSlide(slides[3]), false); // Table slide is never a card slide
    });
  });

  describe('8. Deck Theme & WCAG Contrast Engine', () => {
    test('UT-P22: normalizeHex, contrastRatio, ensureContrast, and readableTextOn enforce WCAG legibility', () => {
      assert.equal(normalizeHex('1a73e8'), '#1A73E8');
      assert.equal(normalizeHex('#abc'), '#AABBCC');
      assert.equal(normalizeHex('invalid'), null);

      const bwRatio = contrastRatio('#000000', '#FFFFFF');
      assert.ok(bwRatio > 20.9 && bwRatio <= 21.1);

      // Low-contrast light gray on white is darkened to achieve >= 4.5:1 contrast
      const adjusted = ensureContrast('#CCCCCC', '#FFFFFF', 4.5);
      assert.ok(contrastRatio(adjusted, '#FFFFFF') >= 4.5);

      // Dark text on bright canvas, white text on dark navy canvas
      assert.equal(readableTextOn('#FFFFFF'), '#111827');
      assert.equal(readableTextOn('#0F172A'), '#FFFFFF');
    });

    test('UT-P23: themeFingerprint, encodePalette, and decodePalette round-trip deck theme tokens', () => {
      const sample = {
        canvasColor: '#FFFFFF',
        themeColors: {
          Accent1: '#1A73E8', Accent2: '#34A853', Accent3: '#FBBC04',
          Accent4: '#EA4335', Accent5: '#9334E6', Accent6: '#00ACC1',
          Dark1: '#111827', Dark2: '#374151', Light1: '#FFFFFF', Light2: '#F3F4F6'
        }
      };
      const fp1 = themeFingerprint(sample);
      const fp2 = themeFingerprint(sample);
      assert.equal(fp1, fp2);

      const encoded = encodePalette({
        canvasColor: '#FFFFFF',
        accentColor: '#1A73E8',
        titleColor: '#111827',
        bodyColor: '#374151',
        mutedColor: '#6B7280',
        cardFill: '#F8FAFC',
        accentRamp: ['#1A73E8', '#34A853', '#FBBC04', '#EA4335'],
        headingFont: 'Plus Jakarta Sans',
        bodyFont: 'Inter'
      }, fp1);

      const decoded = decodePalette(encoded);
      assert.ok(decoded);
      assert.equal(decoded.fingerprint, fp1);
      assert.equal(decoded.canvasColor, '#FFFFFF');
      assert.equal(decoded.accentColor, '#1A73E8');
      assert.deepEqual(decoded.accentRamp, ['#1A73E8', '#34A853', '#FBBC04', '#EA4335']);
      assert.equal(decoded.headingFont, 'Plus Jakarta Sans');
      assert.equal(decoded.bodyFont, 'Inter');
      assert.equal(decodePalette('corrupted|payload'), null);
    });

    test('UT-P26: finalizeSample resolves exact canvas dimensions from PowerPointApi 1.10 pageSetup when available', () => {
      // 10-inch widescreen on Mac/Web (720 x 405)
      const sampleMac = finalizeSample({
        pageSetupWidth: 720,
        pageSetupHeight: 405,
        maxRight: 0,
        maxBottom: 0
      });
      assert.equal(sampleMac.slideWidth, 720);
      assert.equal(sampleMac.slideHeight, 405);

      // 13.33-inch widescreen on Windows (960 x 540)
      const sampleWin = finalizeSample({
        pageSetupWidth: 960,
        pageSetupHeight: 540,
        maxRight: 0,
        maxBottom: 0
      });
      assert.equal(sampleWin.slideWidth, 960);
      assert.equal(sampleWin.slideHeight, 540);
    });

    test('UT-P27: finalizeSample accurately snaps to 720x405 widescreen when pageSetup is unavailable and avoids 960 overflow', () => {
      // Blank or minimal presentation defaults to 720x405 rather than overflowing 960
      const sampleDefault = finalizeSample({
        maxRight: 30,
        maxBottom: 20
      });
      assert.equal(sampleDefault.slideWidth, 720);
      assert.equal(sampleDefault.slideHeight, 405);

      // Existing shapes on a 720x405 slide (e.g. content up to right=680)
      const sample720 = finalizeSample({
        maxRight: 680,
        maxBottom: 380
      });
      assert.equal(sample720.slideWidth, 720);
      assert.equal(sample720.slideHeight, 405);

      // Presentation with wide shapes that exceed 720 (e.g. right=850) snaps to 960x540
      const sample960 = finalizeSample({
        maxRight: 850,
        maxBottom: 500
      });
      assert.equal(sample960.slideWidth, 960);
      assert.equal(sample960.slideHeight, 540);
    });
  });

  describe('9. Slide Builder Utilities & Prompt Enhancer', () => {
    test('UT-P24: fitFontSize scales font down for dense text and isSubstantiveSlideBody filters empty/placeholder strings', () => {
      const shortSize = fitFontSize('Short title', 400, 120, { max: 24, min: 10 });
      const longText = Array.from({ length: 40 }, () => 'Enterprise cloud architecture scalability and resilience').join(' ');
      const denseSize = fitFontSize(longText, 220, 90, { max: 24, min: 10 });

      assert.equal(shortSize, 24);
      assert.ok(denseSize < shortSize);
      assert.ok(denseSize >= 10);

      assert.equal(isSubstantiveSlideBody('• Executive slide content'), false);
      assert.equal(isSubstantiveSlideBody('```'), false);
      assert.equal(isSubstantiveSlideBody('• Hybrid cloud adoption reduced infrastructure latency by 42% across global regions.'), true);
      assert.equal(typeof slideBuilderTest.syncDeadlineFor(4), 'number');
    });

    test('UT-P25: extractRequestedSlideCount and enhancePromptForPowerPoint detect creation counts while ignoring edit counts', () => {
      assert.equal(extractRequestedSlideCount('create 5 slides about cloud security'), 5);
      assert.equal(extractRequestedSlideCount('build a seven-slide presentation on Q4 results'), 7);
      assert.equal(extractRequestedSlideCount('summarize sustainability efforts in 10 slides'), 10);
      // Edit requests on existing slides must NOT trigger multi-slide generation
      assert.equal(extractRequestedSlideCount('shorten these 5 slides'), null);

      const multiSlideEnhanced = enhancePromptForPowerPoint(
        'create 5 slides to summarize our sustainability report',
        { userInstruction: 'create 5 slides to summarize our sustainability report' }
      );
      assert.match(multiSlideEnhanced, /EXACTLY 5 slides/i);
      assert.match(multiSlideEnhanced, /CARD LAYOUT BUDGET: At most 2 of the 5 slides/i);

      const editEnhanced = enhancePromptForPowerPoint('shorten this slide and make it punchy');
      assert.match(editEnhanced, /CRITICAL INSTRUCTIONS FOR SLIDE EDITING/i);
    });
  });
});

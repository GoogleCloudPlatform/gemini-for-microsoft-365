/**
 * Comprehensive Automated Test Suite for Excel Adapter & Taxonomy Categorization
 */

import { test, describe, beforeEach } from 'node:test';
import assert from 'node:assert/strict';
import {
  columnIndexToLetter,
  letterToColumnIndex,
  parseRangeAddress,
  resolveRangeAddress,
  normalizeFormula,
  extractExcelFormula,
  normalizeChartType,
  normalizeChartSpec,
  extractExcelChartSpec,
  normalizeChartEditSpec,
  extractExcelChartEditSpec,
  normalizeRangeFormatSpec,
  extractExcelRangeFormatSpec,
  normalizeTaxonomySpec,
  extractExcelTaxonomySpec,
  extractInsightsSummaryTable,
  normalizeInsightsTableData,
  applyTaxonomyCategorization,
  applyRangeFormatting,
  createAggregatedSummaryTableAndChart,
  insertNativeChart,
  editNativeChart,
  ExcelAdapter
} from '../src/adapters/ExcelAdapter.js';
import { ExcelInsightsService } from '../src/adapters/ExcelInsightsService.js';
import { parseMarkdown, cleanMarkdownEscapes } from '../src/core/markdownParser.js';

describe('Excel Adapter & Taxonomy Engine Test Suite', () => {

  describe('1. Column & Address Helpers', () => {
    test('UT-1: columnIndexToLetter & letterToColumnIndex bidirectional conversions', () => {
      assert.equal(columnIndexToLetter(0), 'A');
      assert.equal(columnIndexToLetter(1), 'B');
      assert.equal(columnIndexToLetter(2), 'C');
      assert.equal(columnIndexToLetter(25), 'Z');
      assert.equal(columnIndexToLetter(26), 'AA');
      assert.equal(columnIndexToLetter(27), 'AB');

      assert.equal(letterToColumnIndex('A'), 0);
      assert.equal(letterToColumnIndex('B'), 1);
      assert.equal(letterToColumnIndex('C'), 2);
      assert.equal(letterToColumnIndex('Z'), 25);
      assert.equal(letterToColumnIndex('AA'), 26);
      assert.equal(letterToColumnIndex('AB'), 27);
    });

    test('UT-2: parseRangeAddress correctly decomposes single cells and 2D bounds', () => {
      const single = parseRangeAddress('B2');
      assert.deepEqual(single, {
        startCol: 'B',
        startColIndex: 1,
        startRow: 2,
        endCol: 'B',
        endColIndex: 1,
        endRow: 2,
        address: 'B2:B2'
      });

      const multi = parseRangeAddress('Sheet1!$B$2:$D$10');
      assert.deepEqual(multi, {
        startCol: 'B',
        startColIndex: 1,
        startRow: 2,
        endCol: 'D',
        endColIndex: 3,
        endRow: 10,
        address: 'B2:D10'
      });
    });

    test('UT-2b: parseRangeAddress and resolveRangeAddress handle column-only and partial ranges', () => {
      const colOnly = parseRangeAddress('A:D');
      assert.deepEqual(colOnly, {
        startCol: 'A',
        startColIndex: 0,
        startRow: null,
        endCol: 'D',
        endColIndex: 3,
        endRow: null,
        address: 'A:D'
      });

      const singleCol = parseRangeAddress('B:B');
      assert.deepEqual(singleCol, {
        startCol: 'B',
        startColIndex: 1,
        startRow: null,
        endCol: 'B',
        endColIndex: 1,
        endRow: null,
        address: 'B:B'
      });

      const resolved = resolveRangeAddress('A:D', { startRow: 1, endRow: 25 });
      assert.equal(resolved.address, 'A1:D25');
      assert.equal(resolved.startRow, 1);
      assert.equal(resolved.endRow, 25);
    });
  });

  describe('2. Formula Extraction & Normalization', () => {
    test('UT-3: normalizeFormula and extractExcelFormula handle standard, fenced, and inline formulas', () => {
      assert.equal(normalizeFormula('=SUMIFS(C:C, D:D, "Approved")'), '=SUMIFS(C:C, D:D, "Approved")');
      assert.equal(normalizeFormula('SUMIFS(C:C, D:D, "Approved")'), '=SUMIFS(C:C, D:D, "Approved")');

      const fencedText = 'Here is your formula:\n```excel\n=XLOOKUP(A2, Products!A:A, Products!B:B, "Not Found")\n```';
      assert.equal(extractExcelFormula(fencedText), '=XLOOKUP(A2, Products!A:A, Products!B:B, "Not Found")');

      const codeText = 'Use `=SUM(B2:B20)` in cell B21';
      assert.equal(extractExcelFormula(codeText), '=SUM(B2:B20)');
    });

    test('UT-3b: normalizeFormula handles smart quotes, non-breaking spaces, and unicode math symbols', () => {
      // Smart quotes to standard double quotes
      assert.equal(normalizeFormula('=SUMIFS(C:C, D:D, “Approved”)'), '=SUMIFS(C:C, D:D, "Approved")');
      assert.equal(normalizeFormula("=COUNTIF(A:A, ‘Active’)"), '=COUNTIF(A:A, "Active")');

      // Preserves single quotes in sheet references while normalizing criteria literals
      assert.equal(normalizeFormula("=SUMIFS('Sales Data'!C:C, 'Sales Data'!D:D, 'Completed')"), '=SUMIFS(\'Sales Data\'!C:C, \'Sales Data\'!D:D, "Completed")');

      // Non-breaking spaces and unicode math operators
      const nbspFormula = '=SUM(A1:A10)\u00A0−\u00A05';
      assert.equal(normalizeFormula(nbspFormula), '=SUM(A1:A10) - 5');

      // Trailing punctuation stripping
      assert.equal(normalizeFormula('=AVERAGE(B2:B20);'), '=AVERAGE(B2:B20)');
    });

    test('UT-3c: extractExcelFormula parses multi-line code blocks and ignores comments', () => {
      const multilineBlock = `Here is the requested formula:
\`\`\`excel
// Calculate total completed revenue
=SUMIFS(
  C:C,
  D:D, “Completed”
)
\`\`\`
This sums column C where column D is Completed.`;

      assert.equal(extractExcelFormula(multilineBlock), '=SUMIFS( C:C, D:D, "Completed" )');

      // Rejects invalid JavaScript code blocks
      const jsBlock = '```javascript\nrange.formulas = [["=SUM(A1:A10)"]];\n```';
      assert.equal(extractExcelFormula(jsBlock), null);
    });
  });

  describe('3. Chart Specification & Editing', () => {
    test('UT-4: extractExcelChartSpec parses fenced excel-chart JSON and normalizes types', () => {
      const response = `Here is the chart:
\`\`\`excel-chart
{
  "type": "column",
  "categoriesRange": "B2:B10",
  "seriesRange": "C2:C10",
  "title": "Quarterly Revenue"
}
\`\`\``;
      const spec = extractExcelChartSpec(response);
      assert.ok(spec);
      assert.equal(spec.type, 'ColumnClustered');
      assert.equal(spec.categoriesRange, 'B2:B10');
      assert.equal(spec.seriesRange, 'C2:C10');
      assert.equal(spec.title, 'Quarterly Revenue');
    });

    test('UT-4b: normalizeChartSpec handles synonyms, arrays, and defaults seriesBy to columns', () => {
      const spec = normalizeChartSpec({
        chartType: 'doughnut',
        xAxis: 'A1:A10',
        series: ['B1:B10', 'C1:C10'],
        title: 'Share by Category'
      });
      assert.ok(spec);
      assert.equal(spec.type, 'Doughnut');
      assert.equal(spec.categoriesRange, 'A1:A10');
      assert.equal(spec.seriesRange, 'B1:B10, C1:C10');
      assert.equal(spec.seriesBy, 'columns');
      assert.equal(spec.title, 'Share by Category');
    });

    test('UT-4c: normalizeChartSpec & extractExcelChartSpec handle aggregate specs and markdown badges', () => {
      const aggResp = `\`\`\`excel-chart
{
  "type": "BarClustered",
  "categoriesRange": "D1:D50",
  "seriesRange": "E1:E50",
  "aggregate": "sum",
  "title": "Distribution of Total Transaction Amount by Category"
}
\`\`\``;
      const spec = extractExcelChartSpec(aggResp);
      assert.ok(spec);
      assert.equal(spec.type, 'BarClustered');
      assert.equal(spec.categoriesRange, 'D1:D50');
      assert.equal(spec.seriesRange, 'E1:E50');
      assert.equal(spec.aggregate, 'sum');
      assert.equal(spec.title, 'Distribution of Total Transaction Amount by Category');

      // Verify Markdown Parser generates the aggregation badge
      const html = parseMarkdown(aggResp);
      assert.ok(html.includes('data-chart-aggregate="sum"'));
      assert.ok(html.includes('Aggregate:'));
      assert.ok(html.includes('SUM'));
    });

    test('UT-5: extractExcelChartEditSpec handles chart modifications', () => {
      const editResp = `\`\`\`excel-chart-edit
{
  "targetChart": "Revenue Chart",
  "chartType": "pie",
  "title": "Updated Share"
}
\`\`\``;
      const editSpec = extractExcelChartEditSpec(editResp);
      assert.ok(editSpec);
      assert.equal(editSpec.action, 'edit');
      assert.equal(editSpec.targetChart, 'Revenue Chart');
      assert.equal(editSpec.chartType, 'Pie');
      assert.equal(editSpec.title, 'Updated Share');
    });
  });

  describe('4. Range Formatting Engine', () => {
    test('UT-6: extractExcelRangeFormatSpec parses JSON and JS pseudo-code formatting', () => {
      const formatJson = `\`\`\`excel-range-format
{
  "targetRange": "B1:D1",
  "fill": { "color": "#0078d4" },
  "font": { "bold": true, "color": "#ffffff" },
  "horizontalAlignment": "Center",
  "autofitColumns": true
}
\`\`\``;
      const spec = extractExcelRangeFormatSpec(formatJson);
      assert.ok(spec);
      assert.equal(spec.targetRange, 'B1:D1');
      assert.equal(spec.fill.color.toLowerCase(), '#0078d4');
      assert.equal(spec.font.bold, true);
      assert.equal(spec.font.color.toLowerCase(), '#ffffff');
      assert.equal(spec.horizontalAlignment, 'Center');
      assert.equal(spec.autofitColumns, true);
    });
  });

  describe('5. Taxonomy Categorization Engine', () => {
    test('UT-7: normalizeTaxonomySpec defaults target column to adjacent column and builds category summary', () => {
      const rawSpec = {
        sourceColumn: 'B',
        header: 'Expense Category',
        startRow: 2,
        categories: [
          { row: 2, sourceValue: 'Uber *Trip San Francisco', category: 'Travel / Transportation' },
          { row: 3, sourceValue: 'AWS Cloud Services', category: 'Cloud Infrastructure' },
          { row: 4, sourceValue: 'Lyft Ride', category: 'Travel / Transportation' }
        ]
      };

      const normalized = normalizeTaxonomySpec(rawSpec);
      assert.ok(normalized);
      assert.equal(normalized.action, 'categorize');
      assert.equal(normalized.sourceColumn, 'B');
      assert.equal(normalized.targetColumn, 'C');
      assert.equal(normalized.insertColumn, true);
      assert.equal(normalized.header, 'Expense Category');
      assert.equal(normalized.startRow, 2);
      assert.equal(normalized.endRow, 4);
      assert.equal(normalized.count, 3);
      assert.deepEqual(normalized.categoryValues, [
        'Travel / Transportation',
        'Cloud Infrastructure',
        'Travel / Transportation'
      ]);
      assert.deepEqual(normalized.taxonomySummary, {
        'Travel / Transportation': 2,
        'Cloud Infrastructure': 1
      });
    });

    test('UT-8: extractExcelTaxonomySpec extracts fenced ```excel-taxonomy JSON', () => {
      const aiResponse = `Here is the taxonomy categorization:
\`\`\`excel-taxonomy
{
  "action": "categorize",
  "sourceColumn": "B",
  "targetColumn": "C",
  "insertColumn": true,
  "header": "Expense Category",
  "startRow": 2,
  "categories": [
    { "row": 2, "sourceValue": "Delta Air Lines", "category": "Travel" },
    { "row": 3, "sourceValue": "Sweetgreen SF", "category": "Meals" },
    { "row": 4, "sourceValue": "GitHub Enterprise", "category": "Software" }
  ],
  "taxonomySummary": {
    "Travel": 1,
    "Meals": 1,
    "Software": 1
  }
}
\`\`\``;

      const spec = extractExcelTaxonomySpec(aiResponse);
      assert.ok(spec);
      assert.equal(spec.sourceColumn, 'B');
      assert.equal(spec.targetColumn, 'C');
      assert.equal(spec.header, 'Expense Category');
      assert.equal(spec.count, 3);
      assert.equal(spec.categoryValues[0], 'Travel');
      assert.equal(spec.categoryValues[1], 'Meals');
      assert.equal(spec.categoryValues[2], 'Software');
    });

    test('UT-9: extractExcelTaxonomySpec falls back to markdown table extraction', () => {
      const tableResponse = `Here is your categorized data:
| Row | Raw Merchant Name | Expense Category |
|---|---|---|
| 2 | Starbucks 1092 | Meals & Entertainment |
| 3 | Hilton Hotels | Travel & Lodging |
| 4 | JetBrains All Products | Software Tooling |
`;

      const spec = extractExcelTaxonomySpec(tableResponse);
      assert.ok(spec);
      assert.equal(spec.sourceColumn, 'B');
      assert.equal(spec.targetColumn, 'C');
      assert.equal(spec.count, 3);
      assert.equal(spec.header, 'Expense Category');
      assert.equal(spec.categoryValues[0], 'Meals & Entertainment');
    });

    test('UT-10: applyTaxonomyCategorization inserts column, sets styled header, and populates data batch', async () => {
      // Mock Office.js context and Worksheet
      const insertedRanges = [];
      const writtenRanges = {};
      const autofitRanges = [];

      global.Excel = {
        InsertShiftDirection: { right: 'Right', down: 'Down' }
      };

      const mockSheet = {
        getRange: (address) => {
          const rangeObj = {
            address,
            values: [],
            format: {
              font: {},
              fill: {},
              horizontalAlignment: '',
              verticalAlignment: '',
              autofitColumns: () => {
                autofitRanges.push(address);
              }
            },
            insert: (dir) => {
              insertedRanges.push({ address, dir });
            }
          };
          writtenRanges[address] = rangeObj;
          return rangeObj;
        }
      };

      const mockContext = {
        sync: async () => {}
      };

      const spec = {
        sourceColumn: 'B',
        targetColumn: 'C',
        insertColumn: true,
        header: 'Expense Category',
        startRow: 2,
        categories: [
          { row: 2, sourceValue: 'Uber', category: 'Travel' },
          { row: 3, sourceValue: 'AWS', category: 'Cloud' }
        ]
      };

      const result = await applyTaxonomyCategorization(mockContext, mockSheet, spec);

      assert.equal(result.sourceColumn, 'B');
      assert.equal(result.targetColumn, 'C');
      assert.equal(result.count, 2);
      assert.equal(result.address, 'C1:C3');

      // Verify column insertion occurred on C:C
      assert.equal(insertedRanges.length, 1);
      assert.equal(insertedRanges[0].address, 'C:C');
      assert.equal(insertedRanges[0].dir, 'Right');

      // Verify header cell C1 was styled
      const headerRange = writtenRanges['C1'];
      assert.ok(headerRange);
      assert.deepEqual(headerRange.values, [['Expense Category']]);
      assert.equal(headerRange.format.font.bold, true);
      assert.equal(headerRange.format.fill.color, '#EBF3FC');

      // Verify data range C2:C3 was written with batch values
      const dataRange = writtenRanges['C2:C3'];
      assert.ok(dataRange);
      assert.deepEqual(dataRange.values, [['Travel'], ['Cloud']]);

      // Verify autofit was called on the full column span C1:C3
      assert.ok(autofitRanges.includes('C1:C3'));
    });

    test('UT-10b: extractExcelTaxonomySpec and normalizeTaxonomySpec handle multi-column taxonomy catalog schemas', () => {
      const response = `Based on the analysis, here is the taxonomy:
\`\`\`json
{
  "version": "1.0",
  "taxonomy": [
    {
      "columnName": "Sales Rep",
      "type": "one-to-one",
      "mapping": [
        { "value": "David Kim", "category": "David Kim" },
        { "value": "Elena Rostova", "category": "Elena Rostova" },
        { "value": "Marcus Vance", "category": "Marcus Vance" },
        { "value": "Sarah Jenkins", "category": "Sarah Jenkins" }
      ]
    }
  ]
}
\`\`\``;

      const spec = extractExcelTaxonomySpec(response, '', { isTaxonomy: true });
      assert.ok(spec);
      assert.equal(spec.action, 'categorize');
      assert.equal(spec.header, 'Sales Rep Category');
      assert.equal(spec.count, 4);
      assert.deepEqual(spec.categoryValues, ['David Kim', 'Elena Rostova', 'Marcus Vance', 'Sarah Jenkins']);
    });
  });

  describe('6. Markdown Parser Visual Components', () => {
    test('UT-11: parseMarkdown renders interactive visual card for ```excel-taxonomy', () => {
      const md = `\`\`\`excel-taxonomy
{
  "sourceColumn": "B",
  "targetColumn": "C",
  "header": "Expense Category",
  "startRow": 2,
  "categories": [
    { "sourceValue": "Uber SF", "category": "Travel" },
    { "sourceValue": "AWS", "category": "Cloud" }
  ],
  "taxonomySummary": {
    "Travel": 1,
    "Cloud": 1
  }
}
\`\`\``;

      const html = parseMarkdown(md);
      assert.ok(html.includes('Taxonomy Categorization'));
      assert.ok(html.includes('Col B ➔ Col C'));
      assert.ok(html.includes('Expense Category'));
      assert.ok(html.includes('Travel'));
      assert.ok(html.includes('Cloud'));
    });

    test('UT-12: parseMarkdown renders visual card for ```excel-range-format', () => {
      const md = `\`\`\`excel-range-format
{
  "targetRange": "A1:E1",
  "fill": { "color": "#0078d4" },
  "font": { "bold": true }
}
\`\`\``;

      const html = parseMarkdown(md);
      assert.ok(html.includes('Excel Range Formatting'));
      assert.ok(html.includes('Target: A1:E1'));
      assert.ok(html.includes('#0078D4') || html.includes('#0078d4'));
    });
  });

  describe('7. Excel Quick Action Chips & Response Actions UI Contracts', () => {
    test('UT-13: selectionChips contains Categorize Taxonomy chip as first chip', async () => {
      const { selectionChips } = await import('../src/taskpanes/excel/chips.js');
      assert.ok(Array.isArray(selectionChips));
      assert.ok(selectionChips.length >= 5);

      const taxonomyChip = selectionChips.find(c => c.id === 'selTaxonomy');
      assert.ok(taxonomyChip, 'selTaxonomy chip must exist');
      assert.equal(taxonomyChip.label, '🏷️ Categorize Taxonomy');
      assert.ok(taxonomyChip.instruction.includes('excel-taxonomy'));
      assert.ok(taxonomyChip.instruction.includes('highlighted column'));
    });

    test('UT-14: Excel taskpane module exports docChips and selectionChips', async () => {
      const excelModule = (await import('../src/taskpanes/excel/index.js')).default;
      assert.ok(excelModule);
      assert.ok(Array.isArray(excelModule.docChips));
      assert.ok(excelModule.docChips.length > 0);
      assert.ok(Array.isArray(excelModule.selectionChips));
      assert.ok(excelModule.selectionChips.length > 0);
    });

    test('UT-15: renderResponseActions renders "🏷️ Insert Categories Column" for taxonomy specifications', async () => {
      const { renderResponseActions } = await import('../src/taskpanes/excel/responseActions.js');
      
      // Setup simple DOM mock if not present
      if (typeof document === 'undefined') {
        globalThis.document = {
          createElement: (tag) => ({
            tagName: tag.toUpperCase(),
            className: '',
            innerHTML: '',
            innerText: '',
            title: '',
            style: {},
            querySelector: () => null,
            appendChild: function(child) { (this.children = this.children || []).push(child); },
            onclick: null
          })
        };
      }

      const taxonomyJson = `\`\`\`excel-taxonomy
{
  "sourceColumn": "B",
  "targetColumn": "C",
  "header": "Expense Category",
  "categories": ["Software", "Travel"]
}
\`\`\``;

      const actionsEl = renderResponseActions({
        rawText: taxonomyJson,
        formattedHtml: parseMarkdown(taxonomyJson),
        textDiv: { innerHTML: parseMarkdown(taxonomyJson) },
        adapter: {},
        chatLog: {}
      });

      assert.ok(actionsEl);
      assert.ok(actionsEl.children && actionsEl.children.length >= 1);
      const insertBtn = actionsEl.children[0];
      assert.equal(insertBtn.className, 'action-btn taxonomy');
      assert.equal(insertBtn.innerHTML, '🏷️ Insert Categories Column');
      assert.ok(insertBtn.title.includes('Insert Column C'));
    });

    test('UT-16: extractInsightsSummaryTable extracts full matrix and does not trigger taxonomy', () => {
      const insightsResponse = `### Key Table Patterns & Trends
- High transaction volume in Meals & Entertainment.

### Structured Insights Summary Table
| Metric / Category | Total Amount | Transaction Count | % of Total |
| Travel | $4,520.00 | 12 | 45.2% |
| Meals | $1,850.50 | 25 | 18.5% |
| Office Supplies | $920.00 | 8 | 9.2% |`;

      const extracted = extractInsightsSummaryTable(insightsResponse);
      assert.ok(extracted);
      assert.equal(extracted.length, 4); // Header + 3 data rows
      assert.equal(extracted[0].length, 4);
      assert.equal(extracted[0][0], 'Metric / Category');
      assert.equal(extracted[1][0], 'Travel');
      assert.equal(extracted[1][1], '$4,520.00');

      // Taxonomy extraction should be guarded against insights summary responses
      const taxonomySpec = extractExcelTaxonomySpec(insightsResponse);
      assert.equal(taxonomySpec, null);
    });

    test('UT-17: renderResponseActions renders "📊 Insert Insights Summary Table" for Insights responses', async () => {
      const { renderResponseActions } = await import('../src/taskpanes/excel/responseActions.js');

      const insightsResponse = `### Structured Insights Summary Table
| Category | Subtotal |
| Travel | $1,200 |
| Meals | $800 |`;

      const actionsEl = renderResponseActions({
        rawText: insightsResponse,
        formattedHtml: parseMarkdown(insightsResponse),
        textDiv: { innerHTML: parseMarkdown(insightsResponse) },
        adapter: {},
        chatLog: {},
        isInsights: true
      });

      assert.ok(actionsEl);
      assert.ok(actionsEl.children && actionsEl.children.length >= 1);
      const insertBtn = actionsEl.children[0];
      assert.equal(insertBtn.innerHTML, '📊 Insert Insights Summary Table');
      assert.ok(insertBtn.title.includes('Structured Insights Summary Table'));
    });

    test('UT-18: Scoped action buttons ensure Taxonomy Categorizer shows category column button and Insights shows insights table button', async () => {
      const { renderResponseActions } = await import('../src/taskpanes/excel/responseActions.js');

      // 1. Taxonomy Categorizer context
      const taxonomyContextEl = renderResponseActions({
        rawText: `| Merchant | Category |\n| Uber | Travel |\n| Starbucks | Meals |`,
        formattedHtml: `<table><tr><th>Merchant</th><th>Category</th></tr><tr><td>Uber</td><td>Travel</td></tr></table>`,
        textDiv: { innerHTML: `<table><tr><th>Merchant</th><th>Category</th></tr><tr><td>Uber</td><td>Travel</td></tr></table>` },
        adapter: {},
        chatLog: {},
        chip: { id: 'chipTaxonomy', label: '🏷️ Taxonomy Categorizer' },
        isTaxonomy: true
      });

      const taxBtn = taxonomyContextEl.children[0];
      assert.equal(taxBtn.innerHTML, '🏷️ Insert Categories Column');

      // 2. Insights context
      const insightsContextEl = renderResponseActions({
        rawText: `### Key Table Patterns & Trends\n- High travel.\n\n### Structured Insights Summary Table\n| Metric | Total |\n| Travel | $5,000 |`,
        formattedHtml: `<h3>Key Table Patterns & Trends</h3><p>High travel.</p><h3>Structured Insights Summary Table</h3><table><tr><th>Metric</th><th>Total</th></tr><tr><td>Travel</td><td>$5,000</td></tr></table>`,
        textDiv: { innerHTML: `` },
        adapter: {},
        chatLog: {},
        chip: { id: 'chipInsightsSheet', label: '📈 Insights (Sheet)', isInsights: true },
        isInsights: true
      });

      const insBtn = insightsContextEl.children[0];
      assert.equal(insBtn.innerHTML, '📊 Insert Insights Summary Table');
    });

    test('UT-19: normalizeInsightsTableData converts currencies, percentages, and numbers to JS numbers and detects formats', () => {
      const rawMatrix = [
        ['Category', 'Total Revenue', 'Transaction Count', 'Growth Rate'],
        ['Travel', '$4,520.00', '12', '45.2%'],
        ['Meals', '$1,850.50', '25', '18.5%'],
        ['Office Supplies', '$920.00', '8', '-9.2%']
      ];

      const result = normalizeInsightsTableData(rawMatrix);
      assert.ok(result);
      assert.equal(result.headers.length, 4);
      assert.deepEqual(result.headers, ['Category', 'Total Revenue', 'Transaction Count', 'Growth Rate']);

      // Check numeric columns detected
      assert.deepEqual(result.numericColumns, [1, 2, 3]);

      // Check column formats
      assert.equal(result.columnFormats[0], null); // Text
      assert.equal(result.columnFormats[1], '$#,##0.00'); // Currency
      assert.equal(result.columnFormats[2], '#,##0'); // Integer count
      assert.equal(result.columnFormats[3], '0.0%'); // Percentage

      // Check cleaned values are genuine numbers
      const cleaned = result.cleanedMatrix;
      assert.equal(cleaned.length, 4);
      assert.equal(cleaned[1][0], 'Travel');
      assert.equal(cleaned[1][1], 4520.00); // parsed float
      assert.equal(typeof cleaned[1][1], 'number');
      assert.equal(cleaned[1][2], 12); // parsed integer
      assert.equal(typeof cleaned[1][2], 'number');
      assert.equal(Math.round(cleaned[1][3] * 1000) / 1000, 0.452); // parsed percentage
      assert.equal(typeof cleaned[1][3], 'number');

      assert.equal(cleaned[3][1], 920.00);
      assert.equal(Math.round(cleaned[3][3] * 1000) / 1000, -0.092);
    });

    test('UT-20: Prompt to update chart renders "🔄 Update Chart" and does not misidentify as taxonomy', async () => {
      const { renderResponseActions } = await import('../src/taskpanes/excel/responseActions.js');

      const chartUpdateResp = `Here is the updated chart specification:
\`\`\`excel-chart-edit
{
  "targetChart": "active",
  "chartType": "Doughnut",
  "title": "Revenue by Category",
  "legend": { "visible": true, "position": "Bottom" },
  "dataLabels": { "visible": true, "showPercentage": true }
}
\`\`\``;

      // Taxonomy extraction should be null
      const taxSpec = extractExcelTaxonomySpec(chartUpdateResp);
      assert.equal(taxSpec, null);

      // Chart edit extraction should succeed
      const editSpec = extractExcelChartEditSpec(chartUpdateResp, '', { userPrompt: 'Change this chart to a Doughnut chart' });
      assert.ok(editSpec);
      assert.equal(editSpec.chartType, 'Doughnut');
      assert.equal(editSpec.title, 'Revenue by Category');

      const actionsEl = renderResponseActions({
        rawText: chartUpdateResp,
        formattedHtml: parseMarkdown(chartUpdateResp),
        textDiv: { innerHTML: parseMarkdown(chartUpdateResp) },
        adapter: {},
        chatLog: {},
        userPrompt: 'Change this chart to a Doughnut chart'
      });

      assert.ok(actionsEl);
      const updateBtn = actionsEl.children[0];
      assert.equal(updateBtn.innerHTML, '🔄 Update Chart');
      assert.ok(updateBtn.className.includes('chart-edit'));
    });

    test('UT-21: Chart edit spec extraction with prompt intent converts regular chart block to edit spec', () => {
      const chartResp = `\`\`\`excel-chart
{
  "type": "Pie",
  "title": "Updated Spend Distribution",
  "categoriesRange": "A2:A10",
  "seriesRange": "B2:B10"
}
\`\`\``;

      // With chart update prompt, extractExcelChartEditSpec should return an edit spec
      const editSpec = extractExcelChartEditSpec(chartResp, '', { userPrompt: 'Update the chart to a pie chart' });
      assert.ok(editSpec);
      assert.equal(editSpec.action, 'edit');
      assert.equal(editSpec.chartType, 'Pie');
      assert.equal(editSpec.title, 'Updated Spend Distribution');

      // extractExcelTaxonomySpec should definitely be null
      const taxSpec = extractExcelTaxonomySpec(chartResp);
      assert.equal(taxSpec, null);
    });

    test('UT-22: Chart JSON containing "categories" does not get misidentified as taxonomy categorization', () => {
      const chartWithCategories = `{
  "type": "ColumnClustered",
  "categories": "A2:A10",
  "header": "Sales Category",
  "dataRange": "A1:B10"
}`;

      const taxSpec = extractExcelTaxonomySpec(chartWithCategories);
      assert.equal(taxSpec, null);
    });

    test('UT-23: Behavior 2 - Markdown parser handles long separator rows and single-cell subheaders correctly', () => {
      const markdown = `### Structured Insights Summary Table

| Metric / Category | Value | Status / Observation |
| :-------------------------- | :---- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| Key Metrics |
| Total Transactions | 7 | Consistent activity |
| Total Spend | $1428.25 | Overall expenditure |
| Category Performance |
| Travel / Transportation | $1054.75 | Highest spending category |
| Cloud & Infrastructure | $300.00 | Significant single expense |
| Trend Status |
| Spending Distribution | Skewed | Heavily weighted towards Travel |
`;

      const parsedHtml = parseMarkdown(markdown);
      assert.ok(parsedHtml.includes('<table'));
      // Separator dashed line must NOT be rendered as a raw paragraph or visible text
      assert.ok(!parsedHtml.includes(':--------------------------'));
      assert.ok(!parsedHtml.includes('--------------------------------------------------------------------------------'));
      assert.ok(parsedHtml.includes('Key Metrics'));
      assert.ok(parsedHtml.includes('Category Performance'));
      assert.ok(parsedHtml.includes('Travel / Transportation'));
    });

    test('UT-24: Behavior 3 - Clean escaped backslashes and detect intelligent categorical chart sub-range', () => {
      const rawMatrix = [
        ['Metric / Category', 'Value', 'Status / Observation'],
        ['Key Metrics', '', ''],
        ['Total Transactions', '7', 'Consistent activity'],
        ['Total Spend', '\\$1428.25', 'Overall expenditure'],
        ['Average Transaction', '\\$204.04', 'Moderate average'],
        ['Category Performance', '', ''],
        ['Travel / Transportation', '\\$1054.75', 'Highest spending category'],
        ['Cloud & Infrastructure', '\\$300.00', 'Significant single expense'],
        ['Meals & Entertainment', '\\$39.00', 'Lowest spending category'],
        ['Office Supplies / Maintenance', '\\$34.50', 'Minor operational expense'],
        ['Trend Status', '', ''],
        ['Spending Distribution', 'Skewed', 'Heavily weighted towards Travel'],
        ['Data Quality', 'Good', 'No missing values, minor naming inconsistency']
      ];

      const normalized = normalizeInsightsTableData(rawMatrix);
      assert.ok(normalized);

      // Verify no backslashes in cleanedMatrix
      for (const row of normalized.cleanedMatrix) {
        for (const cell of row) {
          if (typeof cell === 'string') {
            assert.ok(!cell.includes('\\$'), `Cell "${cell}" should not contain escaped backslash`);
          }
        }
      }

      // Verify Total Spend and Travel / Transportation parsed to JS numbers
      assert.equal(normalized.cleanedMatrix[3][1], 1428.25);
      assert.equal(typeof normalized.cleanedMatrix[3][1], 'number');
      assert.equal(normalized.cleanedMatrix[6][1], 1054.75);
      assert.equal(typeof normalized.cleanedMatrix[6][1], 'number');

      // Verify intelligent sub-section chart detection targets Category Performance (rows 6..9)
      assert.ok(normalized.chartRangeInfo);
      assert.equal(normalized.chartRangeInfo.isSubSection, true);
      assert.equal(normalized.chartRangeInfo.startRowOffset, 6);
      assert.equal(normalized.chartRangeInfo.rowCount, 4);
      assert.equal(normalized.chartRangeInfo.colCount, 2);
      assert.equal(normalized.chartRangeInfo.chartType, 'barClustered');
      assert.ok(normalized.chartRangeInfo.title.includes('Category Performance'));
    });

    test('UT-25: Behavior 1 - Natural language chart prompt produces Update Chart button', async () => {
      const { renderResponseActions } = await import('../src/taskpanes/excel/responseActions.js');

      const naturalLangResponse = `I've updated the chart to represent the Category Performance data as a horizontal Bar Chart with clear data labels.`;

      const actionsEl = renderResponseActions({
        rawText: naturalLangResponse,
        formattedHtml: parseMarkdown(naturalLangResponse),
        textDiv: { innerHTML: parseMarkdown(naturalLangResponse) },
        adapter: {},
        chatLog: {},
        userPrompt: 'Change this chart to a horizontal bar chart'
      });

      assert.ok(actionsEl);
      const firstBtn = actionsEl.children[0];
      assert.equal(firstBtn.innerHTML, '🔄 Update Chart');
      assert.ok(firstBtn.className.includes('chart-edit'));
      // Must NOT be taxonomy
      assert.ok(!firstBtn.className.includes('taxonomy'));
      assert.ok(!firstBtn.innerHTML.includes('Categories Column'));
    });

    test('UT-25b: Formula response action renders Insert Formula, Formula Preview Pill, Copy to Clipboard and suppresses Range Formatting', async () => {
      const { renderResponseActions } = await import('../src/taskpanes/excel/responseActions.js');

      const formulaResponse = `To calculate the sum of all weighted values across your active deal pipeline, you can use a formula that totals the values in Column J (from Row 2 to Row 15).

Here is the exact formula for your target cell P2:

\`\`\`excel
=SUM(J2:J15)
\`\`\`

📊 Calculation Summary
Target Range: J2:J15
Calculated Total: $1,003,700

🎨 Excel Range Formatting  Target: P2
{ "targetRange": "P2", "fill": { "color": "#E0F2FE" }, "font": { "bold": true } }`;

      const actionsEl = renderResponseActions({
        rawText: formulaResponse,
        formattedHtml: parseMarkdown(formulaResponse),
        textDiv: { innerHTML: parseMarkdown(formulaResponse) },
        adapter: {},
        chatLog: { createCopyButton: () => ({ className: 'action-btn copy', innerHTML: '📋 Copy to Clipboard' }) },
        userPrompt: 'create a formula to calculate total of weighted values'
      });

      assert.ok(actionsEl);
      assert.ok(actionsEl.children && actionsEl.children.length >= 2);
      
      // First button must be ➕ Insert Formula
      const insertBtn = actionsEl.children[0];
      assert.equal(insertBtn.innerHTML, '➕ Insert Formula');
      assert.equal(insertBtn.className, 'action-btn formula');
      assert.ok(insertBtn.title.includes('=SUM(J2:J15)'));

      // Second element must be the formula preview pill
      const pill = actionsEl.children[1];
      assert.equal(pill.className, 'formula-preview-pill');
      assert.equal(pill.innerText, '=SUM(J2:J15)');

      // Verify that Range Formatting button (🎨 Apply Range Format) is NOT rendered
      const allButtonsHtml = actionsEl.innerHTML || actionsEl.children.map(c => c.innerHTML || c.innerText).join(' ');
      assert.ok(!allButtonsHtml.includes('Apply Range Format'), 'Must not render Apply Range Format for formula response');
      assert.ok(!allButtonsHtml.includes('New Summary Tab'), 'Must not render New Summary Tab for formula response');
      assert.ok(!allButtonsHtml.includes('Insert Interactive Chart'), 'Must not render Insert Interactive Chart for formula response');
    });

    test('UT-25c: Taxonomy button click renders 🏷️ Insert Categories Column for taxonomy response', async () => {
      const { renderResponseActions } = await import('../src/taskpanes/excel/responseActions.js');

      const taxonomyResponse = `Based on the analysis of the provided spreadsheet data, I have identified the columns containing categorical text and generated the requested excel-taxonomy JSON specification.

\`\`\`json
{
  "version": "1.0",
  "taxonomy": [
    {
      "columnName": "Sales Rep",
      "mapping": [
        { "value": "David Kim", "category": "David Kim" }
      ]
    }
  ]
}
\`\`\``;

      const actionsEl = renderResponseActions({
        rawText: taxonomyResponse,
        formattedHtml: parseMarkdown(taxonomyResponse),
        textDiv: { innerHTML: parseMarkdown(taxonomyResponse) },
        adapter: {},
        chatLog: { createCopyButton: () => ({ className: 'action-btn copy', innerHTML: '📋 Copy to Clipboard' }) },
        userPrompt: 'Categorize this data'
      });

      assert.ok(actionsEl);
      const firstBtn = actionsEl.children[0];
      assert.equal(firstBtn.innerHTML, '🏷️ Insert Categories Column');
      assert.equal(firstBtn.className, 'action-btn taxonomy');
    });
  });

  describe('8. Special Character & Backslash ($500 vs \\$500) Sanitization', () => {
    test('UT-26: cleanCellText strips escaped backslashes from currency, numbers, and markdown symbols', () => {
      // Escaped dollar signs
      assert.equal(ExcelInsightsService.cleanCellText('\\$500'), '$500');
      assert.equal(ExcelInsightsService.cleanCellText('\\\\$500'), '$500');
      assert.equal(ExcelInsightsService.cleanCellText('\\$500\\'), '$500');
      assert.equal(ExcelInsightsService.cleanCellText('\\$1,428.25'), '$1,428.25');
      assert.equal(ExcelInsightsService.cleanCellText('Total: \\$5,230.50'), 'Total: $5,230.50');
      assert.equal(ExcelInsightsService.cleanCellText('**\\$500**'), '$500');

      // Escaped other currency and math symbols
      assert.equal(ExcelInsightsService.cleanCellText('\\€750.00'), '€750.00');
      assert.equal(ExcelInsightsService.cleanCellText('\\£1,200'), '£1,200');
      assert.equal(ExcelInsightsService.cleanCellText('\\¥50,000'), '¥50,000');
      assert.equal(ExcelInsightsService.cleanCellText('\\%15.5'), '%15.5');
      assert.equal(ExcelInsightsService.cleanCellText('15.5\\%'), '15.5%');

      // Escaped digits
      assert.equal(ExcelInsightsService.cleanCellText('\\500'), '500');

      // Bullet points and complex strings
      assert.equal(ExcelInsightsService.cleanCellText('• Total Spend: \\$1,428.25 (Est.)'), '• Total Spend: $1,428.25 (Est.)');
      assert.equal(ExcelInsightsService.cleanCellText('Growth: \\$500 - \\$600'), 'Growth: $500 - $600');

      // Non-string or falsy values
      assert.equal(ExcelInsightsService.cleanCellText(500), '500');
      assert.equal(ExcelInsightsService.cleanCellText(null), '');
      assert.equal(ExcelInsightsService.cleanCellText(undefined), '');
      assert.equal(ExcelInsightsService.cleanCellText(''), '');
    });

    test('UT-27: cleanMarkdownEscapes unescapes markdown text properly', () => {
      assert.equal(cleanMarkdownEscapes('\\$500'), '$500');
      assert.equal(cleanMarkdownEscapes('\\$500\\'), '$500');
      assert.equal(cleanMarkdownEscapes('Revenue was \\$94.5B with \\$39.54B profit.'), 'Revenue was $94.5B with $39.54B profit.');
      assert.equal(cleanMarkdownEscapes(''), '');
      assert.equal(cleanMarkdownEscapes(null), '');
    });

    test('UT-28: parseMarkdown renders HTML with clean unescaped currency symbols', () => {
      const md = `### Executive Summary
- Total Revenue: **\\$1,428.25**
- Profit Margin: **18.5\\%**

| Category | Amount |
|---|---|
| Travel | \\$1,054.75 |
| Meals | \\$39.00 |`;

      const html = parseMarkdown(md);
      assert.ok(!html.includes('\\$'), 'Rendered HTML must not contain backslash-escaped dollar sign');
      assert.ok(!html.includes('\\%'), 'Rendered HTML must not contain backslash-escaped percent sign');
      assert.ok(html.includes('$1,428.25'));
      assert.ok(html.includes('$1,054.75'));
      assert.ok(html.includes('$39.00'));
      assert.ok(html.includes('18.5%'));
    });

    test('UT-29: normalizeInsightsTableData correctly parses escaped \\$ numbers to pure floats', () => {
      const tableWithEscapes = [
        ['Metric / Category', 'Value'],
        ['Total Spend', '\\$1,428.25'],
        ['Average Price', '\\$500'],
        ['Tax', '\\$39.50']
      ];

      const normalized = normalizeInsightsTableData(tableWithEscapes);
      assert.ok(normalized);
      assert.equal(normalized.cleanedMatrix[1][1], 1428.25);
      assert.equal(typeof normalized.cleanedMatrix[1][1], 'number');
      assert.equal(normalized.cleanedMatrix[2][1], 500);
      assert.equal(typeof normalized.cleanedMatrix[2][1], 'number');
      assert.equal(normalized.cleanedMatrix[3][1], 39.5);
      assert.equal(typeof normalized.cleanedMatrix[3][1], 'number');
      assert.equal(normalized.columnFormats[1], '$#,##0.00');
    });

    test('UT-30: parseNumberWithUnit extracts clean number and unit from escaped strings', () => {
      const parsed1 = ExcelInsightsService.parseNumberWithUnit('\\$94.5B');
      assert.ok(parsed1);
      assert.equal(parsed1.num, 94.5);
      assert.equal(parsed1.unit, 'B');

      const parsed2 = ExcelInsightsService.parseNumberWithUnit('\\$500');
      assert.ok(parsed2);
      assert.equal(parsed2.num, 500);

      const parsed3 = ExcelInsightsService.parseNumberWithUnit('**\\$1,200.50M**');
      assert.ok(parsed3);
      assert.equal(parsed3.num, 1200.5);
      assert.equal(parsed3.unit, 'M');
    });
  });

  describe('9. Excel 2-Row Header Layout, Quick Actions Consolidation & Session Reset Lifecycle', () => {
    test('UT-31: Excel taskpane 2-row header layout conformance & removal of redundant yellow quick actions box', async () => {
      const fs = await import('node:fs');
      const path = await import('node:path');
      const htmlPath = path.resolve('src/taskpanes/excel/index.html');
      const htmlContent = fs.readFileSync(htmlPath, 'utf8');

      // 1. Top Preview Container exists with 2-row structure
      assert.ok(htmlContent.includes('id="excelSchemaPreview"'), 'excelSchemaPreview must exist');
      assert.ok(htmlContent.includes('class="schema-row primary-row"'), 'primary-row must exist on Row 1');
      assert.ok(htmlContent.includes('class="schema-row chart-row"'), 'chart-row must exist on Row 2');

      // 2. Row 1 contains Worksheet chip on Left and Target Cell/Range on Right
      const primaryRowStart = htmlContent.indexOf('class="schema-row primary-row"');
      const chartRowStart = htmlContent.indexOf('class="schema-row chart-row"');
      const sheetBadgePos = htmlContent.indexOf('id="sheetSchemaBadge"');
      const rangeContainerPos = htmlContent.indexOf('class="target-range-container"');
      const targetRangePos = htmlContent.indexOf('id="targetRangeBadge"');

      assert.ok(primaryRowStart > 0, 'primaryRow must exist');
      assert.ok(sheetBadgePos > primaryRowStart && sheetBadgePos < chartRowStart, 'sheetSchemaBadge must be in primary-row');
      assert.ok(rangeContainerPos > sheetBadgePos && rangeContainerPos < chartRowStart, 'target-range-container must be after sheetSchemaBadge in primary-row');
      assert.ok(targetRangePos > rangeContainerPos && targetRangePos < chartRowStart, 'targetRangeBadge must be inside target-range-container');

      // 3. Row 2 contains Chart target badge
      const chartBadgePos = htmlContent.indexOf('id="targetChartBadge"');
      const removeChartBtnPos = htmlContent.indexOf('id="removeTargetChartBtn"');
      assert.ok(chartBadgePos > chartRowStart, 'targetChartBadge must be inside chart-row');
      assert.ok(removeChartBtnPos > chartBadgePos, 'removeTargetChartBtn must be inside targetChartBadge');

      // 4. Redundant yellow quick actions box is completely removed from chat container
      assert.ok(!htmlContent.includes('id="selectionToolbar"'), 'selectionToolbar must be removed from HTML');
      assert.ok(!htmlContent.includes('id="selectionChipsContainer"'), 'selectionChipsContainer must be removed from HTML');

      // 5. Chat history contains welcome bubble
      const chatHistoryPos = htmlContent.indexOf('id="chatHistory"');
      const welcomeBubblePos = htmlContent.indexOf('id="welcomeSystemBubble"');
      assert.ok(chatHistoryPos > 0, 'chatHistory must exist');
      assert.ok(welcomeBubblePos > chatHistoryPos, 'welcomeSystemBubble must be inside chatHistory');
    });

    test('UT-32: Spreadsheet Tools & Analysis panel complete tooling verification', async () => {
      const { docChips } = await import('../src/taskpanes/excel/chips.js');

      assert.ok(Array.isArray(docChips), 'docChips must be an array');
      assert.strictEqual(docChips.length, 4, 'docChips must contain 4 streamlined spreadsheet tools');

      const chipIds = docChips.map(c => c.id);

      // Verify all streamlined primary tools are present
      assert.ok(chipIds.includes('chipInsights'), 'chipInsights must be present');
      assert.ok(chipIds.includes('chipChartGen'), 'chipChartGen must be present');
      assert.ok(chipIds.includes('chipFormulaGen'), 'chipFormulaGen must be present');
      assert.ok(chipIds.includes('chipTaxonomy'), 'chipTaxonomy must be present');

      // Verify each chip has valid non-empty fields
      for (const chip of docChips) {
        assert.ok(chip.id && chip.id.length > 0, `Chip ${chip.id} must have a valid id`);
        assert.ok(chip.label && chip.label.length > 0, `Chip ${chip.id} must have a valid label`);
        assert.ok(chip.title && chip.title.length > 0, `Chip ${chip.id} must have a valid title`);
        assert.ok(chip.prompt || typeof chip.handler === 'function', `Chip ${chip.id} must have prompt or handler`);
      }
    });

    test('UT-33: TaskpaneShell.populateDocChips context ingestion with active selection and document text', async () => {
      const { TaskpaneShell } = await import('../src/shell/TaskpaneShell.js');

      const mockChipsContainer = {
        innerHTML: '',
        children: [],
        appendChild(child) {
          this.children.push(child);
        }
      };

      globalThis.document = {
        getElementById: (id) => {
          if (id === 'docToolsChipsContainer') return mockChipsContainer;
          if (id === 'loading') return { style: { display: 'none' } };
          return null;
        },
        querySelector: () => null,
        createElement: (tag) => ({
          tagName: tag.toUpperCase(),
          className: '',
          innerHTML: '',
          title: '',
          dataset: {},
          onclick: null,
          addEventListener: () => {}
        })
      };

      let executedPrompt = '';
      let executedBubble = '';

      const mockAdapter = {
        getSelectedText: async () => 'Product,Sales\nWidget A,1500\nWidget B,2300',
        getFullDocumentText: async () => 'Sheet1 UsedRange A1:B10 containing sales records'
      };

      const mockChatSession = {
        executeWorkflow: async (prompt, bubble) => {
          executedPrompt = prompt;
          executedBubble = bubble;
        }
      };

      const mockCtx = {
        adapter: mockAdapter,
        chatSession: mockChatSession,
        chatLog: { appendSystem: () => {} },
        hostModule: { hostName: 'Excel' },
        selectionWatcher: {
          getCurrentSelection: () => 'Product,Sales\nWidget A,1500\nWidget B,2300'
        }
      };

      const testChips = [
        {
          id: 'chipFormatRange',
          label: '🎨 Style & Format Range',
          prompt: 'Format this selected range with headers and borders.'
        }
      ];

      TaskpaneShell.populateDocChips(testChips, mockCtx);
      assert.equal(mockChipsContainer.children.length, 1);

      // Trigger click on chip button
      await mockChipsContainer.children[0].onclick();

      assert.ok(executedPrompt.includes('Selected Data / Range Context:'), 'Prompt must contain selected data context');
      assert.ok(executedPrompt.includes('Widget A,1500'), 'Prompt must contain selection text');
      assert.ok(executedPrompt.includes('Active Document / Sheet Context:'), 'Prompt must contain sheet context');
      assert.ok(executedPrompt.includes('Format this selected range with headers and borders.'), 'Prompt must contain chip instruction');
    });

    test('UT-34: Excel onSessionReset and removeTargetChartBtn dismiss targetChartRow and reset scroll position', async () => {
      const excelModule = (await import('../src/taskpanes/excel/index.js')).default;

      let targetChartBadgeDisplay = 'block';
      let targetChartRowDisplay = 'flex';
      let scrollTopValue = 350;
      let chatLogScrollTopCalled = false;

      const mockBadge = {
        style: {
          set display(val) { targetChartBadgeDisplay = val; },
          get display() { return targetChartBadgeDisplay; }
        }
      };

      const mockRow = {
        style: {
          set display(val) { targetChartRowDisplay = val; },
          get display() { return targetChartRowDisplay; }
        }
      };

      const mockChatHistory = {
        get scrollTop() { return scrollTopValue; },
        set scrollTop(val) { scrollTopValue = val; }
      };

      globalThis.document = {
        getElementById: (id) => {
          if (id === 'targetChartBadge') return mockBadge;
          if (id === 'targetChartRow') return mockRow;
          if (id === 'chatHistory') return mockChatHistory;
          if (id === 'promptText') return { placeholder: '' };
          return null;
        },
        querySelector: () => null
      };

      const mockAdapter = {
        setLockedChart: () => {},
        getSelectedText: async () => '',
        getSelectionContext: async () => ({ address: 'A1', rowCount: 1, colCount: 1, hasData: false })
      };

      const mockCtx = {
        adapter: mockAdapter,
        chatLog: {
          scrollToTop: () => { chatLogScrollTopCalled = true; }
        }
      };

      excelModule.onSessionReset(mockCtx);

      assert.equal(targetChartBadgeDisplay, 'none', 'targetChartBadge must be hidden');
      assert.equal(targetChartRowDisplay, 'none', 'targetChartRow must be hidden');
      assert.equal(scrollTopValue, 0, 'chatHistory.scrollTop must be reset to 0');
      assert.equal(chatLogScrollTopCalled, true, 'chatLog.scrollToTop must be called');
    });

    test('UT-35: ChatLog.clear() clears messages, inserts welcome bubble, and scrolls to top', async () => {
      const { ChatLog } = await import('../src/shell/ChatLog.js');

      globalThis.document = {
        createElement: (tag) => ({
          tagName: tag.toUpperCase(),
          className: '',
          innerHTML: '',
          id: ''
        }),
        querySelector: () => null
      };

      const mockContainer = {
        children: [],
        innerHTML: '<div class="chat-bubble user">Old message</div>',
        scrollTop: 500,
        appendChild: function(child) {
          this.children.push(child);
        },
        querySelector: () => null
      };

      const chatLog = new ChatLog(mockContainer);
      chatLog.clear('Welcome to Gemini Excel Assistant!');

      assert.equal(mockContainer.scrollTop, 0, 'scrollTop must be reset to 0');
      assert.equal(mockContainer.children.length, 1, 'Only welcome bubble should remain');
      assert.equal(mockContainer.children[0].id, 'welcomeSystemBubble', 'Welcome bubble ID should be set');
    });

    test('UT-36: Table Insights summary table normalizes mixed currency and narrative text into genuine numbers for Excel charting', () => {
      const userResponseTable = [
        ['Metric / Category', 'Performance Analysis', 'Trend Status'],
        ['Total Pipeline Value', '$1,672,000 total pipeline across 14 deals', '🟢 Strong Pipeline'],
        ['Total Weighted Value', '$996,700 expected close value (Avg. Win Prob: 52%)', '🟡 Healthy Outlook'],
        ['Top Sales Performer', 'Elena Rostova ($562,000 across 4 opportunities)', '🔥 Leading Volume'],
        ['Closed-Lost Leakage', '$175,000 lost across 2 major accounts', '🔴 Revenue Leakage'],
        ['Largest Deal Size', '$235,000 (Vanguard Media - OPP-2009)', '⚠️ High Concentration']
      ];

      const normalized = normalizeInsightsTableData(userResponseTable);
      assert.ok(normalized, 'Normalization result should exist');
      assert.equal(normalized.cleanedMatrix.length, 6);
      
      // Check column types
      assert.ok(normalized.numericColumns.includes(1), 'Column 1 (Performance Analysis) must be recognized as numeric');
      assert.equal(normalized.columnFormats[1], '$#,##0');

      // Check converted numeric values
      assert.equal(normalized.cleanedMatrix[1][1], 1672000, 'Total Pipeline Value must be converted to numeric 1672000');
      assert.equal(normalized.cleanedMatrix[2][1], 996700, 'Total Weighted Value must be converted to numeric 996700');
      assert.equal(normalized.cleanedMatrix[3][1], 562000, 'Top Sales Performer must be converted to numeric 562000');
      assert.equal(normalized.cleanedMatrix[4][1], 175000, 'Closed-Lost Leakage must be converted to numeric 175000');
      assert.equal(normalized.cleanedMatrix[5][1], 235000, 'Largest Deal Size must be converted to numeric 235000');

      // Check chart range info
      assert.ok(normalized.chartRangeInfo);
      assert.equal(normalized.chartRangeInfo.colCount, 2, 'Chart data range must include exactly 2 columns (Category and Value), excluding Trend Status');
      assert.equal(normalized.chartRangeInfo.rowCount, 6);
    });

    test('UT-37: extractChartDataset extracts multi-row dataset from structured response and enriches category names', () => {
      const responseText = `📋 Structured Insights Summary Table
| Metric / Category | Performance Analysis | Trend Status |
| Total Pipeline Value | $1,672,000 total pipeline across 14 deals | 🟢 Strong Pipeline |
| Total Weighted Value | $996,700 expected close value | 🟡 Healthy Outlook |
| Top Sales Performer | Elena Rostova ($562,000 across 4 opportunities) | 🔥 Leading Volume |
| Closed-Lost Leakage | $175,000 lost across 2 major accounts | 🔴 Revenue Leakage |
| Largest Deal Size | $235,000 (Vanguard Media) | ⚠️ High Concentration |`;

      const dataset = ExcelInsightsService.extractChartDataset(responseText);
      assert.ok(dataset, 'Dataset should be extracted');
      assert.equal(dataset.matrix.length, 6);
      assert.equal(dataset.matrix[0][0], 'Metric / Category');
      assert.equal(dataset.matrix[0][1], 'Performance Analysis');
      assert.equal(dataset.matrix[1][1], 1672000);
      assert.equal(dataset.matrix[2][1], 996700);
      assert.equal(dataset.matrix[3][1], 562000);
      assert.equal(dataset.matrix[4][1], 175000);
      assert.equal(dataset.matrix[5][1], 235000);
    });

    test('UT-38: Markdown parser renders visual images with only a single Zoom & Review button', () => {
      const mdWithImage = `Here is your visual:
![Table Trend Chart](data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=)`;

      const parsedHtml = parseMarkdown(mdWithImage);
      assert.ok(parsedHtml.includes('img-action-btn-zoom'), 'Must contain Zoom & Review button');
      assert.ok(parsedHtml.includes('🔍 Zoom & Review'));

      // Count occurrences of "Zoom"
      const zoomMatches = parsedHtml.match(/🔍\s*Zoom/g) || [];
      assert.equal(zoomMatches.length, 1, 'There must be exactly one Zoom button in the rendered card');
      assert.ok(!parsedHtml.includes('img-zoom-btn'), 'Must not contain duplicate img-zoom-btn');
    });

    test('UT-39: normalizeTaxonomySpec safely handles column names and taxonomy array format', () => {
      const fullCatalogSpec = {
        version: "1.0",
        taxonomy: [
          {
            columnName: "Sales Rep",
            type: "one-to-one",
            options: { allowOther: true, allowMultiple: false },
            mapping: [
              { value: "David Kim", category: "David Kim" },
              { value: "Elena Rostova", category: "Elena Rostova" },
              { value: "Marcus Vance", category: "Marcus Vance" },
              { value: "Sarah Jenkins", category: "Sarah Jenkins" }
            ]
          }
        ]
      };

      const normalized = normalizeTaxonomySpec(fullCatalogSpec);
      assert.ok(normalized, 'Normalized spec should be produced');
      assert.equal(normalized.action, 'categorize');
      assert.equal(normalized.sourceColumnName, 'Sales Rep');
      assert.equal(normalized.sourceColumn, 'B'); // Clean fallback
      assert.equal(normalized.targetColumn, 'C'); // Valid adjacent letter
      assert.equal(normalized.header, 'Sales Rep Category');
      assert.equal(normalized.categories.length, 4);
      assert.equal(normalized.categoryValues[0], 'David Kim');
      assert.equal(normalized.categoryValues[1], 'Elena Rostova');
    });

    test('UT-40: ExcelInsightsService.colIndexToLetter and letterToColIndex conversion', () => {
      assert.equal(ExcelInsightsService.colIndexToLetter(0), 'A');
      assert.equal(ExcelInsightsService.colIndexToLetter(1), 'B');
      assert.equal(ExcelInsightsService.colIndexToLetter(25), 'Z');
      assert.equal(ExcelInsightsService.colIndexToLetter(26), 'AA');

      assert.equal(ExcelInsightsService.letterToColIndex('A'), 0);
      assert.equal(ExcelInsightsService.letterToColIndex('B'), 1);
      assert.equal(ExcelInsightsService.letterToColIndex('Z'), 25);
      assert.equal(ExcelInsightsService.letterToColIndex('AA'), 26);
    });

    test('UT-41: extractExcelTaxonomySpec extracts taxonomy from JSON code block when isTaxonomy flag is true', () => {
      const responseText = `Here is the taxonomy categorization specification:
\`\`\`json
{
  "version": "1.0",
  "taxonomy": [
    {
      "columnName": "Priority",
      "mapping": [
        { "value": "High", "category": "Urgent" },
        { "value": "Low", "category": "Standard" }
      ]
    }
  ]
}
\`\`\``;

      const extracted = extractExcelTaxonomySpec(responseText, "", { isTaxonomy: true });
      assert.ok(extracted, 'Should extract taxonomy spec');
      assert.equal(extracted.sourceColumnName, 'Priority');
      assert.equal(extracted.header, 'Priority Category');
      assert.equal(extracted.categories.length, 2);
    });

    test('UT-42: applyTaxonomyCategorization correctly maps pipeline values with zeros and large numbers without substring false positives', async () => {
      const mockWrittenValues = [];
      const mockSheet = {
        getUsedRange: () => ({
          load: () => {},
          rowIndex: 0,
          rowCount: 15,
          columnIndex: 0,
          columnCount: 13,
          values: [["Deal ID", "Opportunity Name", "Account", "Sales Rep", "Priority", "Stage", "Deal Amount", "Win Prob.", "Weighted Value", "Status", "Target Close", "Lead Source", "Next Step"]]
        }),
        getRange: (addr) => {
          return {
            load: () => {},
            values: addr.startsWith("I2:I15") ? [
              ["148000"], ["57000"], ["140000"], ["15000"], ["58000"], ["0"], ["168000"],
              ["21700"], ["141000"], ["9600"], ["105000"], ["0"], ["29400"], ["104000"]
            ] : [],
            set values(v) {
              if (addr.startsWith("J2:J15")) {
                mockWrittenValues.push(...v.map(r => r[0]));
              }
            },
            insert: () => {},
            format: {
              font: {},
              fill: {},
              autofitColumns: () => {}
            }
          };
        }
      };

      const mockContext = {
        sync: async () => {}
      };

      const taxonomySpec = {
        sourceColumn: "I",
        targetColumn: "J",
        header: "Value Tier",
        startRow: 2,
        categories: [
          { sourceValue: 148000, category: "High Value" },
          { sourceValue: 57000, category: "Medium Value" },
          { sourceValue: 140000, category: "High Value" },
          { sourceValue: 15000, category: "Low Value" },
          { sourceValue: 58000, category: "Medium Value" },
          { sourceValue: 0, category: "Low Value" },
          { sourceValue: 168000, category: "High Value" },
          { sourceValue: 21700, category: "Low Value" },
          { sourceValue: 141000, category: "High Value" },
          { sourceValue: 9600, category: "Low Value" },
          { sourceValue: 105000, category: "High Value" },
          { sourceValue: 0, category: "Low Value" },
          { sourceValue: 29400, category: "Low Value" },
          { sourceValue: 104000, category: "High Value" }
        ]
      };

      const res = await applyTaxonomyCategorization(mockContext, mockSheet, taxonomySpec);
      assert.ok(res);
      assert.equal(res.targetColumn, "J");
      assert.equal(res.header, "Value Tier");
      assert.equal(mockWrittenValues.length, 14);
      assert.equal(mockWrittenValues[0], "High Value");
      assert.equal(mockWrittenValues[1], "Medium Value");
      assert.equal(mockWrittenValues[2], "High Value");
      assert.equal(mockWrittenValues[3], "Low Value");
      assert.equal(mockWrittenValues[4], "Medium Value");
      assert.equal(mockWrittenValues[5], "Low Value");
      assert.equal(mockWrittenValues[6], "High Value");
      assert.equal(mockWrittenValues[7], "Low Value");
      assert.equal(mockWrittenValues[8], "High Value");
      assert.equal(mockWrittenValues[9], "Low Value");
      assert.equal(mockWrittenValues[10], "High Value");
      assert.equal(mockWrittenValues[11], "Low Value");
      assert.equal(mockWrittenValues[12], "Low Value");
      assert.equal(mockWrittenValues[13], "High Value");
    });

    test('UT-43: Table Insights 5-column Structured Summary Table normalizes numeric value column and correctly isolates chart data columns', () => {
      const sampleInsightsResponse = `
### 📋 Structured Insights Summary Table
| Category / Metric | Value | Trend | Status | Notes |
| :--- | :--- | :--- | :--- | :--- |
| Total Pipeline Value | $1,672,000 | 📈 Positive | Healthy overall pipeline size | |
| Total Weighted Pipeline | $996,700 | 📈 Stable | Highly realistic expected revenue | |
| Average Deal Size | $119,429 | ➖ Neutral | Consistent mid-market to enterprise mix | |
| Closed Won Revenue | $198,000 | ✅ Achieved | Revenue locked in from 2 deals | |
| Closed Lost Revenue | $175,000 | ❌ Realized | Revenue lost from 2 deals | |
| Open Pipeline Value | $1,299,000 | 🔄 Active | Active opportunities requiring attention | |
| Max Active Deal Value | $235,000 | ⚠️ High Risk | Single largest exposure (OPP-2009) | |
`;
      const rawTable = ExcelInsightsService.extractInsightsSummaryTable(sampleInsightsResponse);
      assert.ok(rawTable, 'Table must be extracted');
      assert.equal(rawTable.length, 8, 'Must have 1 header row + 7 data rows');

      const normalized = ExcelInsightsService.normalizeInsightsTableData(rawTable);
      assert.ok(normalized, 'Normalization result must exist');
      
      // Numeric columns should ONLY be column 1 (Value)
      assert.deepEqual(normalized.numericColumns, [1], 'Only Column 1 should be detected as numeric');
      assert.equal(normalized.columnFormats[1], '$#,##0');

      // Narrative text columns (Status, Notes) must NOT be converted to numbers
      assert.equal(normalized.cleanedMatrix[4][3], 'Revenue locked in from 2 deals');
      assert.equal(normalized.cleanedMatrix[7][3], 'Single largest exposure (OPP-2009)');

      // Chart Range Info must only span 2 columns (Category and Value)
      assert.ok(normalized.chartRangeInfo);
      assert.equal(normalized.chartRangeInfo.colCount, 2, 'Chart data range must include exactly 2 columns');
      assert.equal(normalized.chartRangeInfo.rowCount, 8);
    });
  });

  describe('10. ExcelAdapter Execution Engine, Priority Routing, Chart Builders & In-Cell @gemini Commands', () => {

    test('UT-44: ExcelAdapter.insertContent 3-tier formula fallback (formulas -> formulasLocal -> values) and rectangular matrix padding', async () => {
      const adapter = new ExcelAdapter();
      let formulaAttempted = false;
      let formulasLocalAttempted = false;
      let valuesWritten = null;
      let gridRangeWritten = null;
      let gridDims = null;

      global.Excel = {
        run: async (cb) => {
          const mockRange = {
            rowIndex: 2,
            columnIndex: 1,
            load: () => {},
            set formulas(_val) {
              formulaAttempted = true;
              throw new Error('RichApi.Error: InvalidFormula');
            },
            set formulasLocal(_val) {
              formulasLocalAttempted = true;
              throw new Error('RichApi.Error: InvalidFormulaLocal');
            },
            set values(val) {
              valuesWritten = val;
            }
          };

          const mockSheet = {
            getRangeByIndexes: (r, c, rows, cols) => {
              gridDims = { r, c, rows, cols };
              return {
                set values(val) {
                  gridRangeWritten = val;
                },
                format: { autofitColumns: () => {} }
              };
            }
          };

          const context = {
            workbook: {
              getSelectedRange: () => mockRange,
              worksheets: { getActiveWorksheet: () => mockSheet }
            },
            sync: async () => {}
          };
          return await cb(context);
        }
      };

      // Part A: 3-tier formula fallback
      await adapter.insertContent('<pre><code>=SUM(B2:B10)</code></pre>', '```excel\n=SUM(B2:B10)\n```');
      assert.equal(formulaAttempted, true, 'Tier 1 range.formulas must be attempted first');
      assert.equal(formulasLocalAttempted, true, 'Tier 2 range.formulasLocal must be attempted second');
      assert.deepEqual(valuesWritten, [['=SUM(B2:B10)']], 'Tier 3 range.values must succeed as fallback');

      // Part B: Rectangular matrix padding for uneven multi-row content (1-col heading + 3-col TSV row + 2-col TSV row)
      const unevenText = `Q3 Regional Summary\nRegion\tQ1\tQ2\nNA\t100\t200\nEMEA\t150`;
      await adapter.insertContent('', unevenText);

      assert.ok(gridDims, 'getRangeByIndexes must be called for multi-row structured output');
      assert.equal(gridDims.rows, 4, 'Matrix row count must equal 4');
      assert.equal(gridDims.cols, 3, 'Matrix width must equal maxColCount (3)');
      assert.ok(Array.isArray(gridRangeWritten));
      for (const row of gridRangeWritten) {
        assert.equal(row.length, 3, 'Every row in rectangular matrix must be padded to length 3');
      }
      assert.deepEqual(gridRangeWritten[0], ['Q3 Regional Summary', '', '']);
      assert.deepEqual(gridRangeWritten[3], ['EMEA', '150', '']);
    });

    test('UT-45: ExcelAdapter.insertContent 6-tier priority routing (chart_edit -> chart -> range_format -> taxonomy)', async () => {
      const adapter = new ExcelAdapter();
      const actionsTriggered = [];

      global.Excel = {
        InsertShiftDirection: { right: 'Right' },
        run: async (cb) => {
          const mockChart = {
            id: 'chart-1',
            name: 'RevenueChart',
            chartType: 'ColumnClustered',
            title: { text: 'Old Title', visible: true, load: () => {} },
            legend: { visible: true, position: 'Right' },
            series: { getItemAt: () => ({ hasDataLabels: false, dataLabels: {} }) },
            axes: { categoryAxis: { setCategoryNames: () => {} } },
            setData: () => {},
            setPosition: () => {},
            load: () => {}
          };

          const makeMockRange = (addr = 'A1:B5') => ({
            address: addr,
            rowIndex: 0,
            columnIndex: 0,
            rowCount: 5,
            columnCount: 2,
            values: [['Category', 'Sales'], ['Cloud', 100], ['AI', 200]],
            numberFormat: [['General', '$#,##0']],
            load: () => {},
            insert: () => { actionsTriggered.push('taxonomy_insert'); },
            clear: () => {},
            format: {
              fill: { color: '' },
              font: { bold: false, color: '' },
              borders: { getItem: () => ({ style: '', color: '', weight: '' }) },
              autofitColumns: () => {},
              autofitRows: () => {}
            }
          });

          const mockSheet = {
            getUsedRange: () => makeMockRange('A1:B3'),
            getRange: (addr) => {
              const r = makeMockRange(addr);
              return new Proxy(r, {
                set(target, prop, val) {
                  if (prop === 'values' && addr === 'C2:C3') {
                    actionsTriggered.push('taxonomy_values');
                  }
                  target[prop] = val;
                  return true;
                }
              });
            },
            charts: {
              getItem: (name) => {
                actionsTriggered.push(`chart_edit:${name}`);
                return mockChart;
              },
              add: (type) => {
                actionsTriggered.push(`chart_add:${type}`);
                return mockChart;
              },
              load: () => {},
              items: [mockChart]
            }
          };

          const selectedRange = makeMockRange('A1:B5');
          selectedRange.format.fill = new Proxy({}, {
            set(_t, prop, val) {
              if (prop === 'color') actionsTriggered.push(`range_format_fill:${val}`);
              return true;
            }
          });

          const context = {
            workbook: {
              getSelectedRange: () => selectedRange,
              getActiveChartOrNullObject: () => ({ isNullObject: true, load: () => {} }),
              worksheets: { getActiveWorksheet: () => mockSheet }
            },
            sync: async () => {}
          };
          return await cb(context);
        }
      };

      // Priority 1: Chart Edit
      await adapter.insertContent('', '```excel-chart-edit\n{"targetChart":"RevenueChart","chartType":"Line","title":"New Trend"}\n```');
      assert.ok(actionsTriggered.includes('chart_edit:RevenueChart'));

      // Priority 2: Chart Creation
      await adapter.insertContent('', '```excel-chart\n{"type":"ColumnClustered","dataRange":"A1:B3","title":"Sales Chart"}\n```');
      assert.ok(actionsTriggered.includes('chart_add:ColumnClustered'));

      // Priority 3: Range Formatting
      await adapter.insertContent('', '```excel-range-format\n{"fill":{"color":"#E0F2FE"},"font":{"bold":true}}\n```');
      assert.ok(actionsTriggered.includes('range_format_fill:#e0f2fe'));

      // Priority 4: Taxonomy Categorization
      await adapter.insertContent('', '```excel-taxonomy\n{"sourceColumn":"B","targetColumn":"C","header":"Tier","startRow":2,"categories":[{"row":2,"category":"Enterprise"},{"row":3,"category":"SMB"}]}\n```');
      assert.ok(actionsTriggered.includes('taxonomy_values'));
    });

    test('UT-46: ExcelAdapter.checkInDocumentCommands detects @gemini in cell/formula and forceRun plain instruction', async () => {
      const adapter = new ExcelAdapter();
      const recordedStates = [];
      let capturedPrompt = '';

      global.Excel = {
        run: async (cb) => {
          const activeCell = {
            address: 'Sheet1!D2',
            values: [['@gemini sum column B']],
            formulas: [['@gemini sum column B']],
            rowCount: 1,
            columnCount: 1,
            cellCount: 1,
            rowIndex: 1,
            columnIndex: 3,
            load: () => {},
            set values(v) {
              recordedStates.push({ type: 'values', val: v });
            },
            get values() {
              return [['@gemini sum column B']];
            },
            set formulas(f) {
              recordedStates.push({ type: 'formulas', val: f });
            },
            get formulas() {
              return [['@gemini sum column B']];
            }
          };

          const mockSheet = {
            name: 'Sheet1',
            load: () => {},
            getRange: () => activeCell,
            getUsedRange: () => ({
              address: 'Sheet1!A1:C3',
              rowCount: 3,
              columnCount: 3,
              rowIndex: 0,
              columnIndex: 0,
              values: [
                ['Item', 'Revenue', 'Cost'],
                ['Widget A', 500, 200],
                ['Widget B', 800, 300]
              ],
              load: () => {}
            }),
            tables: { items: [], load: () => {} },
            charts: { items: [], load: () => {} }
          };

          const context = {
            workbook: {
              getSelectedRange: () => activeCell,
              getActiveChartOrNullObject: () => ({ isNullObject: true, load: () => {} }),
              worksheets: { getActiveWorksheet: () => mockSheet }
            },
            sync: async () => {}
          };
          return await cb(context);
        }
      };

      await adapter.checkInDocumentCommands(false, {
        onStatus: () => {},
        executePrompt: async (fullPrompt) => {
          capturedPrompt = fullPrompt;
          return '```excel\n=SUM(B2:B3)\n```';
        }
      });

      assert.ok(
        recordedStates.some(s => s.type === 'values' && s.val[0][0].includes('Gemini generating')),
        'Must write temporary loading indicator to active cell before awaiting AI'
      );
      assert.match(capturedPrompt, /Worksheet: "Sheet1"/, 'Prompt must include worksheet schema context');
      assert.match(capturedPrompt, /sum column B/, 'Prompt must include extracted user instruction');
      assert.ok(
        recordedStates.some(s => s.type === 'formulas' && s.val[0][0] === '=SUM(B2:B3)'),
        'Must write evaluated formula back into active cell'
      );
    });

    test('UT-47: Sticky chart locking (setLockedChart/getLockedChart/getSelectionTargetInfo) and 4-tier editNativeChart resolution', async () => {
      const adapter = new ExcelAdapter();
      const stickyMeta = { id: 'c-99', name: 'StickyChart', chartType: 'Pie', title: 'Revenue Share', width: 450, height: 300 };
      adapter.setLockedChart(stickyMeta);
      assert.deepEqual(adapter.getLockedChart(), stickyMeta);

      // Verify getSelectionTargetInfo retains sticky chart when no chart is actively selected
      const mockContextNoActiveChart = {
        workbook: {
          worksheets: { getActiveWorksheet: () => ({ name: 'Sheet1', load: () => {} }) },
          getActiveChartOrNullObject: () => ({ isNullObject: true, load: () => {} }),
          getSelectedRange: () => ({
            address: 'Sheet1!A1:B10',
            rowCount: 10,
            columnCount: 2,
            cellCount: 20,
            load: () => {},
            getTables: () => ({ items: [], load: () => {} })
          })
        },
        sync: async () => {}
      };

      const targetInfo = await adapter.getSelectionTargetInfo(mockContextNoActiveChart);
      assert.equal(targetInfo.type, 'range');
      assert.equal(targetInfo.cleanAddress, 'A1:B10');
      assert.deepEqual(targetInfo.chart, stickyMeta, 'Must fall back to sticky locked chart when no chart is actively selected');

      // Verify 4-tier chart resolution in editNativeChart (falls back to Tier 4 last chart on sheet when ID, Name, and Active are absent)
      const fallbackChart = {
        id: 'last-chart-id',
        name: 'FallbackChart',
        chartType: 'ColumnClustered',
        title: { text: 'Initial', visible: true },
        legend: { visible: true },
        load: () => {}
      };

      const mockSheetFallback = {
        charts: {
          getItem: () => { throw new Error('NotFound'); },
          items: [fallbackChart],
          load: () => {}
        }
      };
      const mockContextFallback = {
        workbook: {
          getActiveChartOrNullObject: () => ({ isNullObject: true, load: () => {} })
        },
        sync: async () => {}
      };

      const edited = await editNativeChart(mockContextFallback, mockSheetFallback, {
        targetChart: 'active',
        chartType: 'Doughnut',
        title: 'Morphed Fallback Chart'
      });

      assert.equal(edited.chartType, 'Doughnut');
      assert.equal(edited.title.text, 'Morphed Fallback Chart');
    });

    test('UT-48: createAggregatedSummaryTableAndChart (=SUMIF summary table) & non-adjacent Pie helper block (=A1, =D1)', async () => {
      const writtenCells = {};
      let addedChartArgs = null;

      const mockSheet = {
        getUsedRange: () => ({
          address: 'A1:D5',
          rowIndex: 0,
          columnIndex: 0,
          rowCount: 5,
          columnCount: 4,
          values: [
            ['Dept', 'Rep', 'Region', 'Sales'],
            ['Cloud', 'Alice', 'NA', 100],
            ['Security', 'Bob', 'EMEA', 200],
            ['Cloud', 'Carol', 'APAC', 300],
            ['Security', 'Dan', 'NA', 150]
          ],
          load: () => {}
        }),
        getRange: (addr) => {
          let currentValues = addr.startsWith('A')
            ? [['Dept'], ['Cloud'], ['Security'], ['Cloud'], ['Security']]
            : [['Sales'], [100], [200], [300], [150]];
          let currentFormulas = null;
          const cellObj = {
            address: addr,
            rowIndex: 0,
            columnIndex: 0,
            rowCount: 5,
            columnCount: 1,
            numberFormat: [['$#,##0']],
            load: () => {},
            format: {
              font: {},
              fill: {},
              borders: { getItem: () => ({}) },
              autofitColumns: () => {}
            },
            get values() {
              return currentValues;
            },
            set values(val) {
              currentValues = val;
              writtenCells[addr] = { ...(writtenCells[addr] || {}), values: val };
            },
            get formulas() {
              return currentFormulas;
            },
            set formulas(val) {
              currentFormulas = val;
              writtenCells[addr] = { ...(writtenCells[addr] || {}), formulas: val };
            }
          };
          return cellObj;
        },
        charts: {
          add: (type, range, seriesBy) => {
            addedChartArgs = { type, rangeAddress: range.address, seriesBy };
            return {
              title: {},
              legend: {},
              series: { getItemAt: () => ({ dataLabels: {} }) },
              axes: { categoryAxis: { setCategoryNames: () => {} } },
              setPosition: () => {},
              load: () => {}
            };
          }
        }
      };

      const mockContext = {
        workbook: {
          getSelectedRange: () => ({ rowIndex: 0, columnIndex: 6, rowCount: 1, columnCount: 1, load: () => {} })
        },
        sync: async () => {}
      };

      // Part A: Aggregated Summary Table (=SUMIF) when categories repeat
      const catResolved = resolveRangeAddress('A1:A5', { startRow: 1, endRow: 5 });
      const seriesResolved = [resolveRangeAddress('D1:D5', { startRow: 1, endRow: 5 })];
      const usedBounds = { startColIndex: 0, endColIndex: 3, startRow: 1, endRow: 5 };

      const aggChart = await createAggregatedSummaryTableAndChart(
        mockContext,
        mockSheet,
        { type: 'ColumnClustered', aggregate: 'sum', hasHeaders: true, title: 'Total Sales by Dept' },
        catResolved,
        seriesResolved,
        usedBounds,
        'ColumnClustered',
        false,
        'Columns'
      );

      assert.ok(aggChart, 'Aggregated chart must be created');
      assert.equal(addedChartArgs.type, 'ColumnClustered');
      assert.deepEqual(writtenCells['G1:H1'].values, [['Dept', 'Total Sales']]);
      assert.deepEqual(writtenCells['G2'].values, [['Cloud']]);
      assert.deepEqual(writtenCells['H2'].formulas, [['=SUMIF($A$2:$A$5, G2, $D$2:$D$5)']]);
      assert.deepEqual(writtenCells['G3'].values, [['Security']]);
      assert.deepEqual(writtenCells['H3'].formulas, [['=SUMIF($A$2:$A$5, G3, $D$2:$D$5)']]);

      // Part B: Non-adjacent Pie chart with unique categories creates dynamic formula-linked helper block (=A1, =D1)
      const uniqueMockSheet = {
        ...mockSheet,
        getRange: (addr) => {
          const base = mockSheet.getRange(addr);
          if (addr.startsWith('A')) {
            base.values = [['Dept'], ['Cloud'], ['Security'], ['AI'], ['Workspace']];
          }
          return base;
        }
      };

      await insertNativeChart(mockContext, uniqueMockSheet, {
        type: 'Pie',
        categoriesRange: 'A1:A5',
        seriesRange: 'D1:D5',
        title: 'Revenue Breakdown'
      });

      assert.equal(addedChartArgs.type, 'Pie');
      assert.equal(addedChartArgs.rangeAddress, 'G1:H5', 'Non-adjacent Pie chart must bind to 2-column helper range G1:H5');
      assert.deepEqual(writtenCells['G1:H5'].formulas[0], ['=A1', '=D1']);
      assert.deepEqual(writtenCells['G1:H5'].formulas[1], ['=A2', '=D2']);
    });

  });
});



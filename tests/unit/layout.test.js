import { test } from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { DEFAULT_CONFIG, resolveOptions } from '../../src/config.js';
import { createLayoutMetrics } from '../../src/layout-metrics.js';
import { createWorkbookWriter } from '../../src/workbook.js';
import { mapWithConcurrency } from '../../src/images.js';

const dependencies = () => {
    const base = { CONFIG: DEFAULT_CONFIG, ExcelJS, debugLog() {}, errorLog() {}, warnLog() {}, mapWithConcurrency };
    return { ...base, ...createLayoutMetrics(base) };
};
test('edge normalization retains ascending finite boundaries', () => {
    const metrics = dependencies();
    assert.deepEqual(metrics.normalizeEdges([100, 0, 1, 50, NaN, 100]), [0, 50, 100]);
    assert.equal(metrics.pxToPoint(96), 72);
});
test('edge normalization never leaves a row thinner than the edge tolerance', () => {
    const metrics = dependencies();
    const tolerance = DEFAULT_CONFIG.EDGE_TOLERANCE_PX;
    // 10 and 11.9 are one cluster; 12 starts the next. Moving the first cluster to its last edge
    // used to leave a 0.1px sliver between 11.9 and 12.
    assert.deepEqual(metrics.normalizeEdges([0, 10, 11.9, 12, 40]), [0, 10, 12, 40]);
    assert.deepEqual(metrics.normalizeEdges([0, 10, 11.5, 12.5, 40]), [0, 10, 12.5, 40]);
    // Edge 0 is the container top and the last edge is the container extent; neither may drift.
    assert.deepEqual(metrics.normalizeEdges([0, 1.5, 30, 31.5]), [0, 31.5]);
    const edges = metrics.normalizeEdges([0, 3, 4.9, 5, 5.1, 7, 8.9, 9, 20, 21.9, 22]);
    for (let index = 1; index < edges.length; index++) {
        assert.ok(edges[index] - edges[index - 1] >= tolerance, 'sliver at ' + edges[index - 1] + '..' + edges[index]);
    }
});
test('page break rows resolve without relying on an implicit scrollY global', async () => {
    const { createLayout } = await import('../../src/layout.js');
    const metrics = dependencies();
    const marker = { getBoundingClientRect: () => ({ top: 100, bottom: 100 }) };
    const layout = createLayout({
        ...metrics,
        collectDomSnapshot: () => ({
            rootLeft: 0, rootTop: 50, rootWidth: 100, rootHeight: 300,
            allCells: [], nestedCells: [], leafCells: [], tableElements: [], textFragments: [],
            pageBreakMarkers: [marker], columnEdges: [0, 100], rowEdges: [0, 100, 200, 300]
        })
    });
    const previous = globalThis.window;
    globalThis.window = { scrollY: 50, scrollX: 0 };
    try {
        // Marker top is 100 + 50 scroll - 50 root top = 100, so the new page starts at row 2.
        assert.deepEqual(layout.buildLayoutModel({ querySelectorAll: () => [] }).pageBreakRows, [2]);
    } finally {
        if (previous === undefined) delete globalThis.window; else globalThis.window = previous;
    }
});

test('OOXML column widths retain physical pixels across narrow and wide columns', () => {
    const base = { CONFIG: { ...DEFAULT_CONFIG, MIN_COL_WIDTH_PX: 1 }, debugLog() {}, errorLog() {} };
    const metrics = createLayoutMetrics(base);
    // Microsoft Open XML example: stored width 8.7109375 is 61 pixels.
    assert.equal(metrics.columnWidthToPixels(8.7109375), 61);
    for (const pixels of [1, 3, 5, 10, 12, 61, 100, 750]) {
        assert.equal(metrics.columnWidthToPixels(metrics.pxToColumnWidth(pixels)), pixels);
    }
    const writer = createWorkbookWriter({ ...base, ...metrics, ExcelJS });
    const sheet = new ExcelJS.Workbook().addWorksheet('Grid');
    writer.applyColumnWidths(sheet, [10.2, 10.2, 10.2, 10.2, 10.2]);
    assert.equal(sheet.columns.reduce((sum, column) => sum + metrics.columnWidthToPixels(column.width), 0), 51);
    writer.applyRowHeights(sheet, [10.2, 10.2, 10.2, 10.2, 10.2]);
    assert.equal([1, 2, 3, 4, 5].reduce((sum, row) => sum + sheet.getRow(row).height, 0), 38.25);
});
test('numeric subtotal cells survive real ExcelJS serialization', async () => {
    const writer = createWorkbookWriter(dependencies());
    const workbook = writer.createWorkbook();
    const sheet = workbook.addWorksheet('Test');
    const cells = [124800, 112000, 329200, 566000].map((value, index) => ({
        value, rowStart: index, rowEnd: index + 1, colStart: 0, colEnd: 2,
        styles: { numFmt: '"HKD" #,##0.00' }
    }));
    writer.writeCellsToSheet(sheet, { cells });
    const bytes = await writer.writeWorkbookToBuffer(workbook);
    const read = new ExcelJS.Workbook();
    await read.xlsx.load(bytes);
    assert.deepEqual([1, 2, 3, 4].map(row => read.getWorksheet('Test').getCell(row, 1).value), [124800, 112000, 329200, 566000]);
    assert.match(read.getWorksheet('Test').getCell('A1').numFmt, /HKD/);
});
test('overlapping non-empty cells fail explicitly', () => {
    const writer = createWorkbookWriter(dependencies());
    const sheet = new ExcelJS.Workbook().addWorksheet('Test');
    const cell = value => ({ value, rowStart: 0, rowEnd: 1, colStart: 0, colEnd: 1, styles: {} });
    assert.throws(() => writer.writeCellsToSheet(sheet, { cells: [cell('first'), cell('second')] }), /Overlapping/);
});
test('export configuration is isolated and rejects invalid numeric tuning', () => {
    const root = { getAttribute: () => null, matches: () => false };
    const first = resolveOptions(root, { currency: 'HKD', pageSetup: { margins: { left: 1 } } });
    const second = resolveOptions(root, {});
    assert.equal(first.config.PAGE_SETUP.margins.left, 1);
    assert.equal(second.config.PAGE_SETUP.margins.left, 0.5);
    assert.throws(() => resolveOptions(root, { edgeTolerancePx: 0 }));
    assert.throws(() => resolveOptions(root, { imageConcurrency: 1.5 }));
    assert.throws(() => resolveOptions(root, { respectPageBreaks: 'true' }));
    assert.throws(() => resolveOptions(root, { minColumnWidthPx: 0 }));
    assert.throws(() => resolveOptions(root, { preserveRichText: 'true' }));
    assert.throws(() => resolveOptions(root, { includeBackgrounds: 0 }));
    assert.throws(() => resolveOptions(root, { currencyDisplay: 'symbol' }));
    assert.throws(() => resolveOptions(root, { normalFont: { name: 'Arial', size: 0 } }));
    const normalFont = { name: 'Arial', size: 9 };
    const tuned = resolveOptions(root, { normalFont });
    normalFont.size = 10;
    assert.equal(tuned.config.NORMAL_FONT.size, 9);
    assert.equal(second.config.NORMAL_FONT, null);
    assert.throws(() => resolveOptions(root, { unknown: true }));
    const exclude = ['.hidden'];
    resolveOptions(root, { exclude });
    exclude.push('.another');
    assert.equal(exclude.length, 2);
});
test('image task concurrency is bounded and result order is stable', async () => {
    let active = 0, peak = 0;
    const result = await mapWithConcurrency([0, 1, 2, 3, 4], 2, async value => {
        active++; peak = Math.max(peak, active);
        await new Promise(resolve => setTimeout(resolve, 2));
        active--; return value * 2;
    });
    assert.equal(peak, 2);
    assert.deepEqual(result, [0, 2, 4, 6, 8]);
});

test('formatted numeric cells fit their original merged geometry without changing values', async () => {
    const writer = createWorkbookWriter(dependencies());
    const book = writer.createWorkbook();
    const styles = { numFmt: '"HKD" #,##0.00', alignment: { horizontal: 'right', wrapText: true } };
    const cells = [1800, -3000, 'Included', '001800'].map((value, rowStart) => ({
        value, rowStart, rowEnd: rowStart + 1, colStart: 0, colEnd: 2, styles
    }));
    const sheet = book.addWorksheet('Test');
    sheet.columns = [{ width: 3 }, { width: 4 }];
    writer.writeCellsToSheet(sheet, { cells });
    const read = new ExcelJS.Workbook();
    await read.xlsx.load(await writer.writeWorkbookToBuffer(book));
    const result = read.getWorksheet('Test');
    for (const address of ['A1', 'A2']) {
        assert.equal(result.getCell(address).alignment.shrinkToFit, true);
        assert.notEqual(result.getCell(address).alignment.wrapText, true);
    }
    assert.deepEqual([1, 2, 3, 4].map(row => result.getCell(row, 1).value), [1800, -3000, 'Included', '001800']);
    assert.equal(result.getCell('A3').alignment.shrinkToFit, undefined);
    assert.equal(result.getCell('A4').alignment.wrapText, true);
    assert.equal(styles.alignment.wrapText, true);
    assert.deepEqual(result.model.merges, ['A1:B1', 'A2:B2', 'A3:B3', 'A4:B4']);
    assert.equal(result.getColumn(1).width, 3);
});

test('marked pages use percentage scaling that preserves breaks and fits each A4 page', async () => {
    const deps = dependencies();
    const writer = createWorkbookWriter(deps);
    const model = { columnWidthsPx: [500, 300], rowHeightsPx: [600, 600, 300, 300, 200],
        cells: [], images: [], pageBreakRows: [3, 3, 5] };
    const book = writer.createWorkbook();
    const sheet = writer.prepareWorksheet(book, model);
    assert.deepEqual(sheet.rowBreaks.map(b => b.id), [2, 4]);
    assert.equal(sheet.pageSetup.fitToPage, false);
    const widthPt = sheet.columns.reduce((sum, c) => sum + deps.columnWidthToPixels(c.width) * 0.75, 0);
    const pageWidthPt = (210 / 25.4 - 1) * 72;
    const pageHeightPt = (297 / 25.4 - 0.2) * 72;
    assert.ok(widthPt * sheet.pageSetup.scale / 100 <= pageWidthPt);
    assert.ok(900 * sheet.pageSetup.scale / 100 <= pageHeightPt);
    assert.ok(Number.isInteger(sheet.pageSetup.scale) && sheet.pageSetup.scale >= 10);
    const read = new ExcelJS.Workbook();
    await read.xlsx.load(await writer.writeWorkbookToBuffer(book));
    assert.equal(read.worksheets[0].pageSetup.fitToPage, false);
    assert.equal(read.worksheets[0].pageSetup.scale, sheet.pageSetup.scale);
    const unmarked = writer.prepareWorksheet(writer.createWorkbook(), { ...model, pageBreakRows: [] });
    assert.equal(unmarked.pageSetup.fitToPage, true);
});

test('explicit scaling and the Fit To opt-out retain their print setup', () => {
    const root = { getAttribute: () => null, matches: () => false };
    const model = { columnWidthsPx: [800], rowHeightsPx: [300, 300],
        cells: [], images: [], pageBreakRows: [2] };
    for (const options of [{ respectPageBreaks: false }, { pageSetup: { fitToPage: false, scale: 65 } }]) {
        const config = resolveOptions(root, options).config;
        const base = { CONFIG: config, ExcelJS, debugLog() {}, errorLog() {}, warnLog() {}, mapWithConcurrency };
        const writer = createWorkbookWriter({ ...base, ...createLayoutMetrics(base) });
        const sheet = writer.prepareWorksheet(writer.createWorkbook(), model);
        assert.equal(sheet.pageSetup.fitToPage, options.respectPageBreaks === false);
        if (options.pageSetup) assert.equal(sheet.pageSetup.scale, 65);
    }
});

test('Normal style serialization is isolated and retains explicit CSS cell fonts', async () => {
    const base = { CONFIG: { ...DEFAULT_CONFIG, NORMAL_FONT: { name: 'Arial', size: 9 } },
        ExcelJS, debugLog() {}, warnLog() {}, errorLog() {} };
    const writer = createWorkbookWriter({ ...base, ...createLayoutMetrics(base) });
    const book = writer.createWorkbook();
    const sheet = book.addWorksheet('Fonts');
    sheet.getCell('A1').value = 'Explicit Calibri';
    sheet.getCell('A1').font = { name: 'Calibri', size: 11, color: { argb: 'FF000000' } };
    const originalAddStyles = book.xlsx.addStyles;
    for (let i = 0; i < 2; i++) {
        const bytes = await writer.writeWorkbookToBuffer(book);
        assert.equal(book.xlsx.addStyles, originalAddStyles);
        const zip = await JSZip.loadAsync(bytes);
        const styles = await zip.file('xl/styles.xml').async('string');
        const normalFont = styles.match(/<fonts\b[^>]*>(<font>.*?<\/font>)/)[1];
        assert.match(normalFont, /<name val="Arial"/);
        assert.match(normalFont, /<sz val="9"/);
        assert.match(styles, /<cellStyle\b[^>]*name="Normal"[^>]*customBuiltin="1"/);
        const read = new ExcelJS.Workbook();
        await read.xlsx.load(bytes);
        assert.equal(read.worksheets[0].getCell('A1').font.name, 'Calibri');
        assert.equal(read.worksheets[0].getCell('A1').font.size, 11);
    }
    const writeBuffer = book.xlsx.writeBuffer;
    book.xlsx.writeBuffer = async () => { throw new Error('Synthetic write failure'); };
    await assert.rejects(writer.writeWorkbookToBuffer(book), /Synthetic write failure/);
    assert.equal(book.xlsx.addStyles, originalAddStyles);
    book.xlsx.writeBuffer = writeBuffer;
});

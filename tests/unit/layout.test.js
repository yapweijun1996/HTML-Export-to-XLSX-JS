import { test } from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
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
    assert.deepEqual(metrics.normalizeEdges([100, 0, 1, 50, NaN, 100]), [1, 50, 100]);
    assert.equal(metrics.pxToPoint(96), 72);
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

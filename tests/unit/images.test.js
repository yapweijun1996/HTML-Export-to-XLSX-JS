import { test } from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import { DEFAULT_CONFIG } from '../../src/config.js';
import { createImages } from '../../src/images.js';
import { createLayoutMetrics } from '../../src/layout-metrics.js';

function placement(alignment, count = 1) {
    const ranges = [];
    const dependencies = {
        CONFIG: { ...DEFAULT_CONFIG, IMAGE_WIDTH_SCALE: 1 }, IMAGE_SCALE: 1,
        debugLog() {}, errorLog() {}, warnLog() {}, isFiniteNumber: Number.isFinite,
        snapToPrecision: value => value, toPositiveNumber: value => value > 0 ? value : null
    };
    dependencies.columnWidthToPixels = createLayoutMetrics(dependencies).columnWidthToPixels;
    const images = createImages(dependencies);
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Images');
    worksheet.columns = [{ width: 7 }, { width: 14 }, { width: 14 }, { width: 14 }];
    for (let row = 1; row <= 3; row++) worksheet.getRow(row).height = 37.5;
    const addImage = worksheet.addImage.bind(worksheet);
    worksheet.addImage = (id, range) => { ranges.push(range); addImage(id, range); };
    const cell = { colStart: 1, colEnd: 4, rowStart: 1, rowEnd: 3,
        sheetPosition: { colStart: 2, rowStart: 2 } };
    const layout = { columnEdges: [0, 50, 150, 250, 350], rowEdges: [0, 20, 70, 120] };
    const jobs = Array.from({ length: count }, () => ({ width: 60, height: 20, base64: 'synthetic', extension: 'png' }));
    images.placeImagesWithAlignment(worksheet, workbook, jobs, cell, layout, alignment);
    return ranges;
}

test('aligned image anchors respect merged-cell geometry without fixed pixel offsets', () => {
    for (const [alignment, col, row] of [['left', 1, 1], ['center', 2.2, 1.8], ['right', 3.4, 1]]) {
        const [range] = placement(alignment);
        assert.ok(Math.abs(range.tl.col - col) < 1e-9, alignment + ' column');
        assert.ok(Math.abs(range.tl.row - row) < 1e-9, alignment + ' row');
        assert.equal(range.editAs, 'oneCell');
        assert.ok(range.br.nativeCol >= Math.floor(col));
        assert.ok(range.br.nativeRow >= Math.floor(row));
    }
});

test('multiple aligned images keep their spacing rather than collapsing onto the right edge', () => {
    const inline = placement('inline', 2);
    assert.equal(inline[0].tl.col, 1);
    assert.ok(Math.abs(inline[1].tl.col - 1.64) < 1e-9);
    const stacked = placement('block', 2);
    assert.equal(stacked[0].tl.row, 1);
    assert.ok(Math.abs(stacked[1].tl.row - 1.48) < 1e-9);
});

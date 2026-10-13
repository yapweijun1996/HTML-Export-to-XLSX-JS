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
    worksheet.columns = [50, 100, 100, 100].map(pixels => ({ width: pixels / 7 }));
    worksheet.getRow(1).height = 15;
    for (let row = 2; row <= 3; row++) worksheet.getRow(row).height = 37.5;
    const addImage = worksheet.addImage.bind(worksheet);
    worksheet.addImage = (id, range) => { ranges.push(range); addImage(id, range); };
    const cell = { colStart: 1, colEnd: 4, rowStart: 1, rowEnd: 3,
        sheetPosition: { colStart: 2, rowStart: 2 } };
    const layout = { columnEdges: [0, 50, 150, 250, 350], rowEdges: [0, 20, 70, 120] };
    const jobs = Array.from({ length: count }, () => ({ width: 60, height: 20, base64: 'synthetic', extension: 'png' }));
    images.placeImagesWithAlignment(worksheet, workbook, jobs, cell, layout, alignment);
    return ranges;
}

function position(marker) {
    return { x: [0, 50, 150, 250, 350][marker.nativeCol] + marker.nativeColOff / 9525,
        y: [0, 20, 70, 120][marker.nativeRow] + marker.nativeRowOff / 9525 };
}

test('aligned image anchors respect merged-cell geometry without fixed pixel offsets', () => {
    for (const [alignment, x, y] of [['left', 50, 20], ['center', 170, 60], ['right', 290, 20]]) {
        const [range] = placement(alignment);
        assert.deepEqual(position(range.tl), { x, y }, alignment + ' position');
        assert.equal(range.editAs, 'oneCell');
        const end = position(range.br);
        assert.equal(end.x - x, 60);
        assert.equal(end.y - y, 20);
    }
});

test('multiple aligned images keep their spacing rather than collapsing onto the right edge', () => {
    const inline = placement('inline', 2);
    assert.equal(position(inline[0].tl).x, 50);
    assert.equal(position(inline[1].tl).x - position(inline[0].tl).x, 64);
    const stacked = placement('block', 2);
    assert.equal(position(stacked[0].tl).y, 20);
    assert.equal(position(stacked[1].tl).y - position(stacked[0].tl).y, 24);
});

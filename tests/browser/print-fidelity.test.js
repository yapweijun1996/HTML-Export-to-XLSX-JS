import { test, expect } from '@playwright/test';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import fs from 'node:fs/promises';

const pixels = width => Math.floor((256 * width + 18) / 256 * 7);
function markerPixels(anchor, sheet) {
    let x = anchor.nativeColOff / 9525, y = anchor.nativeRowOff / 9525;
    for (let c = 1; c <= anchor.nativeCol; c++) x += pixels(sheet.getColumn(c).width);
    for (let r = 1; r <= anchor.nativeRow; r++) y += sheet.getRow(r).height / 0.75;
    return { x, y };
}

for (const mode of ['', 'min']) {
    test('Globe3 print profile retains physical layout, mixed fonts and document currency ' + (mode || 'source'), async ({ page }) => {
        await page.goto('/tests/fixtures/print-fidelity.html?mode=' + mode);
        await page.waitForFunction(() => Boolean(window.exportPrintFormToXlsx));
        const download = page.waitForEvent('download');
        const result = await page.evaluate(async () => {
            const root = document.querySelector('#export');
            const before = root.outerHTML;
            const origin = root.getBoundingClientRect();
            const images = [...root.querySelectorAll('img')].map(img => {
                const r = img.getBoundingClientRect();
                return { x: r.left - origin.left, y: r.top - origin.top, width: r.width, height: r.height };
            });
            const result = await window.exportPrintFormToXlsx();
            return { bytes: [...new Uint8Array(result.buffer)], warnings: result.warnings, images,
                width: root.getBoundingClientRect().width, unchanged: before === root.outerHTML };
        });
        await (await download).saveAs(test.info().outputPath('print-fidelity.xlsx'));
        const book = new ExcelJS.Workbook();
        await book.xlsx.load(Buffer.from(result.bytes));
        const sheet = book.worksheets[0];
        const cells = [];
        sheet.eachRow(row => row.eachCell(cell => {
            if (!cell.isMerged || cell.master === cell) cells.push(cell);
        }));
        expect(result.unchanged).toBe(true);
        expect(result.warnings).toEqual([]);
        expect(sheet.columns.reduce((sum, c) => sum + pixels(c.width), 0)).toBeCloseTo(result.width, 0);
        const images = sheet.getImages();
        expect(images).toHaveLength(3);
        for (let i = 0; i < images.length; i++) {
            const start = markerPixels(images[i].range.tl, sheet);
            const end = markerPixels(images[i].range.br, sheet);
            expect(Math.abs(start.x - result.images[i].x)).toBeLessThan(1.1);
            // Multiline rows reserve Excel's line height; page-relative image
            // offsets and full image dimensions remain fixed after that expansion.
            expect(images[i].range.tl.nativeRowOff).toBe(i ? 2 * 9525 : 0);
            expect(end.x - start.x).toBeCloseTo(result.images[i].width, 3);
            expect(end.y - start.y).toBeCloseTo(result.images[i].height, 3);
        }
        const customer = cells.find(cell => cell.text.startsWith('Customer Name'));
        expect(customer.value.richText.map(run => run.text).join('')).toBe('Customer Name\nFlat 14 Industrial Centre\nAddress line two');
        expect(customer.value.richText.find(run => run.text.includes('Customer Name')).font.bold).toBe(true);
        expect(customer.value.richText.find(run => run.text.includes('Flat 14')).font.bold).not.toBe(true);
        expect(customer.value.richText.every(run => run.font.size === 9)).toBe(true);
        let customerHeight = 0;
        for (let row = customer.row; row <= sheet.rowCount && sheet.getCell(row, customer.col).master === customer; row++) {
            customerHeight += sheet.getRow(row).height;
        }
        expect(customerHeight).toBeGreaterThanOrEqual(36.75);
        const contact = cells.find(cell => cell.text.startsWith('Contact '));
        expect(contact.value.richText.find(run => run.text.includes('Test Person')).font.bold).toBe(true);
        expect(contact.value.richText.find(run => run.text.includes('Second line')).font.italic).toBe(true);
        expect(cells.some(cell => cell.text.includes('PRIVATE'))).toBe(false);
        expect(cells.find(cell => cell.text === 'Summary').fill?.fgColor).toBeUndefined();
        const part = cells.find(cell => cell.text === 'Part A -');
        expect(part.alignment.wrapText).not.toBe(true);
        expect(part.alignment.shrinkToFit).toBe(true);
        expect(cells.filter(cell => typeof cell.value === 'number').map(cell => cell.value)).toEqual([1, 1800, 1800, 1, 1800, 1800, 8700]);
        expect(cells.filter(cell => cell.numFmt).every(cell => cell.numFmt.includes('"HKD"'))).toBe(true);
        expect(sheet.pageSetup.horizontalCentered).not.toBe(true);
        expect(sheet.pageSetup.fitToPage).toBe(false);
        expect(sheet.pageSetup.scale).toBeGreaterThanOrEqual(98);
        expect(sheet.pageSetup.printArea).toMatch(/^A1:/);
        // ExcelJS does not read rowBreaks back; inspect the serialized worksheet.
        const zip = await JSZip.loadAsync(Buffer.from(result.bytes));
        const xml = await zip.file('xl/worksheets/sheet1.xml').async('string');
        const breaks = [...xml.matchAll(/<brk\b[^>]*\bid="(\d+)"/g)].map(match => Number(match[1]));
        expect(breaks).toEqual(images.slice(1).map(image => image.range.tl.nativeRow));
        for (const row of sheet._rows.filter(Boolean)) expect((row.height / 0.75) % 1).toBe(0);
        await fs.writeFile(test.info().outputPath('browser-layout.json'), JSON.stringify({ images: result.images, scale: sheet.pageSetup.scale }));
    });
}

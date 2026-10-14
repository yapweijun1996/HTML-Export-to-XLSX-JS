import { test, expect } from '@playwright/test';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';

async function installPrintFixture(page, mode) {
    await page.goto('/tests/fixtures/quotation.html?mode=' + mode);
    await page.waitForFunction(() => Boolean(window.exportPrintFormToXlsx));
    return page.evaluate(() => {
        const root = document.querySelector('#export');
        // Reproduce a narrow unit-price cell and zero-height markers emitted by the ERP paginator.
        const header = title => '<table><tr style="height:80px"><td colspan="4">' + title + '</td></tr></table>';
        const detail = amount => '<table style="font:12px Arial;width:750px;table-layout:fixed">' +
            '<colgroup><col style="width:550px"><col style="width:35px"><col style="width:35px"><col style="width:130px"></colgroup>' +
            '<tr><td>Description</td><td>Unit</td><td>Price</td><td>Amount</td></tr>' +
            '<tr><td>Equipment</td><td colspan="2" style="text-align:right">$ ' + amount +
            '</td><td>Included</td></tr><tr style="height:700px"><td colspan="4">Terms</td></tr></table>';
        root.innerHTML = header('HEADER ONE') + detail('1,800.00') +
            '<div class="pagebreak_bf" style="font-size:0pt;line-height:0">&nbsp;</div>' +
            header('HEADER TWO') + detail('3,000.00') +
            '<div class="pagebreak_bf_processed" data-xlsx-page-break style="height:0"></div>' +
            header('HEADER THREE') + detail('-3,000.00') +
            '<div style="display:none"><div class="pagebreak_bf"></div><table><tr><td>HIDDEN</td></tr></table></div>';
        return root.outerHTML;
    });
}

for (const mode of ['', 'min']) {
    test('legacy print export preserves three page boundaries and fits numeric currency (' + (mode || 'source') + ')', async ({ page }) => {
        const before = await installPrintFixture(page, mode);
        const download = page.waitForEvent('download');
        const result = await page.evaluate(async () => {
            const result = await window.exportPrintFormToXlsx();
            return { bytes: [...new Uint8Array(result.buffer)], warnings: result.warnings,
                html: document.querySelector('#export').outerHTML };
        });
        await (await download).saveAs(test.info().outputPath('print-layout.xlsx'));
        expect(result.html).toBe(before);
        expect(result.warnings).toEqual([]);
        const bytes = Buffer.from(result.bytes);
        const book = new ExcelJS.Workbook();
        await book.xlsx.load(bytes);
        const sheet = book.worksheets[0];
        const cells = [];
        sheet.eachRow(row => row.eachCell(cell => {
            if (!cell.isMerged || cell.master === cell) cells.push(cell);
        }));
        const headers = ['HEADER ONE', 'HEADER TWO', 'HEADER THREE'].map(value => cells.find(c => c.value === value));
        expect(headers.every(Boolean)).toBe(true);
        const zip = await JSZip.loadAsync(bytes);
        const xml = await zip.file('xl/worksheets/sheet1.xml').async('string');
        const breaks = [...xml.matchAll(/<brk\b[^>]*\bid="(\d+)"/g)].map(match => Number(match[1]));
        expect(breaks).toEqual(headers.slice(1).map(c => c.row - 1));
        expect(xml).toContain('manualBreakCount="2"');
        expect(sheet.pageSetup.fitToPage).toBe(false);
        expect(sheet.pageSetup.scale).toBeGreaterThanOrEqual(10);
        expect(sheet.pageSetup.scale).toBeLessThan(100);
        const widthPt = sheet.columns.reduce((sum, col) => sum +
            Math.floor((256 * col.width + 18) / 256 * 7) * 0.75, 0);
        expect(widthPt * sheet.pageSetup.scale / 100).toBeLessThanOrEqual((210 / 25.4 - 0.5) * 72);
        for (let i = 0; i < headers.length; i++) {
            let heightPt = 0;
            const end = headers[i + 1]?.row - 1 || sheet.rowCount;
            for (let r = headers[i].row; r <= end; r++) heightPt += sheet.getRow(r).height;
            expect(heightPt * sheet.pageSetup.scale / 100).toBeLessThanOrEqual((297 / 25.4 - 0.8) * 72);
        }
        const amounts = cells.filter(c => typeof c.value === 'number');
        expect(amounts.map(c => c.value)).toEqual([1800, 3000, -3000]);
        for (const cell of amounts) {
            expect(cell.numFmt).toContain('"$"');
            expect(cell.alignment.shrinkToFit).toBe(true);
            expect(cell.alignment.wrapText).not.toBe(true);
            expect(cell.isMerged).toBe(true);
        }
        expect(cells.filter(c => c.value === 'Included')).toHaveLength(3);
        expect(cells.some(c => c.value === 'HIDDEN')).toBe(false);
    });

    test('unmarked forms and explicit Fit To selection keep automatic scaling (' + (mode || 'source') + ')', async ({ page }) => {
        await installPrintFixture(page, mode);
        const exports = await page.evaluate(async () => {
            const result = await window.HtmlToXlsx.exportXlsx({ root: '#export', download: false, respectPageBreaks: false });
            document.querySelectorAll('.pagebreak_bf, .pagebreak_bf_processed').forEach(marker => marker.remove());
            const unmarked = await window.HtmlToXlsx.exportXlsx({ root: '#export', download: false });
            return [result, unmarked].map(r => [...new Uint8Array(r.buffer)]);
        });
        for (const bytes of exports) {
            const book = new ExcelJS.Workbook();
            await book.xlsx.load(Buffer.from(bytes));
            expect(book.worksheets[0].pageSetup.fitToPage).toBe(true);
        }
    });
}

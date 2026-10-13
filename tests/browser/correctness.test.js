import { test, expect } from '@playwright/test';
import ExcelJS from 'exceljs';

async function exportHtml(page, mode, html, options = {}) {
    await page.goto('/tests/fixtures/quotation.html?mode=' + mode);
    await page.waitForFunction(() => Boolean(window.HtmlToXlsx));
    const result = await page.evaluate(async ({ html, options }) => {
        const root = document.querySelector('#export');
        root.innerHTML = html;
        const before = root.innerHTML;
        const result = await window.HtmlToXlsx.exportXlsx({ root, download: false, ...options });
        return { bytes: [...new Uint8Array(result.buffer)], warnings: result.warnings,
            unchanged: before === root.innerHTML };
    }, { html, options });
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(Buffer.from(result.bytes));
    const sheet = workbook.worksheets[0];
    const cells = [];
    sheet.eachRow(row => row.eachCell(cell => {
        if (!cell.isMerged || cell.address === cell.master.address) cells.push({ value: cell.value, format: cell.numFmt });
    }));
    expect(result.unchanged).toBe(true);
    return { ...result, sheet, cells };
}

for (const mode of ['', 'min']) {
    test('excluded descendants and hidden fragments stay out of the workbook (' + (mode || 'source') + ')', async ({ page }) => {
        const result = await exportHtml(page, mode, `
            <table><tr><td>PUBLIC <span data-xlsx-ignore>OMIT_CHILD</span><br>SECOND LINE
                <span class="private">OMIT_CUSTOM</span><span hidden>OMIT_HIDDEN</span>
                <span style="visibility:hidden">OMIT_INVISIBLE</span>
                <img data-xlsx-ignore src="./missing-ignore.png" width="20" height="20">
                <img style="visibility:hidden" src="./missing-hidden.png" width="20" height="20">
            </td></tr>
            <tr><td><span style="visibility:hidden">OMIT_NESTED</span><b>KEEP NESTED</b>
                <table><tr><td>KEEP CHILD</td></tr></table></td></tr>
            <tr><td data-xlsx-ignore data-excel-value="OMIT_OVERRIDE">OMIT CELL</td></tr>
            <tr><td style="visibility:hidden" data-excel-value="OMIT_HIDDEN_OVERRIDE">OMIT CELL</td></tr>
            <tr><td>BEFORE<div data-xlsx-ignore><table><tr><td>OMIT TABLE</td></tr></table></div>AFTER</td></tr>
            <tr><td style="visibility:hidden"><span style="visibility:visible">RESTORED VISIBLE</span></td></tr></table>`,
        { exclude: ['[data-xlsx-ignore]', '.private'], imageTimeoutMs: 100 });
        const text = result.cells.map(cell => String(cell.value)).join('\n');
        expect(text).not.toContain('OMIT');
        for (const label of ['PUBLIC', 'SECOND LINE', 'KEEP NESTED', 'KEEP CHILD', 'BEFORE', 'AFTER', 'RESTORED VISIBLE']) {
            expect(text).toContain(label);
        }
        expect(result.cells.some(cell => String(cell.value).includes('PUBLIC\nSECOND LINE'))).toBe(true);
        expect(result.sheet.getImages()).toHaveLength(0);
        expect(result.warnings).toEqual([]);
    });

    test('nested values honor text, numeric overrides and currency without duplication (' + (mode || 'source') + ')', async ({ page }) => {
        const result = await exportHtml(page, mode, `
            <table><tr><td data-excel-type="text">$ 123.00
                <table><tr><td>$ 5.00</td></tr></table></td></tr>
            <tr><td><span data-excel-type="number" data-excel-value="99"><b>Displayed</b> value</span>
                <table><tr><td>CHILD LABEL</td></tr></table></td></tr>
            <tr><td data-excel-currency="KWD"><span data-excel-type="currency"><b>1.</b><i>234</i></span>
                <table><tr><td>$ 10.00</td></tr></table></td></tr>
            <tr><td data-excel-type="text"><b>00</b><i>123</i>
                <table><tr><td>SECOND CHILD</td></tr></table></td></tr></table>`);
        expect(result.cells).toEqual(expect.arrayContaining([
            { value: '$ 123.00', format: undefined }, { value: 99, format: undefined },
            { value: 1.234, format: expect.stringContaining('KWD') },
            { value: '00123', format: undefined },
            { value: 5, format: expect.stringContaining('HKD') },
            { value: 10, format: expect.stringContaining('HKD') }
        ]));
        expect(result.cells.filter(cell => cell.value === 99)).toHaveLength(1);
        expect(result.cells.some(cell => cell.value === 'Displayed' || cell.value === 'value')).toBe(false);
        expect(result.warnings).toEqual([]);
    });

    test('left center and right image anchors survive workbook serialization (' + (mode || 'source') + ')', async ({ page }) => {
        const rows = ['left', 'center', 'right'].map(alignment =>
            `<tr><td colspan="3" style="height:100px;padding:0" data-excel-image-align="${alignment}">
            <img src="./stamp.svg" width="60" height="20"></td></tr>`).join('');
        const result = await exportHtml(page, mode,
            `<table style="width:300px"><colgroup><col style="width:100px"><col style="width:100px"><col style="width:100px"></colgroup>
            ${rows}<tr><td style="padding:0">A</td><td style="padding:0">B</td><td style="padding:0">C</td></tr></table>`,
            { imageWidthScale: 1 });
        const images = result.sheet.getImages();
        expect(images).toHaveLength(3);
        expect(images.map(image => image.range.tl.col)).toEqual([
            expect.closeTo(0, 4), expect.closeTo(1.2, 4), expect.closeTo(2.4, 4)
        ]);
        expect(result.warnings).toEqual([]);
    });
}

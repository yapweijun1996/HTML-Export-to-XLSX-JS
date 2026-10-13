import { test, expect } from '@playwright/test';
import ExcelJS from 'exceljs';

async function exportHtml(page, mode, html) {
    await page.goto('/tests/fixtures/quotation.html?mode=' + mode);
    await page.waitForFunction(() => Boolean(window.HtmlToXlsx));
    const result = await page.evaluate(async html => {
        const root = document.querySelector('#export');
        root.innerHTML = html;
        const style = document.createElement('style');
        style.textContent = '#export td{padding:0;font:12px Arial;vertical-align:top}';
        document.head.appendChild(style);
        const before = root.outerHTML;
        const bounds = Array.from(root.querySelectorAll('td')).map(cell => cell.getBoundingClientRect().toJSON());
        const result = await HtmlToXlsx.exportXlsx({ root, download: false });
        return { bytes: [...new Uint8Array(result.buffer)], warnings: result.warnings,
            unchanged: root.outerHTML === before,
            sameBounds: JSON.stringify(bounds) === JSON.stringify(Array.from(root.querySelectorAll('td')).map(cell => cell.getBoundingClientRect().toJSON())) };
    }, html);
    const book = new ExcelJS.Workbook();
    await book.xlsx.load(Buffer.from(result.bytes));
    const sheet = book.worksheets[0];
    const cells = [];
    sheet.eachRow(row => row.eachCell(cell => {
        if (!cell.isMerged || cell.master === cell) cells.push(cell);
    }));
    expect(result.unchanged).toBe(true);
    expect(result.sameBounds).toBe(true);
    expect(result.warnings).toEqual([]);
    return { sheet, cells };
}

for (const mode of ['', 'min']) {
    test('ERP captions retain descendant fonts and stay on one line (' + (mode || 'source') + ')', async ({ page }) => {
        const { cells } = await exportHtml(page, mode, `
            <table><colgroup><col style="width:50px"><col style="width:700px"></colgroup>
                <tr><td align="center"><font color="darkblue"><b>Part A - </font></b></td>
                    <td><font color="darkblue"><b><u>部分 A -</u></font></b></td></tr>
                <tr><td><u><b>Item</b></u></td><td><u>Description</u></td></tr>
                <tr><td><font color="darkblue"><b>Part B -</b></font></td>
                    <td><font color="darkblue"><b><u>部分 B -</u></b></font></td></tr>
                <tr><td colspan="2"><span style="font-family:Verdana;color:blue"><i><u>ITALIC CAPTION</u></i></span></td></tr>
                <tr><td colspan="2"><span style="text-decoration:underline"><b><span>ANCESTOR UNDERLINE</span></b></span></td></tr>
                <tr><td colspan="2">$ 1,800.00</td></tr></table>`);
        const find = text => cells.find(cell => String(cell.value).trim() === text);
        for (const text of ['Part A -', 'Part B -', '部分 A -', '部分 B -']) {
            expect(find(text).font.color.argb).toBe('FF00008B');
            expect(find(text).font.bold).toBe(true);
            expect(find(text).alignment.wrapText).not.toBe(true);
            expect(find(text).alignment.shrinkToFit).toBe(true);
        }
        for (const text of ['部分 A -', '部分 B -', 'Item', 'Description', 'ANCESTOR UNDERLINE']) {
            expect(find(text).font.underline).toBe(true);
        }
        const italic = find('ITALIC CAPTION');
        expect(italic.font.name).toBe('Verdana');
        expect(italic.font.italic).toBe(true);
        expect(italic.font.underline).toBe(true);
        expect(italic.font.color.argb).toBe('FF0000FF');
        expect(cells.filter(c => c.value === 1800)).toHaveLength(1);
    });

    test('explicit and naturally wrapped styled text remain multiline (' + (mode || 'source') + ')', async ({ page }) => {
        const { cells } = await exportHtml(page, mode, `
            <table style="width:100px"><tr><td><font color="darkblue"><b><u>FIRST<br>SECOND</u></b></font></td></tr>
                <tr><td><font color="darkblue"><b><u>NATURALLY WRAPPED LONG CAPTION</u></b></font></td></tr></table>`);
        for (const cell of cells.filter(c => c.value)) {
            expect(cell.font.underline).toBe(true);
            expect(cell.font.bold).toBe(true);
            expect(cell.font.color.argb).toBe('FF00008B');
            expect(cell.alignment.wrapText).toBe(true);
            expect(cell.alignment.shrinkToFit).not.toBe(true);
        }
        expect(cells.some(c => c.value === 'FIRST\nSECOND')).toBe(true);
    });

    test('hidden, mixed and nested text do not leak formatting (' + (mode || 'source') + ')', async ({ page }) => {
        const { cells } = await exportHtml(page, mode, `
            <table><tr><td><font color="darkblue"><b><u>VISIBLE CAPTION</u></b></font>
                <span hidden style="font-size:100px;color:red">PRIVATE</span></td></tr>
                <tr><td><font color="blue"><u><i>STYLED</i></u></font> PLAIN</td></tr>
                <tr><td><span style="text-decoration:line-through"><b>STRIKE CAPTION</b></span></td></tr>
                <tr><td><font color="darkblue"><b><u>NESTED CAPTION</u></b></font>
                    <table><tr><td>CHILD LABEL</td></tr></table></td></tr></table>`);
        const find = text => cells.find(cell => String(cell.value).trim() === text);
        for (const text of ['VISIBLE CAPTION', 'NESTED CAPTION']) {
            expect(find(text).font.color.argb).toBe('FF00008B');
            expect(find(text).font.underline).toBe(true);
            expect(find(text).alignment.shrinkToFit).toBe(true);
        }
        expect(find('VISIBLE CAPTION').font.size).toBeLessThan(10);
        const mixed = find('STYLED PLAIN');
        expect(mixed.font.color.argb).toBe('FF000000');
        expect(mixed.font.underline).not.toBe(true);
        expect(mixed.font.italic).not.toBe(true);
        expect(find('STRIKE CAPTION').font.strike).toBe(true);
        expect(find('CHILD LABEL').font.color.argb).toBe('FF000000');
        expect(find('CHILD LABEL').font.underline).not.toBe(true);
        expect(cells.some(c => String(c.value).includes('PRIVATE'))).toBe(false);
    });
}

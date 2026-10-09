import { test, expect } from '@playwright/test';
import ExcelJS from 'exceljs';

async function exported(page, mode = '') {
    await page.goto('/tests/fixtures/quotation.html?mode=' + mode);
    await page.waitForFunction(() => Boolean(window.HtmlToXlsx));
    expect(await page.evaluate(() => window.HtmlToXlsx.version)).toBe('0.1.0');
    const result = await page.evaluate(async () => {
        const result = await window.HtmlToXlsx.exportXlsx({ root: '#export', download: false });
        return { bytes: [...new Uint8Array(result.buffer)], warnings: result.warnings, stats: result.stats };
    });
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(Buffer.from(result.bytes));
    return { workbook, sheet: workbook.worksheets[0], ...result };
}
function values(sheet) {
    const result = [];
    sheet.eachRow(row => row.eachCell(cell => {
        if (!cell.isMerged || cell.address === cell.master.address) result.push({ value: cell.value, format: cell.numFmt });
    }));
    return result;
}
for (const mode of ['', 'min']) {
    test('complete editable quotation and images (' + (mode || 'source') + ')', async ({ page }) => {
        const result = await exported(page, mode);
        const cells = values(result.sheet);
        const labels = cells.map(cell => cell.value).filter(value => typeof value === 'string' && /^[ABCEFG]\d+$/.test(value));
        expect(labels).toHaveLength(68);
        expect(new Set(labels).size).toBe(68);
        expect(cells.map(cell => cell.value)).toEqual(expect.arrayContaining([120.25, 80, 300.5, 500.75]));
        const amounts = cells.filter(cell => typeof cell.value === 'number' && cell.format?.includes('HKD'));
        expect(amounts.length).toBeGreaterThanOrEqual(140);
        expect(cells.filter(cell => cell.format?.includes('US$'))).toHaveLength(0);
        expect(cells.some(cell => cell.value === 'INTERNAL OMITTED')).toBe(false);
        expect(cells.some(cell => String(cell.value).includes('測試設備'))).toBe(true);
        expect(result.sheet.getImages()).toHaveLength(1);
        expect(result.warnings).toEqual([]);
        await page.screenshot({ path: test.info().outputPath('fixture.png'), fullPage: true });
    });
}
test('configuration does not leak between concurrent exports and explicit codes win', async ({ page }) => {
    await page.goto('/tests/fixtures/quotation.html');
    await page.waitForFunction(() => Boolean(window.HtmlToXlsx));
    const bytes = await page.evaluate(async () => {
        const results = await Promise.all(['USD', 'SGD'].map(currency =>
            window.HtmlToXlsx.exportXlsx({ root: '#export', currency, download: false })));
        return results.map(result => [...new Uint8Array(result.buffer)]);
    });
    for (const [index, currency] of ['USD', 'SGD'].entries()) {
        const workbook = new ExcelJS.Workbook(); await workbook.xlsx.load(Buffer.from(bytes[index]));
        expect(values(workbook.worksheets[0]).some(cell => cell.format?.includes(currency))).toBe(true);
        expect(values(workbook.worksheets[0]).some(cell => cell.format?.includes('HKD'))).toBe(false);
    }
});
test('generic API reports missing roots and unsupported options', async ({ page }) => {
    await page.goto('/tests/fixtures/quotation.html');
    await page.waitForFunction(() => Boolean(window.HtmlToXlsx));
    const errors = await page.evaluate(async () => {
        const messages = [];
        for (const options of [{ root: '#absent' }, { root: '#export', currency: 'XYZ' }, { root: '#export', edgeTolerancePx: 0 }]) {
            try { await window.HtmlToXlsx.exportXlsx({ ...options, download: false }); }
            catch (error) { messages.push(error.message); }
        }
        return messages;
    });
    expect(errors).toHaveLength(3);
});
test('legacy adapter uses container currency and button binding works after late script load', async ({ page }) => {
    await page.goto('/tests/fixtures/quotation.html?mode=min');
    await page.waitForFunction(() => Boolean(window.exportPrintFormToXlsx));
    const download = page.waitForEvent('download');
    await page.getByRole('button', { name: 'Export XLSX' }).click();
    const file = await download;
    const path = test.info().outputPath('button.xlsx');
    await file.saveAs(path);
    const workbook = new ExcelJS.Workbook(); await workbook.xlsx.readFile(path);
    expect(values(workbook.worksheets[0]).some(cell => cell.format?.includes('HKD'))).toBe(true);
    await expect(page.getByRole('button', { name: 'Export XLSX' })).toBeEnabled();
});

test('regional currencies, typed identifiers and literal formula text survive export', async ({ page }) => {
    await page.goto('/tests/fixtures/quotation.html?mode=min');
    await page.waitForFunction(() => Boolean(window.HtmlToXlsx));
    const bytes = await page.evaluate(async () => {
        const root = document.createElement('div');
        root.style.width = '750px';
        root.innerHTML = '<table><tr><td>EUR 1.234,56</td><td data-excel-type="text">00123</td>' +
            '<td>=SUM(A1:A2)</td><td data-excel-currency="KWD" data-excel-type="currency">1,234</td></tr></table>';
        document.body.appendChild(root);
        const result = await window.HtmlToXlsx.exportXlsx({ root, locale: 'de-DE', download: false });
        return [...new Uint8Array(result.buffer)];
    });
    const workbook = new ExcelJS.Workbook(); await workbook.xlsx.load(Buffer.from(bytes));
    const cells = values(workbook.worksheets[0]);
    expect(cells).toEqual(expect.arrayContaining([
        expect.objectContaining({ value: 1234.56, format: expect.stringContaining('EUR') }),
        expect.objectContaining({ value: 1.234, format: expect.stringContaining('KWD') }),
        expect.objectContaining({ value: '00123' }), expect.objectContaining({ value: '=SUM(A1:A2)' })
    ]));
});

test('image failures are bounded and reported without silently claiming success', async ({ page }) => {
    await page.goto('/tests/fixtures/quotation.html');
    await page.waitForFunction(() => Boolean(window.HtmlToXlsx));
    const result = await page.evaluate(async () => {
        const image = document.createElement('img');
        image.src = './does-not-exist.png'; image.width = 20; image.height = 20;
        document.querySelector('#export td').appendChild(image);
        const result = await window.HtmlToXlsx.exportXlsx({ root: '#export', download: false, imageTimeoutMs: 100 });
        return { warnings: result.warnings, images: result.stats.images };
    });
    expect(result.warnings.length).toBeGreaterThan(0);
    expect(result.images).toBe(1);
});

test('visible cells and merges remain compatible with the original exporter', async ({ page }) => {
    await page.goto('/tests/fixtures/quotation.html?mode=baseline');
    await page.waitForFunction(() => Boolean(window.exportPrintFormToXlsx));
    const download = page.waitForEvent('download');
    await page.evaluate(() => window.exportPrintFormToXlsx());
    const file = await download;
    const path = test.info().outputPath('baseline.xlsx');
    await file.saveAs(path);
    const original = new ExcelJS.Workbook(); await original.xlsx.readFile(path);
    const current = await exported(page);
    const texts = sheet => values(sheet).map(cell => String(cell.value ?? '').trim()).filter(Boolean);
    expect(texts(current.sheet)).toEqual(texts(original.worksheets[0]));
    expect(current.sheet.model.merges).toEqual(original.worksheets[0].model.merges);
    expect(current.sheet.getImages().length).toBe(original.worksheets[0].getImages().length);
});

test('precise image placement and explicit page setup serialize correctly', async ({ page }) => {
    await page.goto('/tests/fixtures/quotation.html?mode=min');
    await page.waitForFunction(() => Boolean(window.HtmlToXlsx));
    const bytes = await page.evaluate(async () => {
        const result = await window.HtmlToXlsx.exportXlsx({
            root: '#export', download: false, imagePlacement: 'precise',
            pageSetup: { orientation: 'landscape', margins: { left: 0.7 } }
        });
        return [...new Uint8Array(result.buffer)];
    });
    const workbook = new ExcelJS.Workbook(); await workbook.xlsx.load(Buffer.from(bytes));
    const sheet = workbook.worksheets[0];
    expect(sheet.getImages()).toHaveLength(1);
    expect(sheet.pageSetup.orientation).toBe('landscape');
    expect(sheet.pageSetup.margins.left).toBe(0.7);
    expect(sheet.pageSetup.fitToPage).toBe(true);
});

test('repeated bindings and stale unbind callbacks do not duplicate downloads', async ({ page }) => {
    await page.goto('/examples/basic.html');
    await page.waitForFunction(() => Boolean(window.HtmlToXlsx));
    await page.evaluate(() => {
        const button = document.querySelector('button');
        const first = window.HtmlToXlsx.bindButton(button, { root: '#quotation' });
        window.HtmlToXlsx.bindButton(button, { root: '#quotation' });
        first();
        window.HtmlToXlsx.bindButton(button, { root: '#quotation' });
    });
    const downloads = [];
    page.on('download', file => downloads.push(file));
    const download = page.waitForEvent('download');
    await page.getByRole('button', { name: 'Export XLSX' }).click();
    await download;
    await expect(page.locator('#status')).toHaveText('Workbook exported.');
    expect(downloads).toHaveLength(1);
});

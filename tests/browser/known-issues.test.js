import { test, expect } from '@playwright/test';
import ExcelJS from 'exceljs';

// Regression coverage for the fixed issues in KNOWN_ISSUES.md.
async function exportHtml(page, mode, html, options = {}) {
    await page.goto('/tests/fixtures/quotation.html?mode=' + mode);
    await page.waitForFunction(() => Boolean(window.HtmlToXlsx));
    const result = await page.evaluate(async ({ html, options }) => {
        const root = document.querySelector('#export');
        root.innerHTML = html;
        const style = document.createElement('style');
        style.textContent = '#export td, #export th{padding:0;font:12px Arial}';
        document.head.appendChild(style);
        await Promise.all(Array.from(root.querySelectorAll('img')).map(image => image.decode()));
        const origin = root.getBoundingClientRect();
        const rendered = Array.from(root.querySelectorAll('img')).map(image => {
            const rect = image.getBoundingClientRect();
            return { x: rect.left - origin.left, y: rect.top - origin.top };
        });
        const before = root.innerHTML;
        const result = await HtmlToXlsx.exportXlsx({ root, download: false, ...options });
        return { bytes: [...new Uint8Array(result.buffer)], warnings: result.warnings, rendered,
            unchanged: before === root.innerHTML };
    }, { html, options });
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(Buffer.from(result.bytes));
    const sheet = workbook.worksheets[0];
    const cells = [];
    sheet.eachRow(row => row.eachCell(cell => {
        if (!cell.isMerged || cell.address === cell.master.address) cells.push(cell);
    }));
    expect(result.unchanged).toBe(true);
    return { ...result, workbook, sheet, cells };
}

function imagePosition(image, sheet, marker = image.range.tl) {
    let x = marker.nativeColOff / 9525, y = marker.nativeRowOff / 9525;
    for (let c = 1; c <= marker.nativeCol; c++) {
        x += Math.floor((256 * sheet.getColumn(c).width + 18) / 256 * 7);
    }
    for (let r = 1; r <= marker.nativeRow; r++) y += sheet.getRow(r).height / 0.75;
    return { x, y };
}

const stamp = (width, height, style = '') => `<img src="./stamp.svg" width="${width}" height="${height}" style="${style}">`;

for (const mode of ['', 'min']) {
    const suffix = ' (' + (mode || 'source') + ')';

    test('amounts and typed numbers display the decimals, grouping and words of their text' + suffix, async ({ page }) => {
        const { cells } = await exportHtml(page, mode, `<table style="width:400px"><tr>
            <td>$ 12.3456</td><td>$ 100</td><td data-excel-type="number">2.00</td>
            <td data-excel-type="number">1,250</td><td>Top 10</td></tr></table>`, { currency: 'HKD' });
        const find = value => cells.find(cell => cell.value === value);
        expect(find(12.3456).numFmt).toBe('"HKD" #,##0.0000;("HKD" #,##0.0000)');
        expect(find(100).numFmt).toBe('"HKD" #,##0;("HKD" #,##0)');
        expect(find(2).numFmt).toBe('0.00');
        expect(find(1250).numFmt).toBe('#,##0');
        expect(find('Top 10')).toBeDefined();
    });

    test('legacy and inner-block text alignment follows the rendered text' + suffix, async ({ page }) => {
        const { cells } = await exportHtml(page, mode, `<table style="width:300px">
            <tr align="center"><td>ROW CENTER</td></tr>
            <tr><td><center>CENTER TAG</center></td></tr>
            <tr><td><div align="right">DIV RIGHT</div></td></tr>
            <tr><th>HEADER</th></tr>
            <tr><td align="center"><b>CELL CENTER</b></td></tr>
            <tr><td>PLAIN</td></tr></table>`);
        const horizontal = value => cells.find(cell => cell.value === value).alignment?.horizontal;
        for (const value of ['ROW CENTER', 'CENTER TAG', 'HEADER', 'CELL CENTER']) expect(horizontal(value), value).toBe('center');
        expect(horizontal('DIV RIGHT')).toBe('right');
        expect(horizontal('PLAIN')).toBe('left');
    });

    for (const imagePlacement of ['alignment', 'precise']) {
        test('images outside table cells keep their rendered position (' + imagePlacement + ')' + suffix, async ({ page }) => {
            const result = await exportHtml(page, mode, `<div style="height:60px">${stamp(40, 40, 'margin:10px 0 0 30px')}</div>
                <table style="width:300px"><tr><td>BODY</td></tr></table>`, { imagePlacement, imageWidthScale: 1 });
            expect(result.sheet.getImages()).toHaveLength(1);
            const position = imagePosition(result.sheet.getImages()[0], result.sheet);
            expect(Math.abs(position.x - result.rendered[0].x)).toBeLessThanOrEqual(1.5);
            expect(Math.abs(position.y - result.rendered[0].y)).toBeLessThanOrEqual(1.5);
            expect(result.warnings).toEqual([]);
        });
    }

    test('aligned images start at their rendered top instead of covering preceding text' + suffix, async ({ page }) => {
        const result = await exportHtml(page, mode, `<table style="width:300px">
            <tr><td>Signature:<br>line two<br>${stamp(40, 40)}</td></tr>
            <tr><td style="text-align:center">Authorised<br>${stamp(40, 40)}</td></tr></table>`, { imageWidthScale: 1 });
        const images = result.sheet.getImages();
        expect(images).toHaveLength(2);
        for (const [index, image] of images.entries()) {
            const position = imagePosition(image, result.sheet);
            expect(result.rendered[index].y).toBeGreaterThan(10);
            expect(Math.abs(position.y - result.rendered[index].y)).toBeLessThanOrEqual(1.5);
        }
        expect(Math.abs(imagePosition(images[1], result.sheet).x - 130)).toBeLessThanOrEqual(1.5);
        expect(result.warnings).toEqual([]);
    });

    test('an image that overflows its cell falls back to its rendered position unless the flow is forced' + suffix, async ({ page }) => {
        const html = attribute => `<table style="width:300px;table-layout:fixed"><tr><td${attribute}>${stamp(400, 20)}</td></tr></table>`;
        const result = await exportHtml(page, mode, html(''), { imageWidthScale: 1 });
        expect(result.sheet.getImages()).toHaveLength(1);
        expect(Math.abs(imagePosition(result.sheet.getImages()[0], result.sheet).x - result.rendered[0].x)).toBeLessThanOrEqual(1.5);
        expect(result.warnings.map(warning => warning.code)).toEqual(['IMAGE_ALIGNMENT_FALLBACK']);
        await expect(exportHtml(page, mode, html(' data-excel-image-align="inline"'), { imageWidthScale: 1 }))
            .rejects.toThrow(/Image alignment exceeds/);
    });

    test('a failed click on an auto-bound button is recorded in the console' + suffix, async ({ page }) => {
        const errors = [];
        page.on('console', message => { if (message.type() === 'error') errors.push(message.text()); });
        await page.goto('/tests/fixtures/quotation.html?mode=' + mode);
        await page.waitForFunction(() => Boolean(window.HtmlToXlsx));
        await page.evaluate(() => { document.querySelector('#export').innerHTML = '<p>No table</p>'; });
        await page.click('[data-xlsx-button]');
        await expect.poll(() => errors.some(text => text.includes('XLSX export failed'))).toBe(true);
    });

    test('an off-screen lazy image costs one timeout, not one per stage' + suffix, async ({ page }) => {
        await page.goto('/tests/fixtures/quotation.html?mode=' + mode);
        await page.waitForFunction(() => Boolean(window.HtmlToXlsx));
        const result = await page.evaluate(async () => {
            const root = document.querySelector('#export');
            const spacer = document.createElement('div');
            spacer.style.height = '30000px';
            root.before(spacer);
            root.innerHTML = '<table><tr><td>LAZY <img loading="lazy" src="./stamp.svg?lazy" width="20" height="20"></td></tr></table>';
            const image = root.querySelector('img');
            const started = performance.now();
            const exported = await HtmlToXlsx.exportXlsx({ root, download: false, imageTimeoutMs: 1000 });
            return { elapsed: performance.now() - started, loaded: image.naturalWidth > 0,
                codes: exported.warnings.map(warning => warning.code) };
        });
        expect(result.loaded).toBe(false);
        expect(result.codes).toEqual(expect.arrayContaining(['IMAGE_PRELOAD_FAILED', 'IMAGE_LOAD_FAILED']));
        expect(result.elapsed).toBeGreaterThanOrEqual(1000);
        expect(result.elapsed).toBeLessThan(1900);
    });

    test('large rasters stay within the canvas limit while keeping their drawing size' + suffix, async ({ page }) => {
        const result = await exportHtml(page, mode, `<table><tr><td>${stamp(4000, 2000)}</td></tr></table>`, { imageWidthScale: 1 });
        const media = Buffer.from(result.workbook.model.media[0].buffer);
        const [width, height] = [media.readUInt32BE(16), media.readUInt32BE(20)];
        expect(width * height).toBeLessThanOrEqual(16777216);
        expect(Math.max(width, height)).toBeLessThanOrEqual(16384);
        expect(width).toBeGreaterThan(4000);
        const image = result.sheet.getImages()[0];
        const start = imagePosition(image, result.sheet), end = imagePosition(image, result.sheet, image.range.br);
        expect(Math.abs(end.x - start.x - 4000)).toBeLessThanOrEqual(1);
        expect(Math.abs(end.y - start.y - 2000)).toBeLessThanOrEqual(1);
    });

    test('a tall HTML row is split into merged rows within Excel\'s 409pt limit' + suffix, async ({ page }) => {
        const { sheet, cells } = await exportHtml(page, mode, '<table style="width:300px"><tr><td style="height:1200px">TALL</td></tr></table>');
        const heights = Array.from({ length: sheet.rowCount }, (_, index) => sheet.getRow(index + 1).height);
        expect(Math.max(...heights)).toBeLessThanOrEqual(409);
        expect(Math.abs(heights.reduce((sum, height) => sum + height, 0) - 900)).toBeLessThanOrEqual(1);
        const tall = cells.find(cell => cell.value === 'TALL');
        expect(sheet.model.merges.some(range => range.startsWith(tall.address + ':'))).toBe(true);
    });

    test('a rowScale below fontScale is reported because rows keep the font scale' + suffix, async ({ page }) => {
        const html = '<table style="width:300px"><tr><td style="height:40px">ROW</td></tr></table>';
        const floored = await exportHtml(page, mode, html, { rowScale: 0.5 });
        expect(floored.warnings.map(warning => warning.code)).toEqual(['ROW_SCALE_FLOOR']);
        expect(floored.sheet.getRow(1).height).toBe(27);
        expect((await exportHtml(page, mode, html)).warnings).toEqual([]);
    });

    test('translucent fills and text are blended over the color beneath them' + suffix, async ({ page }) => {
        const { cells } = await exportHtml(page, mode, `<table style="width:300px;background:#ff0000"><tr>
            <td style="background:rgba(0,0,255,0.5)">MIX</td><td style="background:rgba(0,0,0,0.05)">ZEBRA</td>
            <td style="color:rgba(0,0,0,0.5)">GRAY</td></tr></table>`);
        const find = value => cells.find(cell => cell.value === value);
        expect(find('MIX').fill.fgColor.argb).toBe('FF800080');
        expect(find('ZEBRA').fill.fgColor.argb).toBe('FFF20000');
        expect(find('GRAY').fill.fgColor.argb).toBe('FFFF0000');
        expect(find('GRAY').font.color.argb).toBe('FF808080');
    });

    test('a refined grid keeps the sheet as wide as the HTML' + suffix, async ({ page }) => {
        const result = await exportHtml(page, mode, `<table style="width:300px;table-layout:fixed"><colgroup>
            <col style="width:1px"><col style="width:1px"><col style="width:298px"></colgroup>
            <tr><td>A</td><td>B</td><td>C</td></tr></table>`);
        expect(result.warnings.map(warning => warning.code)).toContain('GRID_REFINED');
        const width = result.sheet.columns.reduce((sum, column) => sum + Math.floor((256 * column.width + 18) / 256 * 7), 0);
        expect(Math.abs(width - 750)).toBeLessThanOrEqual(2);
    });

    test('borders hidden by an !important stylesheet rule stay hidden' + suffix, async ({ page }) => {
        const { cells } = await exportHtml(page, mode, `<style>#export td.plain{border:none !important}</style>
            <table style="width:300px"><tr><td class="plain" style="border:2px solid red">PLAIN</td>
            <td style="border:2px solid red">BOXED</td></tr></table>`);
        const border = value => cells.find(cell => cell.value === value).border || {};
        expect(['top', 'bottom', 'left'].filter(side => border('PLAIN')[side])).toEqual([]);
        expect(['top', 'bottom', 'right'].every(side => border('BOXED')[side])).toBe(true);
    });
}

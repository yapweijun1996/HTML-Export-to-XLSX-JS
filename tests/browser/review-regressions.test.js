import { test, expect } from '@playwright/test';
import ExcelJS from 'exceljs';
import http from 'node:http';

let imageServer;
let imageOrigin;
test.beforeAll(async () => {
    imageServer = http.createServer((request, response) => {
        const color = request.url.startsWith('/blue') ? 'blue' : 'red';
        response.writeHead(200, { 'Content-Type': 'image/svg+xml', 'Access-Control-Allow-Origin': '*' });
        response.end(`<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20"><rect width="20" height="20" fill="${color}"/></svg>`);
    });
    await new Promise(resolve => imageServer.listen(0, '127.0.0.1', resolve));
    imageOrigin = 'http://127.0.0.1:' + imageServer.address().port;
});
test.afterAll(async () => {
    imageServer.closeAllConnections();
    await new Promise(resolve => imageServer.close(resolve));
});

async function prepare(page, mode, html) {
    await page.goto('/tests/fixtures/quotation.html?mode=' + mode);
    await page.waitForFunction(() => Boolean(window.HtmlToXlsx));
    await page.evaluate(html => {
        document.querySelector('#export').innerHTML = html;
        const style = document.createElement('style');
        style.textContent = '#export table{border-collapse:collapse} #export td{padding:0;font:12px Arial}';
        document.head.appendChild(style);
    }, html);
}

async function exported(page, options = {}) {
    const result = await page.evaluate(async options => {
        const root = document.querySelector('#export');
        const before = root.innerHTML;
        const result = await HtmlToXlsx.exportXlsx({ root, download: false, imageScale: 1, imageWidthScale: 1, ...options });
        return { bytes: [...new Uint8Array(result.buffer)], warnings: result.warnings, stats: result.stats,
            unchanged: before === root.innerHTML };
    }, options);
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

function imagePosition(image, sheet) {
    const marker = image.range.tl;
    let x = marker.nativeColOff / 9525, y = marker.nativeRowOff / 9525;
    for (let c = 1; c <= marker.nativeCol; c++) {
        x += Math.floor((256 * sheet.getColumn(c).width + 18) / 256 * 7);
    }
    for (let r = 1; r <= marker.nativeRow; r++) y += sheet.getRow(r).height / 0.75;
    return { x, y };
}

for (const mode of ['', 'min']) {
    const suffix = ' (' + (mode || 'source') + ')';
    test('disjoint text and nested declarations retain each value once' + suffix, async ({ page }) => {
        await prepare(page, mode, `<table><tr><td data-excel-type="text">BEFORE<table><tr><td>INNER</td></tr></table>AFTER</td></tr>
            <tr><td><span data-excel-type="text">OUTER <span data-excel-type="text">CHILD</span> END</span>
            <table><tr><td>SECOND</td></tr></table></td></tr>
            <tr><td data-excel-type="number" data-excel-value="99">FIRST<table><tr><td>THIRD</td></tr></table>LAST</td></tr></table>`);
        const result = await exported(page);
        expect(result.cells.map(cell => cell.value).filter(value => value !== '')).toEqual([
            'BEFORE', 'INNER', 'AFTER', 'OUTER', 'CHILD', 'END', 'SECOND', 99, 'THIRD'
        ]);
        expect(result.warnings).toEqual([]);
    });

    test('overlapping text rejects instead of reporting a partial workbook as success' + suffix, async ({ page }) => {
        await prepare(page, mode, '<table><tr><td style="line-height:14px">OUTER<table style="margin-top:-14px"><tr><td>INNER</td></tr></table></td></tr></table>');
        await expect(exported(page)).rejects.toThrow(/Overlapping|merge/i);
        await prepare(page, mode, '<table><tr><td data-excel-type="number">12<table><tr><td>INNER</td></tr></table>34</td></tr></table>');
        await expect(exported(page)).rejects.toThrow(/numeric declaration spans separate regions/);
    });

    test('filtered text retains rendered case, whitespace and explicit line breaks' + suffix, async ({ page }) => {
        await prepare(page, mode, `<table>
            <tr><td style="text-transform:uppercase">hello <span data-xlsx-ignore>PRIVATE</span> world<br>second line</td></tr>
            <tr><td style="white-space:pre-line">A  B\nC<span hidden>PRIVATE</span></td></tr>
            <tr><td style="white-space:pre"> A  B <span hidden>PRIVATE</span></td></tr>
            <tr><td style="text-transform:capitalize"><span>he</span><span>llo world</span><span hidden>PRIVATE</span></td></tr></table>`);
        const result = await exported(page);
        expect(result.cells.map(cell => cell.value)).toEqual(['HELLO WORLD\nSECOND LINE', 'A B\nC', ' A  B ', 'Hello World']);
    });

    test('visible font styles, background layers and RTL alignment survive serialization' + suffix, async ({ page }) => {
        await prepare(page, mode, `<table style="background:red">
            <tr><td><b>VISIBLE</b><span data-xlsx-ignore style="font-size:100px;font-weight:normal">PRIVATE</span></td></tr>
            <tr><td><b>HIDDEN FONT</b><span style="display:none;font-size:100px;font-weight:normal">PRIVATE</span></td></tr>
            <tr><td><b><span style="font-weight:normal">NORMAL</span></b></td></tr>
            <tr style="background:lime"><td>ROW</td></tr>
            <tr><td style="background:yellow"><table><tr><td>PARENT</td></tr></table></td></tr>
            <tr><td><table style="background:blue"><tr><td>BLUE</td></tr></table></td></tr>
            <tr><td style="direction:rtl;text-align:start">RIGHT</td></tr>
            <tr><td style="direction:rtl;text-align:end">LEFT</td></tr></table>`);
        const result = await exported(page);
        const find = value => result.cells.find(cell => cell.value === value);
        for (const value of ['VISIBLE', 'HIDDEN FONT']) {
            expect(find(value).font.size).toBe(8);
            expect(find(value).font.bold).toBe(true);
        }
        expect(find('NORMAL').font.bold).not.toBe(true);
        for (const [value, color] of [['ROW', 'FF00FF00'], ['PARENT', 'FFFFFF00'], ['BLUE', 'FF0000FF']]) {
            expect(find(value).fill.fgColor.argb).toBe(color);
        }
        expect(find('RIGHT').alignment.horizontal).toBe('right');
        expect(find('LEFT').alignment.horizontal).toBe('left');
    });

    test('table roots retain their own border and background' + suffix, async ({ page }) => {
        await prepare(page, mode, '<table id="root-table" style="background:red;border:2px solid red"><tr><td>ROOT</td></tr></table>');
        const result = await exported(page, { root: '#root-table' });
        const cell = result.cells.find(cell => cell.value === 'ROOT');
        expect(cell.fill.fgColor.argb).toBe('FFFF0000');
        expect(Object.keys(cell.border)).toEqual(expect.arrayContaining(['top', 'bottom', 'left', 'right']));
    });

    test('nested images are collected, four inline images stay distinct and zero-size images are omitted' + suffix, async ({ page }) => {
        const img = '<img src="./stamp.svg" width="20" height="20">';
        await prepare(page, mode, `<table style="width:300px"><tr><td>${img}<table><tr><td>INNER</td></tr></table></td></tr>
            <tr><td style="height:20px;line-height:0" data-excel-image-align="inline">${img.repeat(4)}</td></tr>
            <tr><td>ZERO<img src="./stamp.svg" style="width:0;height:0"></td></tr></table>`);
        const result = await exported(page);
        expect(result.sheet.getImages()).toHaveLength(5);
        const positions = result.sheet.getImages().slice(1).map(image => imagePosition(image, result.sheet).x);
        expect(new Set(positions).size).toBe(4);
        expect(positions.every((position, index) => !index || position > positions[index - 1])).toBe(true);
        expect(result.warnings).toEqual([]);
    });

    test('explicit image alignment wins over default CSS while cell attributes remain authoritative' + suffix, async ({ page }) => {
        await prepare(page, mode, `<table style="width:300px"><tr><td style="height:100px"><img src="./stamp.svg" width="20" height="20"></td></tr>
            <tr><td style="height:100px" data-excel-image-align="left"><img src="./stamp.svg" width="20" height="20"></td></tr></table>`);
        const result = await exported(page, { imageAlignment: 'center' });
        const images = result.sheet.getImages();
        // Alignment sets the horizontal flow; vertically the image keeps its rendered top (top-aligned cell).
        expect(imagePosition(images[0], result.sheet)).toEqual({ x: 140, y: 0 });
        expect(imagePosition(images[1], result.sheet).x).toBe(0);
    });

    test('precise images map through unequal merged columns without fixed compensation' + suffix, async ({ page }) => {
        await prepare(page, mode, `<table style="width:300px;table-layout:fixed"><colgroup><col style="width:100px"><col style="width:200px"></colgroup>
            <tr><td>A</td><td>B</td></tr><tr><td colspan="2" style="height:100px;vertical-align:middle"><img src="./stamp.svg" width="20" height="20" style="margin-left:150px"></td></tr></table>`);
        const original = await page.locator('#export img').evaluate(image => {
            const rect = image.getBoundingClientRect();
            const cell = image.closest('td').getBoundingClientRect();
            return { left: rect.left - cell.left, top: rect.top - cell.top };
        });
        expect(original.left).toBe(150);
        expect(original.top).toBeGreaterThan(0);
        expect(original.top).toBeLessThan(50);
        const result = await exported(page, { imagePlacement: 'precise' });
        const image = result.sheet.getImages()[0];
        const position = imagePosition(image, result.sheet);
        expect(position.x).toBeCloseTo(original.left, 3);
        // Inline image baselines differ by browser; the DOM is the placement contract.
        const previousHeight = result.sheet.getRow(1).height / 0.75;
        expect(position.y - previousHeight).toBeCloseTo(original.top, 1);
    });

    test('localized signed amounts remain editable numbers in the workbook' + suffix, async ({ page }) => {
        for (const [locale, currency, html, values] of [
            ['en-US', 'USD', '<td>RMB 100</td><td>-$1,234.56</td><td data-excel-type="currency">($1,234.56)</td>', [100, -1234.56, -1234.56]],
            ['de-DE', 'EUR', '<td data-excel-type="currency">1.234,56 €</td>', [1234.56]],
            ['sv-SE', 'USD', '<td data-excel-type="number">−1 234,56</td>', [-1234.56]]
        ]) {
            await prepare(page, mode, '<table><tr>' + html + '</tr></table>');
            const result = await exported(page, { locale, currency });
            expect(result.cells.map(cell => cell.value)).toEqual(values);
        }
    });

    test('forced image flows that exceed cell capacity fail explicitly' + suffix, async ({ page }) => {
        await prepare(page, mode, `<table style="width:300px"><tr><td style="height:20px;line-height:0" data-excel-image-align="block">
            <img src="./stamp.svg" width="20" height="20"><img src="./stamp.svg" width="20" height="20"></td></tr></table>`);
        await expect(exported(page)).rejects.toThrow(/Image alignment exceeds/);
    });

    test('real CORS fallback embeds PNG pixels from the selected responsive source' + suffix, async ({ page }) => {
        await prepare(page, mode, `<table><tr><td><img src="${imageOrigin}/red.svg?signature=synthetic" width="20" height="20"></td>
            <td><picture><source srcset="${imageOrigin}/blue.svg"><img src="${imageOrigin}/red.svg" width="20" height="20"></picture></td></tr></table>`);
        const result = await exported(page);
        expect(result.stats.images).toBe(2);
        expect(result.sheet.getImages()).toHaveLength(2);
        expect(result.workbook.model.media.map(media => media.extension)).toEqual(['png', 'png']);
        const colors = await page.evaluate(async data => Promise.all(data.map(async base64 => {
            const image = new Image(); image.src = 'data:image/png;base64,' + base64; await image.decode();
            const canvas = document.createElement('canvas'); canvas.width = canvas.height = 20;
            const context = canvas.getContext('2d'); context.drawImage(image, 0, 0, 20, 20);
            return [...context.getImageData(10, 10, 1, 1).data];
        })), result.workbook.model.media.map(media => Buffer.from(media.buffer).toString('base64')));
        expect(colors).toEqual([[255, 0, 0, 255], [0, 0, 255, 255]]);
        expect(result.warnings.filter(warning => warning.code === 'IMAGE_CANVAS_FALLBACK')).toHaveLength(2);
        expect(result.warnings[0].source).toBe(imageOrigin + '/red.svg');
        expect(result.warnings[0].error).toEqual(expect.any(String));
        expect(JSON.stringify(result.warnings)).not.toContain('signature=');
    });

    test('throwing warning callbacks reject and restore the export button after font timeout' + suffix, async ({ page }) => {
        await prepare(page, mode, '<table><tr><td>FONT</td></tr></table>');
        const errors = [];
        page.on('pageerror', error => errors.push(error.message));
        const result = await page.evaluate(async () => {
            Object.defineProperty(document, 'fonts', { configurable: true, value: { ready: new Promise(() => {}) } });
            const button = document.createElement('button'); document.body.appendChild(button);
            HtmlToXlsx.bindButton(button, { root: '#export', download: false, imageTimeoutMs: 20,
                onWarning: () => { throw new Error('Observer failure'); } });
            const error = new Promise(resolve => button.addEventListener('xlsx:error', event => resolve(event.detail.message), { once: true }));
            button.click();
            const message = await Promise.race([error, new Promise(resolve => setTimeout(() => resolve('hung'), 1000))]);
            return { message, disabled: button.disabled, busy: button.dataset.busy || '' };
        });
        expect(result).toEqual({ message: 'Observer failure', disabled: false, busy: '' });
        expect(errors).toEqual([]);
    });
}

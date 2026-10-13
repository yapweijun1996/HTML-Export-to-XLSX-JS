import { test, expect } from '@playwright/test';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { SaxesParser } from 'saxes';
import fs from 'node:fs/promises';

function xmlTree(xml) {
    const stack = [];
    let root;
    const parser = new SaxesParser();
    parser.on('opentag', node => {
        const element = { name: node.name, attributes: node.attributes, children: [], text: '' };
        if (stack.length) stack.at(-1).children.push(element);
        else root = element;
        stack.push(element);
    });
    parser.on('text', text => { if (stack.length) stack.at(-1).text += text; });
    parser.on('closetag', () => stack.pop());
    parser.write(xml).close();
    return root;
}

function marker(anchor, name) {
    const element = anchor.children.find(child => child.name === name);
    expect(element).toBeDefined();
    return Object.fromEntries(element.children.map(child => [child.name.split(':').at(-1), Number(child.text)]));
}

// Read physical sizes from the package independently of the exporter and ExcelJS's
// fractional Anchor getters, which clamp large offsets to unrelated internal units.
function positionEmu(marker, sheet) {
    const widths = new Map();
    for (const column of sheet.children.find(child => child.name === 'cols').children) {
        const { min, max, width } = column.attributes;
        const w = Number(width);
        // OOXML Column.width conversion, including the Normal font's 7px digit width.
        const pixels = Math.floor((256 * w + Math.floor(128 / 7)) / 256 * 7);
        for (let c = Number(min); c <= Number(max); c++) widths.set(c - 1, pixels * 9525);
    }
    const heights = new Map(sheet.children.find(child => child.name === 'sheetData').children
        .map(row => [Number(row.attributes.r) - 1, Math.round(Number(row.attributes.ht) * 12700)]));
    let x = marker.colOff;
    let y = marker.rowOff;
    for (let c = 0; c < marker.col; c++) x += widths.get(c) ?? 64 * 9525;
    for (let r = 0; r < marker.row; r++) y += heights.get(r) ?? 15 * 12700;
    return { x, y };
}

for (const mode of ['', 'min']) {
    for (const imagePlacement of ['alignment', 'precise']) {
        test('valid drawing anchors retain image sizes and reuse identical media: ' + imagePlacement + ' ' + (mode || 'source'), async ({ page }) => {
            await page.goto('/tests/fixtures/quotation.html?mode=' + mode);
            await page.waitForFunction(() => Boolean(window.HtmlToXlsx));
            const results = await page.evaluate(async imagePlacement => {
                const png = color => {
                    const canvas = document.createElement('canvas');
                    canvas.width = 40; canvas.height = 20;
                    const context = canvas.getContext('2d');
                    context.fillStyle = color; context.fillRect(0, 0, 40, 20);
                    return canvas.toDataURL('image/png');
                };
                const red = png('red'), blue = png('blue'), green = png('green');
                const root = document.querySelector('#export');
                root.innerHTML = '<table style="width:300px;table-layout:fixed;border-collapse:collapse">' +
                    '<colgroup><col style="width:100px"><col style="width:200px"></colgroup>' +
                    [red, red, red, blue, green].map((src, i) =>
                        `<tr><td style="height:80px;padding:0;text-align:center" data-excel-image-align="center">` +
                        `<img src="${src}" width="${i === 4 ? 20 : 40}" height="20"></td>` +
                        `<td>$ ${[900, 1800, 3000, 3000, 8700][i]}.00</td></tr>`).join('') + '</table>';
                const before = root.innerHTML;
                const results = [];
                // Each workbook must register its own media, while repeated placements
                // within either workbook must share only byte-identical images.
                for (let i = 0; i < 2; i++) {
                    const result = await HtmlToXlsx.exportXlsx({ root, download: false, currency: 'HKD',
                        imagePlacement, imageScale: 1, imageWidthScale: 0.8, columnWidthScale: 1.7, rowScale: 1.4 });
                    results.push({ bytes: [...new Uint8Array(result.buffer)], stats: result.stats, warnings: result.warnings });
                }
                return { results, unchanged: root.innerHTML === before };
            }, imagePlacement);
            expect(results.unchanged).toBe(true);
            for (const [index, result] of results.results.entries()) {
                expect(result.warnings).toEqual([]);
                expect(result.stats.images).toBe(5);
                const bytes = Buffer.from(result.bytes);
                const zip = await JSZip.loadAsync(bytes);
                const media = Object.keys(zip.files).filter(name => /^xl\/media\/[^/]+\.png$/.test(name));
                expect(media).toHaveLength(3);
                const hashes = await Promise.all(media.map(name => zip.file(name).async('base64')));
                expect(new Set(hashes).size).toBe(3);
                const sheet = xmlTree(await zip.file('xl/worksheets/sheet1.xml').async('string'));
                const drawing = xmlTree(await zip.file('xl/drawings/drawing1.xml').async('string'));
                expect(drawing.children).toHaveLength(5);
                for (const [i, anchor] of drawing.children.entries()) {
                    expect(anchor.name).toBe('xdr:twoCellAnchor');
                    expect(anchor.attributes).toEqual({ editAs: 'oneCell' });
                    const from = marker(anchor, 'xdr:from'), to = marker(anchor, 'xdr:to');
                    for (const value of [...Object.values(from), ...Object.values(to)]) {
                        expect(Number.isSafeInteger(value)).toBe(true);
                        expect(value).toBeGreaterThanOrEqual(0);
                    }
                    const start = positionEmu(from, sheet), end = positionEmu(to, sheet);
                    expect(end.x - start.x).toBe(Math.round((i === 4 ? 16 : 32) * 9525));
                    expect(end.y - start.y).toBe(16 * 9525);
                }
                const relationships = xmlTree(await zip.file('xl/drawings/_rels/drawing1.xml.rels').async('string'));
                const targets = new Map(relationships.children.map(node => [node.attributes.Id, node.attributes.Target]));
                const imageTargets = drawing.children.map(anchor => {
                    const picture = anchor.children.find(node => node.name === 'xdr:pic');
                    const fill = picture.children.find(node => node.name === 'xdr:blipFill');
                    return targets.get(fill.children.find(node => node.name === 'a:blip').attributes['r:embed']);
                });
                expect(imageTargets[0]).toBe(imageTargets[1]);
                expect(imageTargets[0]).toBe(imageTargets[2]);
                expect(new Set(imageTargets).size).toBe(3);
                const workbook = new ExcelJS.Workbook();
                await workbook.xlsx.load(bytes);
                expect(workbook.worksheets[0].getImages()).toHaveLength(5);
                const amounts = [];
                workbook.worksheets[0].eachRow(row => row.eachCell(cell => {
                    if (!cell.isMerged && typeof cell.value === 'number') amounts.push(cell.value);
                }));
                expect(amounts).toEqual([900, 1800, 3000, 3000, 8700]);
                if (index === 0) await fs.writeFile(test.info().outputPath('images.xlsx'), bytes);
            }
        });
    }
}

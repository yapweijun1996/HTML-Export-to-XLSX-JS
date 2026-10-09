// MIT License. Copyright (c) 2026 YAP WEI JUN.
import { resolveOptions } from './config.js';
import { detectCurrency, escapeRegExp, parseCellValue } from './values.js';
import { createLayoutMetrics } from './layout-metrics.js';
import { createStyles } from './styles.js';
import { createDom } from './dom.js';
import { createLayout } from './layout.js';
import { createImages, mapWithConcurrency } from './images.js';
import { createWorkbookWriter } from './workbook.js';

// The build injects package.json's version; direct source imports identify themselves as development.
export const version = typeof __HTML_XLSX_VERSION__ === 'string' ? __HTML_XLSX_VERSION__ : 'development';

export function downloadBuffer(buffer, filename) {
    const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    try { link.click(); }
    finally {
        link.remove();
        setTimeout(() => URL.revokeObjectURL(url), 1000);
    }
}

function filenameFor(value) {
    const stamp = new Date().toISOString().replace(/[.:]/g, '-');
    const base = String(value || document.title || 'HTML export').replace(/[\\/?%*:|"<>]/g, '-').trim();
    return /\.xlsx$/i.test(base) ? base : base + '_' + stamp + '.xlsx';
}

export async function exportXlsx(options = {}) {
    options = { ...options };
    if (typeof document === 'undefined') throw new Error('exportXlsx requires a browser DOM');
    const root = typeof options.root === 'string'
        ? document.querySelector(options.root) : options.root || document.querySelector('[data-xlsx-export]');
    if (!root || typeof root.getBoundingClientRect !== 'function') throw new TypeError('Export root not found');
    const ExcelJS = options.exceljs || globalThis.ExcelJS;
    if (!ExcelJS?.Workbook) throw new Error('Load ExcelJS separately or supply options.exceljs');
    const { config: CONFIG, currency, locale } = resolveOptions(root, options);
    const warnings = [];
    const warnLog = message => {
        const warning = { code: 'EXPORT_WARNING', message: String(message) };
        warnings.push(warning);
        options.onWarning?.(warning);
    };
    const dependencies = {
        CONFIG, ExcelJS, escapeRegExp, warnLog, errorLog: warnLog,
        debugLog: (...args) => { if (CONFIG.DEBUG) console.debug('[html-to-xlsx]', ...args); },
        detectCurrency: raw => detectCurrency(raw, { currency, locale }),
        parseCellValue: (text, cell) => parseCellValue(text, cell, { currency, locale }),
        mapWithConcurrency
    };
    Object.assign(dependencies, createLayoutMetrics(dependencies));
    Object.assign(dependencies, createStyles(dependencies));
    Object.assign(dependencies, createDom(dependencies));
    Object.assign(dependencies, createLayout(dependencies));
    Object.assign(dependencies, createImages(dependencies));
    Object.assign(dependencies, createWorkbookWriter(dependencies));
    // Measure only after images and fonts settle; neither wait can block indefinitely.
    let fontTimer;
    if (document.fonts?.ready) {
        await Promise.race([
            document.fonts.ready,
            new Promise(resolve => { fontTimer = setTimeout(() => {
                warnLog('Font readiness timed out; using available font metrics');
                resolve();
            }, CONFIG.IMAGE_TIMEOUT_MS); })
        ]);
        clearTimeout(fontTimer);
    }
    await mapWithConcurrency(Array.from(root.querySelectorAll('img')), CONFIG.IMAGE_CONCURRENCY,
        async image => {
            if (dependencies.isElementBlacklisted(image, root)) return;
            try { await dependencies.ensureImageLoaded(image); }
            catch { warnLog('An image did not load before layout measurement'); }
        });
    const layoutModel = dependencies.buildLayoutModel(root);
    if (!layoutModel.cells.length) throw new Error('No visible table cells to export');
    const workbook = dependencies.createWorkbook();
    const worksheet = dependencies.prepareWorksheet(workbook, layoutModel);
    await dependencies.processAndPlaceImages(workbook, worksheet, layoutModel);
    const buffer = await dependencies.writeWorkbookToBuffer(workbook);
    const filename = filenameFor(options.filename);
    if (options.download !== false) downloadBuffer(buffer, filename);
    return {
        buffer, filename, warnings,
        stats: { cells: layoutModel.cells.length, images: worksheet.getImages().length }
    };
}

const bindings = new WeakMap();
export function bindButton(button, options = {}) {
    if (typeof button === 'string') button = document.querySelector(button);
    if (!button) throw new TypeError('Export button not found');
    bindings.get(button)?.();
    const handler = async event => {
        event.preventDefault();
        if (button.dataset.busy === '1') return;
        const previouslyDisabled = button.disabled;
        button.dataset.busy = '1';
        button.disabled = true;
        try {
            const result = await exportXlsx(options);
            button.dispatchEvent(new CustomEvent('xlsx:exported', { detail: result }));
        } catch (error) {
            button.dispatchEvent(new CustomEvent('xlsx:error', { detail: error }));
        } finally {
            button.disabled = previouslyDisabled;
            delete button.dataset.busy;
        }
    };
    button.addEventListener('click', handler);
    const unbind = () => {
        button.removeEventListener('click', handler);
        if (bindings.get(button) === unbind) bindings.delete(button);
    };
    bindings.set(button, unbind);
    return unbind;
}

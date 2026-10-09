// MIT License. Copyright (c) 2026 YAP WEI JUN.
import { validateCurrency } from './values.js';
export const DEFAULT_CONFIG = {
		DEBUG: false,
		ROOT_SELECTOR: '[data-xlsx-export]',
		BUTTON_SELECTOR: '[data-xlsx-button]',
		SHEET_NAME: 'PrintForm',
		MIN_COL_WIDTH_PX: 10,
		MIN_ROW_HEIGHT_PX: 1,
		// EDGE_TOLERANCE_PX clusters nearly-identical DOM edges (in CSS px) into one Excel grid line.
		// Default 2.0 prevents fragmented micro-columns/rows and duplicated inner borders.
		// Tuning range: 1.5–3.0 is typical; lower => more tiny fragments, higher => risk collapsing close columns/rows.
		// See README.md options and normalizeEdges() for details.
		EDGE_TOLERANCE_PX: 2,
		IMAGE_FORMAT: 'image/png',
		LOG_PREFIX: '[export]',
		CSS_DPI: 96,
		TARGET_DPI: 96,
		EXPORT_SCALE: 1,
		// FONT_SCALE keeps Excel text close to HTML size; 0.9 works well for 96 DPI layouts
		FONT_SCALE: 0.9,
		// ROW_SCALE controls row height scaling; falls back to EXPORT_SCALE when null
		ROW_SCALE: 1,
		// IMAGE_SCALE auto-upsamples to device DPI (>=1) so exports stay sharp
		IMAGE_SCALE: 3,
		// COLUMN_WIDTH_SCALE provides a final adjustment factor for column widths.
		// The system auto-calculates the width, but you can use this to fine-tune.
		// Example: 1.05 means "make all columns 5% wider". 0.95 means "5% narrower".
		COLUMN_WIDTH_SCALE: 1.0,
		// IMAGE_WIDTH_SCALE provides a final adjustment factor for all image sizes.
		// Example: 1.1 means "make all images 10% larger". 0.9 means "10% smaller".
		IMAGE_WIDTH_SCALE: 0.92,
		// IMAGE_OFFSET_* tweak image anchors inside cells to account for rendering differences.
		IMAGE_OFFSET_LEFT_PX: 30,
		IMAGE_OFFSET_TOP_PX: 10,
		// IMAGE_PLACEMENT_MODE controls how images are anchored.
		// 'precise' keeps the original pixel offsets, while 'alignment' applies HTML-like flows.
		IMAGE_PLACEMENT_MODE: 'alignment',
		// IMAGE_ALIGNMENT defines the default flow when IMAGE_PLACEMENT_MODE is 'alignment'.
		IMAGE_ALIGNMENT: 'inline', // inline | block | center | left | right | precise
		// IMAGE_ALIGNMENT_ATTR lets individual cells override the flow via a data attribute.
		IMAGE_ALIGNMENT_ATTR: 'data-excel-image-align',
		// IMAGE_SPACING_PX applies consistent gaps between auto-aligned images.
		IMAGE_SPACING_PX: 4,
		// IMAGE_MAX_PER_ROW limits how many images share one row when flowing inline.
		IMAGE_MAX_PER_ROW: 3,
		BLACKLIST_SELECTORS: ['[data-xlsx-ignore]'],
		PAGE_BREAK_SELECTOR: '[data-xlsx-page-break]',
		IMAGE_TIMEOUT_MS: 10000,
		IMAGE_CONCURRENCY: 4,
		PAGE_SETUP: {
			// ExcelJS pageSetup reference: https://github.com/exceljs/exceljs#worksheet-page-setup
			paperSize: 9, // A4 portrait constant (Excel uses numeric paper codes)
			orientation: 'portrait', // 'portrait' or 'landscape'
			fitToPage: true, // enable Excel auto-scaling so sheet fits the printable page
			scale: 100, // only used when fitToPage is false; retained for quick toggling
			fitToWidth: 1, // number of pages to fit horizontally when fitToPage is true
			fitToHeight: 0, // 0 keeps unlimited pages vertically; adjust if needed
			horizontalCentered: true, // center the sheet between left/right margins
			verticalCentered: false, // set true to center content vertically on the page
			margins: {
				left: 0.5, // inches
				right: 0.5, // inches
				top: 0.1, // inches before header text
				bottom: 0.1, // inches before footer text
				header: 0.3, // header gap in inches (add >=0.25 if top rows get clipped)
				footer: 0.3 // footer gap in inches
			}
		}
	};


const mappings = {
    edgeTolerancePx: 'EDGE_TOLERANCE_PX', exportScale: 'EXPORT_SCALE',
    fontScale: 'FONT_SCALE', rowScale: 'ROW_SCALE', imageScale: 'IMAGE_SCALE',
    columnWidthScale: 'COLUMN_WIDTH_SCALE', imageWidthScale: 'IMAGE_WIDTH_SCALE',
    imagePlacement: 'IMAGE_PLACEMENT_MODE', imageAlignment: 'IMAGE_ALIGNMENT',
    imageTimeoutMs: 'IMAGE_TIMEOUT_MS', imageConcurrency: 'IMAGE_CONCURRENCY',
    pageBreakSelector: 'PAGE_BREAK_SELECTOR', exclude: 'BLACKLIST_SELECTORS',
    debug: 'DEBUG', sheetName: 'SHEET_NAME'
};
const allowed = new Set([...Object.keys(mappings), 'root', 'filename', 'currency', 'locale',
    'exceljs', 'download', 'pageSetup', 'onWarning']);
const freezeDeep = object => {
    for (const value of Object.values(object)) {
        if (value && typeof value === 'object') freezeDeep(value);
    }
    return Object.freeze(object);
};
freezeDeep(DEFAULT_CONFIG);

export function resolveOptions(root, options = {}) {
    for (const key of Object.keys(options)) {
        if (!allowed.has(key)) throw new TypeError('Unknown export option: ' + key);
    }
    const currency = validateCurrency(options.currency ?? root.getAttribute('data-excel-currency'));
    const locale = options.locale ?? root.getAttribute('data-excel-locale') ?? 'en-US';
    new Intl.NumberFormat(locale);
    const config = { ...DEFAULT_CONFIG, BLACKLIST_SELECTORS: [...DEFAULT_CONFIG.BLACKLIST_SELECTORS] };
    for (const [key, internal] of Object.entries(mappings)) {
        if (options[key] !== undefined) config[internal] = options[key];
    }
    for (const [key, internal] of Object.entries(mappings)) {
        if (key.endsWith('Scale') || key.endsWith('Ms') || key === 'imageConcurrency' || key === 'edgeTolerancePx') {
            if (!(Number.isFinite(config[internal]) && config[internal] > 0)) {
                throw new RangeError(key + ' must be a positive finite number');
            }
        }
    }
    if (!Number.isInteger(config.IMAGE_CONCURRENCY)) throw new RangeError('imageConcurrency must be an integer');
    if (!['alignment', 'precise'].includes(config.IMAGE_PLACEMENT_MODE)) throw new RangeError('Invalid imagePlacement');
    if (!['inline', 'block', 'center', 'left', 'right', 'precise'].includes(config.IMAGE_ALIGNMENT)) throw new RangeError('Invalid imageAlignment');
    if (!Array.isArray(config.BLACKLIST_SELECTORS) || config.BLACKLIST_SELECTORS.some(value => typeof value !== 'string')) {
        throw new TypeError('exclude must be an array of CSS selectors');
    }
    for (const selector of [...config.BLACKLIST_SELECTORS, config.PAGE_BREAK_SELECTOR]) root.matches(selector);
    if (options.onWarning !== undefined && typeof options.onWarning !== 'function') throw new TypeError('onWarning must be a function');
    config.BLACKLIST_SELECTORS = [...config.BLACKLIST_SELECTORS];
    const page = structuredClone(options.pageSetup || {});
    config.PAGE_SETUP = { ...DEFAULT_CONFIG.PAGE_SETUP, ...page,
        margins: { ...DEFAULT_CONFIG.PAGE_SETUP.margins, ...page.margins } };
    return { config: freezeDeep(config), currency, locale };
}

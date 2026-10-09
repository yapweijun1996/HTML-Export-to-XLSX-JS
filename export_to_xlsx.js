/***************************************************************************************************************************
Version 5.0.1
File Description
No.   Date      By                Log
001.  20251010  WeiJun            Creation of file for Wan Tai(TSK-NA)
***************************************************************************************************************************/
/****  export_to_xlsx.js  ****/

(function () {
	const CONFIG = {
		DEBUG: false,
		ROOT_SELECTOR: '.exp_to_excel_button_content',
		BUTTON_SELECTOR: '.html_to_excel_btn',
		SHEET_NAME: 'PrintForm',
		MIN_COL_WIDTH_PX: 10,
		MIN_ROW_HEIGHT_PX: 1,
		// EDGE_TOLERANCE_PX clusters nearly-identical DOM edges (in CSS px) into one Excel grid line.
		// Default 2.0 prevents fragmented micro-columns/rows and duplicated inner borders.
		// Tuning range: 1.5–3.0 is typical; lower => more tiny fragments, higher => risk collapsing close columns/rows.
		// See README.md "Configuration: EDGE_TOLERANCE_PX" and normalizeEdges() for details.
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
		// BLACKLIST_SELECTORS defines CSS selectors whose elements (and descendants) are skipped during export.
		BLACKLIST_SELECTORS: ['table#mainapproval','table#ext_id','table#myMenu'],
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
	
	const debugLog = (...args) => {
		if (!CONFIG.DEBUG) return;
		console.log(CONFIG.LOG_PREFIX, ...args);
	};
	
	const warnLog = (...args) => {
		if (!CONFIG.DEBUG) {
			console.warn(...args);
			return;
		}
		console.warn(CONFIG.LOG_PREFIX, ...args);
	};
	
	const errorLog = (...args) => {
		if (!CONFIG.DEBUG) {
			console.error(...args);
			return;
		}
		console.error(CONFIG.LOG_PREFIX, ...args);
	};
	
	const POINTS_PER_CSS_PX = 72 / CONFIG.CSS_DPI;
	const toPositiveNumber = (value) => {
		const num = Number(value);
		return Number.isFinite(num) && num > 0 ? num : null;
	};
	
	const resolveExportScale = () => {
		const configured = toPositiveNumber(CONFIG.EXPORT_SCALE);
		if (configured) {
			return configured;
		}
		if (CONFIG.CSS_DPI > 0 && CONFIG.TARGET_DPI > 0) {
			return CONFIG.TARGET_DPI / CONFIG.CSS_DPI;
		}
		return 1;
	};
	const EXPORT_SCALE = resolveExportScale();
	
	const resolveFontScale = () => {
		const configured = toPositiveNumber(CONFIG.FONT_SCALE);
		if (configured) {
			return configured;
		}
		return EXPORT_SCALE;
	};
	const FONT_SCALE = resolveFontScale();
	
	const normalizedBlacklistSelectors = (() => {
		let cachedSelectors = null;
		return () => {
			if (cachedSelectors) return cachedSelectors;
			const raw = CONFIG.BLACKLIST_SELECTORS;
			let list = [];
			if (Array.isArray(raw)) {
				list = raw;
			} else if (typeof raw === 'string') {
				list = [raw];
			}
			cachedSelectors = list
			.map((item) => (typeof item === 'string' ? item.trim() : ''))
			.filter((item) => item.length > 0);
			return cachedSelectors;
		};
	})();
	
	const loggedBlacklistedElements = new WeakSet();
	
	const elementMatchesSelector = (element, selector) => {
		if (!element || !selector) return false;
		const matcher =
		element.matches ||
		element.matchesSelector ||
		element.msMatchesSelector ||
		element.webkitMatchesSelector ||
		element.mozMatchesSelector;
		if (typeof matcher !== 'function') return false;
		try {
			return matcher.call(element, selector);
		} catch (err) {
			if (CONFIG.DEBUG) {
				debugLog('Selector match failed', selector, err);
			}
			return false;
		}
	};
	
	const isElementBlacklisted = (element, root) => {
		if (!element || element.nodeType !== 1) return false;
		const selectors = normalizedBlacklistSelectors();
		if (!selectors.length) return false;
		let current = element;
		while (current && current.nodeType === 1) {
			for (const selector of selectors) {
				if (elementMatchesSelector(current, selector)) {
					if (CONFIG.DEBUG && !loggedBlacklistedElements.has(current)) {
						debugLog('Skipping blacklisted element', selector, current);
						loggedBlacklistedElements.add(current);
					}
					return true;
				}
			}
			if (root && current === root) {
				break;
			}
			current = current.parentElement;
		}
		return false;
	};
	
	const measureAvgCharWidth = (() => {
		let avgCharWidth = null;
		return (fontStyle = '9pt Arial') => {
			if (avgCharWidth !== null) {
				return avgCharWidth;
			}
			if (typeof document === 'undefined') {
				avgCharWidth = 7; // Fallback for non-browser environment
				return avgCharWidth;
			}
			try {
				const canvas = document.createElement('canvas');
				const context = canvas.getContext('2d');
				if (!context) {
					avgCharWidth = 7;
					return avgCharWidth;
				}
				context.font = fontStyle;
				// Using a wider character for a more stable average width
				const metrics = context.measureText('0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz');
				avgCharWidth = metrics.width / 62;
			} catch (e) {
				errorLog('Failed to measure avg char width, falling back to 7px', e);
				avgCharWidth = 7; // Fallback on error
			}
			debugLog(`Measured average character width for font "${fontStyle}":`, avgCharWidth);
			return avgCharWidth;
		};
	})();
	
	const resolveRowScale = () => {
		const configured = toPositiveNumber(CONFIG.ROW_SCALE);
		if (configured) {
			return configured;
		}
		return EXPORT_SCALE;
	};
	const ROW_SCALE = resolveRowScale();
	
	const resolveImageScale = () => {
		const configured = toPositiveNumber(CONFIG.IMAGE_SCALE);
		if (configured) {
			return configured;
		}
		return EXPORT_SCALE;
	};
	const IMAGE_SCALE = resolveImageScale();
	
	const COLOR_CANVAS_CTX = (function () {
		if (typeof document === 'undefined') return null;
		const canvas = document.createElement('canvas');
		return canvas && canvas.getContext ? canvas.getContext('2d') : null;
	})();
	
	const DEFAULT_EXCEL_PADDING_PX = 5;
	
	const clampToByte = (value) => {
		if (!isFiniteNumber(value)) return null;
		if (value <= 0) return 0;
		if (value >= 255) return 255;
		return Math.round(value);
	};
	
	const CURRENCY_DESCRIPTORS = [
		{ label: 'SGD', tokens: ['SGD', 'S$'], decimals: 2 },
		{ label: 'US$', tokens: ['US$', 'USD', '$'], decimals: 2, guard: (raw) => !/\bSGD\b/i.test(raw) && !/^S\$/i.test(raw) },
		{ label: 'RM', tokens: ['RM', 'MYR'], decimals: 2 },
		{ label: 'CNY', tokens: ['CNY', 'RMB', '¥'], decimals: 2, guard: (raw) => !/JPY|JAPAN/i.test(raw) },
		{ label: 'JPY', tokens: ['JPY', 'JP¥', 'Japan yen', 'Japan yan', '¥'], decimals: 0, guard: (raw) => /JPY|JAPAN/i.test(raw) },
		{ label: 'TWD', tokens: ['TWD', 'NT$', 'NTD'], decimals: 2 }
	];
	
	const escapeRegExp = (text) => text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
	
	const detectCurrency = (rawText) => {
		if (!rawText) return null;
		const original = String(rawText).trim();
		if (!original) return null;
		
		for (const descriptor of CURRENCY_DESCRIPTORS) {
			if (descriptor.guard && !descriptor.guard(original)) continue;
			const pattern = descriptor.tokens
			.slice()
			.sort((a, b) => b.length - a.length)
			.map(escapeRegExp)
			.join('|');
			if (!pattern) continue;
			const matcher = new RegExp(`^(${pattern})\\s*(.+)$`, 'i');
			const match = original.match(matcher);
			if (!match) continue;
			
			let numericText = match[2].trim();
			const isNegative = /^\(.*\)$/.test(numericText);
			if (isNegative) {
				numericText = numericText.slice(1, -1).trim();
			}
			numericText = numericText.replace(/[, ]/g, '');
			if (!/^\d+(\.\d+)?$/.test(numericText)) continue;
			
			const numericValue = Number.parseFloat(numericText);
			if (!Number.isFinite(numericValue)) continue;
			
			const decimals = Number.isInteger(descriptor.decimals) ? descriptor.decimals : 2;
			const fraction = decimals > 0 ? `.${'0'.repeat(decimals)}` : '';
			const numFmt = `"${descriptor.label}" #,##0${fraction};("${descriptor.label}" #,##0${fraction})`;
			
			return {
				value: isNegative ? -numericValue : numericValue,
				numFmt
			};
		}
		
		return null;
	};
	
	const toHexByte = (value) => {
		const byte = clampToByte(value);
		if (byte === null) return null;
		return byte.toString(16).padStart(2, '0').toUpperCase();
	};
	
	const parseHexColor = (value) => {
		if (!value || typeof value !== 'string') return null;
		let hex = value.trim();
		if (!hex) return null;
		if (hex[0] === '#') hex = hex.slice(1);
		if (![3, 4, 6, 8].includes(hex.length)) return null;
		const expand = (ch) => ch + ch;
		let r, g, b, a = 255;
		if (hex.length === 3 || hex.length === 4) {
			const chars = hex.split('');
			r = parseInt(expand(chars[0]), 16);
			g = parseInt(expand(chars[1]), 16);
			b = parseInt(expand(chars[2]), 16);
			if (hex.length === 4) {
				a = parseInt(expand(chars[3]), 16);
			}
		} else {
			r = parseInt(hex.slice(0, 2), 16);
			g = parseInt(hex.slice(2, 4), 16);
			b = parseInt(hex.slice(4, 6), 16);
			if (hex.length === 8) {
				a = parseInt(hex.slice(6, 8), 16);
			}
		}
		if ([r, g, b, a].some((component) => Number.isNaN(component))) return null;
		return { r, g, b, a: a / 255 };
	};
	
	const parseAlphaValue = (value) => {
		if (value == null) return 1;
		const trimmed = String(value).trim();
		if (!trimmed) return 1;
		if (trimmed.endsWith('%')) {
			const num = parseFloat(trimmed.slice(0, -1));
			if (!Number.isFinite(num)) return 1;
			return Math.min(1, Math.max(0, num / 100));
		}
		const num = parseFloat(trimmed);
		if (!Number.isFinite(num)) return 1;
		return Math.min(1, Math.max(0, num));
	};
	
	const parseRgbFunction = (value) => {
		if (!value || typeof value !== 'string') return null;
		const match = value.trim().match(/^rgba?\(\s*(.+)\s*\)$/i);
		if (!match) return null;
		const body = match[1];
		let alphaPart;
		let colorsPart = body;
		if (body.includes('/')) {
			const parts = body.split('/');
			colorsPart = parts[0];
			alphaPart = parts[1];
		}
		const colorTokens = colorsPart
		.split(/[\s,]+/)
		.map((token) => token.trim())
		.filter(Boolean);
		if (colorTokens.length < 3) return null;
		const parseComponent = (component) => {
			if (!component) return null;
			if (component.endsWith('%')) {
				const num = parseFloat(component.slice(0, -1));
				if (!Number.isFinite(num)) return null;
				return clampToByte((num / 100) * 255);
			}
			const num = parseFloat(component);
			if (!Number.isFinite(num)) return null;
			return clampToByte(num);
		};
		const r = parseComponent(colorTokens[0]);
		const g = parseComponent(colorTokens[1]);
		const b = parseComponent(colorTokens[2]);
		if ([r, g, b].some((component) => component === null)) return null;
		let alpha = 1;
		if (typeof alphaPart === 'string' && alphaPart.trim()) {
			alpha = parseAlphaValue(alphaPart);
		} else if (colorTokens.length >= 4) {
			alpha = parseAlphaValue(colorTokens[3]);
		}
		return { r, g, b, a: alpha };
	};
	
	const COLOR_CACHE = new Map();
	
	const parseCssColor = (value) => {
		if (!value) return null;
		const raw = String(value).trim();
		if (!raw) return null;
		if (raw.toLowerCase() === 'transparent') {
			return { r: 0, g: 0, b: 0, a: 0 };
		}
		const hex = parseHexColor(raw);
		if (hex) return hex;
		const rgb = parseRgbFunction(raw);
		if (rgb) return rgb;
		if (!COLOR_CANVAS_CTX) return null;
		try {
			COLOR_CANVAS_CTX.fillStyle = '#000000';
			COLOR_CANVAS_CTX.fillStyle = raw;
			const normalized = COLOR_CANVAS_CTX.fillStyle;
			if (typeof normalized !== 'string' || !normalized) {
				return null;
			}
			const viaHex = parseHexColor(normalized);
			if (viaHex) return viaHex;
			return parseRgbFunction(normalized);
		} catch (err) {
			return null;
		}
	};
	
	const colorToARGB = (value) => {
		const key = value || '';
		if (COLOR_CACHE.has(key)) {
			return COLOR_CACHE.get(key);
		}
		const parsed = parseCssColor(key);
		if (!parsed) {
			COLOR_CACHE.set(key, null);
			return null;
		}
		const alphaByte = clampToByte(parsed.a * 255);
		if (alphaByte === null || alphaByte <= 0) {
			COLOR_CACHE.set(key, null);
			return null;
		}
		const parts = [alphaByte, parsed.r, parsed.g, parsed.b]
		.map((component) => toHexByte(component));
		if (parts.some((component) => component === null)) {
			COLOR_CACHE.set(key, null);
			return null;
		}
		const argb = parts.join('');
		COLOR_CACHE.set(key, argb);
		return argb;
	};
	
	const widthToBorderStyle = (widthPx, cssStyle) => {
		if (!widthPx || widthPx <= 0.1) return null;
		const normalizedStyle = (cssStyle || '').toLowerCase();
		if (!normalizedStyle || normalizedStyle === 'none' || normalizedStyle === 'hidden') return null;
		
		const styleByWidth = () => {
			if (widthPx < 1.5) return 'thin';
			if (widthPx < 2.5) return 'medium';
			return 'thick';
		};
		
		switch (normalizedStyle) {
			case 'solid':
			return styleByWidth();
			case 'dashed':
			return widthPx < 2 ? 'dashed' : 'mediumDashed';
			case 'dotted':
			return 'dotted';
			case 'double':
			return 'double';
			case 'groove':
			case 'ridge':
			return widthPx < 2 ? 'medium' : 'double';
			case 'inset':
			case 'outset':
			return widthPx < 1.5 ? 'thin' : 'medium';
			case 'dash-dot':
			case 'dashdot':
			case 'dot-dash':
			return widthPx < 2 ? 'dashDot' : 'mediumDashDot';
			case 'dash-dot-dot':
			case 'dashdotdot':
			case 'dot-dot-dash':
			return widthPx < 2 ? 'dashDotDot' : 'mediumDashDotDot';
			default:
			return styleByWidth();
		}
	};
	
	const BORDER_STYLE_KEYWORDS = new Set([
		'none',
		'hidden',
		'dotted',
		'dashed',
		'solid',
		'double',
		'groove',
		'ridge',
		'inset',
		'outset',
		'dash-dot',
		'dashdot',
		'dot-dash',
		'dash-dot-dot',
		'dashdotdot',
		'dot-dot-dash'
	]);
	
	const parseBorderShorthand = (value) => {
		if (!value) return null;
		const tokens = value
		.split(/\s+/)
		.map((token) => token.trim().replace(/!important$/i, ''))
		.filter((token) => token && token.toLowerCase() !== '!important');
		if (!tokens.length) return null;
		
		let width = null;
		let style = null;
		const colorParts = [];
		
		for (const token of tokens) {
			if (width === null) {
				const widthCandidate = parseFloat(token);
				if (Number.isFinite(widthCandidate)) {
					width = widthCandidate;
					continue;
				}
			}
			
			if (style === null) {
				const lowerToken = token.toLowerCase();
				if (BORDER_STYLE_KEYWORDS.has(lowerToken)) {
					style = token;
					continue;
				}
			}
			
			colorParts.push(token);
		}
		
		return {
			width,
			style,
			color: colorParts.length ? colorParts.join(' ') : null
		};
	};
	
	const matchInlineStyleValue = (styleText, property) => {
		if (!styleText || !property) return null;
		const pattern = new RegExp(`(?:^|;)\\s*${escapeRegExp(property)}\\s*:\\s*([^;]+)`, 'i');
		const match = pattern.exec(styleText);
		if (!match) return null;
		const raw = match[1]?.trim().replace(/;$/, '');
		if (!raw) return null;
		return raw.replace(/\s*!important\s*$/i, '').trim() || null;
	};
	
	const parseInlineBorderFromAttribute = (element, side) => {
		if (!element || typeof element.getAttribute !== 'function') return null;
		const styleAttr = element.getAttribute('style');
		if (!styleAttr) return null;
		if (side) {
			const specific = matchInlineStyleValue(styleAttr, `border-${side}`);
			if (specific) {
				return parseBorderShorthand(specific);
			}
		}
		const generic = matchInlineStyleValue(styleAttr, 'border');
		return generic ? parseBorderShorthand(generic) : null;
	};
	const sanitizeSheetName = (name) => (name || 'Sheet1').replace(/[\\/*?:\[\]]/g, '').slice(0, 31) || 'Sheet1';
	
	const getFilenameBase = () => {
		const title = document.title || 'page';
		return title.replace(/[\\\/?%*:|"<>]/g, '-').trim() || 'page';
	};
	
	
	const pxToPoint = (px, scale = 1) => {
		const effectiveScale = toPositiveNumber(scale) || 1;
		const scaledPx = Math.max(0, px) * effectiveScale;
		return Math.round(scaledPx * POINTS_PER_CSS_PX * 100) / 100;
	};
	
	const pxToColumnWidth = (px, scale = 1) => {
		const targetPx = Math.max(CONFIG.MIN_COL_WIDTH_PX, px);
		const effectiveScale = toPositiveNumber(scale) || 1;
		const scaledPx = Math.max(CONFIG.MIN_COL_WIDTH_PX, targetPx * effectiveScale);
		const avgCharWidth = measureAvgCharWidth();
		const excelPadding = DEFAULT_EXCEL_PADDING_PX; // A reasonable assumption for default cell padding.
		
		// Base calculation using dynamic measurement
		const characterCount = (scaledPx - excelPadding) / avgCharWidth;
		
		// Apply the final manual adjustment factor
		const finalAdjustment = toPositiveNumber(CONFIG.COLUMN_WIDTH_SCALE) || 1.0;
		
		return Math.max(0, characterCount * finalAdjustment);
	};
	
	const pxToColumnWidthForImagePositioning = (px, scale = 1) => {
		// This is a clean version of pxToColumnWidth, WITHOUT the manual adjustment factor.
		// It's used ONLY for calculating the image position grid, so that image sizes
		// are not affected by the final column width tuning.
		const targetPx = Math.max(CONFIG.MIN_COL_WIDTH_PX, px);
		const effectiveScale = toPositiveNumber(scale) || 1;
		const scaledPx = Math.max(CONFIG.MIN_COL_WIDTH_PX, targetPx * effectiveScale);
		const avgCharWidth = measureAvgCharWidth();
		const excelPadding = DEFAULT_EXCEL_PADDING_PX;
		const characterCount = (scaledPx - excelPadding) / avgCharWidth;
		return Math.max(0, characterCount);
	};
	
	const columnWidthToPixels = (widthChars) => {
		if (!isFiniteNumber(widthChars) || widthChars <= 0) return 0;
		if (widthChars < 1) {
			return Math.floor(widthChars * 12);
		}
		return Math.floor(widthChars * 7 + 5);
	};
	
	const rowHeightToPixels = (heightPoints) => {
		if (!isFiniteNumber(heightPoints) || heightPoints <= 0) return 0;
		return Math.round((heightPoints * CONFIG.CSS_DPI) / 72);
	};
	
	const interpolateCoordinate = (valuePx, sourceEdges, targetEdges) => {
		if (!isFiniteNumber(valuePx) || valuePx <= 0) return 0;
		const lastIndex = Math.max(0, sourceEdges.length - 2);
		for (let i = 0; i <= lastIndex; i++) {
			const start = sourceEdges[i];
			const end = sourceEdges[i + 1];
			if (valuePx <= start) {
				return targetEdges[i] || 0;
			}
			if (valuePx > end && i < lastIndex) continue;
			const span = Math.max(end - start, 1e-4);
			const ratio = (valuePx - start) / span;
			const targetStart = targetEdges[i] || 0;
			const targetEnd = targetEdges[i + 1] || targetStart;
			return targetStart + ratio * (targetEnd - targetStart);
		}
		return targetEdges[targetEdges.length - 1] || 0;
	};
	
	const buildWorksheetEdgeMaps = (worksheet, layoutModel, forImagePositioning = false) => {
		const columnWidthFn = forImagePositioning ? pxToColumnWidthForImagePositioning : pxToColumnWidth;
		const columnEdgesPx = [0];
		for (let i = 0; i < layoutModel.columnWidthsPx.length; i++) {
			const column = worksheet.getColumn(i + 1);
			const widthChars =
			toPositiveNumber(column?.width) ||
			columnWidthFn(layoutModel.columnWidthsPx[i], EXPORT_SCALE);
			const widthPx = Math.max(CONFIG.MIN_COL_WIDTH_PX, columnWidthToPixels(widthChars));
			columnEdgesPx.push(columnEdgesPx[i] + widthPx);
			debugLog(`Column width map (forImage: ${forImagePositioning})`, {
				index: i,
				htmlPx: layoutModel.columnWidthsPx[i],
				excelWidthChars: widthChars,
				excelPx: widthPx
			});
		}
		
		const rowEdgesPx = [0];
		for (let i = 0; i < layoutModel.rowHeightsPx.length; i++) {
			const row = worksheet.getRow(i + 1);
			const heightPoints =
			toPositiveNumber(row?.height) ||
			pxToPoint(layoutModel.rowHeightsPx[i], ROW_SCALE);
			const heightPx = Math.max(CONFIG.MIN_ROW_HEIGHT_PX, rowHeightToPixels(heightPoints));
			rowEdgesPx.push(rowEdgesPx[i] + heightPx);
			debugLog('Row height map', {
				index: i,
				htmlPx: layoutModel.rowHeightsPx[i],
				excelPoints: heightPoints,
				excelPx: heightPx
			});
		}
		
		return { columnEdgesPx, rowEdgesPx };
	};
	const SNAP_PRECISION = 4;
	const snapToPrecision = (value) => {
		if (!isFiniteNumber(value)) return 0;
		const factor = 10 ** SNAP_PRECISION;
		const snapped = Math.round(value * factor) / factor;
		return snapped < 0 ? 0 : snapped;
	};
	
	const resolveSize = (...candidates) => {
		for (const candidate of candidates) {
			const snapped = snapToPrecision(candidate);
			if (snapped > 0) {
				return snapped;
			}
		}
		return 0;
	};
	
	const normalizeEdges = (edges) => {
		const sorted = Array.from(edges)
		.map((value) => snapToPrecision(value))
		.filter((value) => isFiniteNumber(value))
		.sort((a, b) => a - b);
		const result = [];
		for (const value of sorted) {
			if (!result.length) {
				result.push(value);
				continue;
			}
			const last = result[result.length - 1];
			if (value - last >= CONFIG.EDGE_TOLERANCE_PX) {
				result.push(value);
			} else if (value > last) {
				result[result.length - 1] = value;
			}
		}
		return result;
	};
	
	const normalizeLooseText = (text) => {
		if (!text) return '';
		return String(text)
		.replace(/\u00A0/g, ' ')
		.replace(/\t+/g, ' ')
		.replace(/\s+\n/g, '\n')
		.replace(/\n{3,}/g, '\n\n')
		.replace(/\s{2,}/g, ' ')
		.trim();
	};
	
	const collectTextFragmentsForCell = (cell, metrics, root) => {
		if (
		!cell ||
		typeof document === 'undefined' ||
		typeof window === 'undefined' ||
		typeof document.createTreeWalker !== 'function'
		) {
			return [];
		}
		
		if (isElementBlacklisted(cell, root)) {
			return [];
		}
		
		const ownerTable = cell.closest('table');
		if (!ownerTable) return [];
		
		const { rootLeft = 0, rootTop = 0, scrollX = 0, scrollY = 0 } = metrics || {};
		const accepted = typeof NodeFilter !== 'undefined' ? NodeFilter.FILTER_ACCEPT : 1;
		const rejected = typeof NodeFilter !== 'undefined' ? NodeFilter.FILTER_REJECT : 2;
		const walker = document.createTreeWalker(
		cell,
		typeof NodeFilter !== 'undefined' ? NodeFilter.SHOW_TEXT : 4,
		{
			acceptNode: (node) => {
				if (!node || !node.textContent) return rejected;
				const parentElement =
				node.parentElement ||
				(node.parentNode && node.parentNode.nodeType === 1 ? node.parentNode : null);
				if (parentElement) {
					if (isElementBlacklisted(parentElement, root)) {
						return rejected;
					}
					const nearestTable = parentElement.closest && parentElement.closest('table');
					if (nearestTable && nearestTable !== ownerTable) {
						return rejected;
					}
				}
				const normalized = normalizeLooseText(node.textContent);
				return normalized ? accepted : rejected;
			}
		},
		false
		);
		
		const fragments = [];
		if (!walker) return fragments;
		
		let current = walker.nextNode();
		while (current) {
			const fragmentText = normalizeLooseText(current.textContent);
			if (fragmentText) {
				const range = document.createRange();
				range.selectNodeContents(current);
				const rect = range.getBoundingClientRect();
				if (typeof range.detach === 'function') {
					range.detach();
				}
				if (rect && rect.width > 0 && rect.height > 0) {
					const sourceElement =
					current.parentElement ||
					(current.parentNode && current.parentNode.nodeType === 1 ? current.parentNode : cell);
					if (isElementBlacklisted(sourceElement, root)) {
						current = walker.nextNode();
						continue;
					}
					fragments.push({
						text: fragmentText,
						cell,
						sourceElement,
						left: snapToPrecision(rect.left + scrollX - rootLeft),
						right: snapToPrecision(rect.right + scrollX - rootLeft),
						top: snapToPrecision(rect.top + scrollY - rootTop),
						bottom: snapToPrecision(rect.bottom + scrollY - rootTop)
					});
				}
			}
			current = walker.nextNode();
		}
		
		return fragments;
	};
	
	const getCellText = (cell) => {
		const text = cell.innerText || '';
		const normalized = text
		.replace(/\u00A0/g, ' ')
		.replace(/\t+/g, ' ')
		.replace(/\s+\n/g, '\n')
		.replace(/\n{3,}/g, '\n\n');
		
		const hasInlineFloat = Array.from(cell.querySelectorAll('*')).some((child) => {
			if (!child || !child.style) return false;
			const floatValue = (child.style.cssFloat || child.style.float || '').toLowerCase();
			return floatValue && floatValue !== 'none';
		});
		
		if (hasInlineFloat) {
			return normalized.replace(/\s*\n\s*/g, ' ').replace(/\s{2,}/g, ' ').trim();
		}
		
		return normalized;
	};
	
	const extractBorders = (cs, element) => {
		const sides = ['top', 'right', 'bottom', 'left'];
		const border = {};
		for (const side of sides) {
			const width = parseFloat(cs.getPropertyValue(`border-${side}-width`)) || 0;
			const type = cs.getPropertyValue(`border-${side}-style`) || '';
			const excelStyle = widthToBorderStyle(width, type);
			if (!excelStyle) continue;
			const color = colorToARGB(cs.getPropertyValue(`border-${side}-color`));
			border[side] = {
				style: excelStyle,
				color: color ? { argb: color } : undefined
			};
		}
		
		if (element && element.style) {
			const capitalize = (text) => text.charAt(0).toUpperCase() + text.slice(1);
			for (const side of sides) {
				if (border[side]) continue;
				const widthStr = element.style[`border${capitalize(side)}Width`];
				const styleStr = element.style[`border${capitalize(side)}Style`];
				const colorStr = element.style[`border${capitalize(side)}Color`];
				const shorthand = element.style[`border${capitalize(side)}`];
				
				let width = Number.parseFloat(widthStr);
				width = Number.isFinite(width) ? width : null;
				let styleValue = styleStr || '';
				let colorValue = colorStr || '';
				
				if ((width == null || width <= 0) || !styleValue || !colorValue) {
					const parsed = parseBorderShorthand(shorthand);
					if (parsed) {
						if ((width == null || width <= 0) && Number.isFinite(parsed.width)) {
							width = parsed.width;
						}
						if (!styleValue && parsed.style) {
							styleValue = parsed.style;
						}
						if ((!colorValue || !colorValue.trim()) && parsed.color) {
							colorValue = parsed.color;
						}
					}
				}
				
				if ((width == null || width <= 0) || !styleValue || !colorValue) {
					const parsedAttr = parseInlineBorderFromAttribute(element, side);
					if (parsedAttr) {
						if ((width == null || width <= 0) && Number.isFinite(parsedAttr.width)) {
							width = parsedAttr.width;
						}
						if (!styleValue && parsedAttr.style) {
							styleValue = parsedAttr.style;
						}
						if ((!colorValue || !colorValue.trim()) && parsedAttr.color) {
							colorValue = parsedAttr.color;
						}
					}
				}
				
				const excelStyle = widthToBorderStyle(width, styleValue);
				if (!excelStyle) continue;
				const color = colorToARGB(colorValue);
				border[side] = {
					style: excelStyle,
					color: color ? { argb: color } : undefined
				};
			}
		}
		
		return Object.keys(border).length ? border : null;
	};
	
	const isFiniteNumber = (value) => (typeof Number.isFinite === 'function' ? Number.isFinite(value) : isFinite(value));
	
	const normalizeFontWeight = (value) => {
		if (typeof value === "number") {
			return isFiniteNumber(value) ? value : 400;
		}
		const str = String(value || '').trim().toLowerCase();
		if (!str) return 400;
		if (str === 'bold' || str === 'bolder') return 700;
		if (str === 'normal') return 400;
		if (str === 'lighter') return 300;
		const parsed = parseInt(str, 10);
		return isFiniteNumber(parsed) ? parsed : 400;
	};
	
	const isCellUniformlyBold = (cell, baseWeight) => {
		if (!cell || typeof document === 'undefined' || typeof window === 'undefined') return false;
		if (typeof document.createTreeWalker !== 'function') return false;
		const showText = typeof NodeFilter !== 'undefined' ? NodeFilter.SHOW_TEXT : 4;
		let walker;
		try {
			walker = document.createTreeWalker(cell, showText, null);
		} catch (err) {
			return false;
		}
		let sawText = false;
		while (walker.nextNode()) {
			const textNode = walker.currentNode;
			if (!textNode || !textNode.textContent || !textNode.textContent.trim()) continue;
			sawText = true;
			let effectiveWeight = baseWeight;
			let ancestor = textNode.parentElement;
			while (ancestor && ancestor !== cell) {
				const ancestorStyle = window.getComputedStyle(ancestor);
				const ancestorWeight = normalizeFontWeight(ancestorStyle.fontWeight);
				if (ancestorWeight >= 600) {
					effectiveWeight = ancestorWeight;
					break;
				}
				ancestor = ancestor.parentElement;
			}
			if (effectiveWeight < 600) {
				return false;
			}
		}
		return sawText;
	};
	
	const resolveFontSizePx = (cell, cellStyle) => {
		const baseSize = parseFloat(cellStyle.fontSize) || 0;
		if (!cell || typeof document === 'undefined' || typeof window === 'undefined') return baseSize;
		const ELEMENT_NODE = typeof Node !== 'undefined' ? Node.ELEMENT_NODE : 1;
		let walker;
		try {
			const showElements = typeof NodeFilter !== 'undefined' ? NodeFilter.SHOW_ELEMENT : ELEMENT_NODE;
			walker = document.createTreeWalker(cell, showElements, null);
		} catch (err) {
			return baseSize;
		}
		let maxSize = baseSize;
		while (walker.nextNode()) {
			const node = walker.currentNode;
			if (!node || (node.nodeType !== ELEMENT_NODE && node.nodeType !== 1)) continue;
			const nodeStyle = window.getComputedStyle(node);
			const childSize = parseFloat(nodeStyle.fontSize) || 0;
			if (childSize > maxSize) {
				maxSize = childSize;
			}
		}
		return maxSize;
	};
	
	const extractFont = (cell, cs) => {
		const families = (cs.fontFamily || '').split(',').map((v) => v.trim().replace(/^"|"$/g, ''));
		const name = families[0] || 'Arial';
		const sizePx = resolveFontSizePx(cell, cs);
		const weight = normalizeFontWeight(cs.fontWeight);
		let bold = weight >= 600;
		if (!bold) {
			try {
				bold = isCellUniformlyBold(cell, weight);
			} catch (err) {
				bold = false;
			}
		}
		
		// 改进字体大小转换逻辑，确保最小字体大小
		let excelFontSize;
		if (sizePx > 0) {
			const pointSize = pxToPoint(sizePx, FONT_SCALE);
			// 确保字体大小至少为8pt（Excel最小值）
			excelFontSize = Math.max(8, Math.round(pointSize * 100) / 100);
		}
		
		const font = {
			name,
			size: excelFontSize,
			bold,
			italic: cs.fontStyle === 'italic'
		};
		const color = colorToARGB(cs.color);
		if (color) font.color = { argb: color };
		
		// 添加字体调试信息
		if (CONFIG.DEBUG && sizePx > 0) {
			debugLog('Font extraction', {
				element: cell?.tagName || 'unknown',
				originalSize: `${sizePx}px`,
				excelSize: excelFontSize,
				fontFamily: name,
				weight,
				bold
			});
		}
		
		return font;
	};
	
	const parseCssPixelValue = (value) => {
		if (value === null || value === undefined) return 0;
		const parsed = parseFloat(value);
		return Number.isFinite(parsed) ? parsed : 0;
	};
	
	const computeEffectiveHorizontalPadding = (cell, cellStyle) => {
		const baseLeft = parseCssPixelValue(cellStyle?.paddingLeft);
		const baseRight = parseCssPixelValue(cellStyle?.paddingRight);
		
		if (!cell || typeof document === 'undefined' || typeof window === 'undefined') {
			return { left: baseLeft, right: baseRight };
		}
		
		let minPositiveExtraLeft = Infinity;
		let minPositiveExtraRight = Infinity;
		
		try {
			const showText = typeof NodeFilter !== 'undefined' ? NodeFilter.SHOW_TEXT : 4;
			const walker = document.createTreeWalker(cell, showText, null);
			while (walker.nextNode()) {
				const node = walker.currentNode;
				if (!node || !node.textContent || !node.textContent.trim()) continue;
				
				let current = node.parentElement;
				let extraLeft = 0;
				let extraRight = 0;
				while (current && current !== cell) {
					const style = window.getComputedStyle(current);
					extraLeft +=
					parseCssPixelValue(style.paddingLeft) +
					parseCssPixelValue(style.marginLeft) +
					parseCssPixelValue(style.borderLeftWidth);
					extraRight +=
					parseCssPixelValue(style.paddingRight) +
					parseCssPixelValue(style.marginRight) +
					parseCssPixelValue(style.borderRightWidth);
					current = current.parentElement;
				}
				if (extraLeft > 0.5) {
					minPositiveExtraLeft = Math.min(minPositiveExtraLeft, extraLeft);
				}
				if (extraRight > 0.5) {
					minPositiveExtraRight = Math.min(minPositiveExtraRight, extraRight);
				}
			}
		} catch (err) {
			if (CONFIG.DEBUG) {
				debugLog('Failed to compute extra padding from descendants', err);
			}
		}
		
		const additionalLeft = Number.isFinite(minPositiveExtraLeft) ? minPositiveExtraLeft : 0;
		const additionalRight = Number.isFinite(minPositiveExtraRight) ? minPositiveExtraRight : 0;
		
		return {
			left: baseLeft + additionalLeft,
			right: baseRight + additionalRight
		};
	};
	
	const computeEffectiveVerticalPadding = (cell, cellStyle) => {
		const baseTop = parseCssPixelValue(cellStyle?.paddingTop);
		const baseBottom = parseCssPixelValue(cellStyle?.paddingBottom);
		
		if (!cell || typeof document === 'undefined' || typeof window === 'undefined') {
			return { top: baseTop, bottom: baseBottom };
		}
		
		let minPositiveExtraTop = Infinity;
		let minPositiveExtraBottom = Infinity;
		
		try {
			const showText = typeof NodeFilter !== 'undefined' ? NodeFilter.SHOW_TEXT : 4;
			const walker = document.createTreeWalker(cell, showText, null);
			while (walker.nextNode()) {
				const node = walker.currentNode;
				if (!node || !node.textContent || !node.textContent.trim()) continue;
				
				let current = node.parentElement;
				let extraTop = 0;
				let extraBottom = 0;
				while (current && current !== cell) {
					const style = window.getComputedStyle(current);
					extraTop +=
					parseCssPixelValue(style.paddingTop) +
					parseCssPixelValue(style.marginTop) +
					parseCssPixelValue(style.borderTopWidth);
					extraBottom +=
					parseCssPixelValue(style.paddingBottom) +
					parseCssPixelValue(style.marginBottom) +
					parseCssPixelValue(style.borderBottomWidth);
					current = current.parentElement;
				}
				if (extraTop > 0.5) {
					minPositiveExtraTop = Math.min(minPositiveExtraTop, extraTop);
				}
				if (extraBottom > 0.5) {
					minPositiveExtraBottom = Math.min(minPositiveExtraBottom, extraBottom);
				}
			}
		} catch (err) {
			if (CONFIG.DEBUG) {
				debugLog('Failed to compute extra vertical padding from descendants', err);
			}
		}
		
		const additionalTop = Number.isFinite(minPositiveExtraTop) ? minPositiveExtraTop : 0;
		const additionalBottom = Number.isFinite(minPositiveExtraBottom) ? minPositiveExtraBottom : 0;
		
		return {
			top: baseTop + additionalTop,
			bottom: baseBottom + additionalBottom
		};
	};
	
	const convertPaddingPxToIndent = (paddingPx) => {
		if (!isFiniteNumber(paddingPx) || paddingPx <= 0.5) return 0;
		const avgCharWidth = measureAvgCharWidth();
		if (!isFiniteNumber(avgCharWidth) || avgCharWidth <= 0) return 0;
		const indentUnits = paddingPx / avgCharWidth;
		if (indentUnits < 0.5) return 0;
		return Math.min(30, Math.round(indentUnits));
	};
	
	const extractAlignment = (cell, cs, textValue) => {
		const normalize = (value) => {
			if (!value) return undefined;
			const normalized = String(value).trim().toLowerCase();
			return normalized || undefined;
		};
		
		const horizontalMap = {
			start: 'left',
			left: 'left',
			end: 'right',
			right: 'right',
			center: 'center',
			justify: 'justify'
		};
		const verticalMap = {
			top: 'top',
			middle: 'middle',
			center: 'middle',
			bottom: 'bottom',
			baseline: 'bottom'
		};
		
		const cssHorizontal = normalize(cs.textAlign);
		const cssVertical = normalize(cs.verticalAlign);
		const attrHorizontal = normalize(cell.getAttribute('align') || cell.align);
		const attrVertical = normalize(cell.getAttribute('valign') || cell.vAlign);
		
		const horizontal = horizontalMap[cssHorizontal] || horizontalMap[attrHorizontal] || undefined;
		const vertical = verticalMap[cssVertical] || verticalMap[attrVertical] || undefined;
		const wrapText = cs.whiteSpace !== 'nowrap' || (textValue && textValue.includes('\n'));
		
		const alignment = {
			horizontal,
			vertical,
			wrapText
		};
		
		const padding = computeEffectiveHorizontalPadding(cell, cs);
		const extraLeft = Math.max(0, padding.left - DEFAULT_EXCEL_PADDING_PX);
		const extraRight = Math.max(0, padding.right - DEFAULT_EXCEL_PADDING_PX);
		
		const preferLeft =
		!alignment.horizontal || alignment.horizontal === 'left' || alignment.horizontal === 'justify';
		if (preferLeft) {
			const indent = convertPaddingPxToIndent(extraLeft);
			if (indent > 0) {
				alignment.indent = indent;
			}
		} else if (alignment.horizontal === 'right') {
			const indent = convertPaddingPxToIndent(extraRight);
			if (indent > 0) {
				alignment.indent = indent;
			}
		}
		
		return alignment;
	};
	
	const buildLayoutModel = (root) => {
		const rootRect = root.getBoundingClientRect();
		const scrollY = window.scrollY || window.pageYOffset || 0;
		const scrollX = window.scrollX || window.pageXOffset || 0;
		
		const rootLeft = snapToPrecision(rootRect.left + scrollX);
		const rootTop = snapToPrecision(rootRect.top + scrollY);
		const rootWidth = snapToPrecision(root.scrollWidth);
		const rootHeight = snapToPrecision(root.scrollHeight);
		
		const allCells = Array.from(root.querySelectorAll('td, th')).filter(
		(cell) => !isElementBlacklisted(cell, root)
		);
		const nestedCells = [];
		const leafCells = [];
		for (const cell of allCells) {
			if (cell.querySelector('table')) {
				nestedCells.push(cell);
			} else {
				leafCells.push(cell);
			}
		}
		const tableElements = Array.from(root.querySelectorAll('table')).filter(
		(tableElement) => !isElementBlacklisted(tableElement, root)
		);
		const textFragments = [];
		const pageBreakMarkers = Array.from(root.querySelectorAll('.pagebreak_bf_processed')).filter(
		(marker) => !isElementBlacklisted(marker, root)
		);
		const colEdgeSet = new Set([0, rootWidth]);
		const rowEdgeSet = new Set([0, rootHeight]);
		const fragmentMetrics = { rootLeft, rootTop, scrollX, scrollY };
		
		for (const cell of allCells) {
			const rect = cell.getBoundingClientRect();
			const width = snapToPrecision(rect.width);
			const height = snapToPrecision(rect.height);
			if (width <= 0 || height <= 0) continue;
			const left = snapToPrecision(rect.left + scrollX - rootLeft);
			const right = snapToPrecision(rect.right + scrollX - rootLeft);
			const top = snapToPrecision(rect.top + scrollY - rootTop);
			const bottom = snapToPrecision(rect.bottom + scrollY - rootTop);
			colEdgeSet.add(left);
			colEdgeSet.add(right);
			rowEdgeSet.add(top);
			rowEdgeSet.add(bottom);
		}
		
		for (const tableElement of tableElements) {
			if (!tableElement || tableElement === root) continue;
			const rect = tableElement.getBoundingClientRect();
			const width = snapToPrecision(rect.width);
			const height = snapToPrecision(rect.height);
			if (width <= 0 || height <= 0) continue;
			const left = snapToPrecision(rect.left + scrollX - rootLeft);
			const right = snapToPrecision(rect.right + scrollX - rootLeft);
			const top = snapToPrecision(rect.top + scrollY - rootTop);
			const bottom = snapToPrecision(rect.bottom + scrollY - rootTop);
			colEdgeSet.add(left);
			colEdgeSet.add(right);
			rowEdgeSet.add(top);
			rowEdgeSet.add(bottom);
		}
		
		for (const cell of nestedCells) {
			const fragments = collectTextFragmentsForCell(cell, fragmentMetrics, root);
			if (!Array.isArray(fragments) || !fragments.length) continue;
			for (const fragment of fragments) {
				if (!fragment) continue;
				textFragments.push(fragment);
				if (Number.isFinite(fragment.left)) {
					colEdgeSet.add(fragment.left);
				}
				if (Number.isFinite(fragment.right)) {
					colEdgeSet.add(fragment.right);
				}
				if (Number.isFinite(fragment.top)) {
					rowEdgeSet.add(fragment.top);
				}
				if (Number.isFinite(fragment.bottom)) {
					rowEdgeSet.add(fragment.bottom);
				}
			}
		}
		
		for (const marker of pageBreakMarkers) {
			if (!marker || typeof marker.getBoundingClientRect !== 'function') continue;
			const rect = marker.getBoundingClientRect();
			const top = snapToPrecision(rect.top + scrollY - rootTop);
			const bottom = snapToPrecision(rect.bottom + scrollY - rootTop);
			if (Number.isFinite(top)) {
				rowEdgeSet.add(Math.max(0, top));
			}
			if (Number.isFinite(bottom)) {
				rowEdgeSet.add(Math.max(0, bottom));
			}
		}
		
		const columnEdges = normalizeEdges(colEdgeSet);
		const rowEdges = normalizeEdges(rowEdgeSet);
		
		const describeCell = (cell) => {
			const rect = cell.getBoundingClientRect();
			const scrollY = window.scrollY || window.pageYOffset || 0;
			const scrollX = window.scrollX || window.pageXOffset || 0;
			const rootLeft = snapToPrecision(root.getBoundingClientRect().left + scrollX);
			const rootTop = snapToPrecision(root.getBoundingClientRect().top + scrollY);
			
			const width = snapToPrecision(rect.width);
			const height = snapToPrecision(rect.height);
			if (width <= 0 || height <= 0) return null;
			const left = snapToPrecision(rect.left + scrollX - rootLeft);
			const right = snapToPrecision(rect.right + scrollX - rootLeft);
			const top = snapToPrecision(rect.top + scrollY - rootTop);
			const bottom = snapToPrecision(rect.bottom + scrollY - rootTop);
			
			const colStartIndex = columnEdges.findIndex((edge) => Math.abs(edge - left) <= CONFIG.EDGE_TOLERANCE_PX);
			const colEndIndex = columnEdges.findIndex((edge) => Math.abs(edge - right) <= CONFIG.EDGE_TOLERANCE_PX);
			const rowStartIndex = rowEdges.findIndex((edge) => Math.abs(edge - top) <= CONFIG.EDGE_TOLERANCE_PX);
			const rowEndIndex = rowEdges.findIndex((edge) => Math.abs(edge - bottom) <= CONFIG.EDGE_TOLERANCE_PX);
			
			if (colStartIndex === -1 || colEndIndex === -1 || rowStartIndex === -1 || rowEndIndex === -1) {
				console.warn('Skipping cell due to unmatched edges', cell);
				return null;
			}
			
			return {
				rect,
				width,
				height,
				left,
				top,
				colStartIndex,
				colEndIndex,
				rowStartIndex,
				rowEndIndex
			};
		};
		
		const resolveEdgeIndex = (edges, value, fallbackIndex) => {
			if (Array.isArray(edges) && edges.length >= 2 && Number.isFinite(value)) {
				const tolerance = Math.max(CONFIG.EDGE_TOLERANCE_PX, 0.5);
				for (let i = 0; i < edges.length; i++) {
					if (Math.abs(edges[i] - value) <= tolerance) {
						return i;
					}
				}
				for (let i = 1; i < edges.length; i++) {
					if (value < edges[i]) {
						return i - 1;
					}
				}
				return edges.length - 1;
			}
			return Number.isInteger(fallbackIndex) ? fallbackIndex : 0;
		};
		
		const descriptorCache = new WeakMap();
		const getDescriptor = (element) => {
			if (!element) return null;
			let cached = descriptorCache.get(element);
			if (!cached) {
				cached = describeCell(element);
				if (cached) {
					descriptorCache.set(element, cached);
				}
			}
			return cached;
		};
		
		const cells = [];
		const images = [];
		const deferredBorders = [];
		const tableBackgroundRegions = [];
		
		for (const cell of leafCells) {
			const descriptor = describeCell(cell);
			if (!descriptor) continue;
			
			const {
				rect,
				width,
				height,
				left,
				top,
				colStartIndex,
				colEndIndex,
				rowStartIndex,
				rowEndIndex
			} = descriptor;
			
			const textValue = getCellText(cell);
			const cs = window.getComputedStyle(cell);
			const font = extractFont(cell, cs);
			const alignment = extractAlignment(cell, cs, textValue);
			const borders = extractBorders(cs, cell);
			const fillColor = colorToARGB(cs.backgroundColor);
			const currencyMatch = detectCurrency(textValue);
			const exportedValue = currencyMatch ? currencyMatch.value : (textValue || '');
			
			const cellModel = {
				element: cell,
				value: exportedValue,
				colStart: colStartIndex,
				colEnd: Math.max(colStartIndex + 1, colEndIndex),
				rowStart: rowStartIndex,
				rowEnd: Math.max(rowStartIndex + 1, rowEndIndex),
				bounds: {
					left,
					top,
					width,
					height
				},
				styles: {
					font,
					alignment,
					border: borders,
					fill: fillColor ? { type: 'pattern', pattern: 'solid', fgColor: { argb: fillColor } } : null,
					numFmt: currencyMatch ? currencyMatch.numFmt : undefined
				},
				images: [],
				sheetPosition: null
			};
			
			const imgList = Array.from(cell.querySelectorAll('img')).filter(
			(img) => !isElementBlacklisted(img, root)
			);
			for (const img of imgList) {
				const imgRect = img.getBoundingClientRect();
				const scrollY = window.scrollY || window.pageYOffset || 0;
				const scrollX = window.scrollX || window.pageXOffset || 0;
				const rootLeft = snapToPrecision(root.getBoundingClientRect().left + scrollX);
				const rootTop = snapToPrecision(root.getBoundingClientRect().top + scrollY);
				
				const imgWidth = resolveSize(imgRect.width, img.width, img.naturalWidth);
				const imgHeight = resolveSize(imgRect.height, img.height, img.naturalHeight);
				if (!imgWidth || !imgHeight) continue;
				const job = {
					element: img,
					width: imgWidth,
					height: imgHeight,
					parentCell: cellModel,
					left: snapToPrecision(imgRect.left + scrollX - rootLeft),
					top: snapToPrecision(imgRect.top + scrollY - rootTop),
					offsetWithinCell: {
						left: snapToPrecision(imgRect.left - rect.left),
						top: snapToPrecision(imgRect.top - rect.top)
					}
				};
				cellModel.images.push(job);
				images.push(job);
				debugLog('Queued image job', img?.src || '(unknown src)', {
					width: job.width,
					height: job.height,
					offsetWithinCell: job.offsetWithinCell,
					parentCellBounds: {
						colStart: cellModel.colStart,
						colEnd: cellModel.colEnd,
						rowStart: cellModel.rowStart,
						rowEnd: cellModel.rowEnd
					}
				});
			}
			
			cells.push(cellModel);
		}
		
		if (textFragments.length) {
			for (const fragment of textFragments) {
				if (!fragment || !fragment.cell || !fragment.text) continue;
				const descriptor = getDescriptor(fragment.cell);
				if (!descriptor) continue;
				const width = snapToPrecision(fragment.right - fragment.left);
				const height = snapToPrecision(fragment.bottom - fragment.top);
				if (width <= 0 || height <= 0) continue;
				
				const colStartIndex = resolveEdgeIndex(columnEdges, fragment.left, descriptor.colStartIndex);
				const colEndIndexRaw = resolveEdgeIndex(columnEdges, fragment.right, descriptor.colEndIndex);
				const rowStartIndex = resolveEdgeIndex(rowEdges, fragment.top, descriptor.rowStartIndex);
				const rowEndIndexRaw = resolveEdgeIndex(rowEdges, fragment.bottom, descriptor.rowEndIndex);
				const colEndIndex = Math.max(colStartIndex + 1, colEndIndexRaw);
				const rowEndIndex = Math.max(rowStartIndex + 1, rowEndIndexRaw);
				
				const duplicate = cells.some(
				(existing) => existing.colStart === colStartIndex && existing.rowStart === rowStartIndex
				);
				if (duplicate) continue;
				
				const referenceElement = fragment.sourceElement || fragment.cell;
				const cs = window.getComputedStyle(referenceElement);
				const font = extractFont(referenceElement, cs);
				const alignment = extractAlignment(referenceElement, cs, fragment.text);
				const fillColor = colorToARGB(cs.backgroundColor);
				const currencyMatch = detectCurrency(fragment.text);
				const exportedValue = currencyMatch ? currencyMatch.value : fragment.text;
				
				cells.push({
					element: fragment.cell,
					value: exportedValue,
					colStart: colStartIndex,
					colEnd: colEndIndex,
					rowStart: rowStartIndex,
					rowEnd: rowEndIndex,
					bounds: {
						left: fragment.left,
						top: fragment.top,
						width,
						height
					},
					styles: {
						font,
						alignment,
						border: null,
						fill: fillColor ? { type: 'pattern', pattern: 'solid', fgColor: { argb: fillColor } } : null,
						numFmt: currencyMatch ? currencyMatch.numFmt : undefined
					},
					images: [],
					sheetPosition: null
				});
			}
		}
		
		const computedStyleCache = new WeakMap();
		const getCachedComputedStyle = (element) => {
			if (!element || typeof window === 'undefined' || typeof window.getComputedStyle !== 'function') {
				return null;
			}
			let cached = computedStyleCache.get(element);
			if (!cached) {
				cached = window.getComputedStyle(element);
				computedStyleCache.set(element, cached);
			}
			return cached;
		};
		
		const pushDeferredBorderInstruction = (descriptor, side, style) => {
			if (!descriptor || !style) return;
			deferredBorders.push({
				side,
				style,
				rowStart: descriptor.rowStartIndex + 1,
				rowEnd: Math.max(descriptor.rowStartIndex + 1, descriptor.rowEndIndex),
				colStart: descriptor.colStartIndex + 1,
				colEnd: Math.max(descriptor.colStartIndex + 1, descriptor.colEndIndex)
			});
		};
		
		const mergeBorderStyle = (target, side, borderStyle) => {
			if (!borderStyle) return;
			target.styles.border = target.styles.border || {};
			// Only apply the border if the target cell doesn't already have one on this side.
			// This prevents outer table borders from overwriting inner cell borders.
			if (!target.styles.border[side]) {
				target.styles.border[side] = { ...borderStyle };
			}
		};
		
		const propagateSharedBorders = (cells) => {
			const tolerance = Math.max(CONFIG.EDGE_TOLERANCE_PX, 1);
			for (const cell of cells) {
				const border = cell.styles?.border;
				if (!border) continue;
				const cellLeft = cell.bounds.left;
				const cellRight = cell.bounds.left + cell.bounds.width;
				const cellTop = cell.bounds.top;
				const cellBottom = cell.bounds.top + cell.bounds.height;
				
				if (border.right) {
					for (const neighbor of cells) {
						if (neighbor === cell) continue;
						const neighborLeft = neighbor.bounds.left;
						const neighborTop = neighbor.bounds.top;
						const neighborBottom = neighbor.bounds.top + neighbor.bounds.height;
						const verticalOverlap = neighborBottom > cellTop && neighborTop < cellBottom;
						if (!verticalOverlap) continue;
						if (Math.abs(neighborLeft - cellRight) > tolerance) continue;
						neighbor.styles.border = neighbor.styles.border || {};
						if (!neighbor.styles.border.left) {
							neighbor.styles.border.left = { ...border.right };
						}
					}
				}
				
				if (border.left) {
					for (const neighbor of cells) {
						if (neighbor === cell) continue;
						const neighborRight = neighbor.bounds.left + neighbor.bounds.width;
						const neighborTop = neighbor.bounds.top;
						const neighborBottom = neighbor.bounds.top + neighbor.bounds.height;
						const verticalOverlap = neighborBottom > cellTop && neighborTop < cellBottom;
						if (!verticalOverlap) continue;
						if (Math.abs(neighborRight - cellLeft) > tolerance) continue;
						neighbor.styles.border = neighbor.styles.border || {};
						if (!neighbor.styles.border.right) {
							neighbor.styles.border.right = { ...border.left };
						}
					}
				}
				
				if (border.bottom) {
					for (const neighbor of cells) {
						if (neighbor === cell) continue;
						const neighborTop = neighbor.bounds.top;
						const neighborLeft = neighbor.bounds.left;
						const neighborRight = neighbor.bounds.left + neighbor.bounds.width;
						const horizontalOverlap = neighborRight > cellLeft && neighborLeft < cellRight;
						if (!horizontalOverlap) continue;
						if (Math.abs(neighborTop - cellBottom) > tolerance) continue;
						neighbor.styles.border = neighbor.styles.border || {};
						if (!neighbor.styles.border.top) {
							neighbor.styles.border.top = { ...border.bottom };
						}
					}
				}
				
				if (border.top) {
					for (const neighbor of cells) {
						if (neighbor === cell) continue;
						const neighborBottom = neighbor.bounds.top + neighbor.bounds.height;
						const neighborLeft = neighbor.bounds.left;
						const neighborRight = neighbor.bounds.left + neighbor.bounds.width;
						const horizontalOverlap = neighborRight > cellLeft && neighborLeft < cellRight;
						if (!horizontalOverlap) continue;
						if (Math.abs(neighborBottom - cellTop) > tolerance) continue;
						neighbor.styles.border = neighbor.styles.border || {};
						if (!neighbor.styles.border.bottom) {
							neighbor.styles.border.bottom = { ...border.top };
						}
					}
				}
			}
		};
		
		const applyBackgroundRegions = (cells, regions) => {
			if (!Array.isArray(regions) || !regions.length) return;
			for (const region of regions) {
				if (!region || !region.fill) continue;
				const { rowStart, rowEnd, colStart, colEnd, fill } = region;
				for (const cell of cells) {
					if (!cell) continue;
					if (cell.rowStart >= rowEnd || cell.rowEnd <= rowStart) continue;
					if (cell.colStart >= colEnd || cell.colEnd <= colStart) continue;
					cell.styles = cell.styles || {};
					if (!cell.styles.fill) {
						cell.styles.fill = {
							type: fill.type,
							pattern: fill.pattern,
							fgColor: fill.fgColor ? { ...fill.fgColor } : undefined
						};
					}
				}
			}
		};
		
		const applyOuterElementBorders = (descriptor, borders) => {
			if (!descriptor || !borders) return;
			
			pushDeferredBorderInstruction(descriptor, 'top', borders.top);
			pushDeferredBorderInstruction(descriptor, 'bottom', borders.bottom);
			pushDeferredBorderInstruction(descriptor, 'left', borders.left);
			pushDeferredBorderInstruction(descriptor, 'right', borders.right);
			
			const tolerance = Math.max(CONFIG.EDGE_TOLERANCE_PX, 2);
			const outerLeft = descriptor.left;
			const outerRight = descriptor.left + descriptor.width;
			const outerTop = descriptor.top;
			const outerBottom = descriptor.top + descriptor.height;
			
			const inRange = (cell) => {
				if (!cell?.bounds) return false;
				const cellLeft = cell.bounds.left;
				const cellRight = cell.bounds.left + cell.bounds.width;
				const cellTop = cell.bounds.top;
				const cellBottom = cell.bounds.top + cell.bounds.height;
				return (
				cellLeft >= outerLeft - tolerance &&
				cellRight <= outerRight + tolerance &&
				cellTop >= outerTop - tolerance &&
				cellBottom <= outerBottom + tolerance
				);
			};
			
			const alignsWithOuterEdge = (cell, side) => {
				if (!cell?.bounds) return false;
				const diff = (() => {
					switch (side) {
						case 'left':
						return Math.abs(cell.bounds.left - outerLeft);
						case 'right':
						return Math.abs(cell.bounds.left + cell.bounds.width - outerRight);
						case 'top':
						return Math.abs(cell.bounds.top - outerTop);
						case 'bottom':
						default:
						return Math.abs(cell.bounds.top + cell.bounds.height - outerBottom);
					}
				})();
				if (diff <= tolerance) return true;
				
				const csCell = getCachedComputedStyle(cell.element);
				if (!csCell) return diff <= tolerance;
				
				if (side === 'left' || side === 'right') {
					const padding = computeEffectiveHorizontalPadding(cell.element, csCell);
					const allowance = side === 'left' ? padding.left : padding.right;
					return diff <= Math.max(tolerance, allowance + CONFIG.EDGE_TOLERANCE_PX);
				}
				
				const padding = computeEffectiveVerticalPadding(cell.element, csCell);
				const allowance = side === 'top' ? padding.top : padding.bottom;
				return diff <= Math.max(tolerance, allowance + CONFIG.EDGE_TOLERANCE_PX);
			};
			
			if (borders.top) {
				const topCells = cells.filter((c) => inRange(c) && alignsWithOuterEdge(c, 'top'));
				for (const cell of topCells) {
					mergeBorderStyle(cell, 'top', borders.top);
				}
			}
			
			if (borders.bottom) {
				const bottomCells = cells.filter(
				(c) => inRange(c) && alignsWithOuterEdge(c, 'bottom')
				);
				for (const cell of bottomCells) {
					mergeBorderStyle(cell, 'bottom', borders.bottom);
				}
			}
			
			if (borders.left) {
				const leftCells = cells.filter((c) => inRange(c) && alignsWithOuterEdge(c, 'left'));
				for (const cell of leftCells) {
					mergeBorderStyle(cell, 'left', borders.left);
				}
			}
			
			if (borders.right) {
				const rightCells = cells.filter(
				(c) => inRange(c) && alignsWithOuterEdge(c, 'right')
				);
				for (const cell of rightCells) {
					mergeBorderStyle(cell, 'right', borders.right);
				}
			}
		};
		
		for (const outerCell of nestedCells) {
			const descriptor = getDescriptor(outerCell);
			if (!descriptor) continue;
			const borders = extractBorders(window.getComputedStyle(outerCell), outerCell);
			if (!borders) continue;
			applyOuterElementBorders(descriptor, borders);
		}
		
		for (const tableElement of tableElements) {
			if (!tableElement || tableElement === root) continue;
			const descriptor = getDescriptor(tableElement);
			if (!descriptor) continue;
			const csTable = window.getComputedStyle(tableElement);
			const borders = extractBorders(csTable, tableElement);
			if (borders) {
				applyOuterElementBorders(descriptor, borders);
			}
			const bgColor = colorToARGB(csTable.backgroundColor);
			if (bgColor) {
				tableBackgroundRegions.push({
					rowStart: descriptor.rowStartIndex,
					rowEnd: Math.max(descriptor.rowStartIndex + 1, descriptor.rowEndIndex),
					colStart: descriptor.colStartIndex,
					colEnd: Math.max(descriptor.colStartIndex + 1, descriptor.colEndIndex),
					fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } }
				});
			}
		}
		
		propagateSharedBorders(cells);
		applyBackgroundRegions(cells, tableBackgroundRegions);
		
		const columnWidthsPx = [];
		for (let i = 0; i < columnEdges.length - 1; i++) {
			const w = snapToPrecision(columnEdges[i + 1] - columnEdges[i]);
			columnWidthsPx.push(w > 0 ? w : CONFIG.MIN_COL_WIDTH_PX);
		}
		
		const rowHeightsPx = [];
		for (let i = 0; i < rowEdges.length - 1; i++) {
			const h = snapToPrecision(rowEdges[i + 1] - rowEdges[i]);
			rowHeightsPx.push(h > 0 ? h : CONFIG.MIN_ROW_HEIGHT_PX);
		}
		
		const collectPageBreakRows = () => {
			if (!pageBreakMarkers.length) return [];
			const tolerance = Math.max(CONFIG.EDGE_TOLERANCE_PX, 1);
			const totalRows = Math.max(0, rowEdges.length - 1);
			const rows = new Set();
			for (const marker of pageBreakMarkers) {
				if (!marker || typeof marker.getBoundingClientRect !== 'function') continue;
				const rect = marker.getBoundingClientRect();
				const markerTop = snapToPrecision(rect.top + scrollY - rootTop);
				if (!Number.isFinite(markerTop)) continue;
				for (let i = 0; i < rowEdges.length; i++) {
					const edge = rowEdges[i];
					if (markerTop <= edge + tolerance) {
						const rowIndex = Math.max(1, i + 1);
						if (rowIndex >= 2 && rowIndex <= totalRows) {
							rows.add(rowIndex);
						}
						break;
					}
				}
			}
			return Array.from(rows).sort((a, b) => a - b);
		};
		
		return {
			columnEdges,
			rowEdges,
			columnWidthsPx,
			rowHeightsPx,
			cells,
			images,
			pageBreakRows: collectPageBreakRows(),
			deferredBorders
		};
	};
	
	const ensureImageLoaded = (img) => {
		if (img.complete) {
			if (img.naturalWidth && img.naturalHeight) return Promise.resolve();
			return Promise.reject(new Error('Image failed to load (already complete with zero size)'));
		}
		return new Promise((resolve, reject) => {
			const cleanup = () => {
				img.removeEventListener('load', onLoad);
				img.removeEventListener('error', onError);
			};
			const onLoad = () => {
				cleanup();
				resolve();
			};
			const onError = (err) => {
				cleanup();
				reject(err);
			};
			img.addEventListener('load', onLoad, { once: true });
			img.addEventListener('error', onError, { once: true });
		});
	};
	
	const blobToBase64 = (blob) => new Promise((resolve, reject) => {
		const reader = new FileReader();
		reader.onerror = () => reject(new Error('Failed to read blob as base64'));
		reader.onload = () => {
			const result = reader.result || '';
			const [, base64] = String(result).split(',');
			if (base64) {
				resolve(base64);
			} else {
				reject(new Error('Empty base64 result'));
			}
		};
		reader.readAsDataURL(blob);
	});
	
	const fetchImageAsBase64 = async (src) => {
		if (!src) return null;
		let url;
		try {
			url = new URL(src, window.location.href);
		} catch (err) {
			warnLog('Invalid image URL, skip fetch fallback', src, err);
			return null;
		}
		
		if (url.protocol === 'file:') {
			warnLog('Skipping fetch fallback for file:// URL (browser security restriction):', url.href);
			return null;
		}
		
		try {
			const response = await fetch(url.href, { mode: 'cors' });
			if (!response.ok) return null;
			const blob = await response.blob();
			const base64 = await blobToBase64(blob);
			const extension = (blob.type && blob.type.split('/')[1]) || 'png';
			debugLog('Fetched image via fallback', url.href, `(${extension})`);
			return { base64, extension };
		} catch (err) {
			warnLog('Fallback fetch-to-base64 failed', src, err);
			return null;
		}
	};
	
	const convertImageJob = async (job) => {
		const { element, width, height } = job;
		if (!element) return { ...job, error: new Error('Image element missing') };
		
		try {
			await ensureImageLoaded(element);
		} catch (err) {
			warnLog('Image failed to load', element.src, err);
			return { ...job, error: err };
		}
		
		const imageScale = toPositiveNumber(IMAGE_SCALE) || 1;
		const renderWidth = Math.max(1, Math.round(width * imageScale));
		const renderHeight = Math.max(1, Math.round(height * imageScale));
		
		const canvas = document.createElement('canvas');
		canvas.width = renderWidth;
		canvas.height = renderHeight;
		const ctx = canvas.getContext('2d');
		if (!ctx) {
			const err = new Error('Canvas 2D context unavailable');
			errorLog(err.message, element.src);
			return { ...job, error: err };
		}
		
		try {
			ctx.drawImage(element, 0, 0, renderWidth, renderHeight);
			const dataUrl = canvas.toDataURL(CONFIG.IMAGE_FORMAT);
			const base64 = dataUrl.split(',')[1];
			if (!base64) throw new Error('Empty base64 from canvas');
			debugLog('Canvas conversion success', element.src, `${renderWidth}x${renderHeight}`, `base64Length=${base64.length}`);
			return {
				...job,
				base64,
				extension: CONFIG.IMAGE_FORMAT === 'image/png' ? 'png' : 'jpeg'
			};
		} catch (err) {
			warnLog('Canvas toDataURL failed, trying fetch fallback', element.src, err);
			const fallback = await fetchImageAsBase64(element.src);
			if (fallback) {
				debugLog('Fallback fetch succeeded for image', element.src, `base64Length=${fallback.base64.length}`);
				return { ...job, base64: fallback.base64, extension: fallback.extension };
			}
			errorLog('Image conversion failed after fallback', element.src, err);
			return { ...job, error: err };
		}
	};
	
	const applyColumnWidths = (worksheet, columnWidthsPx) => {
		worksheet.columns = columnWidthsPx.map((px) => ({ width: pxToColumnWidth(px, EXPORT_SCALE) }));
	};
	
	const applyRowHeights = (worksheet, rowHeightsPx) => {
		rowHeightsPx.forEach((px, idx) => {
			const row = worksheet.getRow(idx + 1);
			const heightByRowScale = pxToPoint(px, ROW_SCALE);
			const heightByFontScale = pxToPoint(px, FONT_SCALE);
			row.height = Math.max(heightByRowScale, heightByFontScale);
			row.alignment = row.alignment || { vertical: 'top' };
		});
	};
	
	const applyPageBreaks = (worksheet, layoutModel) => {
		const breaks = Array.isArray(layoutModel.pageBreakRows) ? layoutModel.pageBreakRows : [];
		if (!breaks.length) return;
		const existingBreaks = worksheet.rowBreaks || (worksheet.rowBreaks = []);
		const seen = new Set(
		existingBreaks
		.map((brk) => (brk && Number.isFinite(brk.id) ? brk.id : null))
		.filter((id) => id !== null)
		);
		for (const rowIndex of breaks) {
			if (!Number.isInteger(rowIndex)) continue;
			const rowId = rowIndex - 1;
			if (rowId < 1) continue;
			if (seen.has(rowId)) continue;
			existingBreaks.push({ id: rowId, max: 16383, man: 1 });
			seen.add(rowId);
		}
		worksheet.rowBreaks = existingBreaks;
	};
	
	const writeCellsToSheet = (worksheet, model) => {
		const occupancy = new Set();
		const sortedCells = model.cells.slice().sort((a, b) => {
			if (a.rowStart !== b.rowStart) return a.rowStart - b.rowStart;
			return a.colStart - b.colStart;
		});
		
		for (const cell of sortedCells) {
			const rowStart = cell.rowStart + 1;
			const rowEnd = cell.rowEnd;
			const colStart = cell.colStart + 1;
			const colEnd = cell.colEnd;
			
			const key = `${rowStart},${colStart}`;
			if (occupancy.has(key)) continue;
			
			const mergeRowEnd = Math.max(rowStart, rowEnd);
			const mergeColEnd = Math.max(colStart, colEnd);
			if (mergeRowEnd > rowStart || mergeColEnd > colStart) {
				worksheet.mergeCells(rowStart, colStart, mergeRowEnd, mergeColEnd);
			}
			
			const targetCell = worksheet.getCell(rowStart, colStart);
			targetCell.value = cell.value;
			if (cell.styles.font) targetCell.font = cell.styles.font;
			if (cell.styles.alignment) targetCell.alignment = cell.styles.alignment;
			if (cell.styles.border) targetCell.border = cell.styles.border;
			if (cell.styles.fill) targetCell.fill = cell.styles.fill;
			if (cell.styles.numFmt) targetCell.numFmt = cell.styles.numFmt;
			
			for (let r = rowStart; r <= mergeRowEnd; r++) {
				for (let c = colStart; c <= mergeColEnd; c++) {
					occupancy.add(`${r},${c}`);
				}
			}
			
			cell.sheetPosition = {
				rowStart,
				rowEnd: mergeRowEnd,
				colStart,
				colEnd: mergeColEnd
			};
		}
	};
	
	const applyDeferredBorders = (worksheet, model) => {
		const instructions = Array.isArray(model.deferredBorders) ? model.deferredBorders : [];
		if (!instructions.length) return;
		
		const applyBorderSide = (cell, side, borderStyle) => {
			if (!cell || !borderStyle) return;
			const existing = cell.border && typeof cell.border === 'object' ? { ...cell.border } : {};
			existing[side] = { ...borderStyle };
			cell.border = existing;
		};
		
		for (const instruction of instructions) {
			if (!instruction || !instruction.style) continue;
			const { side, style, rowStart, rowEnd, colStart, colEnd } = instruction;
			if (!side || rowStart == null || rowEnd == null || colStart == null || colEnd == null) continue;
			const safeRowStart = Math.max(1, Math.floor(rowStart));
			const safeRowEnd = Math.max(safeRowStart, Math.floor(rowEnd));
			const safeColStart = Math.max(1, Math.floor(colStart));
			const safeColEnd = Math.max(safeColStart, Math.floor(colEnd));
			
			switch (side) {
				case 'left':
				for (let r = safeRowStart; r <= safeRowEnd; r++) {
					applyBorderSide(worksheet.getCell(r, safeColStart), 'left', style);
				}
				break;
				case 'right':
				for (let r = safeRowStart; r <= safeRowEnd; r++) {
					applyBorderSide(worksheet.getCell(r, safeColEnd), 'right', style);
				}
				break;
				case 'top':
				for (let c = safeColStart; c <= safeColEnd; c++) {
					applyBorderSide(worksheet.getCell(safeRowStart, c), 'top', style);
				}
				break;
				case 'bottom':
				for (let c = safeColStart; c <= safeColEnd; c++) {
					applyBorderSide(worksheet.getCell(safeRowEnd, c), 'bottom', style);
				}
				break;
				default:
				break;
			}
		}
	};
	
	const locateIndexAndOffset = (edges, px, preferNextEdge = false) => {
		const value = Math.max(0, px);
		const tolerance = CONFIG.EDGE_TOLERANCE_PX;
		for (let i = 0; i < edges.length - 1; i++) {
			const start = edges[i];
			const end = edges[i + 1];
			if (value >= start && value < end - tolerance) {
				return { index: i, offset: value - start };
			}
			const closeToEnd = Math.abs(value - end) <= tolerance;
			if (closeToEnd) {
				if (preferNextEdge && i + 1 < edges.length - 1) {
					return { index: i + 1, offset: 0 };
				}
				return { index: i, offset: end - start };
			}
		}
		const lastIndex = Math.max(0, edges.length - 2);
		return {
			index: lastIndex,
			offset: value - edges[lastIndex]
		};
	};
	
	const clampOffset = (offsetPx, maxPx) => {
		if (!isFinite(offsetPx) || offsetPx < 0) return 0;
		if (isFinite(maxPx) && maxPx > 0) {
			return Math.round(Math.max(0, Math.min(offsetPx, maxPx)));
		}
		return Math.round(offsetPx);
	};
	
	const toImageAlignment = (value) => {
		if (!value) return null;
		const keyword = String(value).trim().toLowerCase();
		switch (keyword) {
			case 'inline':
			case 'block':
			case 'center':
			case 'left':
			case 'right':
			case 'precise':
			return keyword;
			default:
			return null;
		}
	};
	
	const getCellImageAlignment = (cellModel) => {
		if (!cellModel) return null;
		const element = cellModel.element;
		const attrName = CONFIG.IMAGE_ALIGNMENT_ATTR;
		if (attrName && element && typeof element.getAttribute === 'function') {
			const attrValue = element.getAttribute(attrName);
			const resolved = toImageAlignment(attrValue);
			if (resolved) {
				return resolved;
			}
		}
		const horizontal =
		cellModel.styles &&
		cellModel.styles.alignment &&
		cellModel.styles.alignment.horizontal
		? String(cellModel.styles.alignment.horizontal).toLowerCase()
		: null;
		switch (horizontal) {
			case 'center':
			case 'middle':
			return 'center';
			case 'right':
			return 'right';
			case 'left':
			return 'left';
			default:
			return null;
		}
	};
	
	const getCellPlacementMetrics = (cellModel, layoutModel) => {
		if (!cellModel || !layoutModel) return null;
		const sheetPosition = cellModel.sheetPosition;
		if (!sheetPosition) return null;
		const columnEdges = layoutModel.columnEdges || [];
		const rowEdges = layoutModel.rowEdges || [];
		const leftEdge = columnEdges[cellModel.colStart] !== undefined ? columnEdges[cellModel.colStart] : 0;
		const rightEdge = columnEdges[cellModel.colEnd] !== undefined ? columnEdges[cellModel.colEnd] : leftEdge;
		const topEdge = rowEdges[cellModel.rowStart] !== undefined ? rowEdges[cellModel.rowStart] : 0;
		const bottomEdge = rowEdges[cellModel.rowEnd] !== undefined ? rowEdges[cellModel.rowEnd] : topEdge;
		const widthPx = Math.max(0, snapToPrecision(rightEdge - leftEdge));
		const heightPx = Math.max(0, snapToPrecision(bottomEdge - topEdge));
		return {
			widthPx,
			heightPx,
			startColIndex: sheetPosition.colStart - 1,
			startRowIndex: sheetPosition.rowStart - 1
		};
	};
	
	const clampWithin = (offsetPx, spanPx, limitPx) => {
		if (!isFiniteNumber(offsetPx)) return 0;
		const safeSpan = isFiniteNumber(spanPx) && spanPx > 0 ? spanPx : 0;
		const safeLimit = isFiniteNumber(limitPx) && limitPx > 0 ? limitPx : 0;
		if (!safeLimit) return 0;
		if (safeSpan >= safeLimit) return 0;
		return Math.max(0, Math.min(offsetPx, safeLimit - safeSpan));
	};
	
	const logLatestWorksheetImage = (worksheet) => {
		if (typeof worksheet?.getImages !== 'function') {
			return;
		}
		const currentImages = worksheet.getImages();
		if (!Array.isArray(currentImages) || !currentImages.length) {
			return;
		}
		const latest = currentImages[currentImages.length - 1];
		if (latest) {
			debugLog('Worksheet image detail', {
				range: latest.range,
				hyperlink: latest.hyperlink || null
			});
		}
		debugLog('Worksheet image count', currentImages.length);
	};
	
	const addAlignedImage = (worksheet, workbook, job, metrics, offsetLeftPx, offsetTopPx, widthPx, heightPx) => {
		if (!metrics || widthPx <= 0 || heightPx <= 0) return false;
		const extraLeft = 1000;
		const extraTop = 10;
		const rawLeft = (offsetLeftPx || 0) + extraLeft;
		const rawTop = (offsetTopPx || 0) + extraTop;
		const safeLeft = clampWithin(rawLeft, widthPx, metrics.widthPx);
		const safeTop = clampWithin(rawTop, heightPx, metrics.heightPx);
		const colRatio = metrics.widthPx > 0 ? safeLeft / metrics.widthPx : 0;
		const rowRatio = metrics.heightPx > 0 ? safeTop / metrics.heightPx : 0;
		const tl = {
			col: metrics.startColIndex + colRatio,
			row: metrics.startRowIndex + rowRatio
		};
		const imgId = workbook.addImage({
			base64: job.base64,
			extension: job.extension
		});
		const ext = {
			width: widthPx,
			height: heightPx
		};
		worksheet.addImage(imgId, { tl, ext, editAs: 'oneCell' });
		debugLog('Placed image using alignment flow', job.element?.src || '(unknown src)', {
			tl,
			ext,
			offset: { left: safeLeft, top: safeTop },
			cellSize: { widthPx: metrics.widthPx, heightPx: metrics.heightPx }
		});
		logLatestWorksheetImage(worksheet);
		return true;
	};
	
	const placeImagesInline = (worksheet, workbook, images, cellModel, layoutModel, justify) => {
		const metrics = getCellPlacementMetrics(cellModel, layoutModel);
		if (!metrics || metrics.widthPx <= 0 || metrics.heightPx <= 0) {
			return false;
		}
		const spacingPx = Math.max(0, snapToPrecision(CONFIG.IMAGE_SPACING_PX || 0));
		const maxPerRowConfigured = toPositiveNumber(CONFIG.IMAGE_MAX_PER_ROW);
		const maxPerRow = maxPerRowConfigured ? Math.max(1, Math.floor(maxPerRowConfigured)) : Number.POSITIVE_INFINITY;
		const finalImageScale = toPositiveNumber(CONFIG.IMAGE_WIDTH_SCALE) || 1.0;
		
		const rows = [];
		let currentRow = { items: [], width: 0, height: 0 };
		
		for (const job of images) {
			const widthPx = Math.max(1, job.width * finalImageScale);
			const heightPx = Math.max(1, job.height * finalImageScale);
			const item = { job, widthPx, heightPx };
			const needsWrap =
			(currentRow.items.length > 0 && currentRow.width + spacingPx + widthPx > metrics.widthPx) ||
			currentRow.items.length >= maxPerRow;
			if (needsWrap) {
				rows.push(currentRow);
				currentRow = { items: [], width: 0, height: 0 };
			}
			const additionalSpacing = currentRow.items.length > 0 ? spacingPx : 0;
			currentRow.items.push(item);
			currentRow.width += widthPx + additionalSpacing;
			currentRow.height = Math.max(currentRow.height, heightPx);
		}
		if (currentRow.items.length) {
			rows.push(currentRow);
		}
		
		let currentTop = 0;
		for (const row of rows) {
			const rowWidth = row.width;
			let startLeft = 0;
			if (justify === 'center') {
				startLeft = Math.max(0, (metrics.widthPx - rowWidth) / 2);
			} else if (justify === 'end') {
				startLeft = Math.max(0, metrics.widthPx - rowWidth);
			}
			let cursor = startLeft;
			for (const item of row.items) {
				addAlignedImage(worksheet, workbook, item.job, metrics, cursor, currentTop, item.widthPx, item.heightPx);
				cursor += item.widthPx + spacingPx;
			}
			currentTop += row.height + spacingPx;
		}
		return true;
	};
	
	const placeImagesStacked = (worksheet, workbook, images, cellModel, layoutModel) => {
		const metrics = getCellPlacementMetrics(cellModel, layoutModel);
		if (!metrics || metrics.widthPx <= 0 || metrics.heightPx <= 0) {
			return false;
		}
		const spacingPx = Math.max(0, snapToPrecision(CONFIG.IMAGE_SPACING_PX || 0));
		const finalImageScale = toPositiveNumber(CONFIG.IMAGE_WIDTH_SCALE) || 1.0;
		let currentTop = 0;
		for (const job of images) {
			const widthPx = Math.max(1, job.width * finalImageScale);
			const heightPx = Math.max(1, job.height * finalImageScale);
			addAlignedImage(worksheet, workbook, job, metrics, 0, currentTop, widthPx, heightPx);
			currentTop += heightPx + spacingPx;
		}
		return true;
	};
	
	const placeSingleImageCentered = (worksheet, workbook, job, cellModel, layoutModel) => {
		const metrics = getCellPlacementMetrics(cellModel, layoutModel);
		if (!metrics || metrics.widthPx <= 0 || metrics.heightPx <= 0) {
			return false;
		}
		const finalImageScale = toPositiveNumber(CONFIG.IMAGE_WIDTH_SCALE) || 1.0;
		const widthPx = Math.max(1, job.width * finalImageScale);
		const heightPx = Math.max(1, job.height * finalImageScale);
		const offsetLeft = Math.max(0, (metrics.widthPx - widthPx) / 2);
		const offsetTop = Math.max(0, (metrics.heightPx - heightPx) / 2);
		return addAlignedImage(worksheet, workbook, job, metrics, offsetLeft, offsetTop, widthPx, heightPx);
	};
	
	const placeImagesWithAlignment = (worksheet, workbook, images, cellModel, layoutModel, alignment) => {
		const keyword = alignment || 'inline';
		if (keyword === 'precise') return false;
		if (images.length === 1) {
			if (keyword === 'center') {
				return placeSingleImageCentered(worksheet, workbook, images[0], cellModel, layoutModel);
			}
			const justify = keyword === 'right' ? 'end' : keyword === 'center' ? 'center' : 'start';
			return placeImagesInline(worksheet, workbook, images, cellModel, layoutModel, justify);
		}
		
		switch (keyword) {
			case 'block':
			return placeImagesStacked(worksheet, workbook, images, cellModel, layoutModel);
			case 'right':
			return placeImagesInline(worksheet, workbook, images, cellModel, layoutModel, 'end');
			case 'center':
			return placeImagesInline(worksheet, workbook, images, cellModel, layoutModel, 'center');
			case 'left':
			case 'inline':
			default:
			return placeImagesInline(worksheet, workbook, images, cellModel, layoutModel, 'start');
		}
	};
	
	const placeImagesPrecisely = (worksheet, workbook, resolvedJobs, layoutModel) => {
		if (!resolvedJobs.length) return;
		
		// We must use the original, unscaled HTML pixel grid for image positioning.
		// This ensures that COLUMN_WIDTH_SCALE does not affect image size or position.
		const htmlColumnEdges = layoutModel.columnEdges;
		const htmlRowEdges = layoutModel.rowEdges;
		
		for (const job of resolvedJobs) {
			if (!job.base64) continue;
			const imgId = workbook.addImage({
				base64: job.base64,
				extension: job.extension
			});
			
			const parentCell = job.parentCell;
			const sheetPosition = parentCell?.sheetPosition;
			if (!sheetPosition) continue;
			
			const { colStart, rowStart } = sheetPosition;
			const startColIndex = colStart - 1;
			const startRowIndex = rowStart - 1;
			
			// Get the top-left corner of the cell in the original HTML pixel grid
			const cellHtmlLeft = htmlColumnEdges[startColIndex] || 0;
			const cellHtmlTop = htmlRowEdges[startRowIndex] || 0;
			
			// Calculate the image's offset within the cell in pixels and apply final tweaks
			const offsetX = job.left - cellHtmlLeft + (CONFIG.IMAGE_OFFSET_LEFT_PX || 0);
			const offsetY = job.top - cellHtmlTop + (CONFIG.IMAGE_OFFSET_TOP_PX || 0);
			
			// Get the width/height of the anchor cell in the original HTML pixel grid
			const anchorCellHtmlWidth = (htmlColumnEdges[startColIndex + 1] || cellHtmlLeft) - cellHtmlLeft;
			const anchorCellHtmlHeight = (htmlRowEdges[startRowIndex + 1] || cellHtmlTop) - cellHtmlTop;
			
			// Calculate the anchor position as a ratio of the cell's dimensions
			const colRatio = anchorCellHtmlWidth > 0 ? offsetX / anchorCellHtmlWidth : 0;
			const rowRatio = anchorCellHtmlHeight > 0 ? offsetY / anchorCellHtmlHeight : 0;
			
			const tl = {
				col: startColIndex + colRatio,
				row: startRowIndex + rowRatio,
			};
			
			// The extension (size) of the image should also be based on the original pixel size
			const finalImageScale = toPositiveNumber(CONFIG.IMAGE_WIDTH_SCALE) || 1.0;
			const ext = {
				width: job.width * finalImageScale,
				height: job.height * finalImageScale,
			};
			
			worksheet.addImage(imgId, { tl, ext, editAs: 'oneCell' });
			
			debugLog('Placed image using direct pixel mapping', job.element?.src || '(unknown src)', {
				tl,
				ext,
				parentCellBounds: parentCell?.bounds,
			});
			
			
			if (typeof worksheet.getImages === 'function') {
				const currentImages = worksheet.getImages();
				if (Array.isArray(currentImages)) {
					const latest = currentImages[currentImages.length - 1];
					if (latest) {
						debugLog('Worksheet image detail', {
							range: latest.range,
							hyperlink: latest.hyperlink || null
						});
					}
					debugLog('Worksheet image count', currentImages.length);
				}
			}
		}
	};
	
	const placeImages = (worksheet, workbook, resolvedJobs, layoutModel) => {
		if (!resolvedJobs.length) return;
		const placementMode = String(CONFIG.IMAGE_PLACEMENT_MODE || 'alignment').toLowerCase();
		if (placementMode === 'precise') {
			placeImagesPrecisely(worksheet, workbook, resolvedJobs, layoutModel);
			return;
		}
		
		const cellGroups = new Map();
		for (const job of resolvedJobs) {
			const parentCell = job.parentCell;
			if (!parentCell || !parentCell.sheetPosition) {
				continue;
			}
			if (!cellGroups.has(parentCell)) {
				cellGroups.set(parentCell, []);
			}
			cellGroups.get(parentCell).push(job);
		}
		
		const preciseFallbackJobs = [];
		const defaultAlignment = toImageAlignment(CONFIG.IMAGE_ALIGNMENT) || 'inline';
		
		for (const [cellModel, images] of cellGroups.entries()) {
			if (!Array.isArray(images) || !images.length) continue;
			const cellAlignment = getCellImageAlignment(cellModel) || defaultAlignment;
			const placed = placeImagesWithAlignment(worksheet, workbook, images, cellModel, layoutModel, cellAlignment);
			if (!placed) {
				preciseFallbackJobs.push(...images);
			}
		}
		
		// Process any images we could not auto-align via the original precise placement logic.
		if (preciseFallbackJobs.length) {
			placeImagesPrecisely(worksheet, workbook, preciseFallbackJobs, layoutModel);
		}
	};
	
	
	const applyPageSetupOptions = (worksheet) => {
		if (!CONFIG.PAGE_SETUP) return;
		const { margins, ...pageSetupRest } = CONFIG.PAGE_SETUP;
		const pageSetupConfig = { ...pageSetupRest };
		if (pageSetupConfig.fitToPage === true) {
			delete pageSetupConfig.scale;
			if (typeof pageSetupConfig.fitToWidth !== 'number') {
				pageSetupConfig.fitToWidth = 1;
			}
			if (typeof pageSetupConfig.fitToHeight !== 'number') {
				pageSetupConfig.fitToHeight = 0;
			}
		}
		if (Object.keys(pageSetupConfig).length) {
			Object.assign(worksheet.pageSetup, pageSetupConfig);
		}
		if (margins) {
			worksheet.pageSetup.margins = {
				...worksheet.pageSetup.margins,
				...margins
			};
		}
		debugLog('Applied page setup', worksheet.pageSetup);
	};
	
	const createWorkbook = () => {
		const workbook = new ExcelJS.Workbook();
		workbook.created = new Date();
		return workbook;
	};
	
	const prepareWorksheet = (workbook, layoutModel) => {
		const defaultRowHeightPt = Math.max(
		pxToPoint(CONFIG.MIN_ROW_HEIGHT_PX, ROW_SCALE),
		pxToPoint(CONFIG.MIN_ROW_HEIGHT_PX, FONT_SCALE)
		);
		const worksheet = workbook.addWorksheet(sanitizeSheetName(CONFIG.SHEET_NAME), {
			properties: { defaultRowHeight: defaultRowHeightPt }
		});
		applyPageSetupOptions(worksheet);
		debugLog('Applying column widths and row heights');
		applyColumnWidths(worksheet, layoutModel.columnWidthsPx);
		applyRowHeights(worksheet, layoutModel.rowHeightsPx);
		debugLog('Writing cell content');
		writeCellsToSheet(worksheet, layoutModel);
		applyDeferredBorders(worksheet, layoutModel);
		applyPageBreaks(worksheet, layoutModel);
		return worksheet;
	};
	
	const processAndPlaceImages = async (workbook, worksheet, layoutModel) => {
		if (!layoutModel.images.length) return;
		debugLog('Converting images...');
		const imageResults = await Promise.all(layoutModel.images.map(convertImageJob));
		const successfulImages = imageResults.filter((item) => item && item.base64);
		const failedImages = imageResults.filter((item) => item && !item.base64);
		if (successfulImages.length) {
			placeImages(worksheet, workbook, successfulImages, layoutModel);
			debugLog('Embedded images', successfulImages.length);
		}
		if (failedImages.length) {
			failedImages.forEach((item) => {
				warnLog('Image skipped (no base64)', {
					src: item.element?.src || '(unknown src)',
					width: item.width,
					height: item.height,
					error: item.error ? item.error.message || String(item.error) : '(no error message)'
				});
			});
			warnLog('Images could not be embedded', failedImages.map((f) => f.element?.src || '(unknown src)'));
		}
	};
	
	const logWorkbookMediaSummary = (workbook) => {
		if (!CONFIG.DEBUG || !workbook) return;
		try {
			const model = workbook.model;
			const mediaCount = Array.isArray(model?.media) ? model.media.length : 0;
			const worksheetSummaries = Array.isArray(model?.worksheets)
			? model.worksheets.map((ws) => ({
				name: ws.name,
				mediaItems: Array.isArray(ws.media) ? ws.media.length : 0,
				drawingCount: Array.isArray(ws.drawings) ? ws.drawings.length : 0
			}))
			: [];
			debugLog('Workbook media summary', { mediaCount, worksheetSummaries });
		} catch (err) {
			warnLog('Failed to inspect workbook model', err);
		}
	};
	
	const writeWorkbookToBuffer = async (workbook) => {
		try {
			logWorkbookMediaSummary(workbook);
			debugLog('Writing workbook buffer');
			const buffer = await workbook.xlsx.writeBuffer();
			debugLog('Workbook buffer size', buffer && (buffer.byteLength || buffer.length || 0));
			return buffer;
		} catch (err) {
			errorLog('writeBuffer failed', err);
			throw err;
		}
	};
	
	const generateTimestampedFilename = () => {
		const d = new Date();
		const pad = (n) => String(n).padStart(2, '0');
		return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}-${pad(d.getHours())}-${pad(d.getMinutes())}-${pad(d.getSeconds())}`;
	};
	
	const triggerDownload = (buffer, filename) => {
		const blob = new Blob([buffer], {
			type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
		});
		const url = URL.createObjectURL(blob);
		const link = document.createElement('a');
		link.href = url;
		link.download = filename;
		document.body.appendChild(link);
		link.click();
		document.body.removeChild(link);
		URL.revokeObjectURL(url);
	};
	
	const exportPrintFormToXlsx = async () => {
		if (typeof ExcelJS === 'undefined') {
			alert('ExcelJS is not loaded. Please ensure folder_javascript/exceljs.min.js is included.');
			return;
		}
		
		const root = document.querySelector(CONFIG.ROOT_SELECTOR);
		if (!root) {
			alert('Export container not found.');
			return;
		}
		
		debugLog('Building layout model...');
		const layoutModel = buildLayoutModel(root);
		debugLog('Layout ready', `cells=${layoutModel.cells.length}`, `images=${layoutModel.images.length}`);
		
		const workbook = createWorkbook();
		const worksheet = prepareWorksheet(workbook, layoutModel);
		
		await processAndPlaceImages(workbook, worksheet, layoutModel);
		
		const buffer = await writeWorkbookToBuffer(workbook);
		const filename = `${getFilenameBase()}_${generateTimestampedFilename()}.xlsx`;
		debugLog('Triggering download', filename);
		triggerDownload(buffer, filename);
	};
	
	const hookButton = () => {
		const button = document.querySelector(CONFIG.BUTTON_SELECTOR);
		if (!button) return;
		button.addEventListener('click', async (evt) => {
			evt.preventDefault();
			if (button.dataset.busy === '1') return;
			button.dataset.busy = '1';
			button.disabled = true;
			try {
				debugLog('Button clicked, starting export');
				await exportPrintFormToXlsx();
				debugLog('Export completed');
			} catch (err) {
				errorLog('Export failed', err);
				alert('Export failed. See console for details.');
			} finally {
				button.disabled = false;
				delete button.dataset.busy;
			}
		});
	};
	
	document.addEventListener('DOMContentLoaded', () => {
		hookButton();
	});
	
	window.exportPrintFormToXlsx = exportPrintFormToXlsx;
})();

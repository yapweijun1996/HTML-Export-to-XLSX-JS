// MIT License. Copyright (c) 2026 YAP WEI JUN.
export function createStyles(dependencies) {
  const { CONFIG, DEFAULT_EXCEL_PADDING_PX, FONT_SCALE, debugLog, escapeRegExp, isFiniteNumber, measureAvgCharWidth, pxToPoint } = dependencies;
	const COLOR_CANVAS_CTX = (function () {
		if (typeof document === 'undefined') return null;
		const canvas = document.createElement('canvas');
		return canvas && canvas.getContext ? canvas.getContext('2d') : null;
	})();

	const clampToByte = (value) => {
		if (!isFiniteNumber(value)) return null;
		if (value <= 0) return 0;
		if (value >= 255) return 255;
		return Math.round(value);
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


  return { COLOR_CANVAS_CTX, clampToByte, toHexByte, parseHexColor, parseAlphaValue, parseRgbFunction, COLOR_CACHE, parseCssColor, colorToARGB, widthToBorderStyle, BORDER_STYLE_KEYWORDS, parseBorderShorthand, matchInlineStyleValue, parseInlineBorderFromAttribute, extractBorders, normalizeFontWeight, isCellUniformlyBold, resolveFontSizePx, extractFont, parseCssPixelValue, computeEffectiveHorizontalPadding, computeEffectiveVerticalPadding, convertPaddingPxToIndent, extractAlignment };
}

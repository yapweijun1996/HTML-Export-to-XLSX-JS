// MIT License. Copyright (c) 2026 YAP WEI JUN.
export function createLayoutMetrics(dependencies) {
  const { CONFIG, debugLog, errorLog } = dependencies;
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

	const DEFAULT_EXCEL_PADDING_PX = 5;

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

	const isFiniteNumber = (value) => (typeof Number.isFinite === 'function' ? Number.isFinite(value) : isFinite(value));


  return { POINTS_PER_CSS_PX, toPositiveNumber, resolveExportScale, EXPORT_SCALE, resolveFontScale, FONT_SCALE, measureAvgCharWidth, resolveRowScale, ROW_SCALE, resolveImageScale, IMAGE_SCALE, DEFAULT_EXCEL_PADDING_PX, pxToPoint, pxToColumnWidth, pxToColumnWidthForImagePositioning, columnWidthToPixels, rowHeightToPixels, interpolateCoordinate, buildWorksheetEdgeMaps, SNAP_PRECISION, snapToPrecision, resolveSize, normalizeEdges, isFiniteNumber };
}

// MIT License. Copyright (c) 2026 YAP WEI JUN.
export function createWorkbookWriter(dependencies) {
 const { CONFIG, EXPORT_SCALE, EXCEL_MAX_DIGIT_WIDTH, ExcelJS, FONT_SCALE, ROW_SCALE, columnWidthToPixels, convertImageJob, debugLog, errorLog, placeImages, pxToColumnWidth, pxToPoint, warnLog, mapWithConcurrency } = dependencies;
	const blankBorderExclusions = new WeakMap();
	const applyColumnWidths = (worksheet, columnWidthsPx) => {
		let totalPixels = 0, writtenPixels = 0;
		worksheet.columns = columnWidthsPx.map(px => {
			totalPixels += pxToColumnWidth(px, EXPORT_SCALE) * EXCEL_MAX_DIGIT_WIDTH;
			const pixels = Math.max(1, Math.round(totalPixels) - writtenPixels);
			writtenPixels += pixels;
			return { width: Math.floor(pixels / EXCEL_MAX_DIGIT_WIDTH * 256) / 256 };
		});
	};

	const applyRowHeights = (worksheet, rowHeightsPx) => {
		let totalPoints = 0, writtenPoints = 0;
		rowHeightsPx.forEach((px, idx) => {
			const row = worksheet.getRow(idx + 1);
			const heightByRowScale = pxToPoint(px, ROW_SCALE);
			const heightByFontScale = pxToPoint(px, FONT_SCALE);
			// Excel rounds row heights to whole screen pixels. Round cumulative edges
			// so fragmented rows do not progressively move images and page boundaries.
			totalPoints += Math.max(heightByRowScale, heightByFontScale);
			row.height = Math.max(0.75, Math.round(totalPoints / 0.75) * 0.75 - writtenPoints);
			writtenPoints += row.height;
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

	const fitMarkedPagesForPrint = (worksheet) => {
		const setup = worksheet.pageSetup;
		if (!CONFIG.RESPECT_PAGE_BREAKS || !worksheet.rowBreaks?.length || !setup.fitToPage) return;
		// Excel's Fit To mode ignores manual breaks. Use percentage scaling for marked pages.
		// Dimensions in millimeters: Letter, Legal, Executive, A3, A4, A5 and JIS B5.
		const papers = { 1: [215.9, 279.4], 5: [215.9, 355.6], 7: [184.15, 266.7],
			8: [297, 420], 9: [210, 297], 11: [148, 210], 13: [182, 257] };
		const paper = papers[setup.paperSize ?? 1];
		setup.fitToPage = false;
		if (!paper) {
			warnLog('Automatic print scaling is unavailable for this paper size; set pageSetup.scale explicitly',
				{ code: 'PRINT_SCALE_UNAVAILABLE' });
			return;
		}
		const [widthMm, heightMm] = setup.orientation === 'landscape' ? [paper[1], paper[0]] : paper;
		const margins = setup.margins;
		const widthPt = (widthMm / 25.4 - margins.left - margins.right) * 72;
		const heightPt = (heightMm / 25.4 - margins.top - margins.bottom) * 72;
		const sheetWidthPt = worksheet.columns.reduce((sum, col) =>
			sum + columnWidthToPixels(col.width) * 72 / 96, 0);
		const boundaries = new Set(worksheet.rowBreaks.map(brk => brk.id));
		let pageHeightPt = 0, tallestPagePt = 0;
		for (let row = 1; row <= worksheet.rowCount; row++) {
			pageHeightPt += worksheet.getRow(row).height ?? worksheet.properties.defaultRowHeight ?? 15;
			if (boundaries.has(row) || row === worksheet.rowCount) {
				tallestPagePt = Math.max(tallestPagePt, pageHeightPt);
				pageHeightPt = 0;
			}
		}
		const scale = Math.floor(100 * Math.min(1, widthPt / sheetWidthPt, heightPt / tallestPagePt));
		if (!Number.isFinite(scale) || scale < 10) {
			warnLog('Marked pages cannot fit the printable area within Excel\'s minimum 10% scale',
				{ code: 'PRINT_SCALE_TOO_SMALL' });
			setup.scale = 10;
			return;
		}
		setup.scale = scale;
	};

	const writeCellsToSheet = (worksheet, model) => {
		const occupancy = new Map();
		const textFor = cell => String(cell.value ?? '').trim();
		const sortedCells = model.cells.slice().sort((a, b) => {
			// Blank layout regions must not reserve cells before meaningful content is written.
			const contentOrder = Number(Boolean(textFor(b))) - Number(Boolean(textFor(a)));
			if (contentOrder) return contentOrder;
			if (a.rowStart !== b.rowStart) return a.rowStart - b.rowStart;
			return a.colStart - b.colStart;
		});

		for (const cell of sortedCells) {
			const rowStart = cell.rowStart + 1;
			const rowEnd = cell.rowEnd;
			const colStart = cell.colStart + 1;
			const colEnd = cell.colEnd;

			const mergeRowEnd = Math.max(rowStart, rowEnd);
			const mergeColEnd = Math.max(colStart, colEnd);
			const text = textFor(cell);
			let collision;
			for (let r = rowStart; r <= mergeRowEnd && !collision; r++) {
				for (let c = colStart; c <= mergeColEnd; c++) {
					const existing = occupancy.get(`${r},${c}`);
					if (existing) { collision = { existing, row: r, column: c }; break; }
				}
			}
			if (collision && text) {
				const existing = collision.existing;
				const position = existing.sheetPosition;
				// Compare source values, not ExcelJS's serialized rich-text objects.
				if (text === textFor(existing) && position.rowStart === rowStart
					&& position.rowEnd === mergeRowEnd && position.colStart === colStart
					&& position.colEnd === mergeColEnd) {
					cell.sheetPosition = { ...position };
					continue;
				}
				const error = new Error('Overlapping non-empty cells at ' + collision.row + ','
					+ collision.column + '; adjust the HTML table structure or edgeTolerancePx');
				error.code = 'XLSX_CELL_OVERLAP';
				error.conflicts = [existing, cell].map(region => ({
					rowStart: region.rowStart + 1, rowEnd: region.rowEnd,
					colStart: region.colStart + 1, colEnd: region.colEnd,
					bounds: region.bounds,
					tag: region.element?.tagName, id: region.element?.id || undefined,
					valueType: typeof region.value, textLength: textFor(region).length
				}));
				throw error;
			}
			if (collision) {
				// Keep blank-region decoration on free grid cells without merging across content.
				const excludedCells = new Set();
				for (let r = rowStart; r <= mergeRowEnd; r++) {
					for (let c = colStart; c <= mergeColEnd; c++) {
						if (occupancy.has(`${r},${c}`)) {
							excludedCells.add(`${r},${c}`);
							continue;
						}
						const target = worksheet.getCell(r, c);
						if (cell.styles.fill) target.fill = cell.styles.fill;
					}
				}
				blankBorderExclusions.set(cell, excludedCells);
				cell.sheetPosition = { rowStart, rowEnd: mergeRowEnd, colStart, colEnd: mergeColEnd };
				continue;
			}
			if (mergeRowEnd > rowStart || mergeColEnd > colStart) {
				worksheet.mergeCells(rowStart, colStart, mergeRowEnd, mergeColEnd);
			}

			const targetCell = worksheet.getCell(rowStart, colStart);
			targetCell.value = cell.richText ? { richText: cell.richText } : cell.value;
			if (cell.styles.font) targetCell.font = cell.styles.font;
			if (cell.styles.alignment) targetCell.alignment = cell.styles.alignment;
			if (cell.styles.fill) targetCell.fill = cell.styles.fill;
			if (cell.styles.numFmt) targetCell.numFmt = cell.styles.numFmt;
			if (typeof cell.value === 'number' && cell.styles.numFmt) {
				// Currency codes can be wider than the HTML symbol. Keep values and the grid intact.
				targetCell.alignment = { ...targetCell.alignment, wrapText: false, shrinkToFit: true };
			}

			for (let r = rowStart; r <= mergeRowEnd; r++) {
				for (let c = colStart; c <= mergeColEnd; c++) {
					occupancy.set(`${r},${c}`, cell);
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
		const instructions = Array.isArray(model.deferredBorders) ? [...model.deferredBorders] : [];
		// Write exact perimeters after merging. Explicit cell borders take precedence
		// over enclosing table borders without extending onto neighboring regions.
		for (const cell of model.cells) {
			if (!cell.sheetPosition) continue;
			for (const [side, style] of Object.entries(cell.styles.border || {})) {
				if (style) instructions.push({ ...cell.sheetPosition, side, style,
					excludedCells: blankBorderExclusions.get(cell) });
			}
		}
		if (!instructions.length) return;

		const applyBorderSide = (cell, side, borderStyle, excludedCells) => {
			if (!cell || !borderStyle || excludedCells?.has(`${cell.row},${cell.col}`)) return;
			const existing = cell.border && typeof cell.border === 'object' ? { ...cell.border } : {};
			existing[side] = { ...borderStyle };
			// ExcelJS merged cells share a style object. Detach it before changing
			// one edge so a partial border cannot spread across the merged region.
			cell.style = { ...cell.style, border: existing };
		};

		for (const instruction of instructions) {
			if (!instruction || !instruction.style) continue;
			const { side, style, rowStart, rowEnd, colStart, colEnd, excludedCells } = instruction;
			if (!side || rowStart == null || rowEnd == null || colStart == null || colEnd == null) continue;
			const safeRowStart = Math.max(1, Math.floor(rowStart));
			const safeRowEnd = Math.max(safeRowStart, Math.floor(rowEnd));
			const safeColStart = Math.max(1, Math.floor(colStart));
			const safeColEnd = Math.max(safeColStart, Math.floor(colEnd));

			switch (side) {
				case 'left':
				for (let r = safeRowStart; r <= safeRowEnd; r++) {
					applyBorderSide(worksheet.getCell(r, safeColStart), 'left', style, excludedCells);
				}
				break;
				case 'right':
				for (let r = safeRowStart; r <= safeRowEnd; r++) {
					applyBorderSide(worksheet.getCell(r, safeColEnd), 'right', style, excludedCells);
				}
				break;
				case 'top':
				for (let c = safeColStart; c <= safeColEnd; c++) {
					applyBorderSide(worksheet.getCell(safeRowStart, c), 'top', style, excludedCells);
				}
				break;
				case 'bottom':
				for (let c = safeColStart; c <= safeColEnd; c++) {
					applyBorderSide(worksheet.getCell(safeRowEnd, c), 'bottom', style, excludedCells);
				}
				break;
				default:
				break;
			}
		}
	};

	const fitMultilineText = (worksheet, model) => {
		if (!CONFIG.PRESERVE_RICH_TEXT) return;
		for (const cell of model.cells) {
			if (typeof cell.value !== 'string' || !cell.styles.alignment?.wrapText) continue;
			const lines = cell.value.trim().split('\n').length;
			if (lines < 2) continue;
			const fontSize = cell.styles.font?.size || CONFIG.NORMAL_FONT?.size || 11;
			// Native Excel Arial 9 AutoFit uses 12pt per line plus one pixel.
			// Fixed merged rows do not AutoFit; reserve this height explicitly.
			const minimum = lines * Math.ceil(fontSize * 4 / 3 / 0.75) * 0.75 + 0.75;
			let height = 0;
			for (let r = cell.rowStart + 1; r <= cell.rowEnd; r++) height += worksheet.getRow(r).height;
			if (height < minimum) worksheet.getRow(cell.rowEnd).height += minimum - height;
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
		fitMultilineText(worksheet, layoutModel);
		applyDeferredBorders(worksheet, layoutModel);
		applyPageBreaks(worksheet, layoutModel);
		if (!worksheet.pageSetup.printArea && layoutModel.columnWidthsPx.length && layoutModel.rowHeightsPx.length) {
			worksheet.pageSetup.printArea = 'A1:' + worksheet.getColumn(layoutModel.columnWidthsPx.length).letter
				+ layoutModel.rowHeightsPx.length;
		}
		fitMarkedPagesForPrint(worksheet);
		return worksheet;
	};

	const processAndPlaceImages = async (workbook, worksheet, layoutModel) => {
		if (!layoutModel.images.length) return;
		debugLog('Converting images...');
		const imageResults = await mapWithConcurrency(layoutModel.images, CONFIG.IMAGE_CONCURRENCY, convertImageJob);
		const successfulImages = imageResults.filter((item) => item && item.base64);
		const failedImages = imageResults.filter((item) => item && !item.base64);
		if (successfulImages.length) {
			placeImages(worksheet, workbook, successfulImages, layoutModel);
			debugLog('Embedded images', successfulImages.length);
		}
		if (failedImages.length) {
			failedImages.forEach((item) => {
				warnLog('Image skipped (no base64)', {
					code: 'IMAGE_SKIPPED',
					source: item.element?.currentSrc || item.element?.src || '(unknown src)',
					width: item.width,
					height: item.height,
					error: item.error ? item.error.message || String(item.error) : '(no error message)'
				});
			});
			warnLog('Images could not be embedded', { code: 'IMAGE_EMBED_FAILED', count: failedImages.length });
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
			warnLog('Failed to inspect workbook model', { code: 'WORKBOOK_INSPECTION_FAILED', error: err });
		}
	};

	const writeWorkbookToBuffer = async (workbook) => {
		const xlsx = workbook.xlsx;
		const originalAddStyles = xlsx.addStyles;
		try {
			if (CONFIG.NORMAL_FONT) {
				// ExcelJS 4.4 has no public Normal-style setter. Limit this compatibility
				// hook to one writer instance and preserve every explicit cell font.
				if (typeof originalAddStyles !== 'function') throw new Error('ExcelJS Normal-style serialization is unsupported');
				xlsx.addStyles = async (zip, model) => {
					const styles = model.styles;
					const fonts = styles?.model?.fonts;
					if (!Array.isArray(fonts) || !fonts.length || typeof styles.map?.font?.toXml !== 'function') {
						throw new Error('ExcelJS Normal-style serialization is unsupported');
					}
					// Desktop Excel resolves its built-in Normal style from font zero.
					// Extracted CSS cell fonts have their own records, including explicit
					// Calibri cells; changing a named-style XF alone is ignored by Excel.
					fonts[0] = styles.map.font.toXml(CONFIG.NORMAL_FONT);
					const xml = styles.xml;
					const normal = /<cellStyle\b[^>]*\bname="Normal"[^>]*\/>/;
					if (!normal.test(xml)) throw new Error('ExcelJS Normal-style definition is missing');
					zip.append(xml.replace(normal, tag => tag.replace(/\/>$/, ' customBuiltin="1"/>')), { name: 'xl/styles.xml' });
				};
			}
			logWorkbookMediaSummary(workbook);
			debugLog('Writing workbook buffer');
			const buffer = await workbook.xlsx.writeBuffer();
			debugLog('Workbook buffer size', buffer && (buffer.byteLength || buffer.length || 0));
			return buffer;
		} catch (err) {
			errorLog('writeBuffer failed', { code: 'WORKBOOK_WRITE_FAILED', error: err });
			throw err;
		} finally {
			if (CONFIG.NORMAL_FONT) xlsx.addStyles = originalAddStyles;
		}
	};

	const sanitizeSheetName = (name) => (name || 'Sheet1').replace(/[\\/*?:\[\]]/g, '').slice(0, 31) || 'Sheet1';

 return { applyColumnWidths, applyRowHeights, applyPageBreaks, writeCellsToSheet, applyDeferredBorders, applyPageSetupOptions, createWorkbook, prepareWorksheet, processAndPlaceImages, logWorkbookMediaSummary, writeWorkbookToBuffer, sanitizeSheetName };
}

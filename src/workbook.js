// MIT License. Copyright (c) 2026 YAP WEI JUN.
export function createWorkbookWriter(dependencies) {
 const { CONFIG, EXPORT_SCALE, ExcelJS, FONT_SCALE, ROW_SCALE, convertImageJob, debugLog, errorLog, placeImages, pxToColumnWidth, pxToPoint, warnLog, mapWithConcurrency } = dependencies;
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
			if (occupancy.has(key)) {
                const existing = worksheet.getCell(rowStart, colStart).value;
                const text = String(cell.value ?? '').trim();
                if (text && text !== String(existing ?? '').trim()) {
                    throw new Error('Overlapping non-empty cells at ' + key + '; adjust the HTML table structure or edgeTolerancePx');
                }
                continue;
            }

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

	const sanitizeSheetName = (name) => (name || 'Sheet1').replace(/[\\/*?:\[\]]/g, '').slice(0, 31) || 'Sheet1';

 return { applyColumnWidths, applyRowHeights, applyPageBreaks, writeCellsToSheet, applyDeferredBorders, applyPageSetupOptions, createWorkbook, prepareWorksheet, processAndPlaceImages, logWorkbookMediaSummary, writeWorkbookToBuffer, sanitizeSheetName };
}

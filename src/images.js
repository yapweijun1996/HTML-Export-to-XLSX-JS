// MIT License. Copyright (c) 2026 YAP WEI JUN.
export function createImages(dependencies) {
  const { CONFIG, IMAGE_SCALE, columnWidthToPixels, debugLog, errorLog, isFiniteNumber, snapToPrecision, toPositiveNumber, warnLog } = dependencies;
	const MAX_CANVAS_AREA = 16777216, MAX_CANVAS_SIDE = 16384;
	const imageIds = new WeakMap();
	const getImageId = (workbook, job) => {
		let formats = imageIds.get(workbook);
		if (!formats) imageIds.set(workbook, formats = new Map());
		let content = formats.get(job.extension);
		if (!content) formats.set(job.extension, content = new Map());
		if (!content.has(job.base64)) {
			content.set(job.base64, workbook.addImage({ base64: job.base64, extension: job.extension }));
		}
		return content.get(job.base64);
	};

	const getImageRange = (worksheet, tl, ext) => {
		// ExcelJS 4.4 writes an invalid editAs on oneCellAnchor. A twoCellAnchor
		// with editAs="oneCell" keeps the same move-with-cell, fixed-size behavior.
		// Both markers use physical EMUs; ExcelJS fractional offsets otherwise use
		// width/height * 10000, which is not the worksheet's physical cell size.
		const spanEmu = (index, isColumn) => {
			const current = isColumn ? worksheet.getColumn(index + 1) : worksheet.getRow(index + 1);
			return isColumn
				? columnWidthToPixels(toPositiveNumber(current.width) || worksheet.properties.defaultColWidth || 8.43) * 9525
				: Math.round((toPositiveNumber(current.height) || worksheet.properties.defaultRowHeight || 15) * 12700);
		};
		const advance = (index, offset, isColumn) => {
			const limit = isColumn ? 16383 : 1048575;
			for (;;) {
				const span = spanEmu(index, isColumn);
				if (offset < span) return { index, offset };
				if (index >= limit) throw new RangeError('Image extends beyond the worksheet boundary');
				offset -= span;
				index++;
			}
		};
		const start = (coordinate, isColumn) => {
			const index = Math.floor(coordinate);
			let offset = Math.round((coordinate - index) * spanEmu(index, isColumn));
			// Excel prints a thin strip on the preceding page when a picture starts
			// exactly on a manual break. A two-pixel inset keeps it on its own page.
			if (!isColumn && offset < 2 * 9525 && worksheet.rowBreaks?.some(brk => brk.id === index)) {
				offset = 2 * 9525;
			}
			return advance(index, offset, isColumn);
		};
		const fromCol = start(tl.col, true), fromRow = start(tl.row, false);
		const col = advance(fromCol.index, fromCol.offset + Math.round(ext.width * 9525), true);
		const row = advance(fromRow.index, fromRow.offset + Math.round(ext.height * 9525), false);
		return { tl: { nativeCol: fromCol.index, nativeColOff: fromCol.offset,
			nativeRow: fromRow.index, nativeRowOff: fromRow.offset }, br: { nativeCol: col.index, nativeColOff: col.offset,
			nativeRow: row.index, nativeRowOff: row.offset }, editAs: 'oneCell' };
	};
	// A wait that failed once in this export is not repeated, so an image that never loads,
	// such as an off-screen lazy image, costs one timeout rather than one per stage.
	const failedLoads = new WeakMap();
	const ensureImageLoaded = (img) => {
		if (img.complete) {
			if (img.naturalWidth && img.naturalHeight) return Promise.resolve();
			return Promise.reject(new Error('Image failed to load (already complete with zero size)'));
		}
		if (failedLoads.has(img)) return Promise.reject(failedLoads.get(img));
		return new Promise((resolve, reject) => {
			const fail = (error) => {
				cleanup();
				failedLoads.set(img, error);
				reject(error);
			};
			const timer = setTimeout(() => fail(new Error('Image load timed out')), CONFIG.IMAGE_TIMEOUT_MS);
			const cleanup = () => {
				clearTimeout(timer);
				img.removeEventListener('load', onLoad);
				img.removeEventListener('error', onError);
			};
			const onLoad = () => {
				cleanup();
				resolve();
			};
			// The error event carries no message; report an Error so warnings stay readable.
			const onError = () => fail(new Error('Image failed to load'));
			img.addEventListener('load', onLoad, { once: true });
			img.addEventListener('error', onError, { once: true });
		});
	};

	const fetchImageAsBase64 = async (src, width, height) => {
		if (!src) return null;
		let url;
		try {
			url = new URL(src, window.location.href);
		} catch (err) {
			warnLog('Invalid image URL, skip fetch fallback', { code: 'IMAGE_URL_INVALID', source: src, error: err });
			return null;
		}

		if (url.protocol === 'file:') {
			warnLog('Skipping fetch fallback for file:// URL (browser security restriction)', { code: 'IMAGE_FILE_UNSUPPORTED', source: url.href });
			return null;
		}

		try {
			const response = await fetch(url.href, { mode: 'cors', signal: AbortSignal.timeout(CONFIG.IMAGE_TIMEOUT_MS) });
			if (!response.ok) throw new Error('Image fetch returned HTTP ' + response.status);
			const blob = await response.blob();
			const objectUrl = URL.createObjectURL(blob);
			try {
				// Rasterize fetched images too: SVG/WebP MIME suffixes are not ExcelJS image extensions.
				const image = new Image();
				image.src = objectUrl;
				await ensureImageLoaded(image);
				const canvas = document.createElement('canvas');
				canvas.width = width || image.naturalWidth;
				canvas.height = height || image.naturalHeight;
				const context = canvas.getContext('2d');
				if (!context) throw new Error('Canvas 2D context unavailable');
				context.drawImage(image, 0, 0, canvas.width, canvas.height);
				const base64 = canvas.toDataURL('image/png').split(',')[1];
				if (!base64) throw new Error('Empty base64 from fetched image');
				return { base64, extension: 'png' };
			} finally { URL.revokeObjectURL(objectUrl); }

		} catch (err) {
			warnLog('Fallback fetch-to-base64 failed', { code: 'IMAGE_FETCH_FAILED', source: src, error: err });
			return null;
		}
	};

	const convertImageJob = async (job) => {
		const { element, width, height } = job;
		if (!element) return { ...job, error: new Error('Image element missing') };

		try {
			await ensureImageLoaded(element);
		} catch (err) {
			warnLog('Image failed to load', { code: 'IMAGE_LOAD_FAILED', source: element.currentSrc || element.src, error: err });
			return { ...job, error: err };
		}

		const imageScale = toPositiveNumber(IMAGE_SCALE) || 1;
		// Stay within the smallest common canvas limits (Safari allows 16,777,216 pixels).
		// Only the raster resolution drops; the drawing keeps its measured size.
		const fit = Math.min(1, Math.sqrt(MAX_CANVAS_AREA / Math.max(1, width * height * imageScale * imageScale)),
			MAX_CANVAS_SIDE / Math.max(1, width * imageScale), MAX_CANVAS_SIDE / Math.max(1, height * imageScale));
		const rasterSize = value => Math.max(1, fit < 1 ? Math.floor(value * imageScale * fit) : Math.round(value * imageScale));
		const renderWidth = rasterSize(width);
		const renderHeight = rasterSize(height);
		if (fit < 1) debugLog('Reduced image raster to the canvas limit', `${renderWidth}x${renderHeight}`);

		const canvas = document.createElement('canvas');
		canvas.width = renderWidth;
		canvas.height = renderHeight;
		const ctx = canvas.getContext('2d');
		if (!ctx) {
			const err = new Error('Canvas 2D context unavailable');
			errorLog(err.message, { code: 'IMAGE_CANVAS_UNAVAILABLE', source: element.currentSrc || element.src });
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
			warnLog('Canvas toDataURL failed, trying fetch fallback', { code: 'IMAGE_CANVAS_FALLBACK', source: element.currentSrc || element.src, error: err });
			const fallback = await fetchImageAsBase64(element.currentSrc || element.src, renderWidth, renderHeight);
			if (fallback) {
				debugLog('Fallback fetch succeeded for image', element.src, `base64Length=${fallback.base64.length}`);
				return { ...job, base64: fallback.base64, extension: fallback.extension };
			}
			errorLog('Image conversion failed after fallback', { code: 'IMAGE_CONVERSION_FAILED', source: element.currentSrc || element.src, error: err });
			return { ...job, error: err };
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

	const getCellImageAlignment = (cellModel, includeCss = true) => {
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
		if (!includeCss) return null;
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
			topPx: topEdge,
			columnEdges: columnEdges.slice(cellModel.colStart, cellModel.colEnd + 1).map(edge => edge - leftEdge),
			rowEdges: rowEdges.slice(cellModel.rowStart, cellModel.rowEnd + 1).map(edge => edge - topEdge),
			startColIndex: sheetPosition.colStart - 1,
			startRowIndex: sheetPosition.rowStart - 1
		};
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

	const coordinateAt = (edges, offset, startIndex = 0) => {
		const position = locateIndexAndOffset(edges, offset, true);
		const span = edges[position.index + 1] - edges[position.index];
		return startIndex + position.index + (span > 0 ? position.offset / span : 0);
	};

	const ALIGNMENT_OVERFLOW = 'Image alignment exceeds its cell; enlarge the HTML cell, reduce imageWidthScale, or use precise placement';
	const fitsCell = (metrics, left, top, width, height) => {
		const tolerance = CONFIG.EDGE_TOLERANCE_PX;
		return left + width <= metrics.widthPx + tolerance && top + height <= metrics.heightPx + tolerance;
	};

	// Alignment sets the horizontal flow. Vertically the flow starts at the images' rendered top,
	// so it does not cover text above them, and moves up only as far as the cell requires.
	const flowTop = (images, metrics, flowHeight, fallback = 0) => {
		const measured = Math.min(...images.map(job => job.top)) - metrics.topPx;
		if (!Number.isFinite(measured)) return fallback;
		return Math.max(0, Math.min(measured, metrics.heightPx - flowHeight));
	};

	const commitPlacements = (worksheet, workbook, metrics, placements, forced) => {
		if (!placements.every(item => fitsCell(metrics, item.left, item.top, item.width, item.height))) {
			// Reject an impossible forced flow rather than clamping several images to the same point.
			if (forced) throw new Error(ALIGNMENT_OVERFLOW);
			warnLog('Images exceed their cell flow; placed at their rendered positions',
				{ code: 'IMAGE_ALIGNMENT_FALLBACK', count: placements.length });
			return false;
		}
		for (const item of placements) {
			addAlignedImage(worksheet, workbook, item.job, metrics, item.left, item.top, item.width, item.height);
		}
		return true;
	};

	const addAlignedImage = (worksheet, workbook, job, metrics, offsetLeftPx, offsetTopPx, widthPx, heightPx) => {
		if (!metrics || widthPx <= 0 || heightPx <= 0) return false;
		if (!fitsCell(metrics, offsetLeftPx, offsetTopPx, widthPx, heightPx)) throw new Error(ALIGNMENT_OVERFLOW);
		const safeLeft = Math.max(0, offsetLeftPx || 0);
		const safeTop = Math.max(0, offsetTopPx || 0);
		const tl = {
			col: coordinateAt(metrics.columnEdges, safeLeft, metrics.startColIndex),
			row: coordinateAt(metrics.rowEdges, safeTop, metrics.startRowIndex)
		};

		const ext = {
			width: widthPx,
			height: heightPx
		};
		const range = getImageRange(worksheet, tl, ext);
		worksheet.addImage(getImageId(workbook, job), range);
		debugLog('Placed image using alignment flow', job.element?.src || '(unknown src)', {
			tl,
			ext,
			offset: { left: safeLeft, top: safeTop },
			cellSize: { widthPx: metrics.widthPx, heightPx: metrics.heightPx }
		});
		logLatestWorksheetImage(worksheet);
		return true;
	};

	const placeImagesInline = (worksheet, workbook, images, cellModel, layoutModel, justify, forced = true) => {
		const metrics = getCellPlacementMetrics(cellModel, layoutModel);
		if (!metrics || metrics.widthPx <= 0 || metrics.heightPx <= 0) {
			return false;
		}
		const spacingPx = Math.max(0, snapToPrecision(CONFIG.IMAGE_SPACING_PX || 0));
		const finalImageScale = toPositiveNumber(CONFIG.IMAGE_WIDTH_SCALE) || 1.0;

		const rows = [];
		let currentRow = { items: [], width: 0, height: 0 };

		for (const job of images) {
			const widthPx = Math.max(1, job.width * finalImageScale);
			const heightPx = Math.max(1, job.height * finalImageScale);
			const item = { job, widthPx, heightPx };
			const needsWrap =
			currentRow.items.length > 0 && currentRow.width + spacingPx + widthPx > metrics.widthPx;
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

		const flowHeight = rows.reduce((sum, row) => sum + row.height, 0) + spacingPx * Math.max(0, rows.length - 1);
		let currentTop = flowTop(images, metrics, flowHeight);
		const placements = [];
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
				placements.push({ job: item.job, left: cursor, top: currentTop, width: item.widthPx, height: item.heightPx });
				cursor += item.widthPx + spacingPx;
			}
			currentTop += row.height + spacingPx;
		}
		return commitPlacements(worksheet, workbook, metrics, placements, forced);
	};

	const placeImagesStacked = (worksheet, workbook, images, cellModel, layoutModel, forced = true) => {
		const metrics = getCellPlacementMetrics(cellModel, layoutModel);
		if (!metrics || metrics.widthPx <= 0 || metrics.heightPx <= 0) {
			return false;
		}
		const spacingPx = Math.max(0, snapToPrecision(CONFIG.IMAGE_SPACING_PX || 0));
		const finalImageScale = toPositiveNumber(CONFIG.IMAGE_WIDTH_SCALE) || 1.0;
		const sizes = images.map(job => ({ job, width: Math.max(1, job.width * finalImageScale),
			height: Math.max(1, job.height * finalImageScale) }));
		const flowHeight = sizes.reduce((sum, item) => sum + item.height, 0) + spacingPx * Math.max(0, sizes.length - 1);
		let currentTop = flowTop(images, metrics, flowHeight);
		const placements = [];
		for (const item of sizes) {
			placements.push({ ...item, left: 0, top: currentTop });
			currentTop += item.height + spacingPx;
		}
		return commitPlacements(worksheet, workbook, metrics, placements, forced);
	};

	const placeSingleImageCentered = (worksheet, workbook, job, cellModel, layoutModel, forced = true) => {
		const metrics = getCellPlacementMetrics(cellModel, layoutModel);
		if (!metrics || metrics.widthPx <= 0 || metrics.heightPx <= 0) {
			return false;
		}
		const finalImageScale = toPositiveNumber(CONFIG.IMAGE_WIDTH_SCALE) || 1.0;
		const widthPx = Math.max(1, job.width * finalImageScale);
		const heightPx = Math.max(1, job.height * finalImageScale);
		const offsetLeft = Math.max(0, (metrics.widthPx - widthPx) / 2);
		// Without a rendered position, center vertically as well.
		const offsetTop = flowTop([job], metrics, heightPx, Math.max(0, (metrics.heightPx - heightPx) / 2));
		return commitPlacements(worksheet, workbook, metrics,
			[{ job, left: offsetLeft, top: offsetTop, width: widthPx, height: heightPx }], forced);
	};

	// A forced flow (cell attribute or explicit imageAlignment) rejects when it cannot fit;
	// otherwise the caller places the images at their rendered positions.
	const placeImagesWithAlignment = (worksheet, workbook, images, cellModel, layoutModel, alignment, forced = true) => {
		const keyword = alignment || 'inline';
		if (keyword === 'precise') return false;
		if (images.length === 1) {
			if (keyword === 'center') {
				return placeSingleImageCentered(worksheet, workbook, images[0], cellModel, layoutModel, forced);
			}
			const justify = keyword === 'right' ? 'end' : keyword === 'center' ? 'center' : 'start';
			return placeImagesInline(worksheet, workbook, images, cellModel, layoutModel, justify, forced);
		}

		switch (keyword) {
			case 'block':
			return placeImagesStacked(worksheet, workbook, images, cellModel, layoutModel, forced);
			case 'right':
			return placeImagesInline(worksheet, workbook, images, cellModel, layoutModel, 'end', forced);
			case 'center':
			return placeImagesInline(worksheet, workbook, images, cellModel, layoutModel, 'center', forced);
			case 'left':
			case 'inline':
			default:
			return placeImagesInline(worksheet, workbook, images, cellModel, layoutModel, 'start', forced);
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
			const parentCell = job.parentCell;
			if (!parentCell?.sheetPosition) continue;
			const tl = {
				col: coordinateAt(htmlColumnEdges, job.left),
				row: coordinateAt(htmlRowEdges, job.top)
			};

			// The extension (size) of the image should also be based on the original pixel size
			const finalImageScale = toPositiveNumber(CONFIG.IMAGE_WIDTH_SCALE) || 1.0;
			const ext = {
				width: job.width * finalImageScale,
				height: job.height * finalImageScale,
			};

			const range = getImageRange(worksheet, tl, ext);
			worksheet.addImage(getImageId(workbook, job), range);

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
		const preciseFallbackJobs = [];
		for (const job of resolvedJobs) {
			const parentCell = job.parentCell;
			if (!parentCell || !parentCell.sheetPosition) {
				continue;
			}
			// An image outside table cells has no cell flow; its rendered position is authoritative.
			if (parentCell.outsideCell) {
				preciseFallbackJobs.push(job);
				continue;
			}
			if (!cellGroups.has(parentCell)) {
				cellGroups.set(parentCell, []);
			}
			cellGroups.get(parentCell).push(job);
		}

		const defaultAlignment = toImageAlignment(CONFIG.IMAGE_ALIGNMENT) || 'inline';

		for (const [cellModel, images] of cellGroups.entries()) {
			if (!Array.isArray(images) || !images.length) continue;
			const cellAlignment = getCellImageAlignment(cellModel, !CONFIG.IMAGE_ALIGNMENT_EXPLICIT) || defaultAlignment;
			const forced = CONFIG.IMAGE_ALIGNMENT_EXPLICIT || Boolean(getCellImageAlignment(cellModel, false));
			const placed = placeImagesWithAlignment(worksheet, workbook, images, cellModel, layoutModel, cellAlignment, forced);
			if (!placed) {
				preciseFallbackJobs.push(...images);
			}
		}

		// Process any images we could not auto-align via the original precise placement logic.
		if (preciseFallbackJobs.length) {
			placeImagesPrecisely(worksheet, workbook, preciseFallbackJobs, layoutModel);
		}
	};



  return { ensureImageLoaded, fetchImageAsBase64, convertImageJob, locateIndexAndOffset, toImageAlignment, getCellImageAlignment, getCellPlacementMetrics, logLatestWorksheetImage, addAlignedImage, placeImagesInline, placeImagesStacked, placeSingleImageCentered, placeImagesWithAlignment, placeImagesPrecisely, placeImages };
}

// Bounded concurrency prevents large documents from converting every image at once.
export async function mapWithConcurrency(items, limit, mapper) {
    const results = new Array(items.length);
    let next = 0;
    await Promise.all(Array.from({ length: Math.min(limit, items.length) }, async () => {
        while (next < items.length) {
            const index = next++;
            results[index] = await mapper(items[index], index);
        }
    }));
    return results;
}

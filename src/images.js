// MIT License. Copyright (c) 2026 YAP WEI JUN.
export function createImages(dependencies) {
  const { CONFIG, IMAGE_SCALE, debugLog, errorLog, isFiniteNumber, snapToPrecision, toPositiveNumber, warnLog } = dependencies;
	const ensureImageLoaded = (img) => {
		if (img.complete) {
			if (img.naturalWidth && img.naturalHeight) return Promise.resolve();
			return Promise.reject(new Error('Image failed to load (already complete with zero size)'));
		}
		return new Promise((resolve, reject) => {
			const timer = setTimeout(() => { cleanup(); reject(new Error('Image load timed out')); }, CONFIG.IMAGE_TIMEOUT_MS);
			const cleanup = () => {
				clearTimeout(timer);
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
			const response = await fetch(url.href, { mode: 'cors', signal: AbortSignal.timeout(CONFIG.IMAGE_TIMEOUT_MS) });
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



  return { ensureImageLoaded, blobToBase64, fetchImageAsBase64, convertImageJob, locateIndexAndOffset, clampOffset, toImageAlignment, getCellImageAlignment, getCellPlacementMetrics, clampWithin, logLatestWorksheetImage, addAlignedImage, placeImagesInline, placeImagesStacked, placeSingleImageCentered, placeImagesWithAlignment, placeImagesPrecisely, placeImages };
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

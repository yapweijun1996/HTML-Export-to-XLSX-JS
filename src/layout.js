// MIT License. Copyright (c) 2026 YAP WEI JUN.
export function createLayout(dependencies) {
  const { CONFIG, collectDomSnapshot, debugLog, extractAlignment, extractBorders, extractFont, extractRichText, extractBackground, getCellText, parseCellValue, isElementExcluded, snapToPrecision, findEdgeIndex } = dependencies;
	const buildLayoutModel = (root, preserveEdges = false) => {
		const { rootLeft, rootTop, rootWidth, rootHeight, allCells, nestedCells, leafCells, tableElements, textFragments, pageBreakMarkers, columnEdges, rowEdges } = collectDomSnapshot(root, preserveEdges);

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

			const colStartIndex = findEdgeIndex(columnEdges, left);
			const colEndIndex = findEdgeIndex(columnEdges, right);
			const rowStartIndex = findEdgeIndex(rowEdges, top);
			const rowEndIndex = findEdgeIndex(rowEdges, bottom);

			if (colStartIndex === -1 || colEndIndex === -1 || rowStartIndex === -1 || rowEndIndex === -1) {
				dependencies.warnLog('A cell has unmatched grid edges', { code: 'GRID_EDGE_MISMATCH' });
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
				const nearest = findEdgeIndex(edges, value);
				if (nearest !== -1) return nearest;
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

			const textValue = getCellText(cell, root);
			const cs = window.getComputedStyle(cell);
			const font = extractFont(cell, cs);
			const alignment = extractAlignment(cell, cs, textValue);
			const borders = extractBorders(cs, cell);
			const fillColor = extractBackground(cell, root);
			const currencyMatch = parseCellValue(textValue, cell);
			const exportedValue = currencyMatch.value;

			const cellModel = {
				element: cell,
				value: exportedValue,
				richText: extractRichText(cell, exportedValue),
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


				const referenceElement = fragment.sourceElement || fragment.cell;
				const cs = window.getComputedStyle(referenceElement);
				const font = extractFont(referenceElement, cs);
				const alignment = extractAlignment(referenceElement, cs, fragment.text);
				const fillColor = extractBackground(referenceElement, root);
				const currencyMatch = parseCellValue(fragment.text, referenceElement);
				const exportedValue = currencyMatch.value;

				cells.push({
					element: fragment.cell,
					value: exportedValue,
					richText: extractRichText(referenceElement, exportedValue),
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

		// Image ownership is independent of text fragments: nested table containers still own images.
		const imageOwners = new Map();
		for (const img of root.querySelectorAll('img')) {
			if (isElementExcluded(img, root)) continue;
			const rect = img.getBoundingClientRect();
			if (rect.width <= 0 || rect.height <= 0) continue;
			const left = snapToPrecision(rect.left + window.scrollX - rootLeft);
			const top = snapToPrecision(rect.top + window.scrollY - rootTop);
			const ownerCell = img.closest('td, th');
			const owner = ownerCell && root.contains(ownerCell) ? ownerCell : null;
			let parentCell = owner && imageOwners.get(owner);
			if (!owner) {
				// Images outside table cells, such as a letterhead logo, keep their rendered position.
				parentCell = { element: img, outsideCell: true, bounds: { left, top, width: rect.width, height: rect.height },
					sheetPosition: { colStart: resolveEdgeIndex(columnEdges, left, 0) + 1, rowStart: resolveEdgeIndex(rowEdges, top, 0) + 1 } };
			} else if (!parentCell) {
				const descriptor = getDescriptor(owner);
				if (!descriptor) continue;
				parentCell = {
					element: owner, bounds: { left: descriptor.left, top: descriptor.top, width: descriptor.width, height: descriptor.height },
					colStart: descriptor.colStartIndex, colEnd: Math.max(descriptor.colStartIndex + 1, descriptor.colEndIndex),
					rowStart: descriptor.rowStartIndex, rowEnd: Math.max(descriptor.rowStartIndex + 1, descriptor.rowEndIndex),
					styles: { alignment: extractAlignment(owner, window.getComputedStyle(owner), '') },
					sheetPosition: { colStart: descriptor.colStartIndex + 1, rowStart: descriptor.rowStartIndex + 1 }
				};
				imageOwners.set(owner, parentCell);
			}
			images.push({ element: img, parentCell, width: rect.width, height: rect.height, left, top });
		}

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

		const applyOuterElementBorders = (descriptor, borders) => {
			if (!descriptor || !borders) return;
			// Outer borders belong to their measured perimeter, not nearby child cells.
			for (const side of ['top', 'bottom', 'left', 'right']) {
				pushDeferredBorderInstruction(descriptor, side, borders[side]);
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
			if (!tableElement) continue;
			const descriptor = getDescriptor(tableElement);
			if (!descriptor) continue;
			const csTable = window.getComputedStyle(tableElement);
			const borders = extractBorders(csTable, tableElement);
			if (borders) {
				applyOuterElementBorders(descriptor, borders);
			}

		}

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
			const scrollY = window.scrollY || window.pageYOffset || 0;
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


  return { buildLayoutModel };
}

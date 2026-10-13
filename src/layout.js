// MIT License. Copyright (c) 2026 YAP WEI JUN.
export function createLayout(dependencies) {
  const { CONFIG, collectDomSnapshot, computeEffectiveHorizontalPadding, computeEffectiveVerticalPadding, debugLog, extractAlignment, extractBorders, extractFont, extractRichText, extractBackground, getCellText, parseCellValue, isElementExcluded, snapToPrecision, findEdgeIndex } = dependencies;
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
			const owner = img.closest('td, th');
			if (!owner || !root.contains(owner)) continue;
			let parentCell = imageOwners.get(owner);
			if (!parentCell) {
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
			images.push({ element: img, parentCell, width: rect.width, height: rect.height,
				left: snapToPrecision(rect.left + window.scrollX - rootLeft),
				top: snapToPrecision(rect.top + window.scrollY - rootTop) });
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
			if (!tableElement) continue;
			const descriptor = getDescriptor(tableElement);
			if (!descriptor) continue;
			const csTable = window.getComputedStyle(tableElement);
			const borders = extractBorders(csTable, tableElement);
			if (borders) {
				applyOuterElementBorders(descriptor, borders);
			}

		}

		propagateSharedBorders(cells);

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

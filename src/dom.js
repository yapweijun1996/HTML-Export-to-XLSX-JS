// MIT License. Copyright (c) 2026 YAP WEI JUN.
export function createDom(dependencies) {
  const { CONFIG, debugLog, snapToPrecision, normalizeEdges } = dependencies;
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


	const collectDomSnapshot = (root) => {
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
		const pageBreakMarkers = Array.from(root.querySelectorAll(CONFIG.PAGE_BREAK_SELECTOR)).filter(
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

		return { rootLeft, rootTop, rootWidth, rootHeight, allCells, nestedCells, leafCells, tableElements, textFragments, pageBreakMarkers, columnEdges, rowEdges };
	};

  return { normalizedBlacklistSelectors, loggedBlacklistedElements, elementMatchesSelector, isElementBlacklisted, normalizeLooseText, collectTextFragmentsForCell, getCellText, collectDomSnapshot };
}

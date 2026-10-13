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

	const styleCache = new WeakMap();
	const computedStyle = element => {
		if (!styleCache.has(element)) styleCache.set(element, window.getComputedStyle(element));
		return styleCache.get(element);
	};

	const computeExcluded = (element, root) => {
		if (!element || isElementBlacklisted(element, root)) return true;
		const style = computedStyle(element);
		if (style.visibility === 'hidden' || style.visibility === 'collapse') return true;
		// Visibility may be restored by a child, but these ancestor properties hide the whole subtree.
		for (let current = element; current; current = current.parentElement) {
			const currentStyle = computedStyle(current);
			if (currentStyle.display === 'none' || currentStyle.contentVisibility === 'hidden'
				|| currentStyle.opacity === '0') return true;
		}
		return false;
	};
	let snapshotExclusions = null;
	const isElementExcluded = (element, root) => {
		if (!element) return true;
		if (!snapshotExclusions) return computeExcluded(element, root);
		const boundary = root || document.documentElement;
		if (!snapshotExclusions.has(boundary)) snapshotExclusions.set(boundary, new WeakMap());
		const cache = snapshotExclusions.get(boundary);
		if (!cache.has(element)) cache.set(element, computeExcluded(element, root));
		return cache.get(element);
	};

	const hasExportableContent = (element, root) => !isElementExcluded(element, root)
		|| Array.from(element.querySelectorAll('*')).some(child => !isElementExcluded(child, root));

	const textToken = node => {
		const style = computedStyle(node.parentElement);
		const preserve = ['pre', 'pre-wrap', 'break-spaces'].includes(style.whiteSpace);
		let text = node.textContent;
		if (!preserve) text = style.whiteSpace === 'pre-line'
			? text.replace(/[\t\r\f ]+/g, ' ') : text.replace(/[\t\r\n\f ]+/g, ' ');
		let language = node.parentElement.closest('[lang]')?.getAttribute('lang') || undefined;
		try { if (language) Intl.getCanonicalLocales(language); } catch { language = undefined; }
		if (style.textTransform === 'uppercase') text = text.toLocaleUpperCase(language);
		if (style.textTransform === 'lowercase') text = text.toLocaleLowerCase(language);
		return { text, preserve, capitalize: style.textTransform === 'capitalize', language };
	};

	const joinTextTokens = tokens => {
		let text = '';
		let preserveTail = false;
		for (const token of tokens) {
			if (token.block) {
				if (text && !text.endsWith('\n')) text = text.replace(/ +$/, '') + '\n';
				continue;
			}
			let part = token.text;
			if (token.capitalize) part = part.replace(/(^|[^\p{L}\p{M}\p{N}])\p{L}/gu,
				(match, boundary, index) => index === 0 && /[\p{L}\p{M}\p{N}]$/u.test(text)
					? match : match.toLocaleUpperCase(token.language));
			if (!token.preserve) {
				part = part.replace(/ *\n */g, '\n');
				if (!text || /[ \n]$/.test(text)) part = part.replace(/^ +/, '');
				if (part.startsWith('\n')) text = text.replace(/ +$/, '');
			}
			text += part;
			preserveTail = token.preserve;
		}
		return (preserveTail ? text : text.replace(/ +$/, '')).replace(/^\n+|\n+$/g, '').replace(/\u00a0/g, ' ');
	};

	const collectTextFragmentsForCell = (cell, metrics, root) => {
		if (!cell || isElementBlacklisted(cell, root)) return [];
		const { rootLeft = 0, rootTop = 0, scrollX = 0, scrollY = 0 } = metrics || {};
		const fragments = [];
		let active = null;
		const visit = (node, owner) => {
			if (node.nodeType === 1) {
				if (isElementBlacklisted(node, root)) return;
				// A table or another declared value separates runs, even if the outer owner is the same.
				if (node.tagName === 'TABLE') { active = null; return; }
				const declared = node.hasAttribute('data-excel-type') || node.hasAttribute('data-excel-value');
				if (declared) active = null;
				const nextOwner = declared ? node : owner;
				const visible = !isElementExcluded(node, root);
				if (node.tagName === 'BR' && visible && active) active.tokens.push({ text: '\n' });
				const block = node !== cell && /^(block|flex|grid|list-item|flow-root)$/.test(computedStyle(node).display);
				if (block && visible && active) active.tokens.push({ block: true });
				for (const child of node.childNodes) visit(child, nextOwner);
				if (block && visible && active) active.tokens.push({ block: true });
				if (declared) active = null;
				return;
			}
			if (node.nodeType !== 3 || isElementExcluded(node.parentElement, root)) return;
			const token = textToken(node);
			if (!token.text.trim()) {
				if (active && active.sourceElement === owner) active.tokens.push(token);
				return;
			}
			const range = document.createRange();
			range.selectNodeContents(node);
			const rect = range.getBoundingClientRect();
			if (!rect.width || !rect.height) return;
			const bounds = {
				left: snapToPrecision(rect.left + scrollX - rootLeft), right: snapToPrecision(rect.right + scrollX - rootLeft),
				top: snapToPrecision(rect.top + scrollY - rootTop), bottom: snapToPrecision(rect.bottom + scrollY - rootTop)
			};
			if (owner && active?.sourceElement === owner) {
				active.left = Math.min(active.left, bounds.left);
				active.right = Math.max(active.right, bounds.right);
				active.top = Math.min(active.top, bounds.top);
				active.bottom = Math.max(active.bottom, bounds.bottom);
				active.tokens.push(token);
			} else {
				active = { cell, sourceElement: owner || node.parentElement, owner, tokens: [token], ...bounds };
				fragments.push(active);
				if (!owner) active = null;
			}
		};
		visit(cell, null);
		const declaredValues = new Set();
		const numericOwners = new Set();
		return fragments.filter(fragment => {
			fragment.text = joinTextTokens(fragment.tokens).trim();
			delete fragment.tokens;
			const owner = fragment.owner;
			delete fragment.owner;
			if (owner?.hasAttribute('data-excel-value')) {
				if (declaredValues.has(owner)) return false;
				declaredValues.add(owner);
			} else if (['number', 'currency'].includes(owner?.getAttribute('data-excel-type'))) {
				if (numericOwners.has(owner)) throw new Error('A numeric declaration spans separate regions; annotate each value separately');
				numericOwners.add(owner);
			}
			return Boolean(fragment.text);
		});
	};

	const getCellText = (cell, root) => {
		const descendants = Array.from(cell.querySelectorAll('*'));
		let text = cell.innerText || '';
		if (isElementExcluded(cell, root) || descendants.some(child => isElementExcluded(child, root))) {
			const tokens = [];
			const visit = node => {
				if (node.nodeType === 3) {
					if (!isElementExcluded(node.parentElement, root)) tokens.push(textToken(node));
					return;
				}
				if (node.nodeType !== 1 || isElementBlacklisted(node, root)) return;
				const visible = !isElementExcluded(node, root);
				if (node.tagName === 'BR') { if (visible) tokens.push({ text: '\n' }); return; }
				const block = visible && node !== cell && /^(block|flex|grid|table|list-item|flow-root)$/.test(computedStyle(node).display);
				if (block) tokens.push({ block: true });
				for (const child of node.childNodes) visit(child);
				if (block) tokens.push({ block: true });
			};
			visit(cell);
			text = joinTextTokens(tokens);
		}
		const normalized = text.replace(/\u00a0/g, ' ').replace(/\t+/g, ' ');
		const hasInlineFloat = descendants.some(child => !isElementExcluded(child, root)
			&& child.style.cssFloat && child.style.cssFloat !== 'none');
		return hasInlineFloat ? normalized.replace(/\s*\n\s*/g, ' ').replace(/\s{2,}/g, ' ').trim() : normalized;
	};


	const collectDomSnapshot = (root) => {
		// The synchronous snapshot starts after font/image waits; preload visibility remains live.
		snapshotExclusions = new WeakMap();
		const rootRect = root.getBoundingClientRect();
		const scrollY = window.scrollY || window.pageYOffset || 0;
		const scrollX = window.scrollX || window.pageXOffset || 0;

		const rootLeft = snapToPrecision(rootRect.left + scrollX);
		const rootTop = snapToPrecision(rootRect.top + scrollY);
		const rootWidth = snapToPrecision(root.scrollWidth);
		const rootHeight = snapToPrecision(root.scrollHeight);

		const allCells = Array.from(root.querySelectorAll('td, th')).filter(
		(cell) => hasExportableContent(cell, root)
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
		const tableElements = [...(root.matches('table') ? [root] : []), ...root.querySelectorAll('table')].filter(
		(tableElement) => hasExportableContent(tableElement, root)
		);
		const textFragments = [];
		const pageBreakMarkers = Array.from(root.querySelectorAll(CONFIG.PAGE_BREAK_SELECTOR)).filter(
		(marker) => !isElementExcluded(marker, root)
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
			if (!tableElement) continue;
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

  return { normalizedBlacklistSelectors, loggedBlacklistedElements, elementMatchesSelector, isElementBlacklisted, isElementExcluded, getElementStyle: computedStyle, collectTextFragmentsForCell, getCellText, collectDomSnapshot };
}

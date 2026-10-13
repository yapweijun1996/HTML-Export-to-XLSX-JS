/*! HTML Export to XLSX JS v0.1.0 | MIT | Copyright (c) 2026 YAP WEI JUN */
(() => {
  // src/values.js
  var escapeRegExp = (text) => String(text).replace(/[.*+?^$()|[\]\\{}]/g, "\\$&");
  var currencies = typeof Intl.supportedValuesOf === "function" ? new Set(Intl.supportedValuesOf("currency")) : null;
  function validateCurrency(value) {
    if (value === void 0 || value === null || value === "") return null;
    const code = String(value).trim().toUpperCase();
    if (!/^[A-Z]{3}$/.test(code) || currencies && !currencies.has(code)) {
      throw new RangeError("Unsupported currency code: " + code);
    }
    new Intl.NumberFormat("en", { style: "currency", currency: code });
    return code;
  }
  function decimalKey(value) {
    const match = String(value).match(/^([+-]?)(\d+)(?:\.(\d+))?(?:e([+-]?\d+))?$/i);
    if (!match) return null;
    let digits = (match[2] + (match[3] || "")).replace(/^0+/, "");
    if (!digits) return "0";
    let exponent = Number(match[4] || 0) - (match[3] || "").length;
    const trimmed = digits.replace(/0+$/, "");
    exponent += digits.length - trimmed.length;
    digits = trimmed;
    if (digits.length > 15) return null;
    return (match[1] === "-" ? "-" : "") + digits + "e" + exponent;
  }
  var minusSigns = /* @__PURE__ */ new Map();
  function normalizeSigns(text, locale) {
    if (!minusSigns.has(locale)) {
      const minus2 = new Intl.NumberFormat(locale).formatToParts(-1).find((part) => part.type === "minusSign")?.value || "-";
      if (minusSigns.size >= 32) minusSigns.delete(minusSigns.keys().next().value);
      minusSigns.set(locale, minus2);
    }
    const minus = minusSigns.get(locale);
    return minus && minus !== "-" ? text.split(minus).join("-") : text;
  }
  function parseLocalizedNumber(value, locale = "en-US") {
    if (typeof value === "number") {
      return Number.isFinite(value) && Math.abs(value) <= Number.MAX_SAFE_INTEGER && decimalKey(value) !== null ? value : null;
    }
    let text = normalizeSigns(String(value ?? "").trim().replace(/[\u061c\u200e\u200f]/g, ""), locale);
    if (!text) return null;
    const negative = /^\(.*\)$/.test(text);
    if (negative) text = text.slice(1, -1);
    const formatter = new Intl.NumberFormat(locale);
    const parts = formatter.formatToParts(12345.6);
    const group = parts.find((part) => part.type === "group")?.value;
    const decimal = parts.find((part) => part.type === "decimal")?.value || ".";
    const digitFormatter = new Intl.NumberFormat(locale, { useGrouping: false });
    for (let digit = 0; digit < 10; digit++) {
      text = text.split(digitFormatter.format(digit)).join(String(digit));
    }
    const normalizedGroup = group?.replace(/\s/g, " ");
    text = text.replace(/\s/g, " ");
    const decimalParts = text.split(decimal);
    if (decimalParts.length > 2 || decimalParts.length === 2 && !/^\d+$/.test(decimalParts[1])) return null;
    if (normalizedGroup && normalizedGroup !== " " && text.includes(normalizedGroup)) {
      const groupedInteger = text.replace(/^[+-]/, "").split(decimal)[0];
      const sizes = groupedInteger.split(normalizedGroup).map((part) => part.length);
      const sample = formatter.formatToParts(123456789).filter((part) => part.type === "integer").map((part) => part.value.length);
      const rightSize = sample.at(-1);
      const middleSize = sample.length > 2 ? sample.at(-2) : rightSize;
      if (sizes.at(-1) !== rightSize || sizes.slice(1, -1).some((size) => size !== middleSize) || sizes[0] < 1 || sizes[0] > middleSize) return null;
      text = text.split(normalizedGroup).join("");
    }
    if (text.includes(" ")) {
      if (normalizedGroup !== " ") return null;
      const integer = text.replace(/^[+-]/, "").split(decimal)[0];
      if (!/^\d{1,3}( \d{3})+$/.test(integer)) return null;
      text = text.replace(/ /g, "");
    }
    if (decimal !== ".") text = text.split(decimal).join(".");
    if (!/^[+-]?\d+(?:\.\d+)?$/.test(text) || negative && /^[+-]/.test(text)) return null;
    const number = Number(text);
    const key = decimalKey(text);
    return Number.isFinite(number) && Math.abs(number) <= Number.MAX_SAFE_INTEGER && key !== null && key === decimalKey(number) ? negative ? -number : number : null;
  }
  var explicitSymbols = [
    ["HK$", "HKD"],
    ["US$", "USD"],
    ["S$", "SGD"],
    ["NT$", "TWD"],
    ["RMB", "CNY"],
    ["RM", "MYR"],
    ["JP\xA5", "JPY"],
    ["\u20AC", "EUR"],
    ["\xA3", "GBP"],
    ["\u20B9", "INR"]
  ];
  function currencyFormat(currency, locale = "en-US") {
    const decimals = new Intl.NumberFormat(locale, {
      style: "currency",
      currency
    }).resolvedOptions().maximumFractionDigits;
    const fraction = decimals ? "." + "0".repeat(decimals) : "";
    return '"' + currency + '" #,##0' + fraction + ';("' + currency + '" #,##0' + fraction + ")";
  }
  function detectCurrency(rawText, options = {}) {
    const locale = options.locale || "en-US";
    let text = normalizeSigns(String(rawText ?? "").trim().replace(/[\u061c\u200e\u200f]/g, ""), locale);
    if (!text) return null;
    const accounting = /^\(.*\)$/.test(text);
    if (accounting) text = text.slice(1, -1).trim();
    const sign = text.match(/^[+-]/)?.[0] || "";
    if (accounting && sign) return null;
    if (sign) text = text.slice(1).trim();
    let code = null;
    const takeSymbol = (symbol) => {
      if (text.startsWith(symbol)) {
        text = text.slice(symbol.length).trim();
        return true;
      }
      if (text.endsWith(symbol)) {
        text = text.slice(0, -symbol.length).trim();
        return true;
      }
      return false;
    };
    for (const [symbol, currency] of explicitSymbols) {
      if (takeSymbol(symbol)) {
        code = currency;
        break;
      }
    }
    if (!code) {
      const prefix = text.match(/^([A-Za-z]{3})\s*/);
      const suffix = text.match(/\s*([A-Za-z]{3})$/);
      const explicit = prefix || suffix;
      if (explicit) {
        try {
          code = validateCurrency(explicit[1]);
        } catch {
          return null;
        }
        text = (prefix ? text.slice(prefix[0].length) : text.slice(0, -suffix[0].length)).trim();
      }
    }
    if (!code && options.currency) {
      code = validateCurrency(options.currency);
      const symbol = new Intl.NumberFormat(options.locale || "en-US", {
        style: "currency",
        currency: code
      }).formatToParts(1).find((part) => part.type === "currency")?.value;
      const tokens = [.../* @__PURE__ */ new Set([symbol, "$", "\xA5"])].filter(Boolean).sort((a, b) => b.length - a.length);
      const token = tokens.find((candidate) => text.startsWith(candidate) || text.endsWith(candidate));
      if (!token) return null;
      if (token === "\xA5" && !["JPY", "CNY"].includes(code)) return null;
      takeSymbol(token);
    }
    if (!code) return null;
    if ((accounting || sign) && /^[+\-(]/.test(text)) return null;
    const value = parseLocalizedNumber(accounting ? "(" + text + ")" : sign + text, locale);
    return value === null ? null : { value, numFmt: currencyFormat(code, options.locale) };
  }
  function parseCellValue(text, element, options = {}) {
    const attribute = (name) => {
      for (let current = element; current; current = current.parentElement) {
        const value = current.getAttribute?.(name);
        if (value !== null && value !== void 0) return value;
        if (current.tagName === "TD" || current.tagName === "TH") break;
      }
      return null;
    };
    const type = attribute("data-excel-type");
    const raw = attribute("data-excel-value") ?? text;
    const cellCurrency = attribute("data-excel-currency");
    const cellOptions = { ...options, currency: cellCurrency ? validateCurrency(cellCurrency) : options.currency };
    if (type === "text") return { value: String(raw) };
    const currency = detectCurrency(raw, cellOptions);
    if (currency) return currency;
    if (type === "number" || type === "currency") {
      const value = parseLocalizedNumber(raw, cellOptions.locale);
      if (value === null) throw new TypeError("Invalid explicitly typed numeric cell");
      if (type === "currency" && !cellOptions.currency) throw new TypeError("Currency cell requires a currency");
      return { value, numFmt: type === "currency" ? currencyFormat(cellOptions.currency, cellOptions.locale) : void 0 };
    }
    return { value: String(raw ?? "") };
  }

  // src/config.js
  var DEFAULT_CONFIG = {
    DEBUG: false,
    ROOT_SELECTOR: "[data-xlsx-export]",
    BUTTON_SELECTOR: "[data-xlsx-button]",
    SHEET_NAME: "PrintForm",
    MIN_COL_WIDTH_PX: 10,
    MIN_ROW_HEIGHT_PX: 1,
    // EDGE_TOLERANCE_PX clusters nearly-identical DOM edges (in CSS px) into one Excel grid line.
    // Default 2.0 prevents fragmented micro-columns/rows and duplicated inner borders.
    // Tuning range: 1.5–3.0 is typical; lower => more tiny fragments, higher => risk collapsing close columns/rows.
    // See README.md options and normalizeEdges() for details.
    EDGE_TOLERANCE_PX: 2,
    IMAGE_FORMAT: "image/png",
    LOG_PREFIX: "[export]",
    CSS_DPI: 96,
    TARGET_DPI: 96,
    EXPORT_SCALE: 1,
    // FONT_SCALE keeps Excel text close to HTML size; 0.9 works well for 96 DPI layouts
    FONT_SCALE: 0.9,
    // ROW_SCALE controls row height scaling; falls back to EXPORT_SCALE when null
    ROW_SCALE: 1,
    // IMAGE_SCALE auto-upsamples to device DPI (>=1) so exports stay sharp
    IMAGE_SCALE: 3,
    // COLUMN_WIDTH_SCALE provides a final adjustment factor for column widths.
    // The system auto-calculates the width, but you can use this to fine-tune.
    // Example: 1.05 means "make all columns 5% wider". 0.95 means "5% narrower".
    COLUMN_WIDTH_SCALE: 1,
    // IMAGE_WIDTH_SCALE provides a final adjustment factor for all image sizes.
    // Example: 1.1 means "make all images 10% larger". 0.9 means "10% smaller".
    IMAGE_WIDTH_SCALE: 0.92,
    // IMAGE_PLACEMENT_MODE controls how images are anchored.
    // 'precise' keeps the original pixel offsets, while 'alignment' applies HTML-like flows.
    IMAGE_PLACEMENT_MODE: "alignment",
    // IMAGE_ALIGNMENT defines the default flow when IMAGE_PLACEMENT_MODE is 'alignment'.
    IMAGE_ALIGNMENT: "inline",
    // inline | block | center | left | right | precise
    // IMAGE_ALIGNMENT_ATTR lets individual cells override the flow via a data attribute.
    IMAGE_ALIGNMENT_ATTR: "data-excel-image-align",
    // IMAGE_SPACING_PX applies consistent gaps between auto-aligned images.
    IMAGE_SPACING_PX: 4,
    BLACKLIST_SELECTORS: ["[data-xlsx-ignore]"],
    PAGE_BREAK_SELECTOR: "[data-xlsx-page-break]",
    IMAGE_TIMEOUT_MS: 1e4,
    IMAGE_CONCURRENCY: 4,
    PAGE_SETUP: {
      // ExcelJS pageSetup reference: https://github.com/exceljs/exceljs#worksheet-page-setup
      paperSize: 9,
      // A4 portrait constant (Excel uses numeric paper codes)
      orientation: "portrait",
      // 'portrait' or 'landscape'
      fitToPage: true,
      // enable Excel auto-scaling so sheet fits the printable page
      scale: 100,
      // only used when fitToPage is false; retained for quick toggling
      fitToWidth: 1,
      // number of pages to fit horizontally when fitToPage is true
      fitToHeight: 0,
      // 0 keeps unlimited pages vertically; adjust if needed
      horizontalCentered: true,
      // center the sheet between left/right margins
      verticalCentered: false,
      // set true to center content vertically on the page
      margins: {
        left: 0.5,
        // inches
        right: 0.5,
        // inches
        top: 0.1,
        // inches before header text
        bottom: 0.1,
        // inches before footer text
        header: 0.3,
        // header gap in inches (add >=0.25 if top rows get clipped)
        footer: 0.3
        // footer gap in inches
      }
    }
  };
  var mappings = {
    edgeTolerancePx: "EDGE_TOLERANCE_PX",
    exportScale: "EXPORT_SCALE",
    fontScale: "FONT_SCALE",
    rowScale: "ROW_SCALE",
    imageScale: "IMAGE_SCALE",
    columnWidthScale: "COLUMN_WIDTH_SCALE",
    imageWidthScale: "IMAGE_WIDTH_SCALE",
    imagePlacement: "IMAGE_PLACEMENT_MODE",
    imageAlignment: "IMAGE_ALIGNMENT",
    imageTimeoutMs: "IMAGE_TIMEOUT_MS",
    imageConcurrency: "IMAGE_CONCURRENCY",
    pageBreakSelector: "PAGE_BREAK_SELECTOR",
    exclude: "BLACKLIST_SELECTORS",
    debug: "DEBUG",
    sheetName: "SHEET_NAME"
  };
  var allowed = /* @__PURE__ */ new Set([
    ...Object.keys(mappings),
    "root",
    "filename",
    "currency",
    "locale",
    "exceljs",
    "download",
    "pageSetup",
    "onWarning"
  ]);
  var freezeDeep = (object) => {
    for (const value of Object.values(object)) {
      if (value && typeof value === "object") freezeDeep(value);
    }
    return Object.freeze(object);
  };
  freezeDeep(DEFAULT_CONFIG);
  function resolveOptions(root, options = {}) {
    for (const key of Object.keys(options)) {
      if (!allowed.has(key)) throw new TypeError("Unknown export option: " + key);
    }
    const currency = validateCurrency(options.currency ?? root.getAttribute("data-excel-currency"));
    const locale = options.locale ?? root.getAttribute("data-excel-locale") ?? "en-US";
    new Intl.NumberFormat(locale);
    const config = { ...DEFAULT_CONFIG, BLACKLIST_SELECTORS: [...DEFAULT_CONFIG.BLACKLIST_SELECTORS] };
    for (const [key, internal] of Object.entries(mappings)) {
      if (options[key] !== void 0) config[internal] = options[key];
    }
    for (const [key, internal] of Object.entries(mappings)) {
      if (key.endsWith("Scale") || key.endsWith("Ms") || key === "imageConcurrency" || key === "edgeTolerancePx") {
        if (!(Number.isFinite(config[internal]) && config[internal] > 0)) {
          throw new RangeError(key + " must be a positive finite number");
        }
      }
    }
    if (!Number.isInteger(config.IMAGE_CONCURRENCY)) throw new RangeError("imageConcurrency must be an integer");
    if (!["alignment", "precise"].includes(config.IMAGE_PLACEMENT_MODE)) throw new RangeError("Invalid imagePlacement");
    if (!["inline", "block", "center", "left", "right", "precise"].includes(config.IMAGE_ALIGNMENT)) throw new RangeError("Invalid imageAlignment");
    config.IMAGE_ALIGNMENT_EXPLICIT = options.imageAlignment !== void 0;
    if (!Array.isArray(config.BLACKLIST_SELECTORS) || config.BLACKLIST_SELECTORS.some((value) => typeof value !== "string")) {
      throw new TypeError("exclude must be an array of CSS selectors");
    }
    for (const selector of [...config.BLACKLIST_SELECTORS, config.PAGE_BREAK_SELECTOR]) root.matches(selector);
    if (options.onWarning !== void 0 && typeof options.onWarning !== "function") throw new TypeError("onWarning must be a function");
    config.BLACKLIST_SELECTORS = [...config.BLACKLIST_SELECTORS];
    const page = structuredClone(options.pageSetup || {});
    config.PAGE_SETUP = {
      ...DEFAULT_CONFIG.PAGE_SETUP,
      ...page,
      margins: { ...DEFAULT_CONFIG.PAGE_SETUP.margins, ...page.margins }
    };
    return { config: freezeDeep(config), currency, locale };
  }

  // src/layout-metrics.js
  function createLayoutMetrics(dependencies) {
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
    const measureAvgCharWidth = /* @__PURE__ */ (() => {
      let avgCharWidth = null;
      return (fontStyle = "9pt Arial") => {
        if (avgCharWidth !== null) {
          return avgCharWidth;
        }
        if (typeof document === "undefined") {
          avgCharWidth = 7;
          return avgCharWidth;
        }
        try {
          const canvas = document.createElement("canvas");
          const context = canvas.getContext("2d");
          if (!context) {
            avgCharWidth = 7;
            return avgCharWidth;
          }
          context.font = fontStyle;
          const metrics = context.measureText("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz");
          avgCharWidth = metrics.width / 62;
        } catch (e) {
          errorLog("Failed to measure avg char width, falling back to 7px", { code: "FONT_MEASUREMENT_FAILED", error: e });
          avgCharWidth = 7;
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
      const excelPadding = DEFAULT_EXCEL_PADDING_PX;
      const characterCount = (scaledPx - excelPadding) / avgCharWidth;
      const finalAdjustment = toPositiveNumber(CONFIG.COLUMN_WIDTH_SCALE) || 1;
      return Math.max(0, characterCount * finalAdjustment);
    };
    const pxToColumnWidthForImagePositioning = (px, scale = 1) => {
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
      return Math.round(heightPoints * CONFIG.CSS_DPI / 72);
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
        const widthChars = toPositiveNumber(column?.width) || columnWidthFn(layoutModel.columnWidthsPx[i], EXPORT_SCALE);
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
        const heightPoints = toPositiveNumber(row?.height) || pxToPoint(layoutModel.rowHeightsPx[i], ROW_SCALE);
        const heightPx = Math.max(CONFIG.MIN_ROW_HEIGHT_PX, rowHeightToPixels(heightPoints));
        rowEdgesPx.push(rowEdgesPx[i] + heightPx);
        debugLog("Row height map", {
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
      const sorted = Array.from(edges).map((value) => snapToPrecision(value)).filter((value) => isFiniteNumber(value)).sort((a, b) => a - b);
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
    const isFiniteNumber = (value) => typeof Number.isFinite === "function" ? Number.isFinite(value) : isFinite(value);
    return { POINTS_PER_CSS_PX, toPositiveNumber, resolveExportScale, EXPORT_SCALE, resolveFontScale, FONT_SCALE, measureAvgCharWidth, resolveRowScale, ROW_SCALE, resolveImageScale, IMAGE_SCALE, DEFAULT_EXCEL_PADDING_PX, pxToPoint, pxToColumnWidth, pxToColumnWidthForImagePositioning, columnWidthToPixels, rowHeightToPixels, interpolateCoordinate, buildWorksheetEdgeMaps, SNAP_PRECISION, snapToPrecision, resolveSize, normalizeEdges, isFiniteNumber };
  }

  // src/styles.js
  function createStyles(dependencies) {
    const { CONFIG, DEFAULT_EXCEL_PADDING_PX, FONT_SCALE, debugLog, escapeRegExp: escapeRegExp2, getElementStyle, isFiniteNumber, measureAvgCharWidth, pxToPoint } = dependencies;
    const COLOR_CANVAS_CTX = (function() {
      if (typeof document === "undefined") return null;
      const canvas = document.createElement("canvas");
      return canvas && canvas.getContext ? canvas.getContext("2d") : null;
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
      return byte.toString(16).padStart(2, "0").toUpperCase();
    };
    const parseHexColor = (value) => {
      if (!value || typeof value !== "string") return null;
      let hex = value.trim();
      if (!hex) return null;
      if (hex[0] === "#") hex = hex.slice(1);
      if (![3, 4, 6, 8].includes(hex.length)) return null;
      const expand = (ch) => ch + ch;
      let r, g, b, a = 255;
      if (hex.length === 3 || hex.length === 4) {
        const chars = hex.split("");
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
      if (trimmed.endsWith("%")) {
        const num2 = parseFloat(trimmed.slice(0, -1));
        if (!Number.isFinite(num2)) return 1;
        return Math.min(1, Math.max(0, num2 / 100));
      }
      const num = parseFloat(trimmed);
      if (!Number.isFinite(num)) return 1;
      return Math.min(1, Math.max(0, num));
    };
    const parseRgbFunction = (value) => {
      if (!value || typeof value !== "string") return null;
      const match = value.trim().match(/^rgba?\(\s*(.+)\s*\)$/i);
      if (!match) return null;
      const body = match[1];
      let alphaPart;
      let colorsPart = body;
      if (body.includes("/")) {
        const parts = body.split("/");
        colorsPart = parts[0];
        alphaPart = parts[1];
      }
      const colorTokens = colorsPart.split(/[\s,]+/).map((token) => token.trim()).filter(Boolean);
      if (colorTokens.length < 3) return null;
      const parseComponent = (component) => {
        if (!component) return null;
        if (component.endsWith("%")) {
          const num2 = parseFloat(component.slice(0, -1));
          if (!Number.isFinite(num2)) return null;
          return clampToByte(num2 / 100 * 255);
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
      if (typeof alphaPart === "string" && alphaPart.trim()) {
        alpha = parseAlphaValue(alphaPart);
      } else if (colorTokens.length >= 4) {
        alpha = parseAlphaValue(colorTokens[3]);
      }
      return { r, g, b, a: alpha };
    };
    const COLOR_CACHE = /* @__PURE__ */ new Map();
    const parseCssColor = (value) => {
      if (!value) return null;
      const raw = String(value).trim();
      if (!raw) return null;
      if (raw.toLowerCase() === "transparent") {
        return { r: 0, g: 0, b: 0, a: 0 };
      }
      const hex = parseHexColor(raw);
      if (hex) return hex;
      const rgb = parseRgbFunction(raw);
      if (rgb) return rgb;
      if (!COLOR_CANVAS_CTX) return null;
      try {
        COLOR_CANVAS_CTX.fillStyle = "#000000";
        COLOR_CANVAS_CTX.fillStyle = raw;
        const normalized = COLOR_CANVAS_CTX.fillStyle;
        if (typeof normalized !== "string" || !normalized) {
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
      const key = value || "";
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
      const parts = [alphaByte, parsed.r, parsed.g, parsed.b].map((component) => toHexByte(component));
      if (parts.some((component) => component === null)) {
        COLOR_CACHE.set(key, null);
        return null;
      }
      const argb = parts.join("");
      COLOR_CACHE.set(key, argb);
      return argb;
    };
    const widthToBorderStyle = (widthPx, cssStyle) => {
      if (!widthPx || widthPx <= 0.1) return null;
      const normalizedStyle = (cssStyle || "").toLowerCase();
      if (!normalizedStyle || normalizedStyle === "none" || normalizedStyle === "hidden") return null;
      const styleByWidth = () => {
        if (widthPx < 1.5) return "thin";
        if (widthPx < 2.5) return "medium";
        return "thick";
      };
      switch (normalizedStyle) {
        case "solid":
          return styleByWidth();
        case "dashed":
          return widthPx < 2 ? "dashed" : "mediumDashed";
        case "dotted":
          return "dotted";
        case "double":
          return "double";
        case "groove":
        case "ridge":
          return widthPx < 2 ? "medium" : "double";
        case "inset":
        case "outset":
          return widthPx < 1.5 ? "thin" : "medium";
        case "dash-dot":
        case "dashdot":
        case "dot-dash":
          return widthPx < 2 ? "dashDot" : "mediumDashDot";
        case "dash-dot-dot":
        case "dashdotdot":
        case "dot-dot-dash":
          return widthPx < 2 ? "dashDotDot" : "mediumDashDotDot";
        default:
          return styleByWidth();
      }
    };
    const BORDER_STYLE_KEYWORDS = /* @__PURE__ */ new Set([
      "none",
      "hidden",
      "dotted",
      "dashed",
      "solid",
      "double",
      "groove",
      "ridge",
      "inset",
      "outset",
      "dash-dot",
      "dashdot",
      "dot-dash",
      "dash-dot-dot",
      "dashdotdot",
      "dot-dot-dash"
    ]);
    const parseBorderShorthand = (value) => {
      if (!value) return null;
      const tokens = value.split(/\s+/).map((token) => token.trim().replace(/!important$/i, "")).filter((token) => token && token.toLowerCase() !== "!important");
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
        color: colorParts.length ? colorParts.join(" ") : null
      };
    };
    const matchInlineStyleValue = (styleText, property) => {
      if (!styleText || !property) return null;
      const pattern = new RegExp(`(?:^|;)\\s*${escapeRegExp2(property)}\\s*:\\s*([^;]+)`, "i");
      const match = pattern.exec(styleText);
      if (!match) return null;
      const raw = match[1]?.trim().replace(/;$/, "");
      if (!raw) return null;
      return raw.replace(/\s*!important\s*$/i, "").trim() || null;
    };
    const parseInlineBorderFromAttribute = (element, side) => {
      if (!element || typeof element.getAttribute !== "function") return null;
      const styleAttr = element.getAttribute("style");
      if (!styleAttr) return null;
      if (side) {
        const specific = matchInlineStyleValue(styleAttr, `border-${side}`);
        if (specific) {
          return parseBorderShorthand(specific);
        }
      }
      const generic = matchInlineStyleValue(styleAttr, "border");
      return generic ? parseBorderShorthand(generic) : null;
    };
    const extractBorders = (cs, element) => {
      const sides = ["top", "right", "bottom", "left"];
      const border = {};
      for (const side of sides) {
        const width = parseFloat(cs.getPropertyValue(`border-${side}-width`)) || 0;
        const type = cs.getPropertyValue(`border-${side}-style`) || "";
        const excelStyle = widthToBorderStyle(width, type);
        if (!excelStyle) continue;
        const color = colorToARGB(cs.getPropertyValue(`border-${side}-color`));
        border[side] = {
          style: excelStyle,
          color: color ? { argb: color } : void 0
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
          let styleValue = styleStr || "";
          let colorValue = colorStr || "";
          if (width == null || width <= 0 || !styleValue || !colorValue) {
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
          if (width == null || width <= 0 || !styleValue || !colorValue) {
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
            color: color ? { argb: color } : void 0
          };
        }
      }
      return Object.keys(border).length ? border : null;
    };
    const normalizeFontWeight = (value) => {
      if (typeof value === "number") {
        return isFiniteNumber(value) ? value : 400;
      }
      const str = String(value || "").trim().toLowerCase();
      if (!str) return 400;
      if (str === "bold" || str === "bolder") return 700;
      if (str === "normal") return 400;
      if (str === "lighter") return 300;
      const parsed = parseInt(str, 10);
      return isFiniteNumber(parsed) ? parsed : 400;
    };
    const visibleTextCache = /* @__PURE__ */ new WeakMap();
    const visibleTextNodes = (cell) => {
      if (!cell || typeof document === "undefined") return [];
      if (visibleTextCache.has(cell)) return visibleTextCache.get(cell);
      const nodes = [];
      const table = cell.closest("table");
      const walker = document.createTreeWalker(cell, NodeFilter.SHOW_TEXT);
      while (walker.nextNode()) {
        const node = walker.currentNode;
        if (!node.textContent.trim() || node.parentElement.closest("table") !== table || dependencies.isElementExcluded(node.parentElement, cell)) continue;
        nodes.push(node);
      }
      visibleTextCache.set(cell, nodes);
      return nodes;
    };
    const isCellUniformlyBold = (cell, baseWeight) => {
      const nodes = visibleTextNodes(cell);
      return nodes.length ? nodes.every((node) => normalizeFontWeight(getElementStyle(node.parentElement).fontWeight) >= 600) : baseWeight >= 600;
    };
    const resolveFontSizePx = (cell, cellStyle) => {
      const sizes = visibleTextNodes(cell).map((node) => parseFloat(getElementStyle(node.parentElement).fontSize) || 0);
      return sizes.length ? sizes.reduce((maximum, size) => Math.max(maximum, size), 0) : parseFloat(cellStyle.fontSize) || 0;
    };
    const backgroundCaches = /* @__PURE__ */ new WeakMap();
    const extractBackground = (element, root) => {
      if (!element) return null;
      if (!backgroundCaches.has(root)) backgroundCaches.set(root, /* @__PURE__ */ new WeakMap());
      const cache = backgroundCaches.get(root);
      if (cache.has(element)) return cache.get(element);
      let color = dependencies.isElementExcluded(element, root) ? null : colorToARGB(getElementStyle(element).backgroundColor);
      if (!color && element !== root) color = extractBackground(element.parentElement, root);
      cache.set(element, color);
      return color;
    };
    const extractFont = (cell, cs) => {
      const families = (cs.fontFamily || "").split(",").map((v) => v.trim().replace(/^"|"$/g, ""));
      const name = families[0] || "Arial";
      const sizePx = resolveFontSizePx(cell, cs);
      const weight = normalizeFontWeight(cs.fontWeight);
      const bold = isCellUniformlyBold(cell, weight);
      let excelFontSize;
      if (sizePx > 0) {
        const pointSize = pxToPoint(sizePx, FONT_SCALE);
        excelFontSize = Math.max(8, Math.round(pointSize * 100) / 100);
      }
      const font = {
        name,
        size: excelFontSize,
        bold,
        italic: cs.fontStyle === "italic"
      };
      const color = colorToARGB(cs.color);
      if (color) font.color = { argb: color };
      if (CONFIG.DEBUG && sizePx > 0) {
        debugLog("Font extraction", {
          element: cell?.tagName || "unknown",
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
      if (value === null || value === void 0) return 0;
      const parsed = parseFloat(value);
      return Number.isFinite(parsed) ? parsed : 0;
    };
    const computeEffectiveHorizontalPadding = (cell, cellStyle) => {
      const baseLeft = parseCssPixelValue(cellStyle?.paddingLeft);
      const baseRight = parseCssPixelValue(cellStyle?.paddingRight);
      if (!cell || typeof document === "undefined" || typeof window === "undefined") {
        return { left: baseLeft, right: baseRight };
      }
      let minPositiveExtraLeft = Infinity;
      let minPositiveExtraRight = Infinity;
      try {
        const showText = typeof NodeFilter !== "undefined" ? NodeFilter.SHOW_TEXT : 4;
        const walker = document.createTreeWalker(cell, showText, null);
        while (walker.nextNode()) {
          const node = walker.currentNode;
          if (!node || !node.textContent || !node.textContent.trim() || dependencies.isElementExcluded(node.parentElement, cell) || node.parentElement.closest("table") !== cell.closest("table")) continue;
          let current = node.parentElement;
          let extraLeft = 0;
          let extraRight = 0;
          while (current && current !== cell) {
            const style = getElementStyle(current);
            extraLeft += parseCssPixelValue(style.paddingLeft) + parseCssPixelValue(style.marginLeft) + parseCssPixelValue(style.borderLeftWidth);
            extraRight += parseCssPixelValue(style.paddingRight) + parseCssPixelValue(style.marginRight) + parseCssPixelValue(style.borderRightWidth);
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
          debugLog("Failed to compute extra padding from descendants", err);
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
      if (!cell || typeof document === "undefined" || typeof window === "undefined") {
        return { top: baseTop, bottom: baseBottom };
      }
      let minPositiveExtraTop = Infinity;
      let minPositiveExtraBottom = Infinity;
      try {
        const showText = typeof NodeFilter !== "undefined" ? NodeFilter.SHOW_TEXT : 4;
        const walker = document.createTreeWalker(cell, showText, null);
        while (walker.nextNode()) {
          const node = walker.currentNode;
          if (!node || !node.textContent || !node.textContent.trim() || dependencies.isElementExcluded(node.parentElement, cell) || node.parentElement.closest("table") !== cell.closest("table")) continue;
          let current = node.parentElement;
          let extraTop = 0;
          let extraBottom = 0;
          while (current && current !== cell) {
            const style = getElementStyle(current);
            extraTop += parseCssPixelValue(style.paddingTop) + parseCssPixelValue(style.marginTop) + parseCssPixelValue(style.borderTopWidth);
            extraBottom += parseCssPixelValue(style.paddingBottom) + parseCssPixelValue(style.marginBottom) + parseCssPixelValue(style.borderBottomWidth);
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
          debugLog("Failed to compute extra vertical padding from descendants", err);
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
        if (!value) return void 0;
        const normalized = String(value).trim().toLowerCase();
        return normalized || void 0;
      };
      const horizontalMap = {
        start: cs.direction === "rtl" ? "right" : "left",
        left: "left",
        end: cs.direction === "rtl" ? "left" : "right",
        right: "right",
        center: "center",
        justify: "justify"
      };
      const verticalMap = {
        top: "top",
        middle: "middle",
        center: "middle",
        bottom: "bottom",
        baseline: "bottom"
      };
      const cssHorizontal = normalize(cs.textAlign);
      const cssVertical = normalize(cs.verticalAlign);
      const attrHorizontal = normalize(cell.getAttribute("align") || cell.align);
      const attrVertical = normalize(cell.getAttribute("valign") || cell.vAlign);
      const horizontal = horizontalMap[cssHorizontal] || horizontalMap[attrHorizontal] || void 0;
      const vertical = verticalMap[cssVertical] || verticalMap[attrVertical] || void 0;
      const wrapText = cs.whiteSpace !== "nowrap" || textValue && textValue.includes("\n");
      const alignment = {
        horizontal,
        vertical,
        wrapText
      };
      const padding = computeEffectiveHorizontalPadding(cell, cs);
      const extraLeft = Math.max(0, padding.left - DEFAULT_EXCEL_PADDING_PX);
      const extraRight = Math.max(0, padding.right - DEFAULT_EXCEL_PADDING_PX);
      const preferLeft = !alignment.horizontal || alignment.horizontal === "left" || alignment.horizontal === "justify";
      if (preferLeft) {
        const indent = convertPaddingPxToIndent(extraLeft);
        if (indent > 0) {
          alignment.indent = indent;
        }
      } else if (alignment.horizontal === "right") {
        const indent = convertPaddingPxToIndent(extraRight);
        if (indent > 0) {
          alignment.indent = indent;
        }
      }
      return alignment;
    };
    return { COLOR_CANVAS_CTX, clampToByte, toHexByte, parseHexColor, parseAlphaValue, parseRgbFunction, COLOR_CACHE, parseCssColor, colorToARGB, widthToBorderStyle, BORDER_STYLE_KEYWORDS, parseBorderShorthand, matchInlineStyleValue, parseInlineBorderFromAttribute, extractBorders, normalizeFontWeight, isCellUniformlyBold, resolveFontSizePx, extractFont, extractBackground, parseCssPixelValue, computeEffectiveHorizontalPadding, computeEffectiveVerticalPadding, convertPaddingPxToIndent, extractAlignment };
  }

  // src/dom.js
  function createDom(dependencies) {
    const { CONFIG, debugLog, snapToPrecision, normalizeEdges } = dependencies;
    const normalizedBlacklistSelectors = /* @__PURE__ */ (() => {
      let cachedSelectors = null;
      return () => {
        if (cachedSelectors) return cachedSelectors;
        const raw = CONFIG.BLACKLIST_SELECTORS;
        let list = [];
        if (Array.isArray(raw)) {
          list = raw;
        } else if (typeof raw === "string") {
          list = [raw];
        }
        cachedSelectors = list.map((item) => typeof item === "string" ? item.trim() : "").filter((item) => item.length > 0);
        return cachedSelectors;
      };
    })();
    const loggedBlacklistedElements = /* @__PURE__ */ new WeakSet();
    const elementMatchesSelector = (element, selector) => {
      if (!element || !selector) return false;
      const matcher = element.matches || element.matchesSelector || element.msMatchesSelector || element.webkitMatchesSelector || element.mozMatchesSelector;
      if (typeof matcher !== "function") return false;
      try {
        return matcher.call(element, selector);
      } catch (err) {
        if (CONFIG.DEBUG) {
          debugLog("Selector match failed", selector, err);
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
              debugLog("Skipping blacklisted element", selector, current);
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
    const styleCache = /* @__PURE__ */ new WeakMap();
    const computedStyle = (element) => {
      if (!styleCache.has(element)) styleCache.set(element, window.getComputedStyle(element));
      return styleCache.get(element);
    };
    const computeExcluded = (element, root) => {
      if (!element || isElementBlacklisted(element, root)) return true;
      const style = computedStyle(element);
      if (style.visibility === "hidden" || style.visibility === "collapse") return true;
      for (let current = element; current; current = current.parentElement) {
        const currentStyle = computedStyle(current);
        if (currentStyle.display === "none" || currentStyle.contentVisibility === "hidden" || currentStyle.opacity === "0") return true;
      }
      return false;
    };
    let snapshotExclusions = null;
    const isElementExcluded = (element, root) => {
      if (!element) return true;
      if (!snapshotExclusions) return computeExcluded(element, root);
      const boundary = root || document.documentElement;
      if (!snapshotExclusions.has(boundary)) snapshotExclusions.set(boundary, /* @__PURE__ */ new WeakMap());
      const cache = snapshotExclusions.get(boundary);
      if (!cache.has(element)) cache.set(element, computeExcluded(element, root));
      return cache.get(element);
    };
    const hasExportableContent = (element, root) => !isElementExcluded(element, root) || Array.from(element.querySelectorAll("*")).some((child) => !isElementExcluded(child, root));
    const textToken = (node) => {
      const style = computedStyle(node.parentElement);
      const preserve = ["pre", "pre-wrap", "break-spaces"].includes(style.whiteSpace);
      let text = node.textContent;
      if (!preserve) text = style.whiteSpace === "pre-line" ? text.replace(/[\t\r\f ]+/g, " ") : text.replace(/[\t\r\n\f ]+/g, " ");
      let language = node.parentElement.closest("[lang]")?.getAttribute("lang") || void 0;
      try {
        if (language) Intl.getCanonicalLocales(language);
      } catch {
        language = void 0;
      }
      if (style.textTransform === "uppercase") text = text.toLocaleUpperCase(language);
      if (style.textTransform === "lowercase") text = text.toLocaleLowerCase(language);
      return { text, preserve, capitalize: style.textTransform === "capitalize", language };
    };
    const joinTextTokens = (tokens) => {
      let text = "";
      let preserveTail = false;
      for (const token of tokens) {
        if (token.block) {
          if (text && !text.endsWith("\n")) text = text.replace(/ +$/, "") + "\n";
          continue;
        }
        let part = token.text;
        if (token.capitalize) part = part.replace(
          /(^|[^\p{L}\p{M}\p{N}])\p{L}/gu,
          (match, boundary, index) => index === 0 && /[\p{L}\p{M}\p{N}]$/u.test(text) ? match : match.toLocaleUpperCase(token.language)
        );
        if (!token.preserve) {
          part = part.replace(/ *\n */g, "\n");
          if (!text || /[ \n]$/.test(text)) part = part.replace(/^ +/, "");
          if (part.startsWith("\n")) text = text.replace(/ +$/, "");
        }
        text += part;
        preserveTail = token.preserve;
      }
      return (preserveTail ? text : text.replace(/ +$/, "")).replace(/^\n+|\n+$/g, "").replace(/\u00a0/g, " ");
    };
    const collectTextFragmentsForCell = (cell, metrics, root) => {
      if (!cell || isElementBlacklisted(cell, root)) return [];
      const { rootLeft = 0, rootTop = 0, scrollX = 0, scrollY: scrollY2 = 0 } = metrics || {};
      const fragments = [];
      let active = null;
      const visit = (node, owner) => {
        if (node.nodeType === 1) {
          if (isElementBlacklisted(node, root)) return;
          if (node.tagName === "TABLE") {
            active = null;
            return;
          }
          const declared = node.hasAttribute("data-excel-type") || node.hasAttribute("data-excel-value");
          if (declared) active = null;
          const nextOwner = declared ? node : owner;
          const visible = !isElementExcluded(node, root);
          if (node.tagName === "BR" && visible && active) active.tokens.push({ text: "\n" });
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
          left: snapToPrecision(rect.left + scrollX - rootLeft),
          right: snapToPrecision(rect.right + scrollX - rootLeft),
          top: snapToPrecision(rect.top + scrollY2 - rootTop),
          bottom: snapToPrecision(rect.bottom + scrollY2 - rootTop)
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
      const declaredValues = /* @__PURE__ */ new Set();
      const numericOwners = /* @__PURE__ */ new Set();
      return fragments.filter((fragment) => {
        fragment.text = joinTextTokens(fragment.tokens).trim();
        delete fragment.tokens;
        const owner = fragment.owner;
        delete fragment.owner;
        if (owner?.hasAttribute("data-excel-value")) {
          if (declaredValues.has(owner)) return false;
          declaredValues.add(owner);
        } else if (["number", "currency"].includes(owner?.getAttribute("data-excel-type"))) {
          if (numericOwners.has(owner)) throw new Error("A numeric declaration spans separate regions; annotate each value separately");
          numericOwners.add(owner);
        }
        return Boolean(fragment.text);
      });
    };
    const getCellText = (cell, root) => {
      const descendants = Array.from(cell.querySelectorAll("*"));
      let text = cell.innerText || "";
      if (isElementExcluded(cell, root) || descendants.some((child) => isElementExcluded(child, root))) {
        const tokens = [];
        const visit = (node) => {
          if (node.nodeType === 3) {
            if (!isElementExcluded(node.parentElement, root)) tokens.push(textToken(node));
            return;
          }
          if (node.nodeType !== 1 || isElementBlacklisted(node, root)) return;
          const visible = !isElementExcluded(node, root);
          if (node.tagName === "BR") {
            if (visible) tokens.push({ text: "\n" });
            return;
          }
          const block = visible && node !== cell && /^(block|flex|grid|table|list-item|flow-root)$/.test(computedStyle(node).display);
          if (block) tokens.push({ block: true });
          for (const child of node.childNodes) visit(child);
          if (block) tokens.push({ block: true });
        };
        visit(cell);
        text = joinTextTokens(tokens);
      }
      const normalized = text.replace(/\u00a0/g, " ").replace(/\t+/g, " ");
      const hasInlineFloat = descendants.some((child) => !isElementExcluded(child, root) && child.style.cssFloat && child.style.cssFloat !== "none");
      return hasInlineFloat ? normalized.replace(/\s*\n\s*/g, " ").replace(/\s{2,}/g, " ").trim() : normalized;
    };
    const collectDomSnapshot = (root) => {
      snapshotExclusions = /* @__PURE__ */ new WeakMap();
      const rootRect = root.getBoundingClientRect();
      const scrollY2 = window.scrollY || window.pageYOffset || 0;
      const scrollX = window.scrollX || window.pageXOffset || 0;
      const rootLeft = snapToPrecision(rootRect.left + scrollX);
      const rootTop = snapToPrecision(rootRect.top + scrollY2);
      const rootWidth = snapToPrecision(root.scrollWidth);
      const rootHeight = snapToPrecision(root.scrollHeight);
      const allCells = Array.from(root.querySelectorAll("td, th")).filter(
        (cell) => hasExportableContent(cell, root)
      );
      const nestedCells = [];
      const leafCells = [];
      for (const cell of allCells) {
        if (cell.querySelector("table")) {
          nestedCells.push(cell);
        } else {
          leafCells.push(cell);
        }
      }
      const tableElements = [...root.matches("table") ? [root] : [], ...root.querySelectorAll("table")].filter(
        (tableElement) => hasExportableContent(tableElement, root)
      );
      const textFragments = [];
      const pageBreakMarkers = Array.from(root.querySelectorAll(CONFIG.PAGE_BREAK_SELECTOR)).filter(
        (marker) => !isElementExcluded(marker, root)
      );
      const colEdgeSet = /* @__PURE__ */ new Set([0, rootWidth]);
      const rowEdgeSet = /* @__PURE__ */ new Set([0, rootHeight]);
      const fragmentMetrics = { rootLeft, rootTop, scrollX, scrollY: scrollY2 };
      for (const cell of allCells) {
        const rect = cell.getBoundingClientRect();
        const width = snapToPrecision(rect.width);
        const height = snapToPrecision(rect.height);
        if (width <= 0 || height <= 0) continue;
        const left = snapToPrecision(rect.left + scrollX - rootLeft);
        const right = snapToPrecision(rect.right + scrollX - rootLeft);
        const top = snapToPrecision(rect.top + scrollY2 - rootTop);
        const bottom = snapToPrecision(rect.bottom + scrollY2 - rootTop);
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
        const top = snapToPrecision(rect.top + scrollY2 - rootTop);
        const bottom = snapToPrecision(rect.bottom + scrollY2 - rootTop);
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
        if (!marker || typeof marker.getBoundingClientRect !== "function") continue;
        const rect = marker.getBoundingClientRect();
        const top = snapToPrecision(rect.top + scrollY2 - rootTop);
        const bottom = snapToPrecision(rect.bottom + scrollY2 - rootTop);
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

  // src/layout.js
  function createLayout(dependencies) {
    const { CONFIG, collectDomSnapshot, computeEffectiveHorizontalPadding, computeEffectiveVerticalPadding, debugLog, extractAlignment, extractBorders, extractFont, extractBackground, getCellText, parseCellValue: parseCellValue2, isElementExcluded, snapToPrecision } = dependencies;
    const buildLayoutModel = (root) => {
      const { rootLeft, rootTop, rootWidth, rootHeight, allCells, nestedCells, leafCells, tableElements, textFragments, pageBreakMarkers, columnEdges, rowEdges } = collectDomSnapshot(root);
      const describeCell = (cell) => {
        const rect = cell.getBoundingClientRect();
        const scrollY2 = window.scrollY || window.pageYOffset || 0;
        const scrollX = window.scrollX || window.pageXOffset || 0;
        const rootLeft2 = snapToPrecision(root.getBoundingClientRect().left + scrollX);
        const rootTop2 = snapToPrecision(root.getBoundingClientRect().top + scrollY2);
        const width = snapToPrecision(rect.width);
        const height = snapToPrecision(rect.height);
        if (width <= 0 || height <= 0) return null;
        const left = snapToPrecision(rect.left + scrollX - rootLeft2);
        const right = snapToPrecision(rect.right + scrollX - rootLeft2);
        const top = snapToPrecision(rect.top + scrollY2 - rootTop2);
        const bottom = snapToPrecision(rect.bottom + scrollY2 - rootTop2);
        const colStartIndex = columnEdges.findIndex((edge) => Math.abs(edge - left) <= CONFIG.EDGE_TOLERANCE_PX);
        const colEndIndex = columnEdges.findIndex((edge) => Math.abs(edge - right) <= CONFIG.EDGE_TOLERANCE_PX);
        const rowStartIndex = rowEdges.findIndex((edge) => Math.abs(edge - top) <= CONFIG.EDGE_TOLERANCE_PX);
        const rowEndIndex = rowEdges.findIndex((edge) => Math.abs(edge - bottom) <= CONFIG.EDGE_TOLERANCE_PX);
        if (colStartIndex === -1 || colEndIndex === -1 || rowStartIndex === -1 || rowEndIndex === -1) {
          dependencies.warnLog("A cell has unmatched grid edges", { code: "GRID_EDGE_MISMATCH" });
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
          const tolerance = Math.max(CONFIG.EDGE_TOLERANCE_PX, 0.5);
          for (let i = 0; i < edges.length; i++) {
            if (Math.abs(edges[i] - value) <= tolerance) {
              return i;
            }
          }
          for (let i = 1; i < edges.length; i++) {
            if (value < edges[i]) {
              return i - 1;
            }
          }
          return edges.length - 1;
        }
        return Number.isInteger(fallbackIndex) ? fallbackIndex : 0;
      };
      const descriptorCache = /* @__PURE__ */ new WeakMap();
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
        const currencyMatch = parseCellValue2(textValue, cell);
        const exportedValue = currencyMatch.value;
        const cellModel = {
          element: cell,
          value: exportedValue,
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
            fill: fillColor ? { type: "pattern", pattern: "solid", fgColor: { argb: fillColor } } : null,
            numFmt: currencyMatch ? currencyMatch.numFmt : void 0
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
          const currencyMatch = parseCellValue2(fragment.text, referenceElement);
          const exportedValue = currencyMatch.value;
          cells.push({
            element: fragment.cell,
            value: exportedValue,
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
              fill: fillColor ? { type: "pattern", pattern: "solid", fgColor: { argb: fillColor } } : null,
              numFmt: currencyMatch ? currencyMatch.numFmt : void 0
            },
            images: [],
            sheetPosition: null
          });
        }
      }
      const imageOwners = /* @__PURE__ */ new Map();
      for (const img of root.querySelectorAll("img")) {
        if (isElementExcluded(img, root)) continue;
        const rect = img.getBoundingClientRect();
        if (rect.width <= 0 || rect.height <= 0) continue;
        const owner = img.closest("td, th");
        if (!owner || !root.contains(owner)) continue;
        let parentCell = imageOwners.get(owner);
        if (!parentCell) {
          const descriptor = getDescriptor(owner);
          if (!descriptor) continue;
          parentCell = {
            element: owner,
            bounds: { left: descriptor.left, top: descriptor.top, width: descriptor.width, height: descriptor.height },
            colStart: descriptor.colStartIndex,
            colEnd: Math.max(descriptor.colStartIndex + 1, descriptor.colEndIndex),
            rowStart: descriptor.rowStartIndex,
            rowEnd: Math.max(descriptor.rowStartIndex + 1, descriptor.rowEndIndex),
            styles: { alignment: extractAlignment(owner, window.getComputedStyle(owner), "") },
            sheetPosition: { colStart: descriptor.colStartIndex + 1, rowStart: descriptor.rowStartIndex + 1 }
          };
          imageOwners.set(owner, parentCell);
        }
        images.push({
          element: img,
          parentCell,
          width: rect.width,
          height: rect.height,
          left: snapToPrecision(rect.left + window.scrollX - rootLeft),
          top: snapToPrecision(rect.top + window.scrollY - rootTop)
        });
      }
      const computedStyleCache = /* @__PURE__ */ new WeakMap();
      const getCachedComputedStyle = (element) => {
        if (!element || typeof window === "undefined" || typeof window.getComputedStyle !== "function") {
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
        if (!target.styles.border[side]) {
          target.styles.border[side] = { ...borderStyle };
        }
      };
      const propagateSharedBorders = (cells2) => {
        const tolerance = Math.max(CONFIG.EDGE_TOLERANCE_PX, 1);
        for (const cell of cells2) {
          const border = cell.styles?.border;
          if (!border) continue;
          const cellLeft = cell.bounds.left;
          const cellRight = cell.bounds.left + cell.bounds.width;
          const cellTop = cell.bounds.top;
          const cellBottom = cell.bounds.top + cell.bounds.height;
          if (border.right) {
            for (const neighbor of cells2) {
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
            for (const neighbor of cells2) {
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
            for (const neighbor of cells2) {
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
            for (const neighbor of cells2) {
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
        pushDeferredBorderInstruction(descriptor, "top", borders.top);
        pushDeferredBorderInstruction(descriptor, "bottom", borders.bottom);
        pushDeferredBorderInstruction(descriptor, "left", borders.left);
        pushDeferredBorderInstruction(descriptor, "right", borders.right);
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
          return cellLeft >= outerLeft - tolerance && cellRight <= outerRight + tolerance && cellTop >= outerTop - tolerance && cellBottom <= outerBottom + tolerance;
        };
        const alignsWithOuterEdge = (cell, side) => {
          if (!cell?.bounds) return false;
          const diff = (() => {
            switch (side) {
              case "left":
                return Math.abs(cell.bounds.left - outerLeft);
              case "right":
                return Math.abs(cell.bounds.left + cell.bounds.width - outerRight);
              case "top":
                return Math.abs(cell.bounds.top - outerTop);
              case "bottom":
              default:
                return Math.abs(cell.bounds.top + cell.bounds.height - outerBottom);
            }
          })();
          if (diff <= tolerance) return true;
          const csCell = getCachedComputedStyle(cell.element);
          if (!csCell) return diff <= tolerance;
          if (side === "left" || side === "right") {
            const padding2 = computeEffectiveHorizontalPadding(cell.element, csCell);
            const allowance2 = side === "left" ? padding2.left : padding2.right;
            return diff <= Math.max(tolerance, allowance2 + CONFIG.EDGE_TOLERANCE_PX);
          }
          const padding = computeEffectiveVerticalPadding(cell.element, csCell);
          const allowance = side === "top" ? padding.top : padding.bottom;
          return diff <= Math.max(tolerance, allowance + CONFIG.EDGE_TOLERANCE_PX);
        };
        if (borders.top) {
          const topCells = cells.filter((c) => inRange(c) && alignsWithOuterEdge(c, "top"));
          for (const cell of topCells) {
            mergeBorderStyle(cell, "top", borders.top);
          }
        }
        if (borders.bottom) {
          const bottomCells = cells.filter(
            (c) => inRange(c) && alignsWithOuterEdge(c, "bottom")
          );
          for (const cell of bottomCells) {
            mergeBorderStyle(cell, "bottom", borders.bottom);
          }
        }
        if (borders.left) {
          const leftCells = cells.filter((c) => inRange(c) && alignsWithOuterEdge(c, "left"));
          for (const cell of leftCells) {
            mergeBorderStyle(cell, "left", borders.left);
          }
        }
        if (borders.right) {
          const rightCells = cells.filter(
            (c) => inRange(c) && alignsWithOuterEdge(c, "right")
          );
          for (const cell of rightCells) {
            mergeBorderStyle(cell, "right", borders.right);
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
        const tolerance = Math.max(CONFIG.EDGE_TOLERANCE_PX, 1);
        const totalRows = Math.max(0, rowEdges.length - 1);
        const rows = /* @__PURE__ */ new Set();
        for (const marker of pageBreakMarkers) {
          if (!marker || typeof marker.getBoundingClientRect !== "function") continue;
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

  // src/images.js
  function createImages(dependencies) {
    const { CONFIG, IMAGE_SCALE, debugLog, errorLog, isFiniteNumber, snapToPrecision, toPositiveNumber, warnLog } = dependencies;
    const ensureImageLoaded = (img) => {
      if (img.complete) {
        if (img.naturalWidth && img.naturalHeight) return Promise.resolve();
        return Promise.reject(new Error("Image failed to load (already complete with zero size)"));
      }
      return new Promise((resolve, reject) => {
        const timer = setTimeout(() => {
          cleanup();
          reject(new Error("Image load timed out"));
        }, CONFIG.IMAGE_TIMEOUT_MS);
        const cleanup = () => {
          clearTimeout(timer);
          img.removeEventListener("load", onLoad);
          img.removeEventListener("error", onError);
        };
        const onLoad = () => {
          cleanup();
          resolve();
        };
        const onError = (err) => {
          cleanup();
          reject(err);
        };
        img.addEventListener("load", onLoad, { once: true });
        img.addEventListener("error", onError, { once: true });
      });
    };
    const fetchImageAsBase64 = async (src, width, height) => {
      if (!src) return null;
      let url;
      try {
        url = new URL(src, window.location.href);
      } catch (err) {
        warnLog("Invalid image URL, skip fetch fallback", { code: "IMAGE_URL_INVALID", source: src, error: err });
        return null;
      }
      if (url.protocol === "file:") {
        warnLog("Skipping fetch fallback for file:// URL (browser security restriction)", { code: "IMAGE_FILE_UNSUPPORTED", source: url.href });
        return null;
      }
      try {
        const response = await fetch(url.href, { mode: "cors", signal: AbortSignal.timeout(CONFIG.IMAGE_TIMEOUT_MS) });
        if (!response.ok) throw new Error("Image fetch returned HTTP " + response.status);
        const blob = await response.blob();
        const objectUrl = URL.createObjectURL(blob);
        try {
          const image = new Image();
          image.src = objectUrl;
          await ensureImageLoaded(image);
          const canvas = document.createElement("canvas");
          canvas.width = width || image.naturalWidth;
          canvas.height = height || image.naturalHeight;
          const context = canvas.getContext("2d");
          if (!context) throw new Error("Canvas 2D context unavailable");
          context.drawImage(image, 0, 0, canvas.width, canvas.height);
          const base64 = canvas.toDataURL("image/png").split(",")[1];
          if (!base64) throw new Error("Empty base64 from fetched image");
          return { base64, extension: "png" };
        } finally {
          URL.revokeObjectURL(objectUrl);
        }
      } catch (err) {
        warnLog("Fallback fetch-to-base64 failed", { code: "IMAGE_FETCH_FAILED", source: src, error: err });
        return null;
      }
    };
    const convertImageJob = async (job) => {
      const { element, width, height } = job;
      if (!element) return { ...job, error: new Error("Image element missing") };
      try {
        await ensureImageLoaded(element);
      } catch (err) {
        warnLog("Image failed to load", { code: "IMAGE_LOAD_FAILED", source: element.currentSrc || element.src, error: err });
        return { ...job, error: err };
      }
      const imageScale = toPositiveNumber(IMAGE_SCALE) || 1;
      const renderWidth = Math.max(1, Math.round(width * imageScale));
      const renderHeight = Math.max(1, Math.round(height * imageScale));
      const canvas = document.createElement("canvas");
      canvas.width = renderWidth;
      canvas.height = renderHeight;
      const ctx = canvas.getContext("2d");
      if (!ctx) {
        const err = new Error("Canvas 2D context unavailable");
        errorLog(err.message, { code: "IMAGE_CANVAS_UNAVAILABLE", source: element.currentSrc || element.src });
        return { ...job, error: err };
      }
      try {
        ctx.drawImage(element, 0, 0, renderWidth, renderHeight);
        const dataUrl = canvas.toDataURL(CONFIG.IMAGE_FORMAT);
        const base64 = dataUrl.split(",")[1];
        if (!base64) throw new Error("Empty base64 from canvas");
        debugLog("Canvas conversion success", element.src, `${renderWidth}x${renderHeight}`, `base64Length=${base64.length}`);
        return {
          ...job,
          base64,
          extension: CONFIG.IMAGE_FORMAT === "image/png" ? "png" : "jpeg"
        };
      } catch (err) {
        warnLog("Canvas toDataURL failed, trying fetch fallback", { code: "IMAGE_CANVAS_FALLBACK", source: element.currentSrc || element.src, error: err });
        const fallback = await fetchImageAsBase64(element.currentSrc || element.src, renderWidth, renderHeight);
        if (fallback) {
          debugLog("Fallback fetch succeeded for image", element.src, `base64Length=${fallback.base64.length}`);
          return { ...job, base64: fallback.base64, extension: fallback.extension };
        }
        errorLog("Image conversion failed after fallback", { code: "IMAGE_CONVERSION_FAILED", source: element.currentSrc || element.src, error: err });
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
        case "inline":
        case "block":
        case "center":
        case "left":
        case "right":
        case "precise":
          return keyword;
        default:
          return null;
      }
    };
    const getCellImageAlignment = (cellModel, includeCss = true) => {
      if (!cellModel) return null;
      const element = cellModel.element;
      const attrName = CONFIG.IMAGE_ALIGNMENT_ATTR;
      if (attrName && element && typeof element.getAttribute === "function") {
        const attrValue = element.getAttribute(attrName);
        const resolved = toImageAlignment(attrValue);
        if (resolved) {
          return resolved;
        }
      }
      if (!includeCss) return null;
      const horizontal = cellModel.styles && cellModel.styles.alignment && cellModel.styles.alignment.horizontal ? String(cellModel.styles.alignment.horizontal).toLowerCase() : null;
      switch (horizontal) {
        case "center":
        case "middle":
          return "center";
        case "right":
          return "right";
        case "left":
          return "left";
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
      const leftEdge = columnEdges[cellModel.colStart] !== void 0 ? columnEdges[cellModel.colStart] : 0;
      const rightEdge = columnEdges[cellModel.colEnd] !== void 0 ? columnEdges[cellModel.colEnd] : leftEdge;
      const topEdge = rowEdges[cellModel.rowStart] !== void 0 ? rowEdges[cellModel.rowStart] : 0;
      const bottomEdge = rowEdges[cellModel.rowEnd] !== void 0 ? rowEdges[cellModel.rowEnd] : topEdge;
      const widthPx = Math.max(0, snapToPrecision(rightEdge - leftEdge));
      const heightPx = Math.max(0, snapToPrecision(bottomEdge - topEdge));
      return {
        widthPx,
        heightPx,
        columnEdges: columnEdges.slice(cellModel.colStart, cellModel.colEnd + 1).map((edge) => edge - leftEdge),
        rowEdges: rowEdges.slice(cellModel.rowStart, cellModel.rowEnd + 1).map((edge) => edge - topEdge),
        startColIndex: sheetPosition.colStart - 1,
        startRowIndex: sheetPosition.rowStart - 1
      };
    };
    const logLatestWorksheetImage = (worksheet) => {
      if (typeof worksheet?.getImages !== "function") {
        return;
      }
      const currentImages = worksheet.getImages();
      if (!Array.isArray(currentImages) || !currentImages.length) {
        return;
      }
      const latest = currentImages[currentImages.length - 1];
      if (latest) {
        debugLog("Worksheet image detail", {
          range: latest.range,
          hyperlink: latest.hyperlink || null
        });
      }
      debugLog("Worksheet image count", currentImages.length);
    };
    const coordinateAt = (edges, offset, startIndex = 0) => {
      const position = locateIndexAndOffset(edges, offset, true);
      const span = edges[position.index + 1] - edges[position.index];
      return startIndex + position.index + (span > 0 ? position.offset / span : 0);
    };
    const addAlignedImage = (worksheet, workbook, job, metrics, offsetLeftPx, offsetTopPx, widthPx, heightPx) => {
      if (!metrics || widthPx <= 0 || heightPx <= 0) return false;
      const tolerance = CONFIG.EDGE_TOLERANCE_PX;
      if (offsetLeftPx + widthPx > metrics.widthPx + tolerance || offsetTopPx + heightPx > metrics.heightPx + tolerance) {
        throw new Error("Image alignment exceeds its cell; enlarge the HTML cell, reduce imageWidthScale, or use precise placement");
      }
      const safeLeft = Math.max(0, offsetLeftPx || 0);
      const safeTop = Math.max(0, offsetTopPx || 0);
      const tl = {
        col: coordinateAt(metrics.columnEdges, safeLeft, metrics.startColIndex),
        row: coordinateAt(metrics.rowEdges, safeTop, metrics.startRowIndex)
      };
      const imgId = workbook.addImage({
        base64: job.base64,
        extension: job.extension
      });
      const ext = {
        width: widthPx,
        height: heightPx
      };
      worksheet.addImage(imgId, { tl, ext, editAs: "oneCell" });
      debugLog("Placed image using alignment flow", job.element?.src || "(unknown src)", {
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
      const finalImageScale = toPositiveNumber(CONFIG.IMAGE_WIDTH_SCALE) || 1;
      const rows = [];
      let currentRow = { items: [], width: 0, height: 0 };
      for (const job of images) {
        const widthPx = Math.max(1, job.width * finalImageScale);
        const heightPx = Math.max(1, job.height * finalImageScale);
        const item = { job, widthPx, heightPx };
        const needsWrap = currentRow.items.length > 0 && currentRow.width + spacingPx + widthPx > metrics.widthPx;
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
        if (justify === "center") {
          startLeft = Math.max(0, (metrics.widthPx - rowWidth) / 2);
        } else if (justify === "end") {
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
      const finalImageScale = toPositiveNumber(CONFIG.IMAGE_WIDTH_SCALE) || 1;
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
      const finalImageScale = toPositiveNumber(CONFIG.IMAGE_WIDTH_SCALE) || 1;
      const widthPx = Math.max(1, job.width * finalImageScale);
      const heightPx = Math.max(1, job.height * finalImageScale);
      const offsetLeft = Math.max(0, (metrics.widthPx - widthPx) / 2);
      const offsetTop = Math.max(0, (metrics.heightPx - heightPx) / 2);
      return addAlignedImage(worksheet, workbook, job, metrics, offsetLeft, offsetTop, widthPx, heightPx);
    };
    const placeImagesWithAlignment = (worksheet, workbook, images, cellModel, layoutModel, alignment) => {
      const keyword = alignment || "inline";
      if (keyword === "precise") return false;
      if (images.length === 1) {
        if (keyword === "center") {
          return placeSingleImageCentered(worksheet, workbook, images[0], cellModel, layoutModel);
        }
        const justify = keyword === "right" ? "end" : keyword === "center" ? "center" : "start";
        return placeImagesInline(worksheet, workbook, images, cellModel, layoutModel, justify);
      }
      switch (keyword) {
        case "block":
          return placeImagesStacked(worksheet, workbook, images, cellModel, layoutModel);
        case "right":
          return placeImagesInline(worksheet, workbook, images, cellModel, layoutModel, "end");
        case "center":
          return placeImagesInline(worksheet, workbook, images, cellModel, layoutModel, "center");
        case "left":
        case "inline":
        default:
          return placeImagesInline(worksheet, workbook, images, cellModel, layoutModel, "start");
      }
    };
    const placeImagesPrecisely = (worksheet, workbook, resolvedJobs, layoutModel) => {
      if (!resolvedJobs.length) return;
      const htmlColumnEdges = layoutModel.columnEdges;
      const htmlRowEdges = layoutModel.rowEdges;
      for (const job of resolvedJobs) {
        if (!job.base64) continue;
        const imgId = workbook.addImage({
          base64: job.base64,
          extension: job.extension
        });
        const parentCell = job.parentCell;
        if (!parentCell?.sheetPosition) continue;
        const tl = {
          col: coordinateAt(htmlColumnEdges, job.left),
          row: coordinateAt(htmlRowEdges, job.top)
        };
        const finalImageScale = toPositiveNumber(CONFIG.IMAGE_WIDTH_SCALE) || 1;
        const ext = {
          width: job.width * finalImageScale,
          height: job.height * finalImageScale
        };
        worksheet.addImage(imgId, { tl, ext, editAs: "oneCell" });
        debugLog("Placed image using direct pixel mapping", job.element?.src || "(unknown src)", {
          tl,
          ext,
          parentCellBounds: parentCell?.bounds
        });
        if (typeof worksheet.getImages === "function") {
          const currentImages = worksheet.getImages();
          if (Array.isArray(currentImages)) {
            const latest = currentImages[currentImages.length - 1];
            if (latest) {
              debugLog("Worksheet image detail", {
                range: latest.range,
                hyperlink: latest.hyperlink || null
              });
            }
            debugLog("Worksheet image count", currentImages.length);
          }
        }
      }
    };
    const placeImages = (worksheet, workbook, resolvedJobs, layoutModel) => {
      if (!resolvedJobs.length) return;
      const placementMode = String(CONFIG.IMAGE_PLACEMENT_MODE || "alignment").toLowerCase();
      if (placementMode === "precise") {
        placeImagesPrecisely(worksheet, workbook, resolvedJobs, layoutModel);
        return;
      }
      const cellGroups = /* @__PURE__ */ new Map();
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
      const defaultAlignment = toImageAlignment(CONFIG.IMAGE_ALIGNMENT) || "inline";
      for (const [cellModel, images] of cellGroups.entries()) {
        if (!Array.isArray(images) || !images.length) continue;
        const cellAlignment = getCellImageAlignment(cellModel, !CONFIG.IMAGE_ALIGNMENT_EXPLICIT) || defaultAlignment;
        const placed = placeImagesWithAlignment(worksheet, workbook, images, cellModel, layoutModel, cellAlignment);
        if (!placed) {
          preciseFallbackJobs.push(...images);
        }
      }
      if (preciseFallbackJobs.length) {
        placeImagesPrecisely(worksheet, workbook, preciseFallbackJobs, layoutModel);
      }
    };
    return { ensureImageLoaded, fetchImageAsBase64, convertImageJob, locateIndexAndOffset, clampOffset, toImageAlignment, getCellImageAlignment, getCellPlacementMetrics, logLatestWorksheetImage, addAlignedImage, placeImagesInline, placeImagesStacked, placeSingleImageCentered, placeImagesWithAlignment, placeImagesPrecisely, placeImages };
  }
  async function mapWithConcurrency(items, limit, mapper) {
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

  // src/workbook.js
  function createWorkbookWriter(dependencies) {
    const { CONFIG, EXPORT_SCALE, ExcelJS, FONT_SCALE, ROW_SCALE, convertImageJob, debugLog, errorLog, placeImages, pxToColumnWidth, pxToPoint, warnLog, mapWithConcurrency: mapWithConcurrency2 } = dependencies;
    const applyColumnWidths = (worksheet, columnWidthsPx) => {
      worksheet.columns = columnWidthsPx.map((px) => ({ width: pxToColumnWidth(px, EXPORT_SCALE) }));
    };
    const applyRowHeights = (worksheet, rowHeightsPx) => {
      rowHeightsPx.forEach((px, idx) => {
        const row = worksheet.getRow(idx + 1);
        const heightByRowScale = pxToPoint(px, ROW_SCALE);
        const heightByFontScale = pxToPoint(px, FONT_SCALE);
        row.height = Math.max(heightByRowScale, heightByFontScale);
        row.alignment = row.alignment || { vertical: "top" };
      });
    };
    const applyPageBreaks = (worksheet, layoutModel) => {
      const breaks = Array.isArray(layoutModel.pageBreakRows) ? layoutModel.pageBreakRows : [];
      if (!breaks.length) return;
      const existingBreaks = worksheet.rowBreaks || (worksheet.rowBreaks = []);
      const seen = new Set(
        existingBreaks.map((brk) => brk && Number.isFinite(brk.id) ? brk.id : null).filter((id) => id !== null)
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
      const occupancy = /* @__PURE__ */ new Set();
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
          const text = String(cell.value ?? "").trim();
          if (text && text !== String(existing ?? "").trim()) {
            throw new Error("Overlapping non-empty cells at " + key + "; adjust the HTML table structure or edgeTolerancePx");
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
        const existing = cell.border && typeof cell.border === "object" ? { ...cell.border } : {};
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
          case "left":
            for (let r = safeRowStart; r <= safeRowEnd; r++) {
              applyBorderSide(worksheet.getCell(r, safeColStart), "left", style);
            }
            break;
          case "right":
            for (let r = safeRowStart; r <= safeRowEnd; r++) {
              applyBorderSide(worksheet.getCell(r, safeColEnd), "right", style);
            }
            break;
          case "top":
            for (let c = safeColStart; c <= safeColEnd; c++) {
              applyBorderSide(worksheet.getCell(safeRowStart, c), "top", style);
            }
            break;
          case "bottom":
            for (let c = safeColStart; c <= safeColEnd; c++) {
              applyBorderSide(worksheet.getCell(safeRowEnd, c), "bottom", style);
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
        if (typeof pageSetupConfig.fitToWidth !== "number") {
          pageSetupConfig.fitToWidth = 1;
        }
        if (typeof pageSetupConfig.fitToHeight !== "number") {
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
      debugLog("Applied page setup", worksheet.pageSetup);
    };
    const createWorkbook = () => {
      const workbook = new ExcelJS.Workbook();
      workbook.created = /* @__PURE__ */ new Date();
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
      debugLog("Applying column widths and row heights");
      applyColumnWidths(worksheet, layoutModel.columnWidthsPx);
      applyRowHeights(worksheet, layoutModel.rowHeightsPx);
      debugLog("Writing cell content");
      writeCellsToSheet(worksheet, layoutModel);
      applyDeferredBorders(worksheet, layoutModel);
      applyPageBreaks(worksheet, layoutModel);
      return worksheet;
    };
    const processAndPlaceImages = async (workbook, worksheet, layoutModel) => {
      if (!layoutModel.images.length) return;
      debugLog("Converting images...");
      const imageResults = await mapWithConcurrency2(layoutModel.images, CONFIG.IMAGE_CONCURRENCY, convertImageJob);
      const successfulImages = imageResults.filter((item) => item && item.base64);
      const failedImages = imageResults.filter((item) => item && !item.base64);
      if (successfulImages.length) {
        placeImages(worksheet, workbook, successfulImages, layoutModel);
        debugLog("Embedded images", successfulImages.length);
      }
      if (failedImages.length) {
        failedImages.forEach((item) => {
          warnLog("Image skipped (no base64)", {
            code: "IMAGE_SKIPPED",
            source: item.element?.currentSrc || item.element?.src || "(unknown src)",
            width: item.width,
            height: item.height,
            error: item.error ? item.error.message || String(item.error) : "(no error message)"
          });
        });
        warnLog("Images could not be embedded", { code: "IMAGE_EMBED_FAILED", count: failedImages.length });
      }
    };
    const logWorkbookMediaSummary = (workbook) => {
      if (!CONFIG.DEBUG || !workbook) return;
      try {
        const model = workbook.model;
        const mediaCount = Array.isArray(model?.media) ? model.media.length : 0;
        const worksheetSummaries = Array.isArray(model?.worksheets) ? model.worksheets.map((ws) => ({
          name: ws.name,
          mediaItems: Array.isArray(ws.media) ? ws.media.length : 0,
          drawingCount: Array.isArray(ws.drawings) ? ws.drawings.length : 0
        })) : [];
        debugLog("Workbook media summary", { mediaCount, worksheetSummaries });
      } catch (err) {
        warnLog("Failed to inspect workbook model", { code: "WORKBOOK_INSPECTION_FAILED", error: err });
      }
    };
    const writeWorkbookToBuffer = async (workbook) => {
      try {
        logWorkbookMediaSummary(workbook);
        debugLog("Writing workbook buffer");
        const buffer = await workbook.xlsx.writeBuffer();
        debugLog("Workbook buffer size", buffer && (buffer.byteLength || buffer.length || 0));
        return buffer;
      } catch (err) {
        errorLog("writeBuffer failed", { code: "WORKBOOK_WRITE_FAILED", error: err });
        throw err;
      }
    };
    const sanitizeSheetName = (name) => (name || "Sheet1").replace(/[\\/*?:\[\]]/g, "").slice(0, 31) || "Sheet1";
    return { applyColumnWidths, applyRowHeights, applyPageBreaks, writeCellsToSheet, applyDeferredBorders, applyPageSetupOptions, createWorkbook, prepareWorksheet, processAndPlaceImages, logWorkbookMediaSummary, writeWorkbookToBuffer, sanitizeSheetName };
  }

  // src/export.js
  var version = true ? "0.1.0" : "development";
  function downloadBuffer(buffer, filename) {
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    try {
      link.click();
    } finally {
      link.remove();
      setTimeout(() => URL.revokeObjectURL(url), 1e3);
    }
  }
  function filenameFor(value) {
    const stamp = (/* @__PURE__ */ new Date()).toISOString().replace(/[.:]/g, "-");
    const base = String(value || document.title || "HTML export").replace(/[\\/?%*:|"<>]/g, "-").trim();
    return /\.xlsx$/i.test(base) ? base : base + "_" + stamp + ".xlsx";
  }
  async function exportXlsx(options = {}) {
    options = { ...options };
    if (typeof document === "undefined") throw new Error("exportXlsx requires a browser DOM");
    const root = typeof options.root === "string" ? document.querySelector(options.root) : options.root || document.querySelector("[data-xlsx-export]");
    if (!root || typeof root.getBoundingClientRect !== "function") throw new TypeError("Export root not found");
    const ExcelJS = options.exceljs || globalThis.ExcelJS;
    if (!ExcelJS?.Workbook) throw new Error("Load ExcelJS separately or supply options.exceljs");
    const { config: CONFIG, currency, locale } = resolveOptions(root, options);
    const warnings = [];
    const warnLog = (message, context = {}) => {
      const warning = { code: "EXPORT_WARNING", ...context, message: String(message) };
      if (warning.error instanceof Error) warning.error = warning.error.message;
      if (warning.source) {
        try {
          const url = new URL(warning.source, document.baseURI);
          warning.source = ["http:", "https:"].includes(url.protocol) ? url.origin + url.pathname : url.protocol;
        } catch {
          warning.source = "(invalid image URL)";
        }
      }
      warnings.push(warning);
      options.onWarning?.(warning);
    };
    const dependencies = {
      CONFIG,
      ExcelJS,
      escapeRegExp,
      warnLog,
      errorLog: warnLog,
      debugLog: (...args) => {
        if (CONFIG.DEBUG) console.debug("[html-to-xlsx]", ...args);
      },
      detectCurrency: (raw) => detectCurrency(raw, { currency, locale }),
      parseCellValue: (text, cell) => parseCellValue(text, cell, { currency, locale }),
      mapWithConcurrency
    };
    Object.assign(dependencies, createLayoutMetrics(dependencies));
    Object.assign(dependencies, createDom(dependencies));
    Object.assign(dependencies, createStyles(dependencies));
    Object.assign(dependencies, createLayout(dependencies));
    Object.assign(dependencies, createImages(dependencies));
    Object.assign(dependencies, createWorkbookWriter(dependencies));
    let fontTimer;
    if (document.fonts?.ready) {
      let timedOut;
      try {
        timedOut = await Promise.race([
          document.fonts.ready.then(() => false),
          new Promise((resolve) => {
            fontTimer = setTimeout(() => resolve(true), CONFIG.IMAGE_TIMEOUT_MS);
          })
        ]);
      } finally {
        clearTimeout(fontTimer);
      }
      if (timedOut) warnLog("Font readiness timed out; using available font metrics", { code: "FONT_TIMEOUT" });
    }
    await mapWithConcurrency(
      Array.from(root.querySelectorAll("img")),
      CONFIG.IMAGE_CONCURRENCY,
      async (image) => {
        if (dependencies.isElementExcluded(image, root)) return;
        try {
          await dependencies.ensureImageLoaded(image);
        } catch (error) {
          warnLog("An image did not load before layout measurement", {
            code: "IMAGE_PRELOAD_FAILED",
            source: image.currentSrc || image.src,
            error
          });
        }
      }
    );
    const layoutModel = dependencies.buildLayoutModel(root);
    if (!layoutModel.cells.length && !layoutModel.images.length) throw new Error("No visible table cells to export");
    const workbook = dependencies.createWorkbook();
    const worksheet = dependencies.prepareWorksheet(workbook, layoutModel);
    await dependencies.processAndPlaceImages(workbook, worksheet, layoutModel);
    const buffer = await dependencies.writeWorkbookToBuffer(workbook);
    const filename = filenameFor(options.filename);
    if (options.download !== false) downloadBuffer(buffer, filename);
    return {
      buffer,
      filename,
      warnings,
      stats: { cells: layoutModel.cells.length, images: worksheet.getImages().length }
    };
  }
  var bindings = /* @__PURE__ */ new WeakMap();
  function bindButton(button, options = {}) {
    if (typeof button === "string") button = document.querySelector(button);
    if (!button) throw new TypeError("Export button not found");
    bindings.get(button)?.();
    const handler = async (event) => {
      event.preventDefault();
      if (button.dataset.busy === "1") return;
      const previouslyDisabled = button.disabled;
      button.dataset.busy = "1";
      button.disabled = true;
      try {
        const result = await exportXlsx(options);
        button.dispatchEvent(new CustomEvent("xlsx:exported", { detail: result }));
      } catch (error) {
        button.dispatchEvent(new CustomEvent("xlsx:error", { detail: error }));
      } finally {
        button.disabled = previouslyDisabled;
        delete button.dataset.busy;
      }
    };
    button.addEventListener("click", handler);
    const unbind = () => {
      button.removeEventListener("click", handler);
      if (bindings.get(button) === unbind) bindings.delete(button);
    };
    bindings.set(button, unbind);
    return unbind;
  }

  // adapters/globe3.js
  function installGlobe3(api, documentRoot = document) {
    const root = documentRoot.querySelector(".exp_to_excel_button_content");
    if (!root) return;
    const options = {
      root,
      currency: root.getAttribute("data-excel-currency") || "USD",
      locale: root.getAttribute("data-excel-locale") || "en-US",
      exclude: ["[data-xlsx-ignore]", "table#mainapproval", "table#ext_id", "table#myMenu"],
      pageBreakSelector: ".pagebreak_bf_processed"
    };
    window.exportPrintFormToXlsx = () => api.exportXlsx(options);
    const button = documentRoot.querySelector(".html_to_excel_btn");
    if (button && !button.hasAttribute("data-xlsx-button")) {
      api.bindButton(button, options);
      button.addEventListener("xlsx:error", () => {
        alert("Export failed. Check the export container, options and ExcelJS dependency.");
      });
    }
  }

  // src/browser.js
  window.HtmlToXlsx = Object.freeze({ exportXlsx, bindButton, downloadBuffer, version });
  var initialize = () => {
    for (const button of document.querySelectorAll("[data-xlsx-button]")) {
      bindButton(button, { root: button.getAttribute("data-xlsx-target") || "[data-xlsx-export]" });
    }
    installGlobe3(window.HtmlToXlsx);
  };
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", initialize, { once: true });
  else initialize();
})();

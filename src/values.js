// MIT License. Copyright (c) 2026 YAP WEI JUN.
export const escapeRegExp = text => String(text).replace(/[.*+?^$()|[\]\\{}]/g, '\\$&');

const currencies = typeof Intl.supportedValuesOf === 'function'
    ? new Set(Intl.supportedValuesOf('currency')) : null;

export function validateCurrency(value) {
    if (value === undefined || value === null || value === '') return null;
    const code = String(value).trim().toUpperCase();
    if (!/^[A-Z]{3}$/.test(code) || (currencies && !currencies.has(code))) {
        throw new RangeError('Unsupported currency code: ' + code);
    }
    new Intl.NumberFormat('en', { style: 'currency', currency: code });
    return code;
}

// Compare decimal values independently of formatting and exponent notation.
// The export contract limits numeric cells to 15 significant decimal digits.
function decimalKey(value) {
    const match = String(value).match(/^([+-]?)(\d+)(?:\.(\d+))?(?:e([+-]?\d+))?$/i);
    if (!match) return null;
    let digits = (match[2] + (match[3] || '')).replace(/^0+/, '');
    if (!digits) return '0';
    let exponent = Number(match[4] || 0) - (match[3] || '').length;
    const trimmed = digits.replace(/0+$/, '');
    exponent += digits.length - trimmed.length;
    digits = trimmed;
    if (digits.length > 15) return null;
    return (match[1] === '-' ? '-' : '') + digits + 'e' + exponent;
}

const minusSigns = new Map();
function normalizeSigns(text, locale) {
    if (!minusSigns.has(locale)) {
        const minus = new Intl.NumberFormat(locale).formatToParts(-1).find(part => part.type === 'minusSign')?.value || '-';
        if (minusSigns.size >= 32) minusSigns.delete(minusSigns.keys().next().value);
        minusSigns.set(locale, minus);
    }
    const minus = minusSigns.get(locale);
    return minus && minus !== '-' ? text.split(minus).join('-') : text;
}

export function parseLocalizedNumber(value, locale = 'en-US') {
    if (typeof value === 'number') {
        return Number.isFinite(value) && Math.abs(value) <= Number.MAX_SAFE_INTEGER
            && decimalKey(value) !== null ? value : null;
    }
    let text = normalizeSigns(String(value ?? '').trim().replace(/[\u061c\u200e\u200f]/g, ''), locale);
    if (!text) return null;
    const negative = /^\(.*\)$/.test(text);
    if (negative) text = text.slice(1, -1);
    const formatter = new Intl.NumberFormat(locale);
    const parts = formatter.formatToParts(12345.6);
    const group = parts.find(part => part.type === 'group')?.value;
    const decimal = parts.find(part => part.type === 'decimal')?.value || '.';
    const digitFormatter = new Intl.NumberFormat(locale, { useGrouping: false });
    for (let digit = 0; digit < 10; digit++) {
        text = text.split(digitFormatter.format(digit)).join(String(digit));
    }
    const normalizedGroup = group?.replace(/\s/g, ' ');
    text = text.replace(/\s/g, ' ');
    const decimalParts = text.split(decimal);
    if (decimalParts.length > 2 || (decimalParts.length === 2 && !/^\d+$/.test(decimalParts[1]))) return null;
    if (normalizedGroup && normalizedGroup !== ' ' && text.includes(normalizedGroup)) {
        // Reject malformed grouping instead of turning arbitrary text into a plausible number.
        const groupedInteger = text.replace(/^[+-]/, '').split(decimal)[0];
        const sizes = groupedInteger.split(normalizedGroup).map(part => part.length);
        const sample = formatter.formatToParts(123456789)
            .filter(part => part.type === 'integer').map(part => part.value.length);
        const rightSize = sample.at(-1);
        const middleSize = sample.length > 2 ? sample.at(-2) : rightSize;
        if (sizes.at(-1) !== rightSize || sizes.slice(1, -1).some(size => size !== middleSize)
            || sizes[0] < 1 || sizes[0] > middleSize) return null;
        text = text.split(normalizedGroup).join('');
    }
    if (text.includes(' ')) {
        if (normalizedGroup !== ' ') return null;
        const integer = text.replace(/^[+-]/, '').split(decimal)[0];
        if (!/^\d{1,3}( \d{3})+$/.test(integer)) return null;
        text = text.replace(/ /g, '');
    }
    if (decimal !== '.') text = text.split(decimal).join('.');
    if (!/^[+-]?\d+(?:\.\d+)?$/.test(text) || (negative && /^[+-]/.test(text))) return null;
    const number = Number(text);
    const key = decimalKey(text);
    return Number.isFinite(number) && Math.abs(number) <= Number.MAX_SAFE_INTEGER
        && key !== null && key === decimalKey(number)
        ? (negative ? -number : number) : null;
}

const explicitSymbols = [
    ['HK$', 'HKD'], ['US$', 'USD'], ['S$', 'SGD'], ['NT$', 'TWD'],
    ['RMB', 'CNY'], ['RM', 'MYR'], ['JP¥', 'JPY'], ['€', 'EUR'],
    ['£', 'GBP'], ['₹', 'INR']
];

export function currencyFormat(currency, locale = 'en-US') {
    const decimals = new Intl.NumberFormat(locale, {
        style: 'currency', currency
    }).resolvedOptions().maximumFractionDigits;
    const fraction = decimals ? '.' + '0'.repeat(decimals) : '';
    return '"' + currency + '" #,##0' + fraction + ';("' + currency + '" #,##0' + fraction + ')';
}

export function detectCurrency(rawText, options = {}) {
    const locale = options.locale || 'en-US';
    let text = normalizeSigns(String(rawText ?? '').trim().replace(/[\u061c\u200e\u200f]/g, ''), locale);
    if (!text) return null;
    const accounting = /^\(.*\)$/.test(text);
    if (accounting) text = text.slice(1, -1).trim();
    const sign = text.match(/^[+-]/)?.[0] || '';
    if (accounting && sign) return null;
    if (sign) text = text.slice(1).trim();
    let code = null;
    const takeSymbol = symbol => {
        if (text.startsWith(symbol)) { text = text.slice(symbol.length).trim(); return true; }
        if (text.endsWith(symbol)) { text = text.slice(0, -symbol.length).trim(); return true; }
        return false;
    };
    // Aliases such as RMB must be resolved before ISO validation, longest first.
    for (const [symbol, currency] of explicitSymbols) {
        if (takeSymbol(symbol)) { code = currency; break; }
    }
    if (!code) {
        const prefix = text.match(/^([A-Za-z]{3})\s*/);
        const suffix = text.match(/\s*([A-Za-z]{3})$/);
        const explicit = prefix || suffix;
        if (explicit) {
            try { code = validateCurrency(explicit[1]); } catch { return null; }
            text = (prefix ? text.slice(prefix[0].length) : text.slice(0, -suffix[0].length)).trim();
        }
    }
    if (!code && options.currency) {
        code = validateCurrency(options.currency);
        const symbol = new Intl.NumberFormat(options.locale || 'en-US', {
            style: 'currency', currency: code
        }).formatToParts(1).find(part => part.type === 'currency')?.value;
        const tokens = [...new Set([symbol, '$', '¥'])].filter(Boolean).sort((a, b) => b.length - a.length);
        const token = tokens.find(candidate => text.startsWith(candidate) || text.endsWith(candidate));
        if (!token) return null;
        if (token === '¥' && !['JPY', 'CNY'].includes(code)) return null;
        takeSymbol(token);
    }
    if (!code) return null;
    if ((accounting || sign) && /^[+\-(]/.test(text)) return null;
    const value = parseLocalizedNumber(accounting ? '(' + text + ')' : sign + text, locale);
    return value === null ? null : { value, numFmt: currencyFormat(code, options.locale) };
}

export function parseCellValue(text, element, options = {}) {
    // Nested text may be owned by a span, but metadata never crosses into a parent table's cell.
    const attribute = name => {
        for (let current = element; current; current = current.parentElement) {
            const value = current.getAttribute?.(name);
            if (value !== null && value !== undefined) return value;
            if (current.tagName === 'TD' || current.tagName === 'TH') break;
        }
        return null;
    };
    const type = attribute('data-excel-type');
    const raw = attribute('data-excel-value') ?? text;
    const cellCurrency = attribute('data-excel-currency');
    const cellOptions = { ...options, currency: cellCurrency ? validateCurrency(cellCurrency) : options.currency };
    if (type === 'text') return { value: String(raw) };
    const currency = detectCurrency(raw, cellOptions);
    if (currency) return currency;
    if (type === 'number' || type === 'currency') {
        const value = parseLocalizedNumber(raw, cellOptions.locale);
        if (value === null) throw new TypeError('Invalid explicitly typed numeric cell');
        if (type === 'currency' && !cellOptions.currency) throw new TypeError('Currency cell requires a currency');
        return { value, numFmt: type === 'currency' ? currencyFormat(cellOptions.currency, cellOptions.locale) : undefined };
    }
    // Text beginning with "=" stays a string; HTML text is never promoted to formulas.
    return { value: String(raw ?? '') };
}

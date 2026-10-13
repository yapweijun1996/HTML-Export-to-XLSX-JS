import { test } from 'node:test';
import assert from 'node:assert/strict';
import { detectCurrency, parseLocalizedNumber, parseCellValue, validateCurrency, currencyFormat } from '../../src/values.js';

test('ambiguous dollar is text unless a document currency is configured', () => {
    assert.equal(detectCurrency('$ 1,200.00'), null);
    assert.deepEqual(detectCurrency('$ 1,200.00', { currency: 'HKD' }), { value: 1200, numFmt: currencyFormat('HKD') });
    assert.equal(detectCurrency('$ 1,200.00', { currency: 'USD' }).value, 1200);
    assert.equal(detectCurrency('$ 1,200.00', { currency: 'SGD' }).value, 1200);
});

test('source currency display keeps symbols and suffixes without changing numeric values', () => {
    const options = { currency: 'HKD', currencyDisplay: 'source' };
    assert.deepEqual(detectCurrency('$ 1,800.00', options), { value: 1800, numFmt: '"$" #,##0.00;("$" #,##0.00)' });
    assert.deepEqual(detectCurrency('1,800.00 HKD', options), { value: 1800, numFmt: '#,##0.00 "HKD";(#,##0.00 "HKD")' });
    assert.deepEqual(detectCurrency('HK$900.00', options), { value: 900, numFmt: '"HK$"#,##0.00;("HK$"#,##0.00)' });
    assert.equal(detectCurrency('($ 3,000.00)', options).value, -3000);
    assert.equal(detectCurrency('$ 1,800.00', { currency: 'HKD' }).numFmt, currencyFormat('HKD'));
});
test('explicit currency takes precedence and supports currencies beyond the legacy list', () => {
    for (const [raw, code, expected] of [
        ['HK$ 10.50', 'HKD', 10.5], ['USD 20.25', 'USD', 20.25],
        ['CHF 30.00', 'CHF', 30], ['KWD 1.234', 'KWD', 1.234],
        ['100 JPY', 'JPY', 100], ['INR 200.00', 'INR', 200]
    ]) {
        const result = detectCurrency(raw, { currency: 'SGD' });
        assert.equal(result.value, expected);
        assert.match(result.numFmt, new RegExp(code));
    }
    assert.ok(currencyFormat('KWD').includes('0.000'));
    assert.ok(!currencyFormat('JPY').includes('.00'));
});
test('regional separators, negative numbers and localized digits', () => {
    for (const [raw, locale, expected] of [
        ['1.234,56', 'de-DE', 1234.56], ['1\u202f234,56', 'fr-FR', 1234.56],
        ['12,34,567.89', 'en-IN', 1234567.89], ['١٬٢٣٤٫٥٦', 'ar-EG', 1234.56],
        ['(1,234.50)', 'en-US', -1234.5], ['-100.25', 'en-US', -100.25]
    ]) assert.equal(parseLocalizedNumber(raw, locale), expected, raw);
    assert.equal(detectCurrency('EUR 1.234,56', { locale: 'de-DE' }).value, 1234.56);
});
test('invalid numbers, grouping and unknown currencies do not silently become amounts', () => {
    for (const raw of ['1,2,3', '1,23.00', '12,345,67', 'Included', '2811 3096', '', 'Infinity', '9007199254740993']) {
        assert.equal(parseLocalizedNumber(raw), null, raw);
    }
    assert.throws(() => validateCurrency('XYZ'));
    assert.equal(detectCurrency('XYZ 100'), null);
});
test('explicit cell types retain leading zeros and never create formulas', () => {
    const cell = attrs => ({ getAttribute: key => attrs[key] ?? null });
    assert.deepEqual(parseCellValue('=SUM(A1:A2)', cell({})), { value: '=SUM(A1:A2)' });
    assert.deepEqual(parseCellValue('00123', cell({ 'data-excel-type': 'text' })), { value: '00123' });
    assert.equal(parseCellValue('5', cell({ 'data-excel-type': 'number' })).value, 5);
    assert.throws(() => parseCellValue('Included', cell({ 'data-excel-type': 'number' })));
    assert.equal(parseCellValue('$ 10', cell({ 'data-excel-currency': 'USD' }), { currency: 'HKD' }).value, 10);
});

test('fractional grouping separators are rejected in every supported grouping style', () => {
    for (const [raw, locale] of [
        ['1,234.5,6', 'en-US'], ['1234.5,6', 'en-US'], ['1.234,5.6', 'de-DE'],
        ['1\u202f234,5\u202f6', 'fr-FR'], ['12,34,567.8,9', 'en-IN']
    ]) assert.equal(parseLocalizedNumber(raw, locale), null, raw);
    assert.equal(parseLocalizedNumber('1,234.56', 'en-US'), 1234.56);
    assert.equal(parseLocalizedNumber('1\u202f234,56', 'fr-FR'), 1234.56);
});

test('numeric conversion does not silently round significant decimal digits', () => {
    const cell = type => ({ getAttribute: key => key === 'data-excel-type' ? type : null });
    for (const raw of ['9007199254740990.12', '123456789012345.67', '0.1234567890123456']) {
        assert.equal(parseLocalizedNumber(raw), null, raw);
        assert.throws(() => parseCellValue(raw, cell('number')), /numeric/);
        assert.equal(parseCellValue('HKD ' + raw, cell(null)).value, 'HKD ' + raw);
    }
    for (const raw of ['1234567890123.45', '000123.4500', '0.00000000000000123', '1000000000000000']) {
        assert.equal(parseLocalizedNumber(raw), Number(raw), raw);
    }
    assert.equal(parseLocalizedNumber(0.1234567890123456), null);
    assert.equal(parseLocalizedNumber(Number.MAX_SAFE_INTEGER + 1), null);
    assert.equal(parseLocalizedNumber(123.45), 123.45);
});

test('fragment attributes resolve within their own cell boundary', () => {
    const element = (tagName, attrs = {}, parentElement = null) => ({
        tagName, parentElement, getAttribute: key => attrs[key] ?? null
    });
    const outer = element('TD', { 'data-excel-type': 'text', 'data-excel-currency': 'HKD' });
    const label = element('SPAN', {}, outer);
    assert.deepEqual(parseCellValue('$ 123.00', label), { value: '$ 123.00' });
    const amount = element('SPAN', { 'data-excel-type': 'currency', 'data-excel-value': '12.50' }, outer);
    assert.deepEqual(parseCellValue('display', element('B', {}, amount)), {
        value: 12.5, numFmt: currencyFormat('HKD')
    });
    const inner = element('TD', {}, element('TABLE', {}, outer));
    assert.equal(parseCellValue('$ 10', inner, { currency: 'USD' }).numFmt, currencyFormat('USD'));
});

test('currency aliases, suffix symbols and signed accounting amounts retain numeric values', () => {
    for (const [text, options, value, currency] of [
        ['RMB 100', {}, 100, 'CNY'], ['RM 100', {}, 100, 'MYR'],
        ['1.234,56 €', { locale: 'de-DE' }, 1234.56, 'EUR'],
        ['-$1,234.56', { currency: 'USD' }, -1234.56, 'USD'],
        ['($1,234.56)', { currency: 'USD' }, -1234.56, 'USD'],
        ['(HK$ 10.25)', {}, -10.25, 'HKD'], ['USD -10.25', {}, -10.25, 'USD'],
        ['+USD 10.25', {}, 10.25, 'USD'], ['100 RMB', {}, 100, 'CNY']
    ]) {
        const result = detectCurrency(text, options);
        assert.equal(result?.value, value, text);
        assert.match(result.numFmt, new RegExp(currency));
    }
    for (const text of ['(-$100)', '-USD -100', '-USD (100)', 'USD 1,23.0', 'XYZ 100', '100.5,6 €']) {
        assert.equal(detectCurrency(text, { currency: 'USD' }), null, text);
    }
});

test('locale-generated negative numbers and currency strings round-trip without losing their signs', () => {
    for (const locale of ['sv-SE', 'fa-IR', 'ar-EG', 'de-DE', 'en-US']) {
        assert.equal(parseLocalizedNumber(new Intl.NumberFormat(locale).format(-1234.56), locale), -1234.56, locale);
        const amount = new Intl.NumberFormat(locale, { style: 'currency', currency: 'USD' }).format(-1234.56);
        assert.equal(detectCurrency(amount, { locale, currency: 'USD' })?.value, -1234.56, amount);
    }
});

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { detectCurrency, parseLocalizedNumber, parseCellValue, validateCurrency, currencyFormat } from '../../src/values.js';

test('ambiguous dollar is text unless a document currency is configured', () => {
    assert.equal(detectCurrency('$ 1,200.00'), null);
    assert.deepEqual(detectCurrency('$ 1,200.00', { currency: 'HKD' }), { value: 1200, numFmt: currencyFormat('HKD') });
    assert.equal(detectCurrency('$ 1,200.00', { currency: 'USD' }).value, 1200);
    assert.equal(detectCurrency('$ 1,200.00', { currency: 'SGD' }).value, 1200);
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

# HTML Export to XLSX JS

[![CI](https://github.com/yapweijun1996/HTML-Export-to-XLSX-JS/actions/workflows/ci.yml/badge.svg)](https://github.com/yapweijun1996/HTML-Export-to-XLSX-JS/actions/workflows/ci.yml)

A browser library that exports visible HTML tables and nested print layouts to
editable XLSX cells, with borders, fonts, merged cells and embedded images.

**Version 0.1.0 is a preview.** Data completeness and editable numeric amounts are
the primary goals. Excel calculates its own wrapping, font metrics and print
pagination; the result is not a pixel-identical reproduction of a browser PDF.

The library is MIT licensed. ExcelJS is a separately loaded dependency with its
own license. No npm publication is provided for this preview.

## Quick start

Download the JavaScript files from a [versioned release](https://github.com/yapweijun1996/HTML-Export-to-XLSX-JS/releases).
Load a locally hosted ExcelJS browser bundle, then the minified exporter:

~~~html
<button data-xlsx-button data-xlsx-target="#quotation">Export XLSX</button>
<div id="quotation" data-xlsx-export
     data-excel-currency="HKD" data-excel-locale="en-HK">
  <table>
    <tr><th>Item</th><th>Qty</th><th>Amount</th></tr>
    <tr><td>A1</td><td data-excel-type="number">2</td><td>$ 1,200.00</td></tr>
  </table>
</div>
<script src="./exceljs.min.js"></script>
<script src="./export_to_xlsx.min.js"></script>
~~~

Both dist files have the same behavior. The readable file aids debugging.
ExcelJS 4.4.0 is the tested dependency; it is not bundled in our dist files.

## JavaScript API

~~~js
const result = await HtmlToXlsx.exportXlsx({
  root: '#quotation',
  currency: 'HKD',
  locale: 'en-HK',
  filename: 'quotation.xlsx',
  download: false
});
// result.buffer: serialized XLSX bytes
// result.filename: safe download filename
// result.warnings: image/font or measurement warnings
// result.stats: cell and embedded image counts
HtmlToXlsx.downloadBuffer(result.buffer, result.filename);
~~~

The core API is exported from src/export.js. The browser bundle exposes
window.HtmlToXlsx. A root can be an Element or a CSS selector. Without a root option,
the first [data-xlsx-export] element is used.

HtmlToXlsx.bindButton(button, options) returns an unbind function. Binding the same
button again replaces the earlier handler. Successful button exports emit
xlsx:exported with the result in event.detail; failures emit xlsx:error with
the Error. Direct API calls reject on failure. Buttons are restored after either
outcome. Loading the script after DOMContentLoaded is supported.

### Options

| Option | Default | Purpose |
|---|---|---|
| root | first [data-xlsx-export] | Element or selector |
| currency | container attribute, otherwise unset | ISO currency for ambiguous symbols |
| locale | container attribute, otherwise en-US | Input separators and localized digits |
| filename | document title plus timestamp | Optional .xlsx filename |
| sheetName | PrintForm | Excel worksheet name |
| download | true | Return bytes without downloading when false |
| exceljs | globalThis.ExcelJS | Optional explicit dependency |
| exclude | ['[data-xlsx-ignore]'] | Omit elements and descendants |
| pageBreakSelector | [data-xlsx-page-break] | Markers mapped to Excel manual breaks |
| pageSetup | A4 portrait, fit one page wide | ExcelJS setup overrides; margins merge |
| edgeTolerancePx | 2 | Coordinate clustering tolerance |
| exportScale, rowScale | 1 | Export and row size scaling |
| fontScale | 0.9 | Font size scaling |
| columnWidthScale | 1 | Column width adjustment |
| imageScale, imageWidthScale | 3, 0.92 | Raster quality and placement size |
| imagePlacement | alignment | alignment or precise |
| imageAlignment | inline | inline, block, center, left, right, precise |
| imageTimeoutMs, imageConcurrency | 10000, 4 | Bounded image waits and concurrency |
| onWarning | unset | Callback receiving { code, message } |
| debug | false | Diagnostic logs; use non-sensitive fixtures |

Unknown options and invalid explicitly typed values reject. Each export owns its
configuration and caches; concurrent exports can use different currencies without
changing global settings.

### Currency and value rules

Precedence is: explicit currency in cell text, cell currency attribute, API
currency option, then container currency attribute. The generic API leaves a bare
$ as text when no currency is configured; it does not guess a country.

Currency codes and decimal places use the runtime's Intl support. Explicit ISO
prefixes or suffixes support currencies beyond a hardcoded list. Common
unambiguous symbols such as HK$, US$, S$, NT$, EUR's euro symbol, GBP's pound symbol
and INR's rupee symbol are recognized. Other symbols require explicit currency
configuration or a typed cell. Locale describes input formatting, not currency.

~~~html
<td data-excel-type="text">00123</td>
<td data-excel-type="number">12</td>
<td data-excel-type="currency" data-excel-currency="KWD">1.234</td>
<td data-excel-type="number" data-excel-value="1200.50">1,200.50</td>
~~~

The data-excel-value attribute uses the selected locale. Explicit numeric cells
reject invalid values. Unmarked plain numeric-looking text remains text,
preserving phone numbers and identifiers. Strings beginning with = remain text
and never become formulas. This library exports displayed amounts; it does not
calculate quotation totals or discounts.

### Images, merges and warnings

Images are embedded in the workbook. Same-origin or CORS-enabled images work;
browser security restrictions cannot be bypassed. Image failures produce warnings
while valid content continues. Inspect warnings before treating an export as
complete. Image load waits and processing concurrency are bounded.

The grid is reconstructed from visible table geometry. Meaningful cells colliding
at an occupied position reject the export instead of silently dropping a value.
Tiny or overlapping HTML regions may need a clearer table structure or a different
edgeTolerancePx. Arbitrary CSS grid layouts, formula generation and pixel-identical
PDF printing are outside this preview's contract.

## Globe3 compatibility

ERP-specific defaults belong to [adapters/globe3.js](adapters/globe3.js).
The bundled adapter activates only when .exp_to_excel_button_content exists.
It preserves .html_to_excel_btn and window.exportPrintFormToXlsx(), legacy
exclusions and .pagebreak_bf_processed markers.

For dollar-based forms without currency metadata, the adapter retains the legacy
USD interpretation. HKD forms should declare their currency:

~~~html
<div class="exp_to_excel_button_content" data-excel-currency="HKD"
     data-excel-locale="en-HK">
  <!-- Existing printform HTML -->
</div>
~~~

Keep ExcelJS separately loaded. Replace the shared exporter with a pinned release
only during an authorized deployment. Remove any form-local ExcelJS.Workbook
override once the standard currency configuration is used. Generic examples
require no Globe3, CFML, database or authentication.

## Source and distribution

| Path | Responsibility |
|---|---|
| src/config.js, src/values.js | Per-export options, numeric parsing |
| src/dom.js, src/layout.js | DOM snapshot and cell grid |
| src/layout-metrics.js, src/styles.js | Units, coordinates, visual styles |
| src/images.js, src/workbook.js | Images and ExcelJS serialization |
| src/export.js, src/browser.js | Public API and browser integration |
| adapters/ | Application-specific compatibility |
| dist/ | Generated readable/minified bundles and SHA256 manifest |
| tests/ | Unit, serialization, browser and original-layout regressions |

Only source is edited. One esbuild script produces dist; ExcelJS stays external.
Package.json owns the release version, which the build injects into the browser
API and bundle headers. Direct source imports identify their version as development.
CI rebuilds in memory and rejects stale committed dist files.

## Development and verification

~~~sh
npm ci
npm run build
npm test
npx playwright install chrome msedge firefox
npm run test:browser
npm run build:check
~~~

Node.js 22 or later is required for development. Start the local fixture server
with node scripts/serve.mjs and open http://127.0.0.1:19483/examples/basic.html
to try the generic example. Fixtures use synthetic HTML and
images, never ERP sessions or customer files. Tests run on Chrome, Edge and
Playwright Firefox against readable and minified bundles, then inspect serialized
workbooks. They cover 68 item labels, subtotals, currencies, editable quantities,
images, exclusions, button download, locales, failures and concurrent options.

Original source is preserved at Git commit 4d93f11 for regression comparison.
CI fetches full history so the fixture server can read this baseline without a
second maintained runtime copy.

ExcelJS 4.4.0 currently reports a moderate uuid advisory through npm audit.
This project does not vendor that dependency or apply a breaking forced downgrade
to hide the report. Review dependency advisories before upgrading; monthly
Dependabot checks help track them.

See [CONTRIBUTING.md](CONTRIBUTING.md), [CHANGELOG.md](CHANGELOG.md) and [LICENSE](LICENSE).

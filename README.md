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
outcome. Buttons bound automatically through data-xlsx-button also log failures
to the browser console. Loading the script after DOMContentLoaded is supported.

### Options

| Option | Default | Purpose |
|---|---|---|
| root | first [data-xlsx-export] | Element or selector |
| currency | container attribute, otherwise unset | ISO currency for ambiguous symbols |
| currencyDisplay | code | code or source; source retains a detected symbol/code, spacing and position |
| locale | container attribute, otherwise en-US | Input separators and localized digits |
| cellType | unset | Callback (element, text, context) returning text, number, currency or undefined for cells without type metadata |
| filename | document title plus timestamp | Optional .xlsx filename |
| sheetName | PrintForm | Excel worksheet name |
| download | true | Return bytes without downloading when false |
| exceljs | globalThis.ExcelJS | Optional explicit dependency |
| exclude | ['[data-xlsx-ignore]'] | Omit elements and descendants |
| pageBreakSelector | [data-xlsx-page-break] | Markers mapped to Excel manual breaks |
| respectPageBreaks | true | Convert Fit To to percentage scaling for marked pages |
| pageSetup | A4 portrait, fit one page wide | ExcelJS setup overrides; margins merge; marked pages use percentage scaling |
| edgeTolerancePx | 2 | Coordinate clustering tolerance |
| exportScale, rowScale | 1 | Export and row size scaling |
| fontScale | 0.9 | Font size scaling |
| normalFont | null | Optional { name, size } for the workbook Normal style and column metrics |
| columnWidthScale | 1 | Column width adjustment |
| minColumnWidthPx | 10 | Minimum width of a grid fragment |
| preserveSingleLine | false | Keep all visibly single-line labels on one line with Shrink to Fit |
| preserveEmailAddresses | false | Fit a whole email address and optional parenthesized contact note on one line |
| preserveRichText | false | Retain mixed visible inline font runs in string cells |
| includeBackgrounds | true | Retain background fills; false matches browser printing without backgrounds |
| imageScale, imageWidthScale | 3, 0.92 | Raster quality and placement size |
| imagePlacement | alignment | alignment or precise |
| imageAlignment | inline | inline, block, center, left, right, precise |
| imageTimeoutMs, imageConcurrency | 10000, 4 | Bounded image waits and concurrency |
| onWarning | unset | Callback receiving { code, message, source?, error?, count? } |
| debug | false | Diagnostic logs; use non-sensitive fixtures |

Unknown options and invalid explicitly typed values reject. Each export owns its
configuration and DOM caches; concurrent exports can use different currencies without
changing global settings.

### Currency and value rules

Precedence is: explicit currency in cell text, cell currency attribute, API
currency option, then container currency attribute. The generic API leaves a bare
$ as text when no currency is configured; it does not guess a country.

Currency codes use the runtime's Intl support. Amounts display the decimal places
written in their text, so $ 12.3456 shows four decimals and $ 100 none; a typed
currency value without numeric display text uses the currency's standard decimals.
Explicit ISO prefixes or suffixes support currencies beyond a hardcoded list. ISO
codes must be uppercase: mixed-case words that are also codes, such as Top 10 or
pen 2, remain text. Uppercase words such as TOP 10 still match; mark them
data-excel-type="text". Common
unambiguous symbols such as HK$, US$, S$, NT$, EUR's euro symbol, GBP's pound symbol
and INR's rupee symbol are recognized. Other symbols require explicit currency
configuration or a typed cell. Locale describes input formatting, not currency.
RMB is an alias for CNY, while RM denotes MYR. Currency symbols can precede or
follow the amount. Leading signs, accounting parentheses and locale-specific
minus signs are supported, including -$1,234.56, ($1,234.56) and 1.234,56 €
(with locale de-DE). Conflicting signs or malformed grouping remain invalid.

~~~html
<td data-excel-type="text">00123</td>
<td data-excel-type="number">12</td>
<td data-excel-type="currency" data-excel-currency="KWD">1.234</td>
<td data-excel-type="number" data-excel-value="1200.50">1,200.50</td>
~~~

Typed numbers keep the grouping, decimal places and accounting parentheses of their
displayed text, so 2.00 displays as 2.00 and 1,250 keeps its separator. A declared
data-excel-value behind non-numeric display text uses the General format.
The data-excel-value attribute uses the selected locale. Explicit numeric cells
reject invalid grouping (including separators in the fractional part), more than
15 significant decimal digits, values outside the safe integer magnitude, and
values that change when converted to a JavaScript number. Unmarked amounts that
cannot be converted safely remain text. Trailing decimal zeros do not consume
precision. Unmarked plain numeric-looking text remains text,
preserving phone numbers and identifiers. Strings beginning with = remain text
and never become formulas. This library exports displayed amounts; it does not
calculate quotation totals or discounts.

For legacy HTML, cellType can explicitly declare a value's type without changing
the DOM. Existing data-excel-type attributes take precedence. Returning undefined
keeps the normal parsing behavior; numeric declarations retain the same locale and
precision checks as typed HTML cells.
The callback context contains root and is unique to each export, allowing adapters
to cache classification metadata for one settled DOM snapshot.

For text alongside a nested table, metadata is resolved from the text's element
up to its containing td/th, stopping before the parent table's cell. Formatting
nodes belonging to one declared value are exported together, so an explicit
data-excel-value is emitted once at its first visible run. A child table or nested
declaration separates text runs; their rectangles are never merged across the
intervening content. Put separate numeric values in separately annotated spans:
a numeric declaration without an explicit value that crosses those boundaries
rejects the export. Child-table cells retain their own metadata and document defaults.

### Images, merges and warnings

Images are embedded in the workbook. Same-origin or CORS-enabled images work;
browser security restrictions cannot be bypassed. Image failures produce warnings
while valid content continues. Inspect warnings before treating an export as
complete. Image load waits and processing concurrency are bounded. An image that
fails or times out is not waited for again in the same export. Browsers do not load
off-screen loading="lazy" images, so those images are omitted with warnings unless
they are loaded before export. Rasters stay within 16,777,216 pixels and 16,384
pixels per side, the smallest common canvas limits; larger images keep their drawing
size at a lower resolution. Images beside
nested tables are included; images with zero rendered width or height are omitted.
Images outside table cells, such as a logo in a div, are placed at their rendered
position in both placement modes.
Fetch fallback uses the browser-selected currentSrc (including picture/srcset)
and rasterizes the fetched image to PNG, including SVG and WebP sources.
Identical raster bytes and formats share one workbook media entry, including across
different cells; each visible placement remains a separate drawing. Different
raster sizes, pixels or formats are not combined, and reuse is scoped to a workbook.
Images use valid two-cell drawing anchors with `editAs="oneCell"` to preserve
fixed image sizes and avoid ExcelJS 4.4's invalid attribute on one-cell anchors.

Warnings include a specific code and, when available, source, error and count.
Source URLs omit credentials, query parameters and fragments; inline data is not
copied into warnings. IMAGE_CANVAS_FALLBACK indicates a recoverable conversion
attempt; IMAGE_SKIPPED indicates an omitted image. IMAGE_ALIGNMENT_FALLBACK
indicates images placed at their rendered position because an unforced alignment
flow did not fit their cell. If onWarning throws, the export
rejects and a bound button is restored, including after a font-readiness timeout.

Exclusion selectors apply to descendants inside cells as well as whole tables.
Hidden text and images are omitted before value parsing and export image processing;
visible children that explicitly restore CSS visibility are retained. Filtering
preserves explicit line breaks without changing the source DOM. This is an export
filter, not an authorization boundary: protected data must not be sent to the page.
Filtered text retains text case transformations and whitespace modes. Excluded
content does not contribute font size, weight or descendant padding. Backgrounds
resolve from the nearest visible ancestor, and RTL start/end alignment respects
the element's direction. Horizontal alignment follows the rendered text, including
legacy align attributes on rows and inner blocks such as center or div align.
A table used directly as root retains its background and borders.
Uniform inline font families, colors and italics are read from visible text rather
than only the outer cell. Underline and strike-through decorations include inline
ancestors. Mixed inline properties retain the cell-level fallback by default.
With preserveRichText, mixed visible font runs retain the existing normalized
text, line breaks, font sizes, colors and decorations. Numeric values stay numeric;
metadata overrides or text transformations that cannot be matched stay plain.
Hidden content and nested child tables do not influence these
font properties. Bold or underlined captions that occupy one rendered line use
Shrink to Fit within the original grid. Explicit line breaks and naturally wrapped
captions retain text wrapping. preserveSingleLine extends this behavior to plain
labels, including summary descriptions.

Globe3 quantity inference also accepts numbered row IDs marked with _processed
by the legacy print engine, including Summary and Optional rows without rowitem.

preserveEmailAddresses also keeps an email address with an optional parenthesized
contact note on one line using Shrink to Fit. Explicit line breaks remain intact.
The Globe3 profile enables this option; other callers retain the default wrapping.

Cell and enclosing-table borders follow their measured grid perimeters. Borders
are not copied onto neighboring cells or padded inner cells. Perimeter styles are
detached after merging so an edge segment does not spread through a shared merged
style. Explicit cell borders take precedence over enclosing-table borders.

Column dimensions follow the [Open XML Column width definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.column?view=openxml-3.0.1),
using the workbook Normal font's Calibri 11 maximum digit width of 7px at 96 DPI.
Column and row edges use cumulative whole-pixel rounding, avoiding drift across
fragmented grids. Drawing start/end offsets use the same physical dimensions.
Images starting on manual page boundaries have a small inset to prevent clipping
residue on the preceding page. The default print area covers the exported grid;
an explicit pageSetup.printArea takes precedence.
With normalFont, the digit width is measured for that font. ExcelJS 4.4 does not
expose a Normal-style setter, so the exporter uses a per-writer style serialization
hook, restored after success or failure. Explicit CSS cell fonts keep their own
font records. Unsupported ExcelJS serialization rejects rather than producing a
workbook with inconsistent dimensions. preserveRichText also reserves enough
merged-row height for explicit multiline text, preventing clipped address lines.

Alignment-mode images use the merged cell's actual column and row grid for left,
center and right positioning, without fixed pixel compensation. Inline and stacked
images retain their spacing. Alignment precedence is the cell's
data-excel-image-align attribute, an explicitly supplied imageAlignment option,
CSS text alignment, then the default inline flow. Inline images wrap by available
width rather than a fixed image count. Vertically, a flow starts at the images'
rendered top, so it does not cover text above them, and moves up only as far as the
cell requires. If a forced flow (the cell attribute or an explicit imageAlignment)
exceeds the cell's size, the export rejects instead of overlapping images; enlarge
the HTML cell, reduce imageWidthScale or use precise placement. An unforced flow
that does not fit places those images at their rendered positions and warns with
IMAGE_ALIGNMENT_FALLBACK. Precise mode maps the original image
coordinates through all spanned columns and rows, without fixed pixel compensation.

The grid is reconstructed from visible table geometry. Meaningful cells colliding
at an occupied position reject the export instead of silently dropping a value.
Edge clusters are bounded from their first coordinate, preventing chains of close
edges from collapsing beyond edgeTolerancePx. Cell and text-fragment coordinates
select the nearest retained edge within that tolerance, including exact matches.
Meaningful content is written before blank layout regions. Blank regions that
overlap content decorate only free grid cells and never merge across that content.
Collision checks cover the entire merged range and compare source values, including
rich text. If clustered boundaries still collide, the exporter rebuilds once using
all measured edges and emits GRID_REFINED on success. A content collision that
remains in that grid rejects the export; no conflicting value is silently dropped.
Tiny or overlapping HTML regions may need a clearer table structure or a different
edgeTolerancePx. Arbitrary CSS grid layouts, formula generation and pixel-identical
PDF printing are outside this preview's contract.

## Globe3 compatibility

ERP-specific defaults belong to [adapters/globe3.js](adapters/globe3.js).
The bundled adapter activates only when .exp_to_excel_button_content exists.
It preserves .html_to_excel_btn and window.exportPrintFormToXlsx(), legacy
exclusions and .pagebreak_bf_processed markers.
Safe numeric values in legacy table.rowitem and table#rowN quantity cells become editable numbers
when both cell edges match a preceding Qty/Quantity or 數量/数量 header in
table.rowheader, table#rowheader, table.rowitem or table#rowN (N is a numeric row ID).
This includes section headers, Summary rows and their quantity columns.
Other untyped values, including phone numbers
and item identifiers, remain text. This classification does not modify the DOM.
Bound-button failures log the exception message to the browser console before the
compatibility alert. Unresolved collisions include source element IDs, ranges,
geometry and value types, without logging document values or error stacks.
It also recognizes zero-height .pagebreak_bf markers emitted by the current ERP
paginator and explicit [data-xlsx-page-break] markers.

Excel can ignore manual page breaks in Fit To mode. When markers are present and
respectPageBreaks is true, the exporter calculates a percentage scale from the
worksheet width, tallest marked page, paper size and margins. Each marked page
then starts on a separate printed page. Forms without markers keep Fit To mode.
Set respectPageBreaks: false to keep an explicit Fit To choice, or set
pageSetup.fitToPage: false with pageSetup.scale to control the percentage yourself.
Automatic percentage scaling supports Letter, Legal, Executive, A3, A4, A5 and
JIS B5. Unsupported paper sizes and pages needing less than Excel's minimum 10%
scale emit a warning and require explicit print settings or a smaller layout.

Formatted numeric cells use Shrink to Fit instead of text wrapping, so expanded
currency codes fit the original grid while values remain editable numbers.
Text cells, merged regions, column widths and row heights retain their geometry.

The Globe3 adapter uses a print profile: fontScale and imageWidthScale are 1,
the workbook Normal style is Arial 9 to match the usual 12px printform text,
imagePlacement is precise, minimum column fragments are 1px, and visible single
lines and mixed inline fonts are preserved. Amounts use the document currency code
(for example HKD) in their number format without changing numeric values. Background fills are omitted to match browser
PDF printing with backgrounds disabled. Page content is left-aligned; margins
are 0.25in left/right and 0.4in top/bottom. Generic API defaults remain unchanged.
The original DOM and CFML paginator are not modified. Native Excel PDF output
must still be reviewed for the actual form; browser and Excel font metrics differ.
Excel's [automatic paper mapping](https://learn.microsoft.com/en-us/office/vba/api/excel.application.mappapersize)
can additionally resize and reposition output even when the worksheet is A4.
This application preference is outside the browser exporter's control; the
exporter does not change Excel or printer preferences.

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

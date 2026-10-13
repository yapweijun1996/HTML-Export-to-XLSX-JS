# Known Issues

Code review of `main` at `f7556fd`, performed 2026-10-13. Line references in the
details of fixed issues describe the code before the fix. Open issues reference the
current source.

Evidence levels:

- **Verified** — reproduced in Chrome through the public `exportXlsx` API, in Node
  against `src/values.js`, or in a real exported workbook.
- **Code** — found by reading the source; not yet reproduced.

Regression tests for fixed issues are in
[tests/unit/values.test.js](tests/unit/values.test.js) and
[tests/browser/known-issues.test.js](tests/browser/known-issues.test.js).

## Summary

| ID | Severity | Area | Issue | Evidence | Status |
|---|---|---|---|---|---|
| BUG-01 | High | Values | Currency amounts always display the currency's default decimals | Verified | Fixed |
| BUG-02 | Medium | Values | Words that are also ISO codes (Top, Pen, Try, All, Cup) become currency numbers | Verified | Fixed for mixed case |
| BUG-03 | Medium | Values | Typed numbers lose their source decimals and grouping | Verified | Fixed |
| BUG-04 | Medium | Styles | Inherited or legacy alignment keywords are ignored | Verified | Fixed |
| BUG-05 | Medium | Images | Images outside a td/th are dropped without a warning | Verified | Fixed |
| BUG-06 | Medium | Images | Alignment mode places images at the cell's top edge, over preceding text | Verified | Fixed |
| BUG-07 | Medium | Images | An image that overflows its cell rejects the whole export in the default mode | Verified | Fixed |
| BUG-08 | Low | Workbook | Rows below the exported grid are 0.75pt (1px) tall | Verified | Open |
| BUG-09 | Medium | Styles | Semi-transparent colors are written with an alpha byte that Excel ignores | Verified (output) | Open |
| BUG-10 | Low | API | Failures from generic `[data-xlsx-button]` buttons are silent | Code | Open |
| BUG-11 | Low | Images | Off-screen lazy images wait twice the timeout and are then dropped | Code | Open |
| BUG-12 | Low | Images | Image load errors are reported as Event objects | Code | Open |
| BUG-13 | Low | Workbook | Row heights are not capped at Excel's 409pt maximum | Code | Open |
| BUG-14 | Low | Layout | The GRID_REFINED retry widens the sheet | Code | Open |
| BUG-15 | Low | Workbook | A rowScale smaller than fontScale has no effect | Code | Open |
| BUG-16 | Low | Styles | The inline-style border fallback revives borders hidden by `!important` rules | Code | Open |
| BUG-17 | Low | Workbook | Outer table borders overwrite inner cell borders on the same edge | Code | Open |
| BUG-18 | Low | Layout | Shared-border propagation is quadratic in the cell count | Code | Open |
| BUG-19 | Low | Images | Raster canvases have no size cap | Code | Open |
| BUG-20 | Low | Config | sheetName is not validated | Code | Open |

## Details

### BUG-01 Currency amounts always display the currency's default decimals

- **Symptom:** `$ 12.3456` with document currency HKD exports value `12.3456` but displays
  `HKD 12.35`. `$ 100` displays `HKD 100.00`. The displayed unit price differs from the
  printform although the stored value is correct.
- **Cause:** `currencyFormat` ([src/values.js](src/values.js) lines 97-108) takes the
  decimal count from `Intl` for the currency only, never from the source text.
- **Fix:** pass the source fraction length from `detectCurrency`. Use the larger of the
  source length and the currency default, or the source length alone.
- **Status:** fixed. Amounts use the source length alone, so the workbook matches the
  printform (`$ 100` displays `HKD 100`). A typed currency value without numeric display
  text keeps the currency default.

### BUG-02 Words that are also ISO codes become currency numbers

- **Symptom:** an untyped cell `Top 10`, `pen 2`, `Try 3`, `all 5`, `Cup 1` or `10 PEN`
  becomes a number with formats such as `"TOP" #,##0.00`. Excel then shows `TOP 10.00`.
- **Cause:** the prefix and suffix match `/^([A-Za-z]{3})\s*/` is case-insensitive
  ([src/values.js](src/values.js) lines 139-145). TOP, PEN, TRY, ALL, CUP, MAD and BOB are
  valid ISO 4217 codes.
- **Fix:** accept only uppercase codes, and preferably only the configured document or
  cell currency. Other codes would require `data-excel-type="currency"`.
- **Status:** fixed for mixed case. Codes must now be uppercase. Uppercase words such as
  `TOP 10` still match, because restricting codes to the document currency would break the
  documented support for explicit foreign-currency amounts; mark such cells
  `data-excel-type="text"`.

### BUG-03 Typed numbers lose their source decimals and grouping

- **Symptom:** a cell typed as a number, for example a Globe3 quantity `2.00` through
  the adapter's cellType, exports `2` with the General format. `1,000` displays as `1000`.
- **Cause:** the numeric branch of `parseCellValue` returns no `numFmt` for
  `type === 'number'` ([src/values.js](src/values.js) lines 188-193).
- **Fix:** derive a number format from the source text, for example `#,##0.00` when the
  source has grouping and two decimals.
- **Status:** fixed. Typed numbers keep grouping, decimals and accounting parentheses.
  Plain integers keep the General format.

### BUG-04 Inherited or legacy alignment keywords are ignored

- **Symptom:** `<tr align="center"><td>X</td></tr>` exports with no horizontal alignment
  (left), although the browser centers it. The computed `text-align` is `-webkit-center`
  in Chrome (`-moz-center` in Firefox). `<td><center>X</center></td>` and
  `<td><div align="center">` are also left-aligned, because only the cell's own style is read.
- **Cause:** `horizontalMap` in `extractAlignment`
  ([src/styles.js](src/styles.js) lines 663-684) maps only standard keywords. The `align`
  attribute fallback reads only the cell itself.
- **Fix:** map the `-webkit-` and `-moz-` prefixed values, and `match-parent`. Read
  alignment from the uniform text-owning block, as is already done for fonts.
- **Status:** fixed. Prefixed keywords are mapped, and alignment is read from the visible
  text; verified in Chrome, Edge and Firefox. `match-parent` was not needed because
  browsers resolve it in the computed value.

### BUG-05 Images outside a td/th are dropped without a warning

- **Symptom:** a logo inside a `<div>` before the first table is not exported, and
  `warnings` is empty.
- **Cause:** image ownership requires `img.closest('td, th')` and silently `continue`s
  otherwise ([src/layout.js](src/layout.js) lines 188-189).
- **Fix:** anchor such images in precise mode against the root grid, or at least emit a
  warning such as `IMAGE_OUTSIDE_CELL`.
- **Status:** fixed. Such images are placed at their rendered position in both placement
  modes.

### BUG-06 Alignment mode places images at the cell's top edge, over preceding text

- **Symptom:** `<td>Signature:<br>line two<br>line three<br><img></td>` places the image
  at row offset 0, covering the text above it. This is the default configuration. The
  Globe3 adapter uses precise mode and is not affected.
- **Cause:** `placeImagesInline` starts `currentTop` at 0 and ignores the image's measured
  position ([src/images.js](src/images.js) line 354).
- **Fix:** start the flow at the first image's measured top relative to the cell, or use
  precise placement when the cell also contains text.
- **Status:** fixed. Flows start at the images' rendered top and move up only as far as
  the cell requires. Horizontal alignment is unchanged. A single centered image is no
  longer centered vertically when its rendered position is known.

### BUG-07 An image that overflows its cell rejects the whole export in the default mode

- **Symptom:** with `table-layout: fixed`, a 400px image in a 300px cell, which renders
  with visible overflow, makes `exportXlsx` reject with *Image alignment exceeds its cell*.
  The README describes this rejection only for a forced flow.
- **Cause:** `addAlignedImage` throws ([src/images.js](src/images.js) lines 298-300). The
  precise fallback in `placeImages` is never reached.
- **Fix:** return `false` so the image falls back to precise placement, and emit a
  warning. Throw only for an explicitly forced flow.
- **Status:** fixed. All placements of a cell are checked before any is written. An
  unforced flow falls back to precise placement with `IMAGE_ALIGNMENT_FALLBACK`. A cell
  attribute or an explicit `imageAlignment` still rejects.

### BUG-08 Rows below the exported grid are 0.75pt (1px) tall

- **Symptom:** in the real Globe3 export, `sheetFormatPr` contains
  `defaultRowHeight="0.75" customHeight="1"`. Rows a user adds below the printform are
  practically invisible.
- **Cause:** the default row height is derived from `MIN_ROW_HEIGHT_PX`
  ([src/workbook.js](src/workbook.js) lines 281-287). Every grid row already receives an
  explicit height.
- **Fix:** leave Excel's default row height (15pt) unchanged.

### BUG-09 Semi-transparent colors are written with an alpha byte that Excel ignores

- **Symptom:** `background: rgba(0,0,0,0.05)` is written as `argb: 0D000000`. Excel does
  not apply the alpha channel of cell colors, so a faint zebra stripe is expected to
  render as solid black. The written value is verified; Excel's rendering has not been
  tested. Text colors take the same path.
- **Cause:** `colorToARGB` keeps the alpha byte ([src/styles.js](src/styles.js)
  lines 145-156).
- **Fix:** composite the color over the resolved background (or white), and write `FF`
  as the alpha.

### BUG-10 Failures from generic buttons are silent

- **Cause:** `bindButton` only dispatches `xlsx:error`
  ([src/export.js](src/export.js) lines 132-134). Buttons bound automatically in
  [src/browser.js](src/browser.js) lines 7-9 have no listener, so a click that fails
  produces no message and no console output. The Globe3 adapter adds its own listener.
- **Fix:** log to the console when no `xlsx:error` listener calls `preventDefault`.

### BUG-11 Off-screen lazy images wait twice the timeout and are then dropped

- **Cause:** an `<img loading="lazy">` outside the viewport never loads. The preload in
  `exportXlsx` and `convertImageJob` each wait `imageTimeoutMs` (10s) through
  `ensureImageLoaded` ([src/images.js](src/images.js) lines 54-77). With a concurrency
  of 4, many lazy images can delay an export by minutes, and the images are then omitted.
- **Fix:** skip the second wait for an image whose preload already failed, and document
  the lazy-loading limitation. Decoding through `img.decode()` does not force a lazy load.

### BUG-12 Image load errors are reported as Event objects

- **Cause:** `onError` rejects with the DOM `Event`
  ([src/images.js](src/images.js) lines 70-73). `warnLog` keeps it unchanged because it is
  not an `Error`, and `IMAGE_SKIPPED` later reports `[object Event]`.
- **Fix:** reject with `new Error('Image failed to load')`.

### BUG-13 Row heights are not capped at Excel's 409pt maximum

- **Cause:** `applyRowHeights` and `fitMultilineText`
  ([src/workbook.js](src/workbook.js) lines 14-27 and 245) write any height. One HTML
  cell taller than 545px produces a row above 409pt. How Excel treats such a row has not
  been tested, but content beyond 409pt cannot be shown.
- **Fix:** split rows taller than 409pt into several rows inside the same merge.

### BUG-14 The GRID_REFINED retry widens the sheet

- **Cause:** with `preserveEdges`, edges are not clustered
  ([src/dom.js](src/dom.js) line 332). Every sub-pixel sliver column is then raised to
  `minColumnWidthPx` (10px by default) by `pxToColumnWidth`
  ([src/layout-metrics.js](src/layout-metrics.js) lines 94-102). The retried sheet is
  wider than the HTML, and precise image positions drift.
- **Fix:** do not apply the minimum width to columns created by the retry, or merge
  sub-pixel slivers before writing.

### BUG-15 A rowScale smaller than fontScale has no effect

- **Cause:** row height uses `max(ROW_SCALE, FONT_SCALE)`
  ([src/workbook.js](src/workbook.js) lines 18-22). With the default fontScale of 0.9,
  `rowScale: 0.8` is ignored. The README describes rowScale as a row size scale.
- **Fix:** document the floor, or use rowScale alone.

### BUG-16 The inline-style border fallback revives borders hidden by `!important` rules

- **Cause:** when the computed border is `none`, `extractBorders` reads `element.style`
  and the `style` attribute ([src/styles.js](src/styles.js) lines 296-348). A stylesheet
  rule `border: none !important` that hides an inline border is therefore ignored, and
  the export shows a border the browser does not draw.
- **Fix:** trust the computed style, and remove the inline fallback.

### BUG-17 Outer table borders overwrite inner cell borders on the same edge

- **Cause:** `applyDeferredBorders` always replaces the existing side
  ([src/workbook.js](src/workbook.js) lines 190-195), unlike `mergeBorderStyle`. A thin
  table border replaces a thicker or colored cell border on the shared edge.
- **Fix:** keep the wider border, matching the CSS border-collapse rules.

### BUG-18 Shared-border propagation is quadratic in the cell count

- **Cause:** `propagateSharedBorders` scans every cell for every bordered side
  ([src/layout.js](src/layout.js) lines 247-321). Several thousand cells mean tens of
  millions of comparisons on the main thread.
- **Fix:** index the cells by edge coordinate.

### BUG-19 Raster canvases have no size cap

- **Cause:** `convertImageJob` creates a canvas of `width × height × imageScale` (3 by
  default) with no limit ([src/images.js](src/images.js) lines 132-139). Large images can
  exceed Safari's 16.7-megapixel canvas limit and are dropped. Embedded PNGs also grow
  the workbook considerably.
- **Fix:** cap the canvas area and lower the scale for large images.

### BUG-20 sheetName is not validated

- **Cause:** `resolveOptions` does not check `sheetName`, so a non-string value fails
  late in `sanitizeSheetName` ([src/workbook.js](src/workbook.js) line 384) with
  `replace is not a function`. Names that begin or end with an apostrophe, and the
  reserved name `History`, are also not handled.
- **Fix:** validate the option in `resolveOptions`.

## Cleanup

These helpers are exported but not called by production code:
`interpolateCoordinate`, `buildWorksheetEdgeMaps` (and `pxToColumnWidthForImagePositioning`
and `rowHeightToPixels`, which only it uses) and `resolveSize` in
[src/layout-metrics.js](src/layout-metrics.js), and `clampOffset` in
[src/images.js](src/images.js).

## Globe3 notes

- The deployed `folder_javascript/export_to_xlsx.js` predates `f7556fd`.
  Re-exported printforms keep the old behaviour until it is redeployed.
- Printforms with a hardcoded label such as `Total Amount HKD:` now show the code twice,
  because the amount format also contains `HKD`.
- The sample export `Globe3 ERP_2026-10-13T07-04-36-381Z.xlsx` has a print scale of 92%,
  and the limit is page height: its columns fit the A4 printable width at 100%. Its
  margins (0.5in sides) are the generic defaults, not the adapter's 0.25in. Because the
  older deployed build produced it, it is not confirmed which options path ran.

## Recently fixed

Fixed in `f7556fd`: edge clusters leaving sliver rows, an undeclared `scrollY` in
page-break mapping, trailing `<br>` line breaks clipping centered text, and the Globe3
adapter showing `$` instead of the document currency code.

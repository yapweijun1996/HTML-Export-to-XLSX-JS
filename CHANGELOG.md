# Changelog

## Unreleased

- Omit excluded descendants and hidden text/images consistently, preserving visible
  content, explicit line breaks and the source DOM.
- Remove fixed image offsets and map alignment/precise anchors through merged-cell
  column/row geometry; honor explicit alignment options and preserve image spacing.
- Collect images beside nested tables, omit zero-size images and rasterize CORS
  fallback from the selected responsive source to PNG. Reject overflowing forced flows.
- Reject fractional grouping separators and unsafe decimal precision instead of
  silently changing numeric values. Untyped unsafe amounts remain text.
- Honor type, value and currency metadata for nested-table text within its owning
  cell, combining contiguous formatting fragments without duplicating declared
  values or merging across child tables and nested declarations.
- Reject colliding text fragments instead of silently omitting them.
- Preserve filtered text transformations and whitespace, visible font styles,
  background precedence, root-table decoration and RTL alignment.
- Parse RMB aliases, suffix currency symbols, signed/accounting currency amounts
  and localized minus signs while preserving numeric precision safeguards.
- Include warning codes and diagnostic context; reject a throwing warning observer
  without hanging after font timeout, and restore bound buttons.
- Cache repeated visibility, font and background reads within the settled DOM
  snapshot to bound the extra work for large tables.
- Add unit and browser regressions with real XLSX read-back, pixel checks and CORS
  responses against both bundles in Chrome, Edge and Firefox.

## 0.1.0 — Preview

- Split the original exporter into configuration, value, DOM, geometry, style,
  image, workbook and export modules.
- Add a generic HtmlToXlsx API, typed cells, explicit currencies and locale-aware
  numeric parsing using Intl.
- Keep Globe3 compatibility in a separate adapter.
- Bound image waits/concurrency and expose warnings.
- Reject overlapping meaningful cells instead of silently omitting values.
- Support late script loading, replaceable button bindings and concurrent options.
- Generate readable/minified bundles with version and SHA256 metadata.
- Add real XLSX read-back, original-layout comparisons and three-browser CI.

ExcelJS remains external. The preview prioritizes editable data and approximate
HTML layout; identical Excel and browser PDF pagination is not guaranteed.

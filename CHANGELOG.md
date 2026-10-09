# Changelog

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

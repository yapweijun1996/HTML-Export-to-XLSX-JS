# Changelog

## Unreleased

- Add preserveEmailAddresses, enabled by the Globe3 profile, to fit complete email
  addresses and optional contact notes on one line while retaining explicit breaks.
- Write cell and enclosing-table borders only on measured perimeters after merging.
  Detach shared merged styles and remove neighbor/padding border propagation to
  prevent stray border segments at adjacent regions and manual page boundaries.
- Add an optional cellType callback for legacy markup, with HTML type metadata
  taking precedence. Classify safe numeric Globe3 item and Summary quantities by
  matching their column boundaries to quantity headers, including section headers
  inside legacy item/numbered row tables, without changing the source DOM.
  Recognize the print engine's _processed row IDs for Summary and Optional rows.
- Write meaningful content before blank regions, compare source values for rich
  text and check entire merged ranges. Refine colliding grids once with every
  measured edge, retaining rejection for unresolved content collisions.
- Bound coordinate clusters from their first edge and select the nearest retained
  edge for cells and text fragments, avoiding grid drift from close-edge chains.
- Log Globe3 export exception messages before the compatibility alert.
- Correct OOXML column-width conversion and cumulative grid rounding. Use physical
  EMUs for both drawing markers and an explicit print area for the exported grid.
- Add a Globe3 print profile retaining original text/image sizes and positions,
  visible single-line labels, mixed inline fonts and the document currency code.
  Keep numeric values editable and generic API configuration isolated.
- Align the Globe3 print origin and margins and omit print background fills.
  Inset pictures at manual page boundaries to prevent preceding-page residue.
- Set the Globe3 workbook Normal style to Arial 9, preserving explicit CSS cell
  fonts and avoiding unintended horizontal print pages. Reserve Excel line height
  for explicit multiline text so merged address cells do not clip their last line.

- Preserve uniform visible inline font families, colors, italics, underlines and
  strike-through decorations, including ancestor decorations and legacy font tags.
- Keep visibly single-line bold or underlined captions on one line without changing
  the grid, preserving explicitly multiline and naturally wrapped text.
- Fit formatted numeric values to their existing cells without changing values or
  column widths, preventing currency-code expansion from displaying hash marks.
- Recognize current Globe3 zero-height page-break markers. Preserve marked print
  pages with calculated percentage scaling and an explicit Fit To opt-out.
- Use schema-compatible two-cell image anchors with fixed-size behavior, retaining
  existing start positions and physical image dimensions after serialization.
- Reuse byte-identical image media within a workbook without combining placements
  or different raster content. Add raw XLSX drawing and media regression checks.
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
- Display amounts with the decimal places of their text and typed numbers with
  their grouping and decimals. Keep mixed-case words such as Top 10 as text.
- Follow rendered horizontal alignment, including legacy row align attributes and
  inner center/div align blocks. Place images outside table cells at their rendered
  position, start aligned image flows at the images' rendered top, and place an
  unforced flow that overflows its cell at the rendered position with a warning.
- Log failures of automatically bound buttons to the console. Wait at most once for
  an image that never loads, report image load errors as messages, and keep rasters
  within common browser canvas limits.
- Split HTML rows taller than Excel's 409pt limit into merged rows and cap reserved
  multiline height. Keep Excel's default height for rows below the grid, and warn
  with ROW_SCALE_FLOOR when an explicit rowScale is below fontScale.
- Blend translucent colors over their backdrop, read borders from the computed style
  only, keep refined-grid columns at their measured width, and validate sheetName.
- Remove unused layout, image and value helpers.
- Record the code review findings and their status in KNOWN_ISSUES.md.

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

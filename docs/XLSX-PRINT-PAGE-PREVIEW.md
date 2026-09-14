# Read-only XLSX print-page preview (96 DPI)

`compileNativeSheetPrintPagePreviewV1` turns already-qualified saved print areas
and authored page setup into isolated page rectangles at 96 CSS pixels per inch.
It reuses the existing print-area and selected-range planners. It does not add a
printer, invent missing setup, or claim Excel or LibreOffice raster parity.

Grid/cache preview remains the default surface. Print-page preview is additive
and fail-closed: when it is unavailable, consumers keep the grid.

## Public API

```ts
import { compileNativeSheetPrintPagePreviewV1 } from '@injoffice/sheets'

const preview = compileNativeSheetPrintPagePreviewV1(geometries, objects)
if (preview.status !== 'available') {
  // preview.reason explains the refusal. Keep grid preview.
} else {
  for (const page of preview.pages) {
    // Isolated capture size: page.width_css_px × page.height_css_px at 96 DPI.
    // Map viewport-local paint: x * page.scale + page.translate_x_emu, then / 9525.
  }
}
```

Compile one source-qualified geometry per saved print-area rectangle, in source
order, using exact Normal-font geometry or the stored-row approximation policy.
Do not pass an explicit host page policy. Optional `{ repeat_print_titles: true }`
repeats already-qualified saved titles. Optional `paint_plans` attach already
compiled native cell-paint plans; this API never shapes glyphs itself.

Each available page records EMU and CSS-pixel paper, content clip, source clip,
scale, and translation. Letter at 100% is 816×1056 CSS pixels. A4 sizes are the
exact millimetre conversions and may be fractional; they are not rounded to
invent a CSS integer paper.

## Qualification gates

The preview is available only when all of the following already exist:

- Source page settings `status: 'available'` joined to the worksheet part, with
  explicit Letter or A4 paper, orientation, four body margins, and a percentage
  scale from 10 through 400.
- A saved print area or print-area set `status: 'available'` joined to the same
  part. Used-range and A1 fallbacks are refused.
- Compiled geometry whose viewport matches that saved area (and saved titles
  when repetition is requested).

The following refuse with an explicit `reason` and empty `pages`:

- Missing `page_settings`.
- Producer `unavailable` settings, including printer relationships, device DPI,
  custom paper, headers/footers, print options, and manual breaks.
- Authored `fit_to_page`. This lane does not invent a fit scale and is not
  Excel fit-to-page qualification.
- Missing, unsupported, or unjoined saved print areas.
- Planner refusals such as merged cells crossing a page, a row/column larger
  than one page, or the 100-page budget.

Malformed caller input (uncompiled geometry, stale package hash, extra option
keys, paint plans that do not join geometry) still throws. Those are not source
refusals.

## Original workbook limits

Read-only originals under the local fidelity corpus are evidence only and are
never committed. Their current authored print markup is outside this profile:

- `TableStyleTest.xlsx` stores A4 portrait plus printer-dependent
  `horizontalDpi` / `verticalDpi` and a printer relationship. Source page
  settings stay unavailable; print-page preview refuses and the grid remains.
- `chart_hyperlink.xlsx` has page margins and no `pageSetup`. Paper and scale
  are not invented.
- `different-column-width-excel2010.xlsx` includes headers/footers and no
  supported `pageSetup`. Headers are not drawn here.

LibreOffice can emit print pages for those files by applying application
defaults. InjOffice does not. Synthetic page-setup fixtures in unit tests prove
the available Letter/A4 percentage path and the refusal reasons.

## What this is not

- Not Excel printer calibration or device-resolution fidelity.
- Not fit-to-page qualification.
- Not LibreOffice/Excel PDF pixel parity.
- Not a replacement for the grid/cache preview or cached-chart images.
- Not a host paper/margin/scale chooser. Use the existing selected-range page
  preview when an explicit host policy is required.

Paint, wrapping, rotation, headers/footers, and drawing rasterization remain
the host’s responsibility. When already-compiled cell-paint plans are supplied,
they stay in viewport-local EMU; each page carries the clip and 96 DPI map
needed for isolated capture. This module does not run HarfBuzz or emit glyph
commands itself.

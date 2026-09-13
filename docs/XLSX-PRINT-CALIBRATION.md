# XLSX print preview and calibration evidence

The selected-range page APIs produce read-only, approximate pages. Source-qualified
settings, successful browser output and Excel-calibrated pagination are separate
claims. This document defines the evidence needed to compare those pages with an
independent Excel export; it does not declare print fidelity complete.

## Current page-settings contract

The producer in `go/xlsxpatch/native_page_preview.go` accepts explicit Letter or A4
paper, portrait or landscape orientation, all six page-margin attributes, and a
percentage scale from 10 through 400. Only the four body margins are projected.
Optional page order is `downThenOver` or `overThenDown`; omission retains
down-then-over ordering in the planner.

An explicitly active `sheetPr/pageSetUpPr@fitToPage` instead requires both
`fitToWidth` and `fitToHeight`, each from 0 through 100, with at least one positive
dimension. In this bounded profile, zero leaves that dimension unconstrained.
Scale is ignored for fit; an absent scale is represented as 100 for compatibility.
The activation element and its enclosing properties must meet the producer's
exact structural checks. Inactive or additional setup properties are not silently
discarded.

Printer relationships, device resolutions, printer defaults, custom paper sizes,
manual breaks, print options and headers/footers are outside this source profile.
Their presence can make source page settings unavailable. A separately selected
explicit host policy can still supply supported geometry, with host provenance;
it does not resolve the omitted printer or workbook behavior.

The planner in `packages/sheets/src/nativeSheetPagePreviewV1.ts` keeps whole source
rows and columns together, refuses merges crossing page boundaries, and limits a
plan to 100 pages. Fit chooses the greatest whole percentage from 100 down to 10
that satisfies its page targets, merge constraints and page budget. It never
enlarges content. Saved area selection and repeated-title selection are separate
source-bound steps. Multiple areas are fitted independently and have a combined
100-page budget. The effective scale is recorded on each returned page.

These choices are explicit host approximations. They are not an implementation
of Excel's printer-dependent fit or pagination algorithm. Microsoft's
[page setup schema documentation](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.pagesetup?view=openxml-3.0.1)
describes device resolution, printer settings and scale/fit precedence;
[Excel scaling guidance](https://support.microsoft.com/en-us/excel/scale-a-worksheet)
also identifies printer paper-size differences as a source of output changes.

## Independent reference record

Keep original workbooks, reference exports, fonts, screenshots and benchmark
scripts in the local evidence corpus, outside git and CI. Preserve the original
bytes. Record these fields for each comparison:

| Record | Required evidence |
| --- | --- |
| Source identity | URL or origin, revision if available, byte count and SHA-256; checksum again after reference capture |
| Application | Excel version/build, operating system, export method, printer/driver and paper configuration where applicable |
| Selection | Worksheet identity and visibility, exact printed range or area order, titles, hidden bands, page-order setting, and whether export used a selection, active sheets or the workbook |
| Settings | Authored page setup/margins and fit activation; any explicit reference-side overrides recorded separately |
| Fonts | Source font descriptors, installed font versions/hashes and actual PDF font inventory; disclose substitution or unresolved font use |
| Calculation | Calculation/link-update policy and whether cached values changed; do not infer cache freshness from an export |
| Reference | PDF hash, page count, per-page size/rotation and page-to-sheet/area mapping |
| Candidate | Exact commit/tree, source hash, public API/options, geometry/font identity, plan JSON, warnings and screenshots |

A PDF that already exists may support a limited comparison. Its presence does
not supply missing export settings, establish the same range selection, or make
an unrelated fit/title/area feature independently qualified. A reference from a
repaired workbook must not be joined to the original workbook's hash.

## Minimum calibration matrix

Each row needs an unchanged source and matching independent export with the
record above. Several explicitly identified worksheets may share one workbook.
This is an initial matrix, not an exhaustive fidelity suite.

| Case | What the comparison must isolate |
| --- | --- |
| Explicit percentage baseline | Letter and A4, both orientations, asymmetric margins, 100% plus a non-100% scale; known row/column dimensions |
| Fit activation | Width-only, height-only and two-dimensional targets; source scale differs from effective fit; enough content to cross page boundaries |
| Saved single area | Non-A1 rectangle, including blank cells beyond used dimensions; exact selected bounds |
| Saved multiple areas | Disjoint areas of unequal sizes, source order and independent fit; page-to-area mapping |
| Repeated titles | Leading and outside-body row/column titles, corner cells, repeated drawings and merges at region boundaries |
| Pagination boundaries | Rows/columns near a page edge, hidden bands, both page orders; count and boundary positions separately |
| Unsupported settings | Manual breaks, centering/gridlines/headings, printer defaults or device settings; verify explicit refusal or declared host approximation |

Generated unit fixtures can prove protocol joins, bounds, refusal behavior and
deterministic planning. They cannot independently establish Excel's fit scale or
page layout. Additional references must be supplied through the agreed local
evidence workflow; do not automate Office or repair sources to fill this matrix.

## Measurement and acceptance

First verify source identity, selection, settings and reference provenance. Then
compare opening/refusal, printed cells and objects, omissions, page count,
page-to-area order, row/column boundaries, text/object displacement and raster
differences separately. Use the same physical page size and raster DPI; do not
resize outputs to conceal dimension errors. A matching count or low pixel delta
alone is not a layout pass.

An implementation improvement still requires usable public API/demo behavior,
positive and negative checks, relevant WASM/browser proof, independent review,
green required CI and a merged tree equal to the tested tree. A measured
calibration claim additionally needs matching independent references for its
specific feature subset. Preserve all remaining limitations in the report.

# Local spreadsheet authoring

`SpreadsheetEditor` is an independent XLSX editor using the existing `OfficeEditorProps` contract. Its own native worker extracts source models and applies revision-guarded transactions. Only a successful apply **and** extraction emits changed bytes. The owner can retain an editor in a hidden tab: keyboard and clipboard handlers are scoped to the focused editor, not the document.

- Click or use arrows to select; Shift selects a rectangle. Enter/F2 or typing opens direct cell input. Enter/Tab applies and advances; Escape cancels. The formula bar also edits the active cell. Pending input synchronously raises `onDraftChange`; native operations raise `onBusyChange`.
- Copy/paste uses quoted TSV and preserves embedded tabs/newlines. Paste, clear, and formatting are bounded to native limits; the whole transaction is atomic. Text beginning with an apostrophe is literal. Identifier strings with leading zeros and numbers exceeding 15 digits remain strings. Editing/copying existing numeric or formula-looking strings preserves their literal type.
- Formatting includes fonts, size, bold/italic, colors, fill, horizontal/vertical alignment, wrap, and qualified numeric/date formats. Pure number/date display qualification is shared with native sheet paint; it does not pull the filesystem/HarfBuzz runtime into the browser.
- Merge/unmerge obey native content preservation. Merged content is displayed but is read-only until unmerged. Non-anchor content cannot be discarded by merging. Dimensions apply to selected rows/columns via the native adapter.
- Undo/redo retain immutable package/model snapshots (20 entries, 128 MiB per stack). Undo is local session history, not a cloud service.

Pending cell input can be checkpointed through `onRecoveryDraftChange` and restored with `initialRecoveryDraft`. The versioned JSON contains only sheet ID, coordinate, value (maximum 32,767 characters), and exact package revision. Restoration rejects stale packages, unknown keys, out-of-bounds or read-only targets, and merged cells. It never applies text automatically. `registerCommit` gives the host Save flow an asynchronous commit callback; it refuses composition/in-flight operations, returns false after native failure, and clears draft/busy guards synchronously after success.

## Explicit limits

Formula edits and other committed changes recalculate through the existing Univer OSS engine in a dedicated local worker; Recalculate also updates an opened workbook. The bounded scalar profile supports arithmetic/comparisons, explicit local/cross-sheet references, SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, IF, IFERROR, AND, OR, NOT, ABS, ROUND, ROUNDUP, ROUNDDOWN, INT, MOD, POWER and SQRT. A dependency pass detects circular references and propagates unresolved dependencies even through IFERROR and unused IF branches. Source formula caches are never calculation inputs. Supported results/errors persist through a separate native cache transaction requiring the exact source revision and complete matching formula set. Formula XML is preserved; unsupported/circular caches are removed. Snapshot undo includes both edits and calculation results.

Opened caches are labeled stored/unverified until locally calculated. The bounded profile rejects named/structured/external/whole-row/whole-column references, unsupported functions and ambiguous error-like string results. Shared/array/data-table formulas prevent workbook calculation until group expansion is supported. ISO date storage is not converted to an invented serial; dependent formulas are unresolved. Read-only source worksheets cannot receive cache writes. Limits: 100,000 stored cells, 10,000 formulas, 1,000,000 reference visits and a 30-second worker deadline. Calculation failure preserves successful document edits with explicitly unverified caches, never marks them current.

Worksheet Add/Rename/Delete controls use native lifecycle commands with explicit delete confirmation and snapshot undo. New sheets accept absent-cell edits anywhere in Excel bounds. Whole-row and whole-column insertion/deletion is available under Rows & columns. Selection determines the insertion/deletion interval; deletion requires explicit confirmation. Native transactions shift cells, dimensions, merged ranges, view anchors and qualified formula/defined-name references together; deleted direct references become #REF! and partial ranges shrink. Unknown graph/package features, implicit cell addresses, formula groups, split views, and dynamic/external/3-D/structured references remain atomic refusals. Filtering, chart editing, and formula autofill are not implemented. Rename updates explicit sheet qualifiers in formulas and defined names while preserving literal strings and caches. Dynamic/3-D/external/structured references, formula-bearing deletion, defined names during add/delete, protection/extensions, graph dependencies, and unqualified metadata inventories remain refusals; errors leave the package unchanged. Protected, unsupported, and merged-cell edits remain native refusals. There is no arbitrary OOXML fallback.

The authoring grid uses ordinary browser fonts/layout rather than native print geometry. It is not a print-fidelity preview. Existing hidden rows/columns remain hidden. The rendered window is 60 rows by 20 columns; arrows, page controls, and the address box navigate all Excel addresses. There is no full-workbook DOM allocation. Existing unsupported number formats may show their stored lexical value. Horizontal and vertical resize uses explicit dimensions in Cell size, not drag handles.

## Verification

With Node 22 and the native packages built:

```
node --test apps/desktop/tests/spreadsheet-authoring.test.cjs apps/desktop/tests/spreadsheet-bundle.test.cjs apps/desktop/tests/spreadsheet-calculation.test.cjs apps/desktop/tests/spreadsheet-structure.test.cjs
npx vitest run packages/formulas/src
npm run typecheck -w @injoffice/desktop
npm test -w @injoffice/sheets
```

The authoring test executes the distributed worker and Go WASM: paste, formulas, styles/display, dimensions, merge/unmerge, original snapshot preservation, mixed-batch refusal, and stale revision refusal. The browser bundle test rejects Node/HarfBuzz leakage. No GUI automation is used.

The calculation test executes the Vite-built worker without DOM or network, then the actual Go WASM cache writer and reopen path. It checks dependency updates, scalar errors, circular/unsupported results, stale revisions and original snapshot preservation. Native tests additionally enforce typed cache values, complete source coverage, duplicate rejection and separate transaction families. Host menu history registers via `registerHistory`; scoped keyboard history leaves input/browser undo intact while drafting.

Whole-row/column mutation verification includes strict and transitional OOXML, absolute/qualified references, named ranges, dimensions, merged ranges, deletion-generated errors, native stale guards and actual WASM mixed-transaction refusals. Existing row/column metadata and source cell contents are copied exactly except for explicit positional changes/deletion. All source formula caches are invalidated after structural changes and the editor runs local calculation. `onBusyChange` reports only loading/operations; pending text is reported solely through `onDraftChange`.


## Frozen panes

Freeze top row, first column or rows/columns before the active cell; Unfreeze restores normal scrolling. Commands preserve cells, formula caches and dimensions. Native `sheet.freeze` writes one qualified view transaction with exact readback; zero rows/columns unfreezes. V1/V2 sheet models now expose optional paired `frozen_rows` / `frozen_columns`. Existing unqualified/multiple/split views and a boundary across a merged cell are refused. The native print compiler still explicitly inventories `SHEET_VIEW_GEOMETRY`; desktop pinning does not claim native print fidelity.

The desktop pins at most 50 rows and 10 columns, keeps those cells present while navigating the bounded grid window, and measures frozen row heights after layout. Larger saved panes are preserved with a visible display-limit message and can be reduced or removed. Native pane limits remain Excel bounds. The real WASM test `spreadsheet-freeze.test.cjs` checks freeze, content-edit preservation, reopen, unfreeze and mixed transaction refusal.


## Stable record sorting

Sort range supports one selected key column, ascending/descending order and an optional header. It moves complete native cell records, preserving literal values, cell formatting and untouched columns. Equal keys keep their relative order. Empty keys remain last in either direction; numbers, case-folded strings and booleans use deterministic type ordering, without locale or network dependencies.

The native profile accepts literal records and qualified row-local formulas in existing rows (maximum 10,000 rows / 100,000 selected cells), excludes hidden rows and merged records, and refuses formula/date/error sort keys. Row-local formula references move with their record, including absolute markers; references outside the sorted rectangle remain fixed. Moved formula caches are cleared for local recalculation. Cross-record references, outside formulas/defined names referring into the sorted data, dynamic references and unqualified package features are explicit atomic refusals. Disjoint formulas and caches stay unchanged. This deliberately avoids treating stale formula caches as sort authority. Source snapshots support undo. `spreadsheet-sort.test.cjs` verifies stable record/style movement, descending order, unselected-column preservation, dependency refusal and source snapshot restoration through actual WASM.


## Borders

Borders offers all, outside, bottom and clear presets with thin/medium/thick/double/dashed/dotted lines and RGB color. Per-edge native `style.patch` properties (`border_left`, `border_right`, `border_top`, `border_bottom`) carry `{style,color}`; `none` removes a visible border and null restores that edge from the inherited base style. Unspecified edges and other style components remain unchanged. The native writer interns qualified border records, preserves existing style records and checks exact effective-style readback. Unsupported diagonal/theme/extension borders refuse modification. The browser uses CSS strokes; the native decoration projection remains the source for exact border geometry.

Color picker output is normalized to the protocol's canonical uppercase RGB form for font, fill and border colors. `spreadsheet-borders.test.cjs` verifies outer/all/clear presets, inherited/native readback, text/font preservation and undo source bytes through real WASM.

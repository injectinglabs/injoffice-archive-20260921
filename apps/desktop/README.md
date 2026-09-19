# InjOffice Desktop

Private workspace `@injoffice/desktop`. Not published to npm.

This directory currently holds the Node host adapters: opaque file ids, atomic
save, recent-file history, recovery journal, renderer URL policy, blank
DOCX/XLSX/PPTX/PDF seeds, the in-app update state machine, and the DOCX PDF
export host and PDF text-replace host (worker-backed; painters land later). They do not import
Electron. Electron is a desktop host dependency; renderer /src stays offline.
`electron/main.cjs` is tested with a mock Electron. The start page
is a React view with recents and create actions. The updates dialog talks to
the host update state machine over `window.injDesktop`. Preferences persist
fail-closed on this device; the command palette searches workspace actions.
`App.tsx` is the session shell (tabs, recents, recovery, close/update freeze);
editors are still mocked in tests. Formatting, hyperlink, table, paragraph,
and page-layout chrome land here without WASM. Selection offsets for caret
restore are in `document-range.ts`. Paragraph/run appearance follows the
OOXML style cascade (bold/italic toggles, cycle-safe). PNG/JPEG insert sizes
from file headers only. Inline PNG/JPEG preview is cached and bounded.
`OpenError` explains unsupported, encrypted, and extract failures and
replaces the start page when open/create fails closed. `.docm` is openable
zip OOXML (VBA is not executed); blank create stays `.docx`. Hidden apply
is an 80ms debounce that skips IME composition; OfficeEditor wiring follows.
CSV/TSV parse and encode keep fields as literal UTF-8 text. Word/character
counts are per paragraph so a keystroke does not recount the document. PDF
find-in-document is sequential with page and hit caps. PDF recovery drafts
are data, never commands; unknown fields fail closed. Worksheet add/delete
reasons are predicted before a destructive native mutation. Presentation
mode owns its own slide index; Escape exits. CSV/TSV sheet export uses
literal values or formula source; native import is a follow-up. Chart
preview geometry finite-normalizes values and caps series/categories. Slide
shape presets map to SVG ellipse/polygon/rect without a canvas. Slide
duplicate/delete/reorder and grouped preview coordinates live in
presentationCommands.ts. Align/distribute uses unrotated boxes and keeps
rotation; spacing keeps the outer objects fixed. SlideArrangePanel is the
checkbox UI for that selection. DocumentPreview is the flowing HTML paper
with per-run contenteditable; Apply is Ctrl/⌘+Enter, Cancel is Escape.
formatting.ts maps native run/paragraph appearance onto toolbar values.
document-authoring.ts is fail-closed native apply for split/join/replace/table/image.
Spreadsheet selection, clipboard copy/paste, virtualized visible-row windows, and versioned recovery drafts live in spreadsheetCommands.ts.
PresentationPlayer is the fullscreen slide dialog; Escape exits and Home/arrows use presentationMode.
PDF command apply/inspect, byte-snapshot history, and page-range parse live in pdf-commands.ts.
OfficeEditor hosts that preview with the ribbon, find/replace, and Apply for pending run text.
Local spreadsheet calculation projects source values into a worker host and writes a revision-checked cache.
PresentationEditor is the PPTX canvas, inspector, and Present control that launches PresentationPlayer.
PdfEditor is the PDF page and annotation workspace; it applies through pdf-commands and does not import OfficeEditor.
SpreadsheetCharts is the worksheet chart sidebar; empty selections disable insert and a valid range emits chart.insert.
The rest of the renderer, GitHub provider, and packaging workflows land
in follow-up changes. Preview builds never load an updater.

Support, unsigned vs signed builds, and engine lockstep: [docs/DESKTOP.md](../../docs/DESKTOP.md).

```bash
npm test -w @injoffice/desktop
```

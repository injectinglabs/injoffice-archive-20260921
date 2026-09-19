# InjOffice Desktop

Private workspace `@injoffice/desktop`. Not published to npm.

This directory currently holds the Node host adapters: opaque file ids, atomic
save, recent-file history, recovery journal, renderer URL policy, blank
DOCX/XLSX/PPTX/PDF seeds, the in-app update state machine, and the DOCX PDF
export host and PDF text-replace host (worker-backed; painters land later). They do not import
Electron. `electron/main.cjs` is tested with a mock Electron. The start page
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
reasons are predicted before a destructive native mutation.
The rest of the renderer,
the real Electron dependency, GitHub provider, and packaging workflows land
in follow-up changes. Preview builds never load an updater.

Support, unsigned vs signed builds, and engine lockstep: [docs/DESKTOP.md](../../docs/DESKTOP.md).

```bash
npm test -w @injoffice/desktop
```

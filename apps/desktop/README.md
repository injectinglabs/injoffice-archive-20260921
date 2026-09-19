# InjOffice Desktop

Private workspace `@injoffice/desktop`. Not published to npm.

This directory currently holds the Node host adapters: opaque file ids, atomic
save, recent-file history, recovery journal, renderer URL policy, blank
DOCX/XLSX/PPTX/PDF seeds, the in-app update state machine, and the DOCX PDF
export host and PDF text-replace host (worker-backed; painters land later). They do not import
Electron. `electron/main.cjs` is tested with a mock Electron; the real
Electron dependency, renderer, GitHub provider, and packaging workflows land
in follow-up changes. Preview builds never load an updater.

Support, unsigned vs signed builds, and engine lockstep: [docs/DESKTOP.md](../../docs/DESKTOP.md).

```bash
npm test -w @injoffice/desktop
```

# InjOffice Desktop

Private workspace `@injoffice/desktop`. Not published to npm.

This directory currently holds the Node host adapters: opaque file ids, atomic
save, recent-file history, recovery journal, renderer URL policy, blank
DOCX/XLSX/PPTX/PDF seeds, and the in-app update state machine. They do not
import Electron. The Electron shell, editors, GitHub provider, and packaging
workflows land in follow-up changes. Preview builds never load an updater.

Support, unsigned vs signed builds, and engine lockstep: [docs/DESKTOP.md](../../docs/DESKTOP.md).

```bash
npm test -w @injoffice/desktop
```

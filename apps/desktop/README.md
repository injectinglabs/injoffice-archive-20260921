# InjOffice Desktop

Private workspace `@injoffice/desktop`. Not published to npm.

This directory currently holds the Node host adapters: opaque file ids, atomic
save, recent-file history, recovery journal, and renderer URL policy. They do
not import Electron. The Electron shell, editors, and packaging workflows land
in follow-up changes.

Support, unsigned vs signed builds, and engine lockstep: [docs/DESKTOP.md](../../docs/DESKTOP.md).

```bash
npm test -w @injoffice/desktop
```

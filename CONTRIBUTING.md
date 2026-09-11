# Contributing to InjOffice

Thank you for improving InjOffice. Bug reports, focused fixes, tests, documentation, and interoperability fixtures are welcome.

## Before opening a change

- Search existing issues and pull requests.
- Keep changes scoped to one behavior or package when practical.
- Do not submit confidential documents, credentials, customer data, proprietary fonts, or copyrighted fixture files without redistribution permission.
- For fidelity bugs, prefer the smallest synthetic file that reproduces the OOXML structure.

For a security vulnerability, follow [SECURITY.md](SECURITY.md) instead of opening a public issue.

## Local checks

Use Node.js 22 or newer (CI uses Node 24):

```bash
npm ci
npm run typecheck
npm test
npm run build
npm run check:packages
```

Run checks for every affected Go module:

```bash
cd go/xlsxpatch
go test ./...
go vet ./...
```

The repository contains independent Go modules under `go/` (`xlsxpatch`, `docxpatch`, `pptxpatch`, `slidesqc`, `officecompat`, `collab`, `injoffice-server`), so repeat that command in each changed module. Tests should be deterministic and must not require private services or credentials.

## Rendering and browser qualification

PR and main CI retain builds, typechecks, unit/contract/security checks, package-consumer checks, and the installed-WASM browser smoke test. Full rendering qualification, exhaustive browser scenarios, and performance benchmarks are deliberately not merge gates.

For rendering or UI changes, run the relevant heavier checks locally before opening the PR. After `npm ci` and `npm run build`:

```bash
npm run build:renderer -w apps/playground -- --base=/injoffice-smoke/
npm run test:showcase-browser -- --built
node scripts/smoke-document-first-browser.mjs
node scripts/qualify-rendering-corpus.mjs
npm run qualify:office-performance
```

Browser checks require Chrome or Chromium (`CHROME_BIN` can select the executable); native qualification also requires Go. The complete built/development, responsive/scroll, documentation, rendering, and performance matrix is available on demand in GitHub Actions → **Manual rendering qualification** → **Run workflow**, selecting the branch to test. Attach relevant results to the PR; a green normal CI run does not mean the full rendering matrix ran.

External fidelity comparisons and Microsoft Office/reference exports remain local-only. The manual workflow uses only the repository's synthetic qualification corpus; do not add private documents, external benchmark downloads, or proprietary fonts to it.

## Design expectations

- Keep persisted and wire-facing specs plain JSON.
- Put editor-independent logic in pure functions and keep DOM/editor adapters thin.
- Treat Univer as an optional editor shell and native paint as preview. Neither is the OOXML file authority.
- Follow the native file API: Go `Extract*` → native JSON → TypeScript paint (preview) → mutation JSON → Go `Apply*`.
- Treat OOXML/PDF input as untrusted and bound allocations and recursion.
- Preserve unknown archive parts when patching an existing file. A writer must fail when it cannot prove that an operation is safe.
- Add a regression test for behavior changes and malformed-input tests for parsers.
- Do not add a required hosted backend. An optional in-repo `injoffice-server` and WASM runtime are in charter; they must remain optional, use documented interfaces, and stay free of consumer-specific backend, authentication, tenant, and credential code. Hosts may still inject their own services.
- Do not claim Microsoft Office parity. Native completion remains partial.

## Pull requests

Explain the user-visible behavior, compatibility impact, and verification performed. Call out changes to exported types, wire formats, archive output, or security boundaries. Coordinate version changes with maintainers and follow the [release checklist](docs/PUBLIC-RELEASE.md); do not overwrite already-published versions.

By participating, you agree to follow [CODE_OF_CONDUCT.md](CODE_OF_CONDUCT.md).

# Contributing to InjOffice

Thank you for improving InjOffice. Bug reports, focused fixes, tests, documentation, and interoperability fixtures are welcome once the public repository opens.

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

Explain the user-visible behavior, compatibility impact, and verification performed. Call out changes to exported types, wire formats, archive output, or security boundaries. Maintainers may ask for a changeset or version bump after the release process is finalized.

By participating, you agree to follow [CODE_OF_CONDUCT.md](CODE_OF_CONDUCT.md).

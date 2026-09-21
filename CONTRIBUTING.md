# Contributing to InjOffice

Thank you for improving InjOffice. Bug reports, focused fixes, tests, documentation, and interoperability fixtures are welcome.

## Before opening a change

- Search existing issues and pull requests.
- Keep changes scoped to one behavior or package when practical.
- Do not submit confidential documents, credentials, customer data, proprietary fonts, or copyrighted fixture files without redistribution permission.
- For fidelity bugs, prefer the smallest synthetic file that reproduces the OOXML structure. Save write-back tests follow [docs/SAVE-FIDELITY.md](docs/SAVE-FIDELITY.md).

For a security vulnerability, follow [SECURITY.md](SECURITY.md) instead of opening a public issue.

Before submitting source, stage the intended files and run `npm run check:secrets`
with Gitleaks 8.30.1 installed (`GITLEAKS_BIN` can select its executable). Required
CI uses the same version and verifies its download checksum. The scan checks
current tracked file contents, including bounded archive/encoding traversal;
it does not audit all Git history or guarantee that every secret format is detected.
Output is fully redacted. The only allowlist entry is an exact synthetic OOXML
font-obfuscation key in its fixture and historical evidence copy; inline
`gitleaks:allow` comments do not bypass the gate. Do not broaden the exception
for a real credential: remove it and follow the private reporting policy.

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

PR CI is deliberately short (a few minutes) and keeps only the fast gates as the required `CI required` status: tracked-source secret scan, build, typecheck, `check:docs-api`, `check:unicode13`, `check:typescript-version`, `test:dependency-integrity`, `test:declaration-specifiers`, the native Office production-E2E contract checks, `check:packages`, the workspace unit tests (sharded across the `ts-tests` matrix from `scripts/ci-test-shards.json`, plus `test:native-office-completion`), the Go modules, the WASM size ceilings and contracts, and the officecompat fuzz shards. Full rendering qualification, exhaustive browser scenarios, and performance benchmarks are not merge gates either.

The slower and audit-style checks currently do not run automatically after merge. They remain in the **Main audit** workflow (`.github/workflows/main-audit.yml`) and can be dispatched on demand via **Run workflow**. They were moved out of the PR job, not removed, so run them locally before opening a PR that touches the office pipeline, packaging, or the qualification harness. After `npm ci` and `npm run build`:

```bash
npm run test                       # every workspace test plus the root script tests, sequentially
npm run test:chrome-cdp-startup
npm run test:release-packages
npm run check:office-architecture
npm run check:office-provenance
npm run qualify:office-reproducibility
npm run check:native-office-production-e2e
npm run test:native-office-production-e2e
npm run check:native-office-completion  # needs full git history (baseline SHA 420424b)
npm run check:consumer
go -C go/docxpatch test -count=1 ./cmd/nativepreviewfixture
node --test scripts/showcase-smoke-server.test.mjs scripts/qualify-rendering-corpus.test.mjs
npm run build:renderer -w apps/playground -- --base=/injoffice-smoke/
node scripts/smoke-xlsx-source-style-browser.mjs --rich --skip-build   # XLSX read-only source grids in Chrome
```

To reproduce one PR test shard exactly as CI runs it, use `node scripts/run-ci-test-shard.mjs <shard>` (`--list` prints the shard names, `--check` verifies every workspace test script is assigned). The PPTX chart and graphic-frame browser qualification workflows (`pptx-*-qualification.yml`) likewise run on push to `main` and on demand rather than on pull requests.

For rendering or UI changes, also run the relevant heavier checks locally before opening the PR:

```bash
npm run build:renderer -w apps/playground -- --base=/injoffice-smoke/
npm run test:showcase-browser -- --built
node scripts/smoke-document-first-browser.mjs
node scripts/qualify-rendering-corpus.mjs
npm run qualify:office-performance
```

Browser checks require Chrome or Chromium (`CHROME_BIN` can select the executable); native qualification also requires Go. The complete built/development, responsive/scroll, documentation, rendering, and performance matrix is available on demand in GitHub Actions → **Manual rendering qualification** → **Run workflow**, selecting the branch to test. Attach relevant results to the PR; a green normal CI run does not mean the full rendering matrix ran.

External fidelity comparisons and Microsoft Office/reference exports remain local-only. The manual workflow uses only the repository's synthetic qualification corpus; do not add private documents, external benchmark downloads, or proprietary fonts to it.

## Desktop app

The optional Electron workspace `apps/desktop` is a local editor around the same native extract/apply engines. It is not a required hosted backend and is not published to npm. Desktop packaging, signed updates, and native WASM editor tests are not PR merge gates. See [docs/DESKTOP.md](docs/DESKTOP.md).

- Unsigned developer previews: GitHub Actions → **Desktop builds** (`desktop.yml`, `workflow_dispatch`) for mac-arm64, mac-x64, win-x64, and linux-x64. Those artifacts must not drive the public updater.
- Signed public drafts: tag `desktop-v*` and the **Desktop release draft** workflow (`desktop-release.yml`, environment `desktop-release`). Drafts are never auto-published.
- Keep the desktop WASM snapshot on current `main`; do not freeze a lagging snapshot as the advertised product.
- Host tests for `apps/desktop` belong in the `core` shard of `scripts/ci-test-shards.json`. Do not treat Electron, TipTap, or Univer as file authority.
- Do not commit `apps/desktop/release/` installer output.

## Design expectations

- Keep persisted and wire-facing specs plain JSON.
- Put editor-independent logic in pure functions and keep DOM/editor adapters thin.
- Treat Univer as an optional editor shell and native paint as preview. Neither is the OOXML file authority. Electron is a host shell only; it is not native authority either.
- Follow the native file API: Go `Extract*` → native JSON → TypeScript paint (preview) → mutation JSON → Go `Apply*`.
- Treat OOXML/PDF input as untrusted and bound allocations and recursion.
- Preserve unknown archive parts when patching an existing file. A writer must fail when it cannot prove that an operation is safe.
- Add a regression test for behavior changes and malformed-input tests for parsers.
- Do not add a required hosted backend. An optional in-repo `injoffice-server` and WASM runtime are in charter; they must remain optional, use documented interfaces, and stay free of consumer-specific backend, authentication, tenant, and credential code. Hosts may still inject their own services.
- Do not claim Microsoft Office parity. Native completion remains partial.

## Pull requests

Explain the user-visible behavior, compatibility impact, and verification performed. Call out changes to exported types, wire formats, archive output, or security boundaries. Coordinate version changes with maintainers and follow the [release checklist](docs/PUBLIC-RELEASE.md); do not overwrite already-published versions.

By participating, you agree to follow [CODE_OF_CONDUCT.md](CODE_OF_CONDUCT.md).

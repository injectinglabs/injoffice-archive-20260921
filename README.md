<p align="center">
  <img src="logo.svg" width="64" height="64" alt="Injecting" />
</p>

# InjOffice

InjOffice is a set of browser and server libraries for editing office artifacts. It adds plain-JSON models, pure transformation engines, provider-neutral agent change sets, optional editor integrations, and surgical OOXML patching. A hosted InjOffice service is not required: consumers can use the browser-local WASM engine, run offline, inject their own backend, or opt into the in-repo `injoffice-server`.

The repository is a monorepo with independently consumable TypeScript packages, private apps, and Go modules. The TypeScript packages target Node.js 22 or newer and modern bundlers. UI integrations use React and, where noted, Univer OSS as an optional editor shell. Univer is not the file authority. Native paint compilers are a preview mode.

> Release status (2026-09-08): the source is licensed under Apache-2.0; 25 of 26 npm packages are published at 0.1.0. `@injoffice/xlsx-wasm` is not yet published. See the [release checklist](docs/PUBLIC-RELEASE.md). The scoped v3 Native Office completion matrix is complete; broader format coverage remains partial and this is not Microsoft Office parity.

## Native file API

The original OOXML bytes remain the authority. The supported path is XLSX-first and the same shape for documents and presentations:

1. Go `Extract*` reads package bytes (`xlsxpatch.ExtractNativeWorkbookV1` / `ExtractNativeWorkbookV2`, `docxpatch.ExtractNativeDocumentV1`, `pptxpatch.ExtractNativePPTX`).
2. The result is versioned native JSON consumed by `@injoffice/sheets`, `@injoffice/docs`, and `@injoffice/pptx-native`.
3. TypeScript paint compilers emit renderer-neutral preview commands. Paint is not a save path.
4. Editors emit mutation JSON.
5. Go `Apply*` writes those mutations back into the original archive.

React, Konva, DOM/HTML layout, screenshots, and Univer must not decide file identity or pass/fail semantics.

## Packages

Twenty-six TypeScript packages under `packages/`:

| Package | Purpose |
|---|---|
| `@injoffice/agent-tools` | Model-neutral agent sessions, bounded reads, immutable change sets, approval hooks, CAS commits, verification, and a JSON tool dispatcher |
| `@injoffice/agent-office` | Capability-scoped XLSX, DOCX, native/authored PPTX, and PDF adapters for the agent change-set lifecycle |
| `@injoffice/sheets` | Native XLSX JSON contracts, sheet paint compilers, and fail-closed mutation batches |
| `@injoffice/docs` | Native DOCX JSON contracts, pagination, and page-paint compilers |
| `@injoffice/pptx-native` | Versioned native PPTX JSON contract shared by parsers, patchers, and renderers |
| `@injoffice/pptx-authored` | Deterministic compiler from authored `DeckSpec`/`WireDeck` to native PPTX v1 |
| `@injoffice/pptx-render` | DOM-free native PPTX RenderTree compiler and renderer-neutral paint commands |
| `@injoffice/charts` | Chart specs, extraction, ECharts options, and optional Univer UI |
| `@injoffice/pivots` | Pivot specs, deterministic aggregation, and optional Univer UI |
| `@injoffice/shapes` | Shape specs, SVG geometry, XLSX conversion, and optional Univer UI |
| `@injoffice/outlines` | Validated nested row/column groups with optional Univer commands and undo |
| `@injoffice/sparklines` | Line, column, and win/loss models with optional Univer lifecycle commands and undo |
| `@injoffice/print` | Spreadsheet print configuration, lifecycle events, native conversion, and host adapters |
| `@injoffice/collab` | Presence, ordered edits, durable outbound journal, causal collaborative undo/redo, permissions, and live share |
| `@injoffice/connectors` | Host-injected data sources, atomic lifecycle/refresh commands, and deterministic range preprocessing |
| `@injoffice/slides` | `DeckSpec` authoring, wire compilation, and legacy React preview surfaces |
| `@injoffice/pdf` | PDF page, text, image, annotation, form, OCR, and redaction operations |
| `@injoffice/font-metrics` | Font discovery, native text-layout contracts, and HarfBuzz shaping |
| `@injoffice/history` | Host-backed durable version lifecycle, isolated previews, restore lineage, and grid/text diffs |
| `@injoffice/formulas` | Audited client formula facade, host-injected calculation jobs, and revisioned collaborative result distribution |
| `@injoffice/native-runtime` | Browser-safe worker protocol and lifecycle shared by optional native WASM engines |
| `@injoffice/xlsx-exchange` | Host-neutral XLSX snapshot/server-unit jobs, progress, cancellation, safe loading, and save-as |
| `@injoffice/xlsx-wasm` | Optional browser worker distribution of the native Go XLSX extract/apply engine |
| `@injoffice/pptx-wasm` | Optional browser worker distribution of the native Go PPTX extract/apply engine |
| `@injoffice/docx-wasm` | Optional browser-worker distribution of the native DOCX extract/apply engine |
| `@injoffice/univer-sheets` | Configurable composition of the complete public Univer Sheets plugin surface |

Private workspace apps (not published):

| App | Purpose |
|---|---|
| `apps/playground` | Browser engine proofs, browser-local native XLSX, DOCX, and PPTX round trips, and collaboration demos. |
| `apps/docx-page-paint-worker` | Native DOCX page-paint worker |

Go modules under `go/` are surgical file writers, native extract/apply engines, validators, an optional in-process collaboration hub, and an optional HTTP server:

| Module | Purpose |
|---|---|
| `github.com/injectinglabs/injoffice/go/xlsxpatch` | Native XLSX extract/apply, charts, pivots, drawings, and archive-preserving patches |
| `github.com/injectinglabs/injoffice/go/docxpatch` | Native DOCX extract/apply and surgical document edits |
| `github.com/injectinglabs/injoffice/go/pptxpatch` | Native PPTX extract/apply and bounded PPTX generation |
| `github.com/injectinglabs/injoffice/go/slidesqc` | Presentation quality checks |
| `github.com/injectinglabs/injoffice/go/officecompat` | Format-neutral OPC preservation, mutation envelope, and native corpus |
| `github.com/injectinglabs/injoffice/go/collab` | In-process room hub and linear operation log |
| `github.com/injectinglabs/injoffice/go/injoffice-server` | Optional XLSX/DOCX/PPTX extract/mutations HTTP server, opaque artifact store, and collab HTTP+SSE |

Each TypeScript package has its own README and five-minute example.

For npm consumers using Univer 0.25.1 (including the collaboration package's
peer dependency), add `"overrides": { "@univerjs/core": { "nanoid": "5.1.16" } }`
to your application's root `package.json`, run `npm install`, and check `npm audit`.
The repository's override does not propagate to consumers. See the
[dependency advisory and mitigation](docs/DEPENDENCY-TRANSPARENCY.md#security-advisory-snapshot).

## Development

Prerequisites: Node.js 22+ and the Go versions declared by each module's `go.mod`.

```bash
npm ci
npm run typecheck
npm test
npm run build
npm run check:packages
```

`check:packages` imports every built ESM entry and inspects every `npm pack` payload so source-only or incomplete releases fail in CI.

Run the browser playground with:

```bash
npm run dev
```

Open http://127.0.0.1:3100 directly into Sheets, the first of four comprehensive workspaces: Sheets, Docs, Slides, and PDF. The introduction is removed; legacy `#/overview` links open Sheets. Existing capabilities from the twenty-six TypeScript packages are grouped inside those examples, not separate sidebar pages. Advanced views use explicitly labelled independent samples where they do not share an artifact model.

Create and edit:

- `#/sheets` live workbook plus native XLSX inspection and bounded write-back
- `#/docs` browser-local real DOCX extraction, guarded text write-back, exact-byte readback, and explicit preservation/refusal evidence; it does not claim Word pagination or page-paint
- `#/slides` DeckSpec authoring, themes, editable transitions, layout QC, and canvas editing
- `#/pdf` real-file PDF viewing, high-DPI rendering, navigation, zoom, text search, outline inspection, current-page rotation, and download

Within Sheets, Analyze and More tools include:

- `#/sheets?feature=charts`, `pivots`, `shapes`, `connectors`, and `formulas` (the old standalone hashes remain compatible bookmarks).

Other capabilities remain inside their relevant tool workspace; these legacy bookmarks still select the corresponding internal view:

- `#/agent` provider-independent capability discovery, bounded inspection, immutable planning, preview/diff, validation, human approval, atomic commit, verification, and refusal proofs across XLSX, DOCX, PPTX, and PDF-shaped artifacts
- `#/history` structured workbook and document diffs
- `#/pptx-authored`, `#/pptx-native`, and `#/pptx-render`, including browser-local bounded PPTX text and AutoShape mutation, exact-byte re-extraction verification, and download
- `#/font-metrics` browser-safe layout contracts with an explicit Node shaping boundary
- `#/collab` zero-setup, two-editor browser collaboration simulation for sheets, documents, decks, and PDF annotations, with the HTTP + SSE sidecar proof available as an integration mode

The playground labels the runtime boundary on every route. Most proofs run entirely in the browser. Native XLSX, DOCX, and PPTX extract/apply use version-matched Go WASM engines in browser Workers by default; their server fallbacks are explicit and never automatic. The server-integration collaboration mode stays on the optional Go sidecar. The React shell is not a second OPC writer.

An optional `injoffice-server` reuses those same extract/mutation handlers and
stores packages under opaque artifact IDs (localhost by default):

```bash
cd go/injoffice-server
go run ./cmd/injoffice-server
```

Or, from the repository root:

```bash
docker compose up --build
```

Vite proxies `/healthz`, `/v1/capabilities`, `/v1/xlsx`, `/v1/docx`, `/v1/pptx`, and `/v1/collab` to the server. Compose publishes the
server only; it does not run the playground. Start the Vite playground separately,
set `VITE_INJOFFICE_API_BASE` to an explicit server origin to enable XLSX,
DOCX, and PPTX fallback, then select it under `#/sheets`, `#/docs`, or
`#/pptx-native`; or switch `#/collab` to its HTTP + SSE integration mode. An
empty API base never enables a same-origin fallback. The server is a local
development sidecar with no authentication, authorization, tenant isolation,
or rate limiting; do not expose it directly to the internet.

Static deployments do not include that local proxy, but the default XLSX,
DOCX, and PPTX browser paths need no API. A self-hosted build can set
`VITE_INJOFFICE_API_BASE` to the
HTTPS origin of a secured, API-compatible host for the explicit server fallback.

The optional `@injoffice/xlsx-wasm`, `@injoffice/docx-wasm`, and
`@injoffice/pptx-wasm` packages wrap the corresponding Go extract/apply
functions in browser workers, allowing supported local workflows without the
Go sidecar. The playground uses XLSX, DOCX, and PPTX browser paths by default
and does not upload package bytes unless the user selects an explicitly
configured server fallback. See each package README and
the matching `go/*patch/cmd/*nativewasm/README.md` binding documentation.

Run a Go module independently, for example:

```bash
cd go/xlsxpatch
go test ./...
```

## Design boundaries

- **Native file API.** Go `Extract*` → native JSON (`@injoffice/sheets`, `@injoffice/docs`, `@injoffice/pptx-native`) → TypeScript paint compilers (preview) → mutation JSON → Go `Apply*`.
- **Plain data contracts.** Native JSON, `ChartSpec`, `PivotSpec`, `DeckSpec`, and the wire types are JSON-compatible and independent of a particular host.
- **Agent change sets.** `@injoffice/agent-tools` owns a provider-neutral inspect → plan → preview/diff → validate → approve → commit → verify lifecycle. Format adapters advertise exact JSON-schema capabilities and refuse unsupported work; applications retain model, prompt, authorization, storage, and UI control.
- **Pure core, optional shell.** Aggregation, extraction, diffing, pagination, paint compilation, and mutations can run without React, Univer, or the DOM. Univer is an optional editor shell, not the file authority.
- **Optional runtime.** Libraries work offline or with a host-injected backend. The XLSX, DOCX, and PPTX WASM runtimes are the playground's browser-local defaults; the in-repo `injoffice-server` remains optional for storage, collaboration, and centralized trust or policy enforcement. Public releases contain only generic library and server components, never consumer-specific backend code, authentication configuration, tenant data, or credentials.
- **Fail-closed fidelity.** Surgical writers start from original OOXML bytes and preserve untouched ZIP parts. If a requested edit cannot be performed safely, it should return an error rather than silently rebuild and discard unsupported content.

The capabilities enumerated by the scoped v3 completion matrix are complete, but native format coverage outside that scope remains partial. Do not assume Microsoft Office, LibreOffice, or unrestricted Excel round-trip parity. See [Native Office completion](docs/NATIVE-OFFICE-COMPLETION.md) and [Public release](docs/PUBLIC-RELEASE.md) for the precise status.

## Documentation

- [Native Office completion matrix](docs/NATIVE-OFFICE-COMPLETION.md)
- [Native Office production E2E contract kit](docs/NATIVE-OFFICE-PRODUCTION-E2E.md)
- [Collaboration protocol](docs/COLLABORATION-PROTOCOL.md)
- [Agent change sets](docs/AGENT-CHANGESETS.md)
- [Public-release checklist](docs/PUBLIC-RELEASE.md)
- [Office roadmap](docs/ROADMAP.md)
- [PDF roadmap](docs/ROADMAP-PDF.md)
- [Formula compatibility](docs/FUNCTIONS.md)
- [Univer Sheets compatibility](docs/UNIVER-SHEETS-COMPATIBILITY.md)
- [Contributing](CONTRIBUTING.md)
- [Security policy](SECURITY.md)

## Project status

InjOffice is pre-1.0. APIs may evolve between minor versions until the public contracts stabilize. Capabilities and limitations are documented per package; roadmap text is not a compatibility guarantee.

# Choose your packages

Start with the document operation you need, then add presentation or integration layers. You do not need to embed the demo or install all 26 packages.

## File workflows

| Goal | Packages | Important boundary |
| --- | --- | --- |
| Validate spreadsheet models and mutations | `@injoffice/sheets` | Contracts do not themselves write XLSX bytes |
| Browser-local XLSX extraction and edits | `@injoffice/xlsx-wasm`, `@injoffice/sheets/browser` | Worker assets and source-bound supported mutations |
| Validate native document models | `@injoffice/docs/native-docx` | Narrow browser entry; not a DOCX generator |
| Browser-local DOCX text edits | `@injoffice/docx-wasm` plus the native DOCX contract | Guarded run/eligible paragraph replacement |
| Browser-local native PPTX | `@injoffice/pptx-wasm`, `@injoffice/pptx-native` | Exact parsed text and AutoShape subset |
| Author a new presentation model | `@injoffice/slides/authoring`, `@injoffice/pptx-authored` | Separate from editing an imported PPTX |
| Compile native slide previews | `@injoffice/pptx-render` | Needs explicit font/layout providers; not a save path |
| PDF inspection and page operations | `@injoffice/pdf/browser` | Pure byte transforms; viewer needs a PDF.js Worker |

Package names and subpaths are distinct: install `@injoffice/docs`, then import from `@injoffice/docs/native-docx`. Do not run `npm install @injoffice/docs/native-docx`.

## Agent workflows

Add `@injoffice/agent-tools` for the common session/tool protocol and one entry from `@injoffice/agent-office`: `/xlsx`, `/docx`, `/pptx`, or `/pdf`. Native Office adapters need host callbacks for source snapshots, isolated previews, and atomic application. The PDF adapter works with byte arrays directly.

The framework does not select a model or authenticate users. See [agent integration](../agents/integration).

## Optional workspace features

| Package | Responsibility |
| --- | --- |
| `charts`, `pivots` | Renderer-neutral chart data and deterministic aggregation; optional editor integrations |
| `shapes` | Geometry, shape models, native conversion, optional editor controls |
| `outlines`, `sparklines`, `print` | Spreadsheet models, lifecycle commands, and host integration |
| `connectors` | Host-injected data-source lifecycle and range preprocessing |
| `formulas` | Audited formula facade and host-owned calculation jobs |
| `collab`, `history` | Collaboration protocols and history lifecycles; storage remains host-owned |
| `font-metrics` | Explicit font discovery, identity, shaping, and layout contracts |
| `xlsx-exchange`, `native-runtime` | Host-neutral jobs and the shared Worker lifecycle |
| `univer-sheets` | Optional composition of public Univer Sheets plugins |

All shorthand names above use the `@injoffice/` scope. The [reference index](../reference/) links to each package's maintained README and exported-name inventory.

## Check a release before adopting it

```sh
npm view @injoffice/pdf version
npm view @injoffice/pdf exports --json
```

Pin the version you have tested, commit your lockfile, and keep native Worker/runtime/WASM assets from the same package build. These docs describe the current repository; they do not imply that every listed package is already on npm.

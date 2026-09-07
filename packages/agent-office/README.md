# @injoffice/agent-office

Provider-neutral document adapters for the InjOffice agent changeset workflow.
The package does not contain an agent loop, prompts, model SDK, hosted backend,
authentication, or UI. Applications keep control of those choices.

Each format is an independent entrypoint, so a PDF-only application does not
load or install the spreadsheet, DOCX, or slide engines:

```ts
import { createPdfAgentAdapter } from '@injoffice/agent-office/pdf'
```

The adapters provide the same lifecycle: identify the exact source, inspect or
read bounded context, validate operations, preview, diff, revision-bound commit,
and verify the committed result.

## Format support

| Entrypoint | Artifact | Mutation authority |
| --- | --- | --- |
| `/xlsx` | Native XLSX v1 or v2 projection plus host callbacks | The host applies an `injoffice.xlsx.mutations` v1 batch atomically. Stable operation ids and the changeset idempotency key are retained. |
| `/docx` | Native DOCX v1 projection plus host callbacks | The adapter resolves a stable run or eligible paragraph target and sends a source-anchored text-replacement envelope to the host. |
| `/pptx` | Authored `DeckSpec` or parsed native PPTX v1 projection | `createAuthoredDeckAgentAdapter` provides pure immutable edits, wire compilation, and layout QC. `createNativePptxAgentAdapter` sends exact text/AutoShape requests to a source-bound host writer. |
| `/pdf` | `Uint8Array` PDF bytes | Local fail-closed page transforms with fresh parse verification. |

XLSX, DOCX, and native PPTX intentionally use injected `snapshot`, `preview`, and `apply`
callbacks. That keeps native storage/WASM/HTTP deployment host-owned while the
adapter enforces source identity, CAS, validation, and idempotency plumbing.

Destructive operations are advertised explicitly. Native Office whole-file
writes, PDF delete/reorder/crop/resize/n-up, and authored slide removal require
the changeset's confirmation hook. Hosts should still apply their own access
control and durable idempotency store at the execution boundary.

## Honest boundaries

- Native XLSX support is limited to the public v1 mutation protocol. The agent adapter intentionally refuses null style fields because the effective-style projection cannot prove that a direct property was cleared.
- Native DOCX support is limited to complete source-anchored run or eligible single-run paragraph text replacement.
- Native PPTX support is limited to exact parsed text and AutoShape mutation; the authored adapter remains a separate `DeckSpec` workflow.
- PDF support here is page operations; other PDF tools can be added as separate,
  capability-scoped operations without widening the existing contract.
- Validation never grants mutation authority that the source projection or
  native host refuses.

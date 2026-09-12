# `@injoffice/pptx-wasm`

Browser-worker distribution of InjOffice's Go PPTX native extractor and surgical mutation engine. It runs entirely in the browser; the optional InjOffice HTTP server is not required.

The package truthfully exposes the native mutation subset: exact text replacement and exact AutoShape property replacement on parsed source elements. It refuses slide insertion/removal, pictures, charts, tables, connectors, groups, animation, transitions, and generic render-model write-back.

`await client.inspectTables(original, { signal })` optionally returns plain table
source text, local cell rectangles and explicit omission warnings. It accepts
only bounded, unmerged, unrotated top-level source tables; source table styles,
borders (including dash patterns), fills and text formatting are never rendered
or translated into Office defaults. Render paragraphs as plain text, and label
any host geometry guides as guides. Positioned text requires a host-selected
font and size. The result has no deck, passthrough tokens or mutation authority.

Inspection snapshots the caller's bytes, computes their SHA-256, and joins the
response to a separate strict extraction of that snapshot. Slide identity,
source part, XML-root fingerprint and unique source-node fingerprint membership
must agree. `part_sha256` hashes the complete XML part; `slide_source_sha256`
hashes its XML root and is the strict-deck join. They are distinct. For refused
tables the strict deck retains only an opaque node fingerprint, so the original
object ID and whole-part hash remain native-producer evidence, not independent
JavaScript ZIP/XML verification. The client checks canonical unique object IDs
and consistent whole-part hashes. The strict source warnings remain applicable.

Limits include 256 tables plus omissions, 4,096 cells/paragraphs, 65,536 UTF-16
text units, and 32×32 cells per table. Native inspection additionally budgets
20,000 source nodes and 16,384 runs before grammar recursion or concatenation,
with 256 paragraphs per cell and 256 runs per paragraph. The response decoder
checks exact object/array shapes, prototypes, data properties and complete grids.
Always terminate the client in `finally` when its host no longer needs it.

```ts
import { createPptxWasmClient } from '@injoffice/pptx-wasm'

const client = createPptxWasmClient()
const original = new Uint8Array(await file.arrayBuffer())
const deck = await client.extract(original)

const target = deck.slides
  .flatMap((slide) => slide.elements)
  .find((element) => element.kind === 'text' && element.source)
if (!target || target.kind !== 'text' || !target.source) throw new Error('No editable text')

const edited = await client.apply(original, deck, {
  expectedSourceRevision: deck.sourceRevision!,
  operations: [{
    operationId: 'replace-title',
    kind: 'text.replace',
    elementId: target.id,
    expectedFingerprintSha256: target.source.fingerprintSha256,
    paragraphs: [{
      align: 'left', level: 0, bullet: false,
      runs: [{
        text: 'Edited locally', bold: false, italic: false,
        fontSizeHundredthPt: 2400, color: '000000', fontFamily: 'Arial',
      }],
    }],
  }],
})
client.terminate()
```

Defaults are package-relative URLs for `pptxnative.worker.js`, `pptxnative.wasm`, and the matching `wasm_exec.js`. Modern bundlers such as Vite emit those assets from the npm tarball. Construction is SSR-safe; a Worker is created lazily on the first operation. In Node or another non-browser runtime, supply `workerFactory`.

The default Worker constructor generally requires the worker URL to be same-origin. If assets are served from another origin or a CDN, provide an application-owned `workerFactory` that satisfies your deployment's Worker, CSP, and CORS policies, plus explicit `workerUrl`, `wasmUrl`, and `goRuntimeUrl` values. CSP must permit the worker, the Go runtime script loaded by `importScripts`, and WebAssembly execution.

Inputs are copied before asynchronous queueing. The client defaults to a 32 MiB compressed package limit and 1 MiB mutation JSON limit; consumers can raise these only up to the native 512 MiB and 3 MiB hard limits. Native validation/CAS failures are recoverable. Malformed binding responses, schema drift, panics, worker errors, timeouts, and Go runtime exits terminate the worker.

CI installs the packed tarball into a clean Vite consumer at a non-root base, then runs the compiled Go engine in Chrome through extract, exact-text apply, re-extract, and text readback without an InjOffice server request.

The package includes Apache-2.0 project terms, NOTICE, and the Go BSD-3-Clause runtime license. Building it requires Node 22+ and Go 1.23.

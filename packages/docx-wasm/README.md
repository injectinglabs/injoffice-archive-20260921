# @injoffice/docx-wasm

Optional browser distribution of InjOffice's native DOCX extraction and text
mutation engine. It runs the same `go/docxpatch` implementation used by the
server inside a Web Worker; TypeScript does not become a second OOXML writer.

```bash
npm install @injoffice/docx-wasm @injoffice/docs
```

```ts
import { createDocxWasmClient } from '@injoffice/docx-wasm'
import { adaptNativeDocxProseMirrorTransactionV1 } from '@injoffice/docs'

const originalBytes = new Uint8Array(await file.arrayBuffer())
const client = createDocxWasmClient()
const document = await client.extract(originalBytes)

const adapted = adaptNativeDocxProseMirrorTransactionV1(document, transaction)
if (!adapted.ok) throw new Error(adapted.issues[0]?.message)

const saved = await client.apply(originalBytes, document, adapted.value.envelope)
const download = new Blob([saved], {
  type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
})
client.terminate()
```

`document.source.package_sha256` binds the exact original DOCX bytes. Every
text mutation also binds a modeled run or eligible paragraph ID and its source
XML SHA-256. `apply`
strictly validates the public Office mutation envelope, exact document CAS,
target anchors, XML text, duplicate/overlapping targets, and native bounds before it sends
only the payload understood by Go. The Go engine remains final save authority.

The v1 writer supports exact text-run replacement and a narrow paragraph
target: `target_kind: 'paragraph'` is valid only when that native paragraph
contains exactly one run and that run is text. Its `target_id` and
`expected_xml_sha256` bind the paragraph ID and paragraph anchor, while Go
splices the sole run's text node. A paragraph target and its underlying run
target may not appear together. Use the public `@injoffice/docs` adapter to
turn a supported ProseMirror transaction projection into the usual run-target
envelope. Unsupported structures and ambiguous, overlapping, stale, or
reordered edits are refused atomically.

Importing the package does not access `Worker`, `window`, or `document`, so it
is safe during Node.js and SSR module evaluation. The default worker is created
only by the first operation. Defaults accept at most 32 MiB DOCX packages and
3 MiB UTF-8 mutation payloads, below Go's 128 MiB and 8 MiB hard ceilings.
Applications can configure lower limits or raise them to the native ceilings:

```ts
const client = createDocxWasmClient({
  maxPackageBytes: 16 * 1024 * 1024,
  maxMutationPayloadBytes: 1024 * 1024,
})
```

These checks throw synchronously before the shared runtime copies data into a
transferable buffer. Caller-owned arrays are never transferred. Extraction
JSON is validated through `@injoffice/docs`; invalid engine output permanently
terminates that client. A settled Go runtime emits a fatal lifecycle
notification and is discarded instead of being reused.

## Assets, origins, and CSP

`docxnative.worker.js`, `docxnative.wasm`, and `wasm_exec.js` are one
version-matched unit. Default URLs use `new URL(asset, import.meta.url)`. If a
bundler does not copy non-code package assets, copy all three to a public
directory and provide explicit URLs:

```ts
const client = createDocxWasmClient({
  workerUrl: '/injoffice/docxnative.worker.js',
  wasmUrl: '/injoffice/docxnative.wasm',
  goRuntimeUrl: '/injoffice/wasm_exec.js',
})
```

The normal `new Worker(url)` path requires the worker script to be same-origin.
Do not point `workerUrl` directly at a cross-origin CDN. Copy it to the
application origin or provide a CSP-reviewed `workerFactory` that creates an
application-owned wrapper. CDN URLs for WASM and the Go runtime must permit
CORS; the worker uses `fetch` and `importScripts` respectively.

A restrictive Content Security Policy must allow the worker in `worker-src`,
the runtime load in the applicable script directive, the WASM fetch in
`connect-src`, and WebAssembly compilation (commonly `'wasm-unsafe-eval'`).
Test the exact production policy in supported browsers.

CI installs the packed tarball into a clean Vite consumer at a non-root base,
then runs the compiled Go engine in Chrome through extract, exact-run apply,
re-extract, and text readback without an InjOffice server request. Unit tests
separately cover the client protocol and worker failure envelopes. The package
contains `wasm_exec.js` from the same Go toolchain that compiled
`docxnative.wasm`; the corresponding Go BSD license is included.

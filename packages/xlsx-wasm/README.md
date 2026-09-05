# @injoffice/xlsx-wasm

Optional browser distribution of InjOffice's native XLSX extraction and
mutation engine. It runs the same `go/xlsxpatch` implementation used by the
server inside a Web Worker; TypeScript does not become a second OOXML writer.

```bash
npm install @injoffice/xlsx-wasm
```

```ts
import {
  adaptWorkbookMutationBatchV1,
  createXlsxWasmClient,
} from '@injoffice/xlsx-wasm'
import type { WorkbookMutationBatchV1 } from '@injoffice/sheets/browser'

const originalBytes = new Uint8Array(await file.arrayBuffer())
const client = createXlsxWasmClient()
const workbook = await client.extract(originalBytes)

// The public Sheets batch uses the outer exact-package CAS.
const batch = {
  protocol: 'injoffice.xlsx.mutations',
  version: 1,
  batch_id: 'save-1',
  expected_revision: workbook.source.package_sha256,
  operations: [{
    operation_id: 'edit-1',
    sheet_id: workbook.sheets[0].id,
    kind: 'cell.set_value',
    cell: { row: 0, column: 0 },
    value: 'Hello',
  }],
} satisfies WorkbookMutationBatchV1

const transaction = adaptWorkbookMutationBatchV1(workbook, batch)
const saved = await client.apply(originalBytes, workbook, transaction)
const download = new Blob([saved], {
  type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
})

client.terminate()
```

`workbook.source.package_sha256` is the outer `sha256:<digest>` CAS for the
exact XLSX bytes. `workbook.revision` is the native contract's inner
`rev:<same-digest>` CAS. The adapter requires the batch's outer revision and
creates the inner revision; `apply` validates both against the extracted
workbook and sends the outer revision to Go. It also validates untrusted
transaction objects again at the call boundary.

The native transaction currently supports cell value/formula changes,
`style.patch`, `row.set_height`, and `column.set_width`. Although the public
Sheets v1 schema recognizes `range.merge` and `range.unmerge`, this adapter
refuses both because the native XLSX transaction deliberately preserves merge
topology. To preserve the public batch's strict order, the adapter also requires
the native engine's family order: cells, then styles, then row/column layout.
`batch_id` is validated but is not forwarded: the in-memory browser call has no
persistence/idempotency store.

Importing the package does not access `Worker`, `window`, or `document`, so it
is safe during Node.js and SSR module evaluation. The default worker is only
created by the first operation. The package defaults to 32 MiB XLSX inputs and
1 MiB UTF-8 mutation JSON, below Go's 128 MiB and 3 MiB hard limits. Applications
may lower or raise the browser limits up to those native ceilings:

```ts
const client = createXlsxWasmClient({
  maxPackageBytes: 16 * 1024 * 1024,
  maxMutationPayloadBytes: 512 * 1024,
})
```

Package and mutation limits throw synchronously before the shared runtime
copies input into a transferable buffer. Caller-owned arrays are never
transferred. Extraction JSON is validated with `@injoffice/sheets`; invalid
engine output permanently terminates that client. A settled Go runtime emits a
fatal lifecycle notification and is discarded instead of being reused.

## Assets, origins, and CSP

The three files `xlsxnative.worker.js`, `xlsxnative.wasm`, and `wasm_exec.js`
are one version-matched unit. Default URLs use `new URL(asset, import.meta.url)`;
the packed-package smoke test verifies that Vite copies and resolves them. If
another bundler does not copy non-code package assets, copy all three to your
public directory and provide explicit URLs:

```ts
const client = createXlsxWasmClient({
  workerUrl: '/injoffice/xlsxnative.worker.js',
  wasmUrl: '/injoffice/xlsxnative.wasm',
  goRuntimeUrl: '/injoffice/wasm_exec.js',
})
```

The normal `new Worker(url)` path requires the worker script to be same-origin.
Do not point `workerUrl` directly at a cross-origin CDN. Either copy it to the
application origin or supply a CSP-reviewed `workerFactory` that creates an
application-owned wrapper. CDN URLs for the WASM and Go runtime must permit
CORS; the worker uses `fetch` for WASM and `importScripts` for `wasm_exec.js`.

A restrictive Content Security Policy must allow the worker in `worker-src`,
the Go runtime load in the applicable script directive, the WASM fetch in
`connect-src`, and WebAssembly compilation (commonly `'wasm-unsafe-eval'`).
Exact directives vary by browser and deployment, so test the production CSP.

CI installs the packed tarball into a clean Vite consumer at a non-root base,
then runs the compiled Go engine in Chrome through extract, apply, re-extract,
and value readback without an InjOffice server request. Unit tests separately
cover the worker protocol and refusal paths.

The package contains `wasm_exec.js` from the same Go toolchain that compiled
`xlsxnative.wasm`. The Go runtime is distributed under its BSD license, which
is included in the package.

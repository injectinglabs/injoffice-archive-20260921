# Browser workers and assets

Native XLSX, DOCX, and PPTX operations run through optional Go WASM distributions in Web Workers. They do not require an InjOffice server. Reading this documentation never instantiates those Workers.

## Choose browser-safe imports

| Use | Entry |
| --- | --- |
| XLSX contracts | `@injoffice/sheets/browser` |
| DOCX contracts and text envelope | `@injoffice/docs/native-docx` |
| Presentation authoring | `@injoffice/slides/authoring` |
| PDF page operations and viewing | `@injoffice/pdf/browser` |
| Native execution | The corresponding `@injoffice/*-wasm` package |

The root Docs/Sheets/PDF entries can include Node-qualified providers. SSR-safe module evaluation does not mean a Worker operation can execute on the server without a supplied host implementation.

## Ship a version-matched asset set

Each native package contains a Worker script, `.wasm` binary, and matching Go `wasm_exec.js`. Keep them together. With a supported bundler, package-relative URLs are emitted as assets. For an explicit DOCX configuration:

```ts
import { createDocxWasmClient } from '@injoffice/docx-wasm'

const client = createDocxWasmClient({
  workerUrl: '/injoffice/docxnative.worker.js',
  wasmUrl: '/injoffice/docxnative.wasm',
  goRuntimeUrl: '/injoffice/wasm_exec.js',
  maxPackageBytes: 16 * 1024 * 1024,
  maxMutationPayloadBytes: 1024 * 1024,
})
```

These paths are deployment examples, not files automatically created in your application. Copy the assets from the exact built/installed package if your bundler does not emit them. If you serve several engine builds with different Go runtimes, give each version its own asset directory.

## Origin and content type

The default Worker constructor expects a same-origin script. Do not point it directly to a cross-origin CDN. Use same-origin assets or an application-owned, policy-reviewed Worker wrapper. Cross-origin WASM/runtime requests need the appropriate CORS policy.

Serve `.wasm` as `application/wasm` and JavaScript as a JavaScript MIME type. A missing Worker route must return an error, not your application's HTML shell.

Your CSP must permit the intended Worker, Go runtime script, WASM fetch, and WebAssembly compilation. Test the exact production CSP and supported browsers rather than adding broad wildcard allowances.

## Lifecycle and recovery

Create a client for a bounded workflow and call `terminate()` in `finally` or your component's disposal path. Input arrays are copied by the native runtime; callers do not lose their original buffers to transfer.

Validation/stale-source refusals can be recoverable. Malformed engine responses, panics, Worker failures, timeouts, and a settled Go runtime can terminate a client. Create a fresh client after a fatal failure; do not keep retrying a dead Worker.

Apply host limits below or within each engine's published ceilings. Worker defaults are not a complete security policy.

## Test your production bundle

The repository's `npm run test:wasm-tarball-browser` exercises installed tarballs, emitted assets, and real Go execution at a non-root URL base. Run equivalent tests for your own bundler and deployment. Do not silently fall back to uploading the file when local execution fails.

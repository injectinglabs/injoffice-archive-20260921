# docxnative WebAssembly binding

This command exposes the native DOCX v1 extractor and mutation writer through
standard Go `syscall/js`. It is a transport binding over the same `docxpatch`
implementation used natively; it is not a second DOCX reader or writer.

Build the module and its matching Go runtime shim:

```bash
./go/docxpatch/cmd/docxnativewasm/build.sh
```

Pass an output directory as the first argument to keep generated artifacts out
of the source tree. The default is the ignored `dist/` directory beside this
README. `wasm_exec.js` must come from the same Go installation that builds the
module.

After `go.run(instance)` installs the binding, JavaScript waits for
`globalThis.docxnativeOnReady` and calls:

```js
const extracted = globalThis.docxnative.extract(docxBytes)
const produced = globalThis.docxnative.apply(docxBytes, payload, packageSHA256)
```

`docxBytes` and binary `payload` may be `Uint8Array` or `ArrayBuffer`; payload
may also be a JSON string. Every call returns one in-band result:

```js
{ ok: true, value: nativeV1JSON }
{ ok: true, value: mutatedDOCXUint8Array }
{ ok: false, error: "catchable refusal", fatal: false }
```

The third `apply` argument is the full `source.package_sha256` value from the
extract contract. The shorter opaque document `revision` is not save
authority. The binding enforces exact argument counts, copies caller-owned
bytes, delegates package bounds and fail-closed validation to `docxpatch`, and
recovers panics inside each callback so one refused request cannot kill the
worker instance. Expected validation and CAS refusals set `fatal: false`; a
recovered panic sets `fatal: true` and the caller must discard the worker.
Consumers may treat a missing marker from an older binding as `false`.

Run the native/WASM parity and refusal-survival contract:

```bash
cd go/docxpatch/cmd/docxnativewasm
go test -v
```

To run the same explicit build and non-skippable contract as CI:

```bash
./go/docxpatch/cmd/docxnativewasm/build.sh
(cd go/docxpatch && \
  DOCXNATIVE_WASM_REQUIRED=1 DOCXNATIVE_WASM_SKIP_BUILD=1 \
  go test -v -count=1 ./cmd/docxnativewasm)
```

The test builds a temporary module when Node and the Go WASM runtime are
available. Set `DOCXNATIVE_WASM` and optionally `DOCXNATIVE_WASM_EXEC` to test
prebuilt artifacts. `DOCXNATIVE_WASM_REQUIRED=1` converts missing Node or WASM
prerequisites into failures; combine it with `DOCXNATIVE_WASM_SKIP_BUILD=1` to
require the explicitly built artifact, as CI does.

The optional `@injoffice/docx-wasm` package supplies a version-matched browser
worker, module, Go runtime, typed `@injoffice/docs` boundary, and host-side
limits. Other production consumers should provide equivalent isolation,
terminate workers that exceed host time or memory policy, lazy-load/cache the
artifacts, and validate untrusted collaborative output again at every external
trust boundary.

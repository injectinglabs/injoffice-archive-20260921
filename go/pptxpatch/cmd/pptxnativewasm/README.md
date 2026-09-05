# pptxnative Go/WASM binding

`pptxnativewasm` wraps the same authoritative native-v1 operations used by the
Go host so local browser clients can extract and mutate PPTX bytes without an
InjOffice server:

- `pptxpatch.ExtractNativePPTX`
- `pptxpatch.MarshalNativePPTXJSON`
- `pptxpatch.ApplyNativePPTXMutationPayload`

The command is a low-level binding intended to be loaded in a Web Worker by a
separately packaged JavaScript client. It does not contain an HTTP adapter,
storage, credentials, or a server fallback.

## JavaScript contract

Once `go.run(instance)` has installed the callbacks, the module calls the
optional `globalThis.pptxnativeOnReady` function and exposes:

```js
pptxnative.extract(pptxUint8Array)
pptxnative.apply(originalUint8Array, mutationJsonStringOrBytes, expectedRevision)
// -> { ok: true, value } | { ok: false, error: string, fatal: boolean }
```

Extraction returns canonical `pptx-native/v1` JSON. Apply returns a
`Uint8Array`. `expectedRevision` is the shared envelope's exact-byte
`sha256:<digest>` CAS; the payload must contain the matching
`rev-<digest>` `expectedSourceRevision`.

Errors are returned in-band. Expected validation and CAS refusals use
`fatal: false` and leave the instance reusable. A recovered panic uses
`fatal: true`; callers must discard that worker because native state may be
unknown.

## Browser-local passthrough identifiers

Native PPTX extraction preserves unsupported source content through opaque
passthrough references. A server normally provides capability tokens that may
be resolved by a trusted artifact host. The WASM binding must not and does not
mint those capabilities.

Instead it supplies a command-local transactional factory that returns
`browser-local-v1-<sha256>` identifiers. The digest uses length-framed values
for the source revision, owner part, object ID, source fingerprint, byte length,
reason, and exact payload digest. Therefore identifiers are deterministic for
the same source package, distinct across changed preservation material, and do
not contain the source bytes or package path.

These identifiers have no resolver and authorize no remote read. A server must
reject the `browser-local-v1-` namespace anywhere it expects a host-issued
capability. Browser-local save remains safe because the surgical applier works
from the original PPTX bytes and independently re-extracts and verifies the
result; it does not resolve the JSON token.

## Build

Standard Go WASM is used, not TinyGo:

```bash
./go/pptxpatch/cmd/pptxnativewasm/build.sh
```

The scoped script writes ignored artifacts to `dist/` beside this README by
default. Set `PPTXNATIVE_WASM_OUT` to an explicit distribution directory (the
npm package build does this):

- `pptxnative.wasm`
- the matching Go toolchain's `wasm_exec.js`
- `pptxnative.worker.js` when the npm worker source is present

No generated artifact is checked in or copied into the playground.

## Contract tests

```bash
./go/pptxpatch/cmd/pptxnativewasm/build.sh
(cd go/pptxpatch && \
  PPTXNATIVE_WASM_REQUIRED=1 PPTXNATIVE_WASM_SKIP_BUILD=1 \
  go test -v -count=1 ./cmd/pptxnativewasm)
```

The tests build the module when necessary, instantiate it with Node and the
matching `wasm_exec.js`, then verify:

- canonical extract parity with in-process Go, including a refused shape with
  a browser-local passthrough identifier;
- exact mutation-byte parity with in-process Go;
- empty input, invalid payload, and stale-CAS refusals remain catchable; and
- successful extraction still works on the same instance after all refusals.

The Go tests use `dist/pptxnative.wasm` when present and otherwise try a
temporary Go WASM build. An attempted automatic build always fails the test if
it fails. `PPTXNATIVE_WASM_REQUIRED=1` turns normally skippable prerequisites,
such as missing Node, into failures. Combine it with
`PPTXNATIVE_WASM_SKIP_BUILD=1` to require the explicitly built artifact, as CI
does. Use `PPTXNATIVE_WASM` and `PPTXNATIVE_WASM_EXEC` to test artifacts at
custom paths.

## Size gate

The initial stripped standard-Go build is 6,189,217 bytes uncompressed. CI
allows up to 7 MiB, leaving limited toolchain headroom while catching accidental
dependency or payload growth. JavaScript package compression is measured
separately by the npm packaging layer.

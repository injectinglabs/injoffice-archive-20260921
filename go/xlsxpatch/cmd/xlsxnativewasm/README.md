# xlsxnative WASM browser engine

Feature-flagged syscall/js wrapper around the same Go functions as
`cmd/xlsxnative`:

- `xlsxpatch.ExtractNativeWorkbookV2`
- `xlsxpatch.ApplyNativeWorkbookMutationPayloadV1`

`@injoffice/xlsx-wasm` distributes this engine, its matching Go runtime, and a
protocol worker for server-independent local file workflows. The playground
default remains the localhost helper
(`go run ./go/xlsxpatch/cmd/xlsxnative serve`).

## Bindings

After the WASM module boots it installs `globalThis.xlsxnative`. Calls return an
in-band envelope so a refused extract/apply cannot panic `js.FuncOf` and exit
the instance:

```js
xlsxnative.extract(uint8Array)
xlsxnative.apply(originalUint8Array, payloadJsonStringOrBytes, expectedRevision)
// -> { ok: true, value } | { ok: false, error, fatal }
```

`value` is native v2 JSON (`extract`) or a `Uint8Array` (`apply`). The playground
worker and `node_contract.mjs` throw `Error(error)` on `{ok:false}`. Expected
validation and CAS refusals set `fatal: false` and keep the module alive. A
recovered Go panic sets `fatal: true`; clients must discard that worker. The
field is additive for compatibility: consumers of older modules should treat a
missing `fatal` as `false`. `expectedRevision` is the outer CAS
(`sha256:<digest>`). The native payload still carries `rev:<digest>` as required by
`ApplyNativeWorkbookMutationPayloadV1`.

## Build

Standard Go WASM (not TinyGo):

```bash
./scripts/build-xlsxnative-wasm.sh
```

That compiles with:

```bash
GOOS=js GOARCH=wasm go build -trimpath -ldflags='-s -w' \
  -o go/xlsxpatch/cmd/xlsxnativewasm/dist/xlsxnative.wasm \
  ./go/xlsxpatch/cmd/xlsxnativewasm
```

and copies `$GOROOT/misc/wasm/wasm_exec.js` (or `$GOROOT/lib/wasm/wasm_exec.js`
on newer toolchains) next to the module. Artifacts are also copied to
`apps/playground/public/xlsxnative/`. Building `@injoffice/xlsx-wasm` copies
the binary, matching runtime, and worker into its publishable `dist/`. Generated
artifacts are gitignored.

## Playground

Default: localhost helper.

Optional spike:

```bash
./scripts/build-xlsxnative-wasm.sh
npm run dev
```

The current playground does not select this runtime. Consume
`@injoffice/xlsx-wasm` directly or provide it through an application runtime
selector.

## Contract test

Happy-tree extract JSON and one-cell apply bytes are compared to in-process Go.
The same instance must also refuse empty bytes / bad CAS with the Go error text
and `fatal: false`, then extract happy-tree afterwards. The Node harness checks
the `fatal: true` consumer path with a synthetic envelope instead of exposing a
production callback that deliberately panics:

```bash
./scripts/build-xlsxnative-wasm.sh
(cd go/xlsxpatch && \
  XLSXNATIVE_WASM_REQUIRED=1 XLSXNATIVE_WASM_SKIP_BUILD=1 \
  go test -count=1 ./cmd/xlsxnativewasm)
```

or directly:

```bash
node go/xlsxpatch/cmd/xlsxnativewasm/node_contract.mjs extract \
  --wasm go/xlsxpatch/cmd/xlsxnativewasm/dist/xlsxnative.wasm \
  --wasm-exec go/xlsxpatch/cmd/xlsxnativewasm/dist/wasm_exec.js \
  --input go/xlsxpatch/testdata/excel-authored/happy-tree.xlsx
```

The Go tests use `dist/xlsxnative.wasm` when present and otherwise try a
temporary Go WASM build. `XLSXNATIVE_WASM_REQUIRED=1` turns missing Node or
WASM prerequisites into failures; combine it with
`XLSXNATIVE_WASM_SKIP_BUILD=1` to require the explicitly built artifact, as CI
does. The tests do not switch TinyGo or rewrite extract/apply.

## Size / notes

Recorded after `./scripts/build-xlsxnative-wasm.sh` with Go 1.23
`GOOS=js GOARCH=wasm` and `-ldflags='-s -w'`:

- `xlsxnative.wasm`: 6,582,926 bytes (6.28 MiB uncompressed)
- `wasm_exec.js`: 16,687 bytes, copied from `$GOROOT/misc/wasm/wasm_exec.js` (not vendored)

Standard Go WASM is used; TinyGo is not. The server remains available for
storage, collaboration, and trust enforcement, but is not required for local
browser-owned files.

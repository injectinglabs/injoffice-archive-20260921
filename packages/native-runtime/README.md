# @injoffice/native-runtime

Browser-safe worker protocol and lifecycle shared by the optional InjOffice
WASM format packages. This package contains no Office parser and no WASM
binary. Install a format package such as `@injoffice/xlsx-wasm` to process
files.

The runtime deliberately requires an injected worker factory. Importing it is
safe in Node.js and server-rendered applications, and format packages remain
responsible for resolving their own worker and WASM assets.

```ts
import { createNativeWasmClient } from '@injoffice/native-runtime'

const client = createNativeWasmClient({
  format: 'xlsx',
  workerFactory: () => new Worker(workerUrl),
  assets: { wasmUrl, goRuntimeUrl },
})

const contractJson = await client.extract(fileBytes)
const savedBytes = await client.apply(fileBytes, mutationJson, expectedRevision)
client.terminate()
```

Calls are serialized because each worker owns one native engine instance.
Input buffers are snapshotted when an operation is called and the snapshots are
transferred later, so queued work cannot observe caller mutations and the
caller's authoritative bytes are never detached. A timeout, abort, fatal worker
error, or malformed matching response terminates the worker; the next call
starts a fresh instance.
Workers may also send the versioned unsolicited `fatal` notification when an
already-ready native engine exits, so the runtime discards it while idle rather
than posting the next operation to a dead worker.

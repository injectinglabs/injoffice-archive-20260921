# @injoffice/xlsx-exchange

Host-neutral jobs for the XLSX snapshot and server-unit lifecycle: import,
optional safe workbook loading, export, save-as, progress, cancellation, and
structured failures. Hosts inject conversion, workspace, and storage adapters;
this package is not another OOXML writer.

```bash
npm install @injoffice/xlsx-exchange @injoffice/xlsx-wasm
```

```ts
import { createXlsxWasmClient } from '@injoffice/xlsx-wasm'
import {
  XlsxExchangeManager,
  createXlsxWasmImportCodec,
} from '@injoffice/xlsx-exchange'

const native = createXlsxWasmClient()
const exchange = new XlsxExchangeManager({
  codec: createXlsxWasmImportCodec(native),
  workspace: {
    async loadSnapshot(snapshot, options) {
      // Verify options.targetUnitId again when mode is "replace".
      return { unitId: options.targetUnitId ?? createWorkbook(snapshot) }
    },
    async loadServerUnit(unitId) { return { unitId } },
    async captureSnapshot(unitId) { return snapshotWorkbook(unitId) },
  },
})

const job = exchange.importSnapshot({
  source: { kind: 'bytes', bytes: new Uint8Array(await file.arrayBuffer()), name: file.name },
  load: { mode: 'new-unit' },
})
const { loadedUnitId } = await job.result
```

The WASM adapter intentionally supplies snapshot import only. InjOffice's Go
XLSX engine extracts the native workbook and writes explicit, fail-closed
mutation batches; it does not serialize arbitrary editor snapshots. Configure
`exportSnapshot` or `exportServerUnit` on a host codec that delegates to an
appropriate native service. An unconfigured direction fails with
`UNSUPPORTED_OPERATION` instead of silently rebuilding a lossy workbook.

`createInjOfficeServerImportCodec` uses the existing explicit
`POST /v1/xlsx/extract` boundary. Snapshot import returns the validated native
model; server-unit import returns the opaque artifact ID minted by the server's
host-configured store. It never falls back from local WASM to a server and
requires a nonempty server origin.

Storage references and destinations are generic host types. Byte sources are
copied before conversion, and exported bytes are copied before `saveAs`.
Conversion must finish before `loadSnapshot` / `loadServerUnit` runs, and
serialization must finish before `saveAs` runs. A replace request requires the
expected `targetUnitId`, allowing the workspace adapter to reject a stale tab
rather than replacing whichever workbook happens to be active later.

Every operation returns a job with immutable snapshots, monotonic progress,
an `AbortSignal` propagated to adapters, and a result promise. Failures are
`XlsxExchangeError` values with code, phase, job ID, operation kind,
retryability, and cause. Cancellation prevents later load/save boundaries once
the signal is observed; host adapters must honor the signal to stop work that
is already in progress.

## Collaborative import

`XlsxCollaborativeImportCoordinator` attaches an imported server unit to a
durable collaboration room through a host-injected boundary. It is deliberately
separate from `loadServerUnit`: the editor activates only the server's committed
receipt, never an uncommitted conversion result.

```ts
const coordinator = new XlsxCollaborativeImportCoordinator(exchange, {
  authorizeAndBegin: collaborationApi.authorizeAndBegin,
  quiesce: collaborationApi.quiesce,
  commitImport: collaborationApi.commitImport,
  abortImport: collaborationApi.abortImport,
  resume: collaborationApi.resume,
  activate: collaborationWorkspace.activate,
})

await coordinator.import({
  source: { kind: 'bytes', bytes, name: 'replacement.xlsx' },
  target: {
    mode: 'replace-room',
    room: 'book-1',
    expectedArtifactVersion: 'version-7',
    expectedHeadSequence: 12,
  },
})
```

Authorization happens before upload. Replacement commits require the local room
to be quiesced, have no pending outbound edits, have a fully saved head, and
match the caller's expected artifact version and operation sequence. The host's
`commitImport` must check the same identities atomically while binding the new
unit and resetting the operation epoch. Pre-commit failures resume the unchanged
room and abandon staging. A failure after the server returns a valid commit
receipt raises `XlsxCollaborativeImportRecoveryError`; it is never presented as
a rollback, and the host can use the receipt to retry idempotent local
activation.

This is an integration contract, not a bundled collaboration server. The host
still owns authenticated authorization, transaction durability, staging cleanup,
operation-log reset, and reconnect recovery.

## Collaborative export

`XlsxCollaborativeExportCoordinator` exports one declared collaboration head
under protocol `injoffice.xlsx-collaborative-export.v1`. The server boundary
must atomically authorize the caller, verify the expected artifact version, log
epoch, and sequence, quiesce a fully saved room with no outbound edits, and mint
an immutable server export unit and revision. The coordinator never captures an
arbitrary browser snapshot.

```ts
const coordinator = new XlsxCollaborativeExportCoordinator(
  exchange,
  collaborationApi,
  {
    idFactory: crypto.randomUUID,
    maxFileBytes: 64 * 1024 * 1024,
    onRecoveryRequired: ({ requestId }) => scheduleExportResync(requestId),
  },
)

const exported = await coordinator.export({
  room: 'book-1',
  expectedArtifactVersion: 'version-7',
  expectedLogEpoch: 'epoch-3',
  expectedHeadSequence: 12,
  destination: 'downloads',
})
```

The exchange manager checks the captured revision, XLSX media type, metadata,
and configured byte limit before `saveAs`. Completion records the exact revision,
byte length, SHA-256 digest, and optional destination location, and a valid
receipt must resume the same head unchanged. A lost, malformed, canceled, or
otherwise uncertain response blocks later exports. The host must call `resync()`;
only an exact committed receipt or an exact aborted-and-resumed response for the
same request ID clears that block. Protocol-bound destination references must be
bounded JSON data; credentials do not belong in them.

This package supplies the bounded client contract and validation, not a
production collaboration service or XLSX exporter. The host still owns
authentication and authorization, durable transaction/lease storage,
single-writer quiescence, immutable capture lifetime, idempotent resolution,
export persistence, and the native serializer.

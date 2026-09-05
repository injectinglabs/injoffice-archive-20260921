# @injoffice/connectors

Host-neutral data connector specifications, tabular normalization, and a Univer-bound range manager.

```bash
npm install @injoffice/connectors
```

```ts
import { csvToGrid, jsonToGrid } from '@injoffice/connectors'

const rows = jsonToGrid({ results: [{ city: 'Oslo', temp: 14 }, { city: 'Lima', temp: 22 }] }, 'results')
const csvRows = csvToGrid('city,temp\nOslo,14\nLima,22')
```

Applications provide a `SourceFetcher` that applies their own authentication, authorization, URL policy, and secret handling. The package never stores credentials or performs arbitrary network requests on its own.

`ConnectorManager` accepts optional host authorization, scheduling, time, and
preprocessing adapters. Specs can declare manual, on-open, or interval refresh,
an in-memory TTL cache, and positional cell schemas. Every fetch receives an
`AbortSignal`; `cancel()` and removal prevent late results from writing.

```ts
const manager = new ConnectorManager(univerAPI, fetchSource, {
  authorize: ({ connector }) => currentUserCanRefresh(connector.id),
  preprocessing: pipeline,
})

manager.add({
  id: 'weather', name: 'Weather',
  source: { kind: 'http', url: '/api/weather', format: 'json' },
  target: { sheetId: 'sheet-1', startRow: 1, startColumn: 0 },
  refresh: 'interval', schedule: { intervalMs: 60_000 },
  cache: { mode: 'memory', ttlMs: 30_000 },
  schema: { columns: [{ index: 0, type: 'string' }, { index: 1, type: 'number', nullable: true }] },
})
```

Scheduling is opt-in and timers are disposed with the connector. The memory
cache never stores credentials and is process-local; applications that need a
shared or durable cache should implement that policy in the host fetcher.

Ranges can also pass through a host-configured preprocessing pipeline before calculation or downstream use:

```ts
import { RangePreprocessPipeline } from '@injoffice/connectors'

const pipeline = new RangePreprocessPipeline()
pipeline.register({
  id: 'trim-text', version: '1', order: 10, deterministic: true,
  process: (grid) => grid.map((row) => row.map((cell) => typeof cell === 'string' ? cell.trim() : cell)),
})

const manifest = pipeline.manifest()
const result = await pipeline.run([[' Ada ']], {
  range: { sheetId: 'people', startRow: 0, startColumn: 0, rowCount: 1, columnCount: 1 },
  revision: 'server-revision-42', mode: 'collaborative',
}, { expectedManifest: manifest })
```

Stage order and the SHA-256 pipeline fingerprint do not depend on registration
order. `manifest()` returns a strict, credential-free descriptor that hosts can
persist with workbook or room metadata; `assertCompatibleManifest()` and the
run option refuse missing, reordered, version-drifted, or tampered contracts.
Manifest v1 fingerprints the UTF-8 bytes of
`injoffice:range-preprocess-manifest:v1\n` followed by compact JSON with
descriptor keys in `id`, `version`, `order`, `deterministic` order. This is a
coordinated protocol migration from the earlier `fnv1a32:` fingerprint: mixed
versions must be upgraded together or continue using their prior package set.
The manifest never serializes executable processor code. Collaborative runs
reject nondeterministic stages. Cancellation is passed into every stage through
an `AbortSignal`; inputs are frozen, outputs are copied and validated, and
lifecycle/error events are host-observable. The connector manager can run this
pipeline before schema validation and bound-range writes. Processors remain
host code and must avoid side effects if callers need replayable results.

## Lifecycle commands and undo

`ConnectorCommandController` adds live handles plus create, add, update, remove,
and refresh operations. Undo records contain the connector model and only the
bounded worksheet cells affected by that operation. A refresh captures the
union of its old and new extents immediately before writing. If a shrink or
write fails partway through, both the prior cells and connector runtime state
are restored. Redo replays the captured result and never fetches the external
source again.

```ts
import { registerUniverConnectorCommands } from '@injoffice/connectors/browser'

const registration = registerUniverConnectorCommands(univer, univerAPI, fetchSource, {
  create: true,
  refresh: true,
  update: false,
  remove: false,
  createMenu: true,
  refreshMenu: true,
  createInput: () => connectorPicker.open(),
  activeConnectorId: () => connectorSidebar.activeId(),
})
```

Each command and each Data-ribbon control is independently optional. Controls
are omitted unless the host provides their input/selection callback. The
browser adapter uses Univer's public command and undo/redo services; it does
not patch editor internals.

Command snapshots are deliberately credential-free. They refuse source URLs
with userinfo, query parameters, or fragments. Use an opaque gateway path in
the `ConnectorSpec`; keep tokens, cookies, authorization checks, egress rules,
and secret rotation inside the injected `SourceFetcher` and `authorize`
callback. The lower-level manager remains available to hosts that own their
own persistence and undo policy.

## Native XLSX persistence

`ConnectorNativePersistence` bridges a host binding for xlsxpatch's
`ReadConnectorDefinitions` / `SetConnectorDefinitions` functions. It validates
and copies both the model and workbook bytes. Hydration restores definitions
without automatically fetching on-open sources; the host initiates refresh only
after its authorization and egress policy are ready.

```ts
const persistence = new ConnectorNativePersistence(nativeConnectorCodec)
await hydrateConnectorManagerFromNative(manager, persistence, originalBytes)

const savedBytes = await persistConnectorManagerToNative(
  manager,
  persistence,
  originalBytes,
)
```

The explicitly supported custom OPC extension is documented in
`docs/XLSX-CONNECTOR-EXTENSION.md`. It persists only credential-free definitions
and preserves unrelated package content. It does not store fetched data, cache
contents, credentials, authorization state, or refresh status, and it is not
presented as Excel Power Query/external-connection interoperability.

Native hydration uses the manager's ordinary restore path and is intentionally
limited to client-authority managers. A manager configured with
`executionAuthority: 'server'` must initialize definitions and owned cells from
a room-bound collaboration `resync`; the authoritative live manager may still
be persisted back to native bytes.

## Collaborative ownership and refresh

`ConnectorCollaborationSession` provides ordered create, update, remove, and
server-owned refresh operations. Construct its manager with
`executionAuthority: 'server'`; this disables local lifecycle calls, on-open
and interval timers, browser fetching, and ordinary local undo restoration.
The session requires an atomic `applySnapshot` host adapter, which should call
`controller.restore(snapshot, { source: 'collaboration' })` inside the host's
transaction/from-collaboration boundary.

Clients request a refresh with the current SHA-256 model precondition and an
optional preprocessing fingerprint, but never submit fetched cells. The
authenticated server enforces connector ownership and single-flight policy,
fetches/preprocesses the source, verifies the current room sequence and target
cell SHA-256, then appends the bounded primitive-only result. Exact old/new
extent unions clear shrink remnants. Schema, workbook bounds, payload quotas,
non-overlapping connector ownership, unchanged acknowledgements, ordered
sequences, and post-apply state are checked before a client advances. Any
mismatch blocks until a room-bound, connector-owned `resync` succeeds.

The package supplies the transport contract and reducers, not an authenticated
fetch service, lease store, durable log, reconnect queue, or database
transaction. The host remains responsible for those systems and for marking
applied range writes so a generic cell collaboration stream does not echo
them.

Remaining production work includes a full source/schema/schedule browser,
permission and credential-handoff UI, durable shared caches, pagination,
production collaboration infrastructure, Excel-native connection mapping,
and multi-client/Office-load conformance.

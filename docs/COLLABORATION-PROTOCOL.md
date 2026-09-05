# InjOffice collaboration protocol

Status: draft v1

This specification defines the small server contract consumed by `@injoffice/collab`. It is transport-independent: a host can use WebSocket RPC, HTTP plus server-sent events, or another authenticated channel as long as the behavior below is preserved.

The server is responsible for authentication, room authorization, presence fan-out, and a linear operation log. The client library is responsible for editor-specific operations, catch-up, conflict transforms, and rendering collaborator state.

## Terms

- **Path**: the host's stable identifier for one artifact. It must not be accepted as an unchecked filesystem path.
- **Room**: the server's canonical room identifier for a path.
- **Client ID**: a unique identifier for one live connection. Reconnecting creates a new client ID.
- **Sequence** (`seq`): a monotonically increasing integer assigned to one accepted operation batch.
- **Head**: the latest sequence in a room.
- **Saved sequence** (`saved_seq`): the newest sequence known to be represented by the persisted artifact bytes.
- **Reset**: a signal that the persisted file was replaced and the existing operation history can no longer be replayed onto the client's snapshot.

## Transport interface

The TypeScript boundary is:

```ts
interface CollabTransport<TSelection> {
  join(path: string, name: string): Promise<JoinResult<TSelection>>
  leave(path: string): Promise<void>
  presence(path: string, selection: TSelection): Promise<void>
  onEvent(handler: (event: CollabEvent) => void): () => void
  onReconnect(handler: () => void): () => void
  opSubmit?(path: string, ops: OpRecord[], baseSeq: number): Promise<number>
  opSince?(path: string, sinceSeq: number): Promise<{
    ops: OpEntry[]
    head: number
    reset: boolean
  }>
}
```

Method names below are recommendations, not a required framing format. Field names in a raw JSON protocol use `snake_case`; an adapter may expose the camel-cased TypeScript methods above.

## RPCs

### `collab.join`

Request:

```json
{ "path": "artifact-id", "name": "Ada" }
```

Response:

```json
{
  "room": "room-id",
  "self": {
    "client_id": "client-1",
    "user_id": "user-1",
    "name": "Ada",
    "color": "#5b6cff",
    "selection": null,
    "joined_at": 1770000000000
  },
  "peers": [],
  "file": { "version": "opaque-version", "mtime": 1770000000000, "size": 4200 },
  "log": { "seq": 12, "saved_seq": 10 }
}
```

Joining is scoped to the authenticated connection. The server must authorize the current user for the artifact before revealing metadata or peers. Repeated joins for the same connection and path should be idempotent.

### `collab.leave`

Request: `{ "path": "artifact-id" }`.

The server removes this connection from the room and emits `collab.peer.left`. Disconnecting must have the same effect even when no explicit leave arrives.

### `collab.presence`

Request:

```json
{
  "path": "artifact-id",
  "selection": { "sheet": "sheet-id", "ranges": [[0, 0, 2, 3]], "active": [0, 0] }
}
```

Selection data is opaque to the server. It must be size-limited and valid JSON. The server associates it with the authenticated connection's client ID and emits `collab.presence` to other current room members. Presence is ephemeral and must not advance the operation sequence.

### `collab.op.submit`

Request:

```json
{
  "path": "artifact-id",
  "base_seq": 12,
  "ops": [{ "id": "editor.operation", "params": { "plain": "json" } }]
}
```

The server accepts the batch only when `base_seq` equals the room head. Acceptance is atomic: it increments the head once, stores one `OpEntry`, emits one `collab.op`, and returns the assigned sequence.

If the base is stale, the server rejects the request with a stable error code of `STALE_BASE` and does not append or broadcast anything. A TypeScript adapter must surface an `Error` whose message contains `STALE_BASE`; the client catches up with `op.since`, transforms pending edits, and retries.

The server treats operation IDs and params as opaque JSON. It must still enforce batch count, encoded size, nesting depth, and room authorization limits.

### `collab.op.since`

Request: `{ "path": "artifact-id", "since_seq": 10 }`.

Normal response:

```json
{
  "ops": [
    { "room": "room-id", "seq": 11, "client_id": "client-2", "ops": [] },
    { "room": "room-id", "seq": 12, "client_id": "client-3", "ops": [] }
  ],
  "head": 12,
  "reset": false
}
```

Entries must be ordered by sequence, contain no duplicates, and cover every retained sequence after `since_seq` through `head`. If the requested history was cleared, compacted, or belongs to replaced file bytes, return `reset: true`. A reset response may omit old entries; the client reloads the artifact and rejoins before applying further operations.

## Server events

Every event frame has an `event` string and `payload` object.

```ts
type CollabEvent =
  | { event: 'collab.peer.joined'; payload: { room: string; peer: PeerInfo } }
  | { event: 'collab.peer.left'; payload: { room: string; client_id: string } }
  | { event: 'collab.presence'; payload: { room: string; client_id: string; selection: unknown } }
  | { event: 'collab.file.changed'; payload: FileChange }
  | { event: 'collab.op'; payload: OpEntry }
```

`collab.op` must be emitted only after the entry is committed. Servers may echo it to the submitting connection; clients identify their own entry by client ID and sequence. Events for rooms a connection has not joined must never be delivered.

A file-change payload is:

```json
{
  "room": "room-id",
  "path": "artifact-id",
  "author": "user",
  "action": "saved",
  "mtime": 1770000000000,
  "size": 4300,
  "version": "opaque-version-2",
  "saved_seq": 12,
  "reset": false
}
```

Set `reset: true` when the artifact was replaced wholesale rather than saved from the current operation history.

## Ordering and reconnect behavior

1. A client joins and records the returned log head.
2. Local edit batches submit against the last applied head.
3. The server serializes accepted batches and assigns each exactly one sequence.
4. On `STALE_BASE`, the client requests missing entries, applies or transforms them, and resubmits.
5. After reconnect, the client rejoins because its client ID changed, catches up from its last applied sequence, and only then flushes queued edits.
6. On `reset: true`, the client discards replay assumptions and reloads the latest persisted artifact.

Sequence numbers provide ordering, not durable identity across resets. A server should use transactions or an equivalent single-writer mechanism so two submissions cannot both be accepted against the same base.

## Importing a replacement artifact

`@injoffice/xlsx-exchange` exposes a transport-independent collaborative import
coordinator under protocol `injoffice.xlsx-collaborative-import.v1`. Its server
adapter has three durable phases:

1. `authorizeAndBegin` authenticates the caller, authorizes the requested new or
   replacement room, and returns a short-lived opaque transaction ID before XLSX
   bytes are uploaded.
2. For replacement, the client quiesces its room and supplies the exact room,
   artifact version, log epoch, head, saved sequence, and pending count observed.
3. `commitImport` atomically rechecks the transaction and replacement head,
   binds the imported server unit, resets the operation epoch, and returns a
   durable receipt representing an active, fully saved room.

A pre-commit failure calls `resume` for the unchanged quiesced room and
`abortImport` for staged resources. Once `commitImport` returns, clients must not
issue a compensating rollback: local activation is idempotently retried from the
receipt. Malformed commit responses are an uncertain-outcome recovery condition,
not permission to delete either artifact. Servers must expire transactions,
deduplicate commit retries by transaction ID, and authorize every phase.

## Exporting a declared collaboration head

`@injoffice/xlsx-exchange` also exposes protocol
`injoffice.xlsx-collaborative-export.v1`. Its client coordinator requires a
server boundary with three idempotent operations:

1. `authorizeAndCapture` authenticates and authorizes the caller, atomically
   compares the declared artifact version, log epoch, and head sequence,
   quiesces a room with no pending outbound operations and `saved = head`, and
   returns a short-lived lease for an immutable export unit and revision.
2. The client exports only that immutable server unit. Before `saveAs`, it
   enforces the configured byte quota, XLSX media type, bounded metadata, and
   exact captured revision. It then submits the revision, byte length, SHA-256
   digest, and optional destination location to `completeExport`.
3. `completeExport` durably records that byte/head binding and resumes the same
   unchanged room. Its receipt must reproduce the lease and an active room at
   the exact exported head.

Any ambiguous begin, export, save, or completion outcome blocks further exports
for that coordinator. `resolveExport` must look up the request ID idempotently
and return either the exact committed receipt or confirm that the transaction
was abandoned and the exact unchanged room resumed. Malformed or different-head
resolution remains blocked and fail-closed.

This is a transport-neutral contract. The repository does not provide the
authenticated server, durable lease/receipt store, single-writer quiescence,
immutable server capture lifecycle, destination service, or production XLSX
serializer.

## Connector operation subprotocol

`@injoffice/connectors` defines `injoffice.connector-ops.v1` for connector
definitions and owned output ranges. Lifecycle submissions contain no fetched
result cells or credentials; create/update operations do carry the bounded,
credential-free source descriptor. A refresh is a separate authenticated request
carrying the connector ID, client request ID, current SHA-256 model precondition,
and optional deterministic preprocessing fingerprint. The browser must never
fetch and submit the result in collaborative mode.

The server verifies room sequence, authorization, connector ownership and
single-flight policy, model SHA-256, preprocessing version, schema, target-cell
SHA-256, output quotas, and overlap policy. It fetches the external source under
server-held credentials and appends one operation containing a primitive-only
rectangular result. The rectangle covers the union of old and new extents, with
nulls in every shrink remainder. The append and authoritative workbook/model
update must be one transaction; only its committed entry is acknowledged and
broadcast.

Clients use a manager in server-authority mode, which disables local schedules,
fetches, lifecycle mutations, and ordinary undo restores. They require an atomic
host apply adapter and verify exact envelopes, contiguous sequences, unchanged
request identity, model/range preconditions, schema, bounds, overlap, and final
postconditions. A conflict or gap blocks the stream. Recovery accepts only a
room-bound snapshot whose primitive ranges cover every and only connector-owned
extent. The library does not provide the authenticated fetch worker, lease
database, durable log, or workbook transaction.

## Collaborative undo and redo

Undo is a client operation, not a server rollback. A client records the forward
and inverse mutations for each of its own actions and rebases both stacks over
every later remote entry before planning an undo or redo. The resulting inverse
batch is submitted as a new local operation at the current head; historical log
entries are never removed or rewritten.

`CollaborativeUndoManager` implements this bounded policy for sheet cell and
row/column range mutations. Later remote writes neutralize overlapping cells in
both inverse and forward batches. Structural changes remap mutation geometry and
saved selections. Planning and commit are separate and revision-checked so a
plan cannot execute after a remote entry invalidates it. Content-only
move-range and reorder-range mutations use exact later-writer cell ownership;
reorders protect both destination cells and the live source cells they read.
Structure-versus-content-range interactions, unsupported object mutations,
malformed geometry, and ambiguous single-range splits return typed blocking
reasons; clients must resync or leave that history entry unavailable rather
than applying a guessed inverse.

`CollaborativeUndoJournalCoordinator` connects those plans to the durable
outbound journal. Forward local actions and subsequent inverse/redo batches use
normal append-only operations; the server never rewrites history. Undo/redo
batches are durably enqueued behind an in-process send gate, atomically applied
by a host executor, committed in the local planner, and only then released to
transport. A remote batch ordered before still-pending local work moves its
inverse geometry but does not claim later local cells. Once the local batch is
acknowledged, later remote writes use the normal ownership-neutralization rules.

The resync implementation must deliver every retained remote operation to the
coordinator in sequence before advancing the journal head. Sequence gaps,
execution refusal, and acknowledgement gaps block further undo/redo. A hard
authoritative reload can clear that block only after durable pending entries
are resolved, and clears local undo history because snapshot reload cannot
prove old inverse preconditions. This protocol does not persist the undo stack,
define proprietary Univer command internals, or add object-specific inverses.

## Security and operational requirements

- Authenticate every RPC and event connection; authorize every path on every call.
- Derive `user_id` and `client_id` server-side. Never trust those fields from a client payload.
- Use an opaque artifact identifier or normalize and constrain path resolution to an allowed root.
- Bound operation bytes, operation count, presence bytes, JSON depth, room members, replay length, and request rate.
- Do not log operation params or selections by default; they can contain document content.
- Apply backpressure or disconnect slow event consumers rather than growing an unbounded queue.
- Expire presence promptly after disconnect and prevent name/color fields from carrying markup into a UI.
- Encrypt network traffic outside a trusted local boundary.

Authentication mechanisms, database schemas, retention periods, and artifact storage are deliberately outside this protocol.

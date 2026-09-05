# @injoffice/collab

Client-side presence and ordered-operation synchronization for sheets, documents, and `DeckSpec` presentations.

Sheet presence includes colored remote ranges, a cell-anchored collaborator name flag, a live `editing` state, and a throttled plain-text `draft` while the edit is uncommitted. Drafts are ephemeral presence; only committed mutations enter the ordered log. Pass `peerLabelComponent` to `PresenceManager` to replace the flag with your own React component, or `false` to render ranges only. The default component exposes stable `ioc-sheet-presence-label*` class names and `--ioc-peer-color` for host styling.

```bash
npm install @injoffice/collab
```

```ts
import { COLLAB_EVENTS, type CollabTransport } from '@injoffice/collab'

const transport: CollabTransport = {
  join: (path, name) => rpc('collab.join', { path, name }),
  leave: (path) => rpc('collab.leave', { path }),
  presence: (path, selection) => rpc('collab.presence', { path, selection }),
  onEvent: (handler) => socketEvents((frame) => {
    if (COLLAB_EVENTS.has(frame.event)) handler(frame)
  }),
  onReconnect: (handler) => reconnectEvents(handler),
  opSubmit: (path, ops, baseSeq) => rpc('collab.op.submit', { path, ops, base_seq: baseSeq }),
  opSince: (path, sinceSeq) => rpc('collab.op.since', { path, since_seq: sinceSeq }),
}
```

`rpc`, `socketEvents`, and `reconnectEvents` above are host functions. This package includes no server, authentication, storage, or network implementation. See the repository's collaboration protocol specification for the complete contract.

## Permissions and live share

`CollabPermissionManager` projects a trusted, revisioned host snapshot into
deny-by-default `view`, `edit`, `comment`, `present`, `follow`, and
`manage-members` decisions. Owner, editor, commenter, and viewer roles provide
defaults; per-user allow/deny rules are supported with deny taking precedence.
`enforce()` is a hook for host mutations. Invalid, stale, conflicting, wrong-room,
and out-of-sequence role changes leave the current snapshot untouched.

`LiveShareSession` models presenter/follower state over a host-supplied
`LiveShareTransport`. A server-ordered event stream starts and stops presenters
and carries validated worksheet viewports. Followers apply those viewports
through a host adapter. The session ignores stale events, freezes on sequence
gaps until `resync()`, binds events to the current permission revision, clears
presentation/following when roles are revoked, and releases subscriptions on
leave or disposal. Local sends become state only after the authoritative event
is echoed, so a rejected send cannot create a phantom presentation.

These are client enforcement and lifecycle primitives, not an authorization
boundary. The host must authenticate client/user identity, issue permission
snapshots, authorize every server mutation, assign room event revisions, and
provide replay/durability. Untrusted clients can bypass JavaScript checks.

## Offline outbound journal

`DurableOutboundJournal` persists ordered local mutation batches through an
injected `OutboundJournalStorage`. Each entry gets a stable key composed from a
persisted ordinal and a host token; the server must scope deduplication by room,
client ID, and key. `connect()` always calls the host's `resync` boundary before
replaying. That boundary can return exact-key transformed batches after applying
remote changes, or block unsafe replay without dropping local data.

Submissions are single-flight and acknowledged in order. The journal persists
attempt counts before sending and removes an entry only after its acknowledgement
is durable. If acknowledgement persistence fails, the same key is retried so an
idempotent server can return the original result. Retry uses configurable
exponential delays through an injected scheduler; the package never sleeps or
owns a timer. Disconnect aborts in-flight I/O, cancellation is explicit, and
disposal rejects local waiters while retaining durable entries for a later
instance. Unknown versions, malformed JSON, invalid entries, and storage bound
to another room/client are refused without overwrite.

`createJournalSubmitter(journal, getBaseSeq)` plugs the journal into
`SubmitQueue` through that queue's existing public send callback. The storage,
scheduler, resync/rebase implementation, and idempotent transport remain host
owned. A stable client ID must survive reloads. This is client crash recovery,
not proof of server durability, production offline conflict coverage, encryption,
quota management, or authenticated delivery.

`enqueue(..., { deferSend: true })` separates durable acceptance from transport
release. The returned receipt exposes the stable key, acknowledgement promise,
and an idempotent `release()`. This is an in-process transaction gate used by
collaborative undo. A restarted journal releases every retained entry because
the undo stack itself is not crash-persisted.

Hosts must mark an acknowledgement as deduplicated when its sequence is not
newer than the resynchronized head; otherwise the journal blocks instead of
assuming an old sequence is safe. Only one active journal may own a given
storage key at a time.

The sheet operation transformer rebases sparse cell writes, range lists, and
peer selections through row/column insertion, removal, and contiguous block
moves. When a block move would split an operation whose wire shape permits
only one range, the result is flagged as lossy instead of widening the target;
hosts should use their normal resync path. Content-only move-range and
reorder-range mutations participate in exact cell ownership: later local cells
are removed from remote matrices, and reorder mutations are split into safe
column bands while protecting both their destination and live source cells.
Interactions between those content mutations and later structural geometry,
plus host object-operation transforms, remain outside this bounded slice and
fail closed.

`CollaborativeUndoManager` is a host-neutral undo/redo rebase layer. Record a
local action's already-computed inverse and forward mutation batches, pass each
ordered remote batch to `rebaseRemote`, execute the mutations returned by
`planUndo` or `planRedo` as normal local mutations, and call `commit` only after
that batch succeeds. Remote row/column insertion, removal, and block moves
rebase cell/range inverses and saved selections. A later remote write to the
same cell removes that cell from both sides of the local history entry, so undo
does not overwrite newer remote content. Plans carry a revision and cannot be
committed after an intervening record, clear, or remote rebase.

The manager fails closed with typed reasons when a transform would split a
single-range mutation, when a remote write follows a local structural action,
or when structure intersects a content-range mutation, or when rich-text,
drawing, image, chart, shape, or malformed operations do not have an
ownership-preserving inverse. Content-only move/reorder inverses use the same
later-writer ownership rules as cell writes. It does not patch
Univer's private undo stack or manufacture inverses; hosts must bridge their
command service to this public planner.

`CollaborativeUndoJournalCoordinator` is that bounded public bridge. It records
already-applied local actions into both the planner and durable journal,
executes remote batches through an injected atomic host, and sends undo/redo
inverses as new journal entries. An inverse is durably enqueued but held from
transport until host execution succeeds and the revision-checked undo plan
commits. Neutralized plans move between stacks without emitting empty network
entries.

The coordinator tracks which history entries still have outbound work. A
remote cell write ordered before one of those pending actions does not take
ownership from the later local write, while remote structural edits still move
its inverse geometry and saved selection. After acknowledgement, a later remote
write can neutralize overlapping cells normally. Stream or acknowledgement
gaps and atomic execution failures block the coordinator until the host loads
authoritative state; hard resync deliberately clears local undo history.

Hosts must route every missing remote entry through `receiveRemote` in sequence
before allowing the journal resync boundary to report its head, and must avoid
capturing coordinator-executed batches a second time. The host still supplies
atomic command execution, inverse construction, selection restoration, durable
storage, transport, authentication, and server transforms for pending outbound
batches. Undo history is memory-only; crash recovery replays durable mutations
but does not restore the prior undo stack. Rich/floating objects and ambiguous
structure/content interactions remain fail-closed.

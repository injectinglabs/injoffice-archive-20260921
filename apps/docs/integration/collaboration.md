# Collaboration and history

Collaboration state, saved file revisions, and history snapshots are related but distinct. Decide which host is authoritative for each before wiring UI events to persistence.

## Collaboration

`@injoffice/collab` supplies client lifecycle and transport-neutral contracts. The optional Go hub and HTTP/SSE server provide an integration path; a browser simulator demonstrates behavior without implementing a production multi-user service.

A real host owns authenticated room membership, permissions, durable operation ordering, reconnect/replay, and storage. Keep collaboration sequence numbers separate from exact native package fingerprints.

Server-ordered changes should not be replayed into a local-only undo stack as if the user made them. On gaps, conflicts, or rejected authority checks, stop and resynchronize from a validated authoritative snapshot instead of silently merging arbitrary models.

The [collaboration protocol](../reference/generated/contracts/collaboration-protocol) documents sessions, presence, operation submission, and resynchronization. The [package reference](../reference/generated/packages/collab) covers client helpers and lifecycle behavior.

## History

`@injoffice/history` manages host-backed version lifecycles, isolated previews, restore lineage, and structured diffs. It does not provide durable storage by itself.

A production host should store source identity and immutable version content together, enforce access checks when listing or reading history, and atomically compare the current revision before restoring. A restore creates a new head/lineage event; do not destroy the previous head just because a user previews an older version.

See [history integration](../reference/generated/packages/history).

## Object-level features

Chart and pivot collaboration have their own bounded conflict semantics. A chart's object revision and its sheet's layer revision are not interchangeable. A pivot conflict can be at whole-object granularity. Do not assume all feature packages implement a shared general-purpose CRDT.

Connectors and calculated results also need source-revision checks. A delayed refresh or formula job must not overwrite a newer workbook snapshot.

## Integration checklist

- Assign stable document, object, operation, and actor IDs.
- Define the authoritative order and durable idempotency boundary.
- Treat remote edits separately from local undo intent.
- Reconcile native file saves against the exact byte revision.
- Test reconnect, stale operations, permission changes, and recovery.
- Keep previewing history non-mutating and guard restore explicitly.

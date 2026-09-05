# @injoffice/pivots

Renderer-neutral pivot specifications, deterministic aggregation, a Univer manager, and XLSX wire conversion.

```bash
npm install @injoffice/pivots react react-dom
```

```ts
import { bakePivot } from '@injoffice/pivots'

const result = bakePivot(
  [
    ['Region', 'Sales'],
    ['West', 20],
    ['East', 35],
    ['West', 15],
  ],
  { rows: ['Region'], columns: [], values: [{ field: 'Sales', agg: 'sum' }] },
)
```

The aggregation engine is pure and does not require a spreadsheet editor. It applies page selections and include/exclude member filters before aggregation and supports deterministic ascending/descending row and column label order. `PivotManager` and `PivotPanel` provide optional Univer/React integration. The repository's Go patch library reads and writes native XLSX pivot definitions; `pivotsFromFile` converts its JSON result to editable specs and refuses native features it cannot represent.

`PivotCommandController` exposes live handles plus create, add, update, remove,
and distinct row, column, value, filter, page-field, and sort operations. Every
successful operation captures the complete managed pivot collection before and
after mutation, including stable native identities. Browser hosts can import
`@injoffice/pivots/browser` and call `registerUniverPivotCommands`; every public
command and the optional Insert ribbon item has an independent feature flag,
while the restore mutation connects snapshots to Univer's undo/redo service.
Pass the returned `controller` to `<PivotPanel controller={controller} />` so
row, column, value, filter, totals, and remove controls enter that same undo
path. The legacy `manager` prop remains available for hosts that deliberately
do not install an undo sink. Untrusted updates cannot replace a pivot id or
native identity.

`PivotCollaborationSession` supplies server-first create, update, and remove
operations for the supported pivot model. Its shared pure reducer validates
exact envelopes, stable native identity, full-object fingerprints, and ordered
base sequences. The manager and its deterministic baked grid change only
after an unchanged authoritative acknowledgement; gaps and conflicts block
the session until a validated `resync()` succeeds. Hosts whose generic cell
collaboration also observes baked-grid writes should inject `applySnapshot` to
apply the model and derived cells transactionally while marking the derived
writes local-only/from-collaboration.

The host still owns authenticated authorization, operation-ID idempotency,
durable append/broadcast, offline replay, presence, and collaborative undo
rebasing. Conflicting changes to separate pivots serialize cleanly, while two
changes to the same pivot conflict at object granularity rather than silently
merging field lists or cache/source state.

Hydrated specs carry a stable `nativeIdentity` based on the pivot-table part. `toWirePivotUpdate` and `toWirePivotRemove` build identity-bound lifecycle requests, and `PivotManager` exposes the same builders for managed specs. Native conversion supports ordered page fields with zero or one selected member, hidden-member filters, row/column label sorts, and the shared-item kinds string, finite number, boolean, and blank. Stateful writes require a complete `fieldMembersOf` inventory; ambiguous or unsupported members are rejected. The Go layer resolves the complete relationship graph before updating or deleting, preserves shared caches, and fails closed on stale or ambiguous identities.

Value/label conditions, value-based autosort, nested presentation, multiple column fields, arbitrary and external sources, formatting, slicers, timelines, durable/offline collaboration infrastructure, and rebased multi-user undo remain outside this bounded surface. The package does not claim general Excel pivot semantic parity.

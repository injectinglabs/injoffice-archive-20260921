# `@injoffice/sparklines`

Renderer-neutral spreadsheet sparklines for InjOffice. The package owns plain-JSON models, validation, data extraction, deterministic line/column/win-loss geometry, SVG serialization, and create/update/remove/group/ungroup lifecycle APIs. The main entry has no DOM, React, Univer, or proprietary-package dependency.

## Five-minute example

```ts
import { SparklineManager, sparklineGeometryToSvg } from '@injoffice/sparklines'

const manager = new SparklineManager()
const sparkline = manager.create({
  type: 'line',
  source: { sheetId: 'sheet-1', startRow: 1, startColumn: 1, endRow: 1, endColumn: 6 },
  target: { sheetId: 'sheet-1', row: 1, column: 7 },
  options: { showHigh: true, showLow: true, emptyCells: 'connect' },
})

const geometry = manager.render(
  sparkline.id,
  () => [[4, 8, null, 3, 12, 6]],
  { width: 96, height: 20, padding: 2 },
)

document.querySelector('#sparkline')!.innerHTML = sparklineGeometryToSvg(geometry)
```

Call `manager.group([firstId, secondId])` to give related sparklines a shared value axis. `serialize()` and atomic `hydrate()` provide a versioned host-owned persistence boundary.

`SparklineCommandController` provides live handles and snapshot-based undo
records for create, update, remove, group, and ungroup operations. The optional
`@injoffice/sparklines/browser` entry registers matching commands and restore
mutations with Univer's public command and undo/redo services. A host can opt
individual commands out. The Insert-ribbon item is added only when the host
provides `createInput`, keeping dialog and selection policy application-owned.

The optional `@injoffice/sparklines/react` entry mounts an accessible insertion
panel with editing controls, group/ungroup workflow, and a host-feedable
`SparklineSvg` cell renderer. Mutations go through `SparklineCommandController`
so the same snapshot undo path as public Univer commands is used. The host still
owns Univer cell painting, localization, and visual placement.

For shared editing, `SparklineCollaborationSession` exposes server-first
create, update, remove, group, and ungroup operations. An authoritative host
must serialize operations by room, reject stale base sequences, enforce
authorization and operation-ID idempotency, and apply the exported pure
`applySparklineCollaborationOperation` reducer before append. Clients verify
stable object fingerprints, exact group-mutation closures, ordered sequences,
and unchanged acknowledgements; a gap, stale base, or conflict blocks further
edits until the host supplies a validated snapshot through `resync()`.

This protocol is deliberately object-level rather than a last-writer-wins
snapshot broadcast. It does not provide a server, durable log, identity
service, reconnect queue, presence UI, or collaborative undo rebase. Hosts
must connect those policies and should route collaborative edits through the
session rather than mutating the manager in parallel.

## Integration and persistence boundary

The manager deliberately accepts a `SparklineValueReader`, so a host can connect it to Univer OSS, another grid, or an in-memory workbook without coupling this package to an editor. Geometry is renderer-neutral; `sparklineGeometryToSvg` is a convenience renderer.

The package exports `sparklinesFromFile` and `toWireSparklines` as the typed
bridge to `go/xlsxpatch.ReadSparklines` and `SetSparklines`. The native layer
round-trips Excel's x14 line, column, and win/loss groups, source and target
references, RGB colors, markers, empty-cell behavior, direction, line weight,
and supported axis bounds. It preserves unrelated worksheet extensions and
all untouched OPC entries. External references, non-RGB colors, date axes,
hidden-cell plotting, explicit x-axis display, mixed unrepresentable axis
modes, malformed XML, and unequal per-member options are reported and refused
instead of being silently discarded.

Native group IDs are deterministic import identities, not persistent OOXML
object IDs: the x14 format does not store one. Native round-trip is not a claim
that a Univer cell renderer is auto-mounted; hosts import
`@injoffice/sparklines/react` when they want the insertion panel.

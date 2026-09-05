# @injoffice/shapes

Plain-JSON spreadsheet shape specifications, SVG geometry, Univer floating-shape integration, and XLSX wire conversion.

```bash
npm install @injoffice/shapes react react-dom
```

```ts
import { roundRectPath, shapeLabel, toWireShapeUpdate } from '@injoffice/shapes'

const path = roundRectPath(240, 100, 0.16)
console.log(shapeLabel('roundRect'), path)
```

Shapes hydrated by `specsFromFileShapes` carry a stable native identity made
from their drawing-part path and DrawingML `cNvPr` object id. Use
`toWireShapeUpdate`, `toWireShapeRemove`, or the corresponding `ShapeManager`
request builders to target exactly that native object. Browser-created shapes
have no native identity until the host saves and rehydrates them, so lifecycle
builders refuse to guess.

Geometry helpers are pure and renderer-independent. `ShapeManager`,
`ShapeFloat`, and `ShapePanel` provide the optional editor shell. Grouped shapes
remain opaque; unknown XLSX drawing objects must remain under the host's
fail-closed preservation policy.

`ShapeCommandController` adds live handles and complete versioned snapshots
for create, add, update, and remove. Browser hosts can import
`@injoffice/shapes/browser` and call `registerUniverShapeCommands` to register
independently configurable commands plus Univer undo/redo restore mutations.
The optional Insert ribbon item is present only when the host supplies a shape
picker. Snapshot restore validates kinds, ids, anchors, and native identities,
and attempts to restore the prior snapshot if mounting fails. Inline text edits
made directly by `ShapeFloat`, transforms, connectors, groups, and z-order are
not yet routed through this command controller.

`ShapeCollaborationSession` provides server-first create, update, remove, and
persisted cell-anchor operations for that currently supported lifecycle. The
host's authoritative service must authenticate and serialize submissions,
reject stale bases, deduplicate operation IDs, run the exported pure reducer,
and durably append before acknowledgement. Clients validate exact envelopes,
stable IDs/native identities, object fingerprints, and contiguous sequences;
conflicts block edits until the host supplies a validated `resync()` snapshot.
Mounted state is restored only after acknowledgement, and a failed mount rolls
back before the session blocks.

The package does not bundle that service, offline replay, presence, or
collaborative undo rebasing. Concurrent edits to different properties of one
shape conflict at object granularity. Pixel-level float-DOM transforms,
connectors, groups, and z-order do not yet have model operations, so the
collaboration contract intentionally does not claim them.

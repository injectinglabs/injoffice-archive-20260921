# @injoffice/outlines

Renderer-neutral row and column outline groups for InjOffice spreadsheets.
Groups use zero-based inclusive ranges, may be disjoint or properly nested,
and are limited to Excel's eight outline levels. Crossing or duplicate ranges
are rejected instead of being silently normalized.

```ts
import { OutlineManager } from '@injoffice/outlines'
import { createUniverOutlineVisibilityAdapter } from '@injoffice/outlines/browser'

const outlines = new OutlineManager(createUniverOutlineVisibilityAdapter(univerAPI))
outlines.add({
  id: 'detail-rows',
  sheetId: 'sheet-1',
  axis: 'row',
  start: 2,
  end: 12,
  collapsed: false,
})
outlines.setCollapsed('detail-rows', true)
```

`UniverOutlineController` can create row or column groups directly from the
active selection and clear groups fully contained by that selection, making it
suitable for host-provided Group/Ungroup commands.

`OutlineCommandController` adds live handles and complete before/after snapshot
records for add, update, collapse, expand, remove, and ungroup. The optional
`registerUniverOutlineCommands` browser adapter registers those operations with
Univer's public command and undo/redo services and can add configurable Group
Rows, Group Columns, Ungroup Rows, and Ungroup Columns items to the Data ribbon.
Import it from `@injoffice/outlines/browser`; the main entry does not load
Univer. Every command and menu item can be disabled.

`layoutOutlineGutter` builds a nested row or column margin with 1-based outline
levels. The optional `@injoffice/outlines/react` entry mounts `OutlineGutter`,
an accessible expand/collapse gutter with level buttons, `aria-expanded`
controls, and keyboard movement. Hosts still own Univer header painting.

## Collaborative outline commands

`OutlineCollaborationSession` is a transport-neutral server-first command
target for create, update, collapse/expand, remove, selection clear, and undo
restore. Each operation carries a stable-ID complete snapshot, exact SHA-256
precondition, and declared-action diff. The manager changes only after an
unchanged consecutive authoritative acknowledgement. Remote entries and
room-bound resync do not create local undo records.

```ts
import { OutlineCollaborationSession, OutlineManager } from '@injoffice/outlines'
import {
  createUniverOutlineVisibilityAdapter,
  registerUniverOutlineCommands,
} from '@injoffice/outlines/browser'

const manager = new OutlineManager(
  createUniverOutlineVisibilityAdapter(univerAPI),
  { mutationAuthority: 'collaboration' },
)

const registration = registerUniverOutlineCommands(univer, univerAPI, {
  commandTargetFactory(controller) {
    const session = new OutlineCollaborationSession(controller, transport, {
      room: 'workbook-1',
      clientId,
      idFactory: crypto.randomUUID,
      authorize: canEditOutlines,
    })
    session.start()
    return session
  },
}, manager)
```

Collaboration authority mode rejects direct `OutlineManager` mutations so
registered commands cannot bypass ordering. The server must authenticate and
authorize the actor, enforce the exact base sequence, deduplicate operation
IDs, run `applyOutlineCollaborationOperation`, durably append, then acknowledge
and broadcast. Resync envelopes must come from that same authenticated room
authority. Client permission hooks are only an immediate UI policy boundary.
Hosts that combine outline visibility with another renderer or persisted state
can inject `applySnapshot`; it must apply completely or roll back before
returning false.

`manager.hydrate()` intentionally refuses in collaboration authority mode.
Clients should initialize from an already validated `OutlineStore` passed to
the constructor before the session starts, or apply the authenticated room's
`OutlineCollaborationResync` envelope. Native XLSX decoding by itself is not a
room authority and must not bypass that resync boundary.

`OutlineStore` provides validation, querying, nesting depth, atomic hydration,
range clearing, and updates without requiring a browser or Univer.
`transformOutline` shifts outline ranges through row or column insertion and
removal for collaboration and undo systems.

## Native XLSX persistence

`createNativeXlsxOutlineWrite` converts a validated group snapshot to the
snake-case contract accepted by Go's `xlsxpatch.ApplyWorksheetOutline`.
`decodeNativeXlsxOutlineSnapshot` validates and converts the corresponding
`xlsxpatch.ReadWorksheetOutline` result. The native implementation preserves
unrelated OPC parts byte-for-byte and rewrites only `outlineLevel`, `hidden`,
`collapsed`, and `sheetPr/outlinePr` summary direction metadata in the target
worksheet.

SpreadsheetML persists outline levels `0..7`; its eighth visible outline state
is the ungrouped level. The renderer-neutral store permits eight logical nested
groups, but the native XLSX helper rejects a snapshot deeper than seven instead
of emitting an unrepresentable file. Independently hidden level-zero rows and
columns are preserved. A collapsed native group requires a valid adjacent
summary row/column and fully hidden detail; malformed or ambiguous source
metadata is refused.

The browser adapter owns visibility for outline-managed ranges. Applications
that also hide rows or columns independently should provide an adapter that
combines those visibility sources. Native file parity covers the supported
group, collapse, and summary-direction semantics; Excel's outline header gutter
controls remain follow-up UI work, so this package does not claim full Univer
Pro outline parity yet.

The collaboration package does not bundle identity, transport, durable event
storage, offline replay, presence, structural-edit integration, property-level
merging, collaborative undo rebasing, or atomic native-XLSX transactions. Its
full collection precondition deliberately makes concurrent outline edits one
conflict domain.

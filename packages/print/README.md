# @injoffice/print

Renderer-neutral spreadsheet print configuration and lifecycle APIs.

```ts
import { PrintManager } from '@injoffice/print'

const print = new PrintManager({
  print: async (snapshot) => renderAndPrint(snapshot),
  screenshot: async (request) => renderRangeToDataUrl(request),
  writeClipboardImage: async (dataUrl) => copyDataUrl(dataUrl),
}, 'sheet-1')

print.updatePrintConfig({ paperSize: 'A4', direction: 'Landscape', scale: 'FitPage' })
print.updatePrintRenderConfig({ gridlines: true })
await print.print()
```

## Persisted configuration commands

`PrintConfigurationSnapshotV1` persists layout and render settings together;
dialog visibility is intentionally ephemeral. `PrintConfigurationCommandController`
serializes concurrent layout, render, full-replacement, and undo/redo restore
operations. Every candidate is fully validated and passed to the injected
`PrintConfigurationPersistence.save` before local state changes. A rejected
durable save therefore leaves both the manager and undo history untouched.
The persistence adapter must provide an atomic complete-snapshot replacement;
this package does not silently select browser storage or a server database.

Browser hosts can import `@injoffice/print/browser` and call
`registerUniverPrintConfigurationCommands`. The adapter registers workbook-
scoped update/replace commands and a restore mutation against Univer's public
command and undo services. Layout, render, and full-replacement commands can be
omitted independently.

## Collaborative persisted configuration

`PrintCollaborationSession` is a transport-neutral, server-first command target
for the persisted configuration. It submits complete validated replacements
with an exact SHA-256 precondition and changes the local manager only after an
unchanged, consecutive server acknowledgement. Remote entries and explicit
resynchronization persist through the same serialized controller without adding
local undo; acknowledged local edits add one complete snapshot inverse. Client
permission hooks are useful for immediate UI refusal, but do not replace server
authentication and authorization.

```ts
import { PrintCollaborationSession } from '@injoffice/print'
import { registerUniverPrintConfigurationCommands } from '@injoffice/print/browser'

const registration = registerUniverPrintConfigurationCommands(
  univer,
  'workbook-1',
  print,
  persistence,
  {
    commandTargetFactory(controller) {
      const session = new PrintCollaborationSession(controller, transport, {
        room: 'workbook-1',
        clientId,
        idFactory: crypto.randomUUID,
        authorize: canEditPrintConfiguration,
      })
      session.start()
      return session
    },
  },
)
```

The authoritative service must authenticate the actor, enforce policy and the
exact base sequence, deduplicate operation IDs, apply the exported
`applyPrintCollaborationOperation` reducer, durably append the accepted entry,
and only then acknowledge and broadcast it. The host must obtain resync
snapshots and sequence numbers from that same authenticated authority.
`registration.dispose()` also disposes a factory-created command target.

## Accessible React print workspace

`@injoffice/print/react` exports a host-mountable `PrintWorkspace`. It provides
a responsive, keyboard-contained print workflow for area, paper, orientation,
scaling, custom margins and paper, repeated rows/columns, page content,
alignment, header/footer tokens, text watermarks, page ranges, preview refresh,
cancel, and print. Every durable setting goes through the supplied
`PrintConfigurationCommandController`, so persistence and snapshot undo keep
the same semantics as programmatic edits. Preview refreshes cancel superseded
sessions, and errors, empty output, page counts, and rendering progress have
screen-reader states.

```tsx
import { PrintWorkspace } from '@injoffice/print/react'

<PrintWorkspace
  controller={commands}
  previewManager={previews}
  renderPage={(page) => <img src={page.payload.imageUrl} alt="" />}
  config={{ controls: { watermark: false }, paperSizes: ['A4', 'Letter'] }}
/>
```

The host still owns the portal/application shell and turns each validated
preview payload into React content. `registerUniverPrintUI` from
`@injoffice/print/browser` optionally adds a Start/Layout ribbon button and
public open command; its injected `open()` callback mounts or reveals the
workspace. The command, ribbon item, action buttons, preview, and individual
control groups can be omitted independently. No hidden browser print call,
DOM-to-canvas conversion, or storage choice is made by this package.

Hosts can provide their own page renderer without coupling the package to a
DOM, canvas, PDF engine, or browser permission:

```ts
import { PrintPreviewManager } from '@injoffice/print'

const previews = new PrintPreviewManager({
  renderPreview: async (request, { signal }) => paginateWithMyRenderer(request, signal),
})
const session = previews.start('preview-42', print.snapshot(), {
  selection: { kind: 'range', startPage: 2, endPage: 4 },
})
const preview = await session.result
```

`PrintPreviewManager` validates one-based all/page-list/page-range selection,
page count and point geometry, structured-cloneable host payloads, lifecycle
events, cancellation, and timeouts. Cell-range selection continues to travel
in `PrintSnapshot.layout.subUnitIds`. It is a preview-session contract, not a
browser or Excel-compatible renderer.

The public model covers the current documented area, paper, orientation,
scaling, repeated-heading, margin, alignment, gridline, header/footer,
screenshot, dialog, and cancellable-event vocabulary. Actual rendering,
browser printing, screenshots, and clipboard access are host adapters: this
package neither requires a DOM nor claims that a browser granted permission.

`hydrateNativePrintSetup` maps the bounded output of
`xlsxpatch.ReadPrintSetups` into editable layout/render configuration. Native
print areas, orientation, named paper, fit/scale, margins, print options,
centered headers/footers, and exact repeated row/column title ranges are
recognized; unsupported native syntax is returned as warnings.

`toNativePrintSetup` converts only settings supported by
`xlsxpatch.SetPrintSetup`. It returns `setup: null` with diagnostics whenever a
requested setting would be lost. Named paper sizes, explicit scaling,
gridlines, headings, and centered page alignment are persisted. Repeated title
rows/columns, custom paper dimensions, end alignment, watermarks, and six-zone
custom headers/footers still require native writer expansion.

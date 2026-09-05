# @injoffice/history

Host-backed durable version lifecycle plus pure, deterministic diff helpers for
spreadsheet grids and document text.

```bash
npm install @injoffice/history
```

```ts
import { HistoryCommandController, HistoryManager, diffGrids, diffText } from '@injoffice/history'

const history = new HistoryManager(xlsxHistoryHost, 'workbook-42')

await history.capture({
  snapshot: xlsxBytes,
  author: { id: session.userId, kind: 'user', displayName: session.name },
  contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  expectedHeadVersionId: currentVersionId,
})

const page = await history.listVersions({
  limit: 50,
  filter: { authorKinds: ['user'], createdAfter: startOfMonth },
})

// A detached copy for a read-only historical viewer. Do not replace the live
// editor with this value.
const preview = await history.loadPreview(page.versions[0].id)

// Restore loads the immutable source and asks the host to append a new head.
const restored = await history.restore({
  sourceVersionId: preview.version.id,
  expectedHeadVersionId: currentVersionId,
  author: { id: session.userId, kind: 'user' },
})

const sheetDiff = diffGrids([['Name', 'Score'], ['Ada', 8]], [['Name', 'Score'], ['Ada', 10]])
const textDiff = diffText('draft one', 'draft two')
```

For a live editor, wrap the manager in a serialized workflow controller. The
workspace owns snapshot capture, detached preview presentation, and activation
of the newly created restore version:

```ts
const commands = new HistoryCommandController(history, {
  captureSnapshot: async ({ signal }) => workbookCodec.capture({ signal }),
  showIsolatedPreview: (preview) => historicalViewer.open(preview),
  activateRestoredVersion: async (version) => collaboration.reloadAndRejoin(version.id),
})

await commands.capture({ author: { id: session.userId, kind: 'user' } })
await commands.preview('version-17')
await commands.restore({ sourceVersionId: 'version-17', author: { id: session.userId, kind: 'user' } })
```

Applications using Univer can opt into any subset of the public commands:

```ts
import { registerUniverHistoryCommands } from '@injoffice/history/browser'

const registration = registerUniverHistoryCommands(univer, commands, {
  capture: canSave,
  restore: canRestore,
  preview: true,
  list: true,
  cancel: true,
})
```

The optional React entry provides a host-mountable, accessible timeline with
change/author filters, cursor pagination, isolated preview actions, explicit
restore confirmation, retry/empty states, and permission-gated controls:

```tsx
import { HistoryTimeline } from '@injoffice/history/react'

<HistoryTimeline
  controller={commands}
  author={{ id: session.userId, kind: 'user', displayName: session.name }}
  canCapture={permissions.canSaveVersion}
  canRestore={permissions.canRestoreVersion}
/>
```

The component inherits the editor font and exposes stable `ioc-history` class
names plus `--ioc-history-*` color properties for host theming. Supplying
`initialPage` supports preloaded and server-rendered sidebars.

`mountHistoryTimeline` and `attachHistorySidebar` place that timeline in a host
sidebar node. `registerUniverHistoryUI` from `@injoffice/history/browser` adds a
Start/History ribbon command and, by default, calls `host.open()` immediately so
the durable timeline is mounted when the Univer adapter registers. Restore
confirmation traps Tab/Escape. Production permission UX, historical workbook
rendering, and localization remain host-owned.

`HistoryHost` is intentionally injected. Authentication, tenant isolation,
durable blob/database storage, quota enforcement, and the retention scheduler
belong to the application server. The host returns effective per-version
retention metadata (`policyId`, `expiresAt`, and `legalHold`). It must apply
filters before cursor pagination and keep snapshots immutable.

`HistoryManager` validates both requests and host responses, preserves author
attribution and restore lineage, supports optimistic head checks, returns
structured-cloned previews, and refuses a restore that overwrites its source
instead of creating a new version. The package does not own permission policy,
a production storage service, or the local undo stack. Restore first appends a durable version
and then calls `activateRestoredVersion`; if activation fails,
`HistoryActivationError` retains the created version so the host can retry
reload/rejoin without repeating the durable restore.

## Collaborative restore

`HistoryCollaborationCoordinator` supplies the reset boundary needed when a
versioned workbook is also an active collaboration room:

```ts
import { HistoryCollaborationCoordinator } from '@injoffice/history'

const collaboration = new HistoryCollaborationCoordinator(roomId, {
  authorizeRestore: ({ sourceVersionId, signal }) => permissions.canRestore(sourceVersionId, signal),
  quiesce: ({ signal }) => room.pauseAndFlush(signal),
  reloadAndReset: ({ version }) => room.replaceFileAndResetLog(version.id),
  rejoin: ({ versionId }) => room.joinVersion(versionId),
  resume: () => room.resume(),
  onRecoveryRequired: ({ version }) => showReloadRequired(version.id),
})

const commands = new HistoryCommandController(history, {
  ...workspace,
  prepareRestore: (context) => collaboration.prepare(context),
})
```

The coordinator checks restore permission before pausing, then requires a
quiesced room with no pending outbound edits, a fully saved operation head,
and (when supplied) the expected artifact version. Durable version creation
happens only after those checks. A successful restore atomically reloads the
new version, resets the old operation epoch, and rejoins; a failed version
creation resumes the unchanged room. The host remains responsible for
authenticated authorization, atomic server storage/log replacement, and
recovery UI if reload or rejoin fails after the durable version was created.

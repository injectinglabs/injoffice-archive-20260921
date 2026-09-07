import { useEffect, useRef, useState } from 'react'
import { ICommandService, LocaleType, mergeLocales } from '@univerjs/core'
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import sheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US'
import '@univerjs/preset-sheets-core/lib/index.css'
import {
  PresenceManager,
  type CommandServiceLike,
} from '../../../../packages/collab/src/index.js'
import { createHttpCollabTransport, type HttpCollabTransport } from '../collabTransport'
import { COLLAB_COPY, parseCollabQuery } from '../collabScope'
import { createOssUniver } from '@injoffice/univer-sheets'
import { bindUniverColorScheme, univerDarkMode } from '../univerColorScheme'
import { deferNestedReactRootStart } from '../nestedReactRootLifecycle'
import {
  COLLAB_API_BASE,
  CollabRoomChrome,
  defaultCollabName,
  fetchSampleWorkbook,
  mintCollabArtifact,
  XLSX_TYPE,
  type ConnectionState,
} from './roomChrome'

const SHEET = 's1'

const SEED = [
  ['Collaborative plan', 'Owner', 'Status', 'Next step'],
  ['Native XLSX', 'Mira', 'In progress', 'Review write-back'],
  ['PDF viewer', 'Noah', 'Planned', 'Add upload'],
  ['DOCX renderer', 'Rin', 'In progress', 'Run fidelity corpus'],
]

function buildCellData() {
  const cellData: Record<number, Record<number, { v: string }>> = {}
  SEED.forEach((row, rowIndex) => {
    cellData[rowIndex] = {}
    row.forEach((value, columnIndex) => {
      cellData[rowIndex]![columnIndex] = { v: value }
    })
  })
  return cellData
}

export function CollabSheetsPanel() {
  const editorRef = useRef<HTMLDivElement | null>(null)
  const uploadRef = useRef<HTMLInputElement | null>(null)
  const transportRef = useRef<HttpCollabTransport | null>(null)
  const managerRef = useRef<PresenceManager | null>(null)
  const apiRef = useRef<ReturnType<typeof createOssUniver>['univerAPI'] | null>(null)
  const commandServiceRef = useRef<CommandServiceLike | null>(null)
  const [name, setName] = useState(defaultCollabName)
  const [artifact, setArtifact] = useState(() => parseCollabQuery(window.location.href).artifact)
  const [manager, setManager] = useState<PresenceManager | null>(null)
  const [connection, setConnection] = useState<ConnectionState>('idle')
  const [editorReady, setEditorReady] = useState(false)
  const [busy, setBusy] = useState(false)
  const [status, setStatus] = useState('Mint or paste an artifact ID, then join. Open the same URL in a second tab.')
  const [error, setError] = useState('')
  const [metrics, setMetrics] = useState({ peers: 0, applied: 0, pending: 0 })

  useEffect(() => deferNestedReactRootStart(() => {
    if (!editorRef.current) return
    const { univer, univerAPI } = createOssUniver({
      locale: LocaleType.EN_US,
      locales: { [LocaleType.EN_US]: mergeLocales(sheetsCoreEnUS) },
      presets: [UniverSheetsCorePreset({ container: editorRef.current })],
      darkMode: univerDarkMode(),
    })
    const unbindScheme = bindUniverColorScheme(univerAPI)
    univerAPI.createWorkbook({
      id: `collab-${crypto.randomUUID()}`,
      name: 'collaboration-protocol-proof.xlsx',
      sheets: {
        [SHEET]: {
          id: SHEET,
          name: 'Plan',
          cellData: buildCellData(),
          columnData: { 0: { w: 170 }, 1: { w: 90 }, 2: { w: 110 }, 3: { w: 180 } },
        },
      },
    })
    apiRef.current = univerAPI
    commandServiceRef.current = univer.__getInjector().get(ICommandService) as CommandServiceLike
    setEditorReady(true)

    return () => {
      unbindScheme()
      managerRef.current?.stop()
      transportRef.current?.close()
      managerRef.current = null
      transportRef.current = null
      apiRef.current = null
      commandServiceRef.current = null
      // Univer owns a nested React root. Dispose it after the outer reset or
      // route commit so React never tears down both roots synchronously.
      window.setTimeout(() => univer.dispose(), 0)
    }
  }), [])

  const disconnect = () => {
    managerRef.current?.stop()
    transportRef.current?.close()
    managerRef.current = null
    transportRef.current = null
    setManager(null)
    setMetrics({ peers: 0, applied: 0, pending: 0 })
  }

  const joinRoom = async (path: string) => {
    const api = apiRef.current
    const commandService = commandServiceRef.current
    if (!api || !commandService) {
      setError('The Univer editor is still starting. Try Join again in a moment.')
      return
    }
    setBusy(true)
    setConnection('joining')
    setError('')
    disconnect()
    try {
      const transport = await createHttpCollabTransport(COLLAB_API_BASE)
      transportRef.current = transport
      const nextManager = new PresenceManager(api, transport, { path, name, commandService })
      managerRef.current = nextManager
      nextManager.onChange(() => {
        setMetrics({
          peers: nextManager.peers().length,
          applied: nextManager.appliedSeq,
          pending: nextManager.pendingOps,
        })
      })
      nextManager.onResync((reason) => {
        setConnection('error')
        setError(`The operation log requires a workbook reload (${reason}). Remint the artifact to restart this proof.`)
      })
      nextManager.onFileChanged(() => {
        setStatus('The backing XLSX changed outside this editor. Reload from the file authority before continuing.')
      })
      nextManager.onPeerSaved(() => {
        setStatus('A peer saved a newer XLSX version; this editor already has its ordered operations.')
      })
      await nextManager.start()
      if (!nextManager.clientId) throw new Error('The sidecar did not join the collaboration room.')
      setManager(nextManager)
      setArtifact(path)
      setMetrics({ peers: nextManager.peers().length, applied: nextManager.appliedSeq, pending: nextManager.pendingOps })
      setConnection('live')
      setStatus(`Live as ${nextManager.self?.name ?? name}. Edit any cell, then watch it appear in the second tab.`)
    } catch (cause) {
      disconnect()
      setConnection('error')
      setError(cause instanceof Error ? cause.message : String(cause))
    } finally {
      setBusy(false)
    }
  }

  const onMint = async (file: Blob) => {
    setBusy(true)
    setError('')
    try {
      const id = await mintCollabArtifact(file)
      setArtifact(id)
      await joinRoom(id)
    } catch (cause) {
      setConnection('error')
      setError(cause instanceof Error ? cause.message : String(cause))
      setBusy(false)
    }
  }

  return (
    <CollabRoomChrome
      title="Live Univer protocol proof"
      connection={connection}
      name={name}
      onNameChange={setName}
      artifact={artifact}
      onArtifactChange={setArtifact}
      busy={busy}
      joinDisabled={!editorReady}
      onJoin={() => void joinRoom(artifact)}
      presence={manager}
      status={status}
      error={error}
      proof={COLLAB_COPY.proof}
      editorTitle="Shared editor"
      editorHint="Edit cells directly in Univer. The uploaded XLSX identifies the room; this proof does not import its contents."
      editorMeta={editorReady ? 'Editor ready' : 'Starting editor'}
      metrics={metrics}
      tryIt={(
        <ol>
          <li>Mint the sample or upload an XLSX.</li>
          <li>Open the same URL in another tab.</li>
          <li>Edit a cell and move the selection.</li>
          <li>Confirm the ordered edit and remote cursor in both tabs.</li>
        </ol>
      )}
      actions={(
        <>
          <button
            type="button"
            className="workbench-button"
            disabled={busy || !editorReady}
            onClick={() => {
              void fetchSampleWorkbook()
                .then((blob) => onMint(blob))
                .catch((cause: unknown) => setError(cause instanceof Error ? cause.message : String(cause)))
            }}
          >
            Mint sample
          </button>
          <button type="button" className="workbench-button" disabled={busy || !editorReady} onClick={() => uploadRef.current?.click()}>
            Open XLSX
          </button>
          <input
            ref={uploadRef}
            className="visually-hidden"
            type="file"
            accept={`.xlsx,${XLSX_TYPE}`}
            disabled={busy || !editorReady}
            aria-label="Open an XLSX file for a collaboration room"
            onChange={(event) => {
              const file = event.target.files?.[0]
              if (file) void onMint(file)
              event.target.value = ''
            }}
          />
        </>
      )}
    >
      <div ref={editorRef} className="collab-univer" data-testid="collab-univer-editor" />
    </CollabRoomChrome>
  )
}

export default CollabSheetsPanel

import { useEffect, useRef, useState, useSyncExternalStore } from 'react'
import { ICommandService, LocaleType, mergeLocales } from '@univerjs/core'
import type { FUniver } from '@univerjs/core/lib/facade'
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import sheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US'
import '@univerjs/preset-sheets-core/lib/index.css'
import {
  PresenceManager,
  type CommandServiceLike,
} from '../../../packages/collab/src/index.js'
import { BrowserCollabHub, createBrowserCollabHub } from './browserCollabTransport'
import { RemoteCellGhost } from './cellGhost'
import { createOssUniver } from '@injoffice/univer-sheets'
import { bindUniverColorScheme, univerDarkMode } from './univerColorScheme'
import { parseCollabQuery, writeCollabQuery, type CollabFormat } from './collabScope'
import { SIM_EDITORS as EDITORS, SIM_FORMAT_HINT, SIM_ROOMS, SimEditorFrame, SimFormatTabs } from './collabSimChrome'
import { CollabSimulatorDocs } from './collabSimulatorDocs'
import { CollabSimulatorPdf } from './collabSimulatorPdf'
import { CollabSimulatorSlides } from './collabSimulatorSlides'

const ROOM = SIM_ROOMS.sheets
const SHEET = 's1'

const WORKBOOK = [
  ['Launch plan', 'Owner', 'Status', 'Next step'],
  ['Native XLSX', 'Mira', 'In progress', 'Review write-back'],
  ['PDF viewer', 'Noah', 'Planned', 'Add upload flow'],
  ['DOCX renderer', 'Rin', 'In progress', 'Run fidelity corpus'],
  ['Collaboration', 'Mira + Noah', 'Live', 'Edit this row together'],
]

function cellData() {
  return Object.fromEntries(WORKBOOK.map((row, rowIndex) => [
    rowIndex,
    Object.fromEntries(row.map((value, columnIndex) => [columnIndex, { v: value }])),
  ]))
}

function SimulatedEditor({ hub, profile }: { hub: BrowserCollabHub; profile: (typeof EDITORS)[number] }) {
  const editorRef = useRef<HTMLDivElement | null>(null)
  const stageRef = useRef<HTMLDivElement | null>(null)
  const [status, setStatus] = useState<'starting' | 'live' | 'error'>('starting')
  const [metrics, setMetrics] = useState({ peers: 0, applied: 0, pending: 0 })
  const [presence, setPresence] = useState<PresenceManager | null>(null)
  const [editorApi, setEditorApi] = useState<FUniver | null>(null)

  useEffect(() => {
    let disposed = false
    let stopEditor: (() => void) | null = null
    const startEditor = () => {
      if (disposed || !editorRef.current) return
      const { univer, univerAPI } = createOssUniver({
        locale: LocaleType.EN_US,
        locales: { [LocaleType.EN_US]: mergeLocales(sheetsCoreEnUS) },
        presets: [UniverSheetsCorePreset({ container: editorRef.current })],
        darkMode: univerDarkMode(),
      })
      const unbindScheme = bindUniverColorScheme(univerAPI)
      univerAPI.createWorkbook({
        id: `collab-simulator-${profile.id}-${crypto.randomUUID()}`,
        name: 'collaboration-simulator.xlsx',
        sheets: {
          [SHEET]: {
            id: SHEET,
            name: 'Plan',
            cellData: cellData(),
            columnData: { 0: { w: 155 }, 1: { w: 100 }, 2: { w: 105 }, 3: { w: 175 } },
          },
        },
      })
      const commandService = univer.__getInjector().get(ICommandService) as CommandServiceLike
      const transport = hub.connect({ userId: profile.id, color: profile.color })
      const manager = new PresenceManager(univerAPI, transport, {
        path: ROOM,
        name: profile.name,
        commandService,
        publishIntervalMs: 60,
        peerLabelComponent: false,
      })
      const update = () => setMetrics({
        peers: manager.peers().length,
        applied: manager.appliedSeq,
        pending: manager.pendingOps,
      })
      manager.onChange(update)

      stopEditor = () => {
        unbindScheme()
        manager.stop()
        transport.close()
        // Univer owns a nested React root. Dispose it after React finishes the
        // outer cleanup so Strict Mode cannot synchronously unmount both roots.
        window.setTimeout(() => univer.dispose(), 0)
      }

      void manager.start().then(() => {
        if (disposed) return
        update()
        setStatus(manager.clientId ? 'live' : 'error')
        if (manager.clientId) {
          setPresence(manager)
          setEditorApi(univerAPI)
        }
      })
    }

    // Effects are intentionally deferred: Strict Mode's development probe can
    // cancel the first pass before either nested Univer root is constructed.
    // Staggering Noah also avoids two render engines claiming containers in the
    // same frame on slower browsers.
    const startTimer = window.setTimeout(startEditor, profile.id === 'noah' ? 75 : 0)

    return () => {
      disposed = true
      window.clearTimeout(startTimer)
      stopEditor?.()
    }
  }, [hub, profile])

  return (
    <SimEditorFrame profile={profile} status={status} presence={presence} metrics={metrics} stageRef={stageRef}>
      <div ref={editorRef} className="collab-sim-editor__canvas" data-testid={`collab-editor-${profile.id}`} />
      {presence && editorApi ? (
        <RemoteCellGhost api={editorApi} manager={presence} editorRef={editorRef} stageRef={stageRef} />
      ) : null}
    </SimEditorFrame>
  )
}

function activityLabel(kind: string) {
  if (kind === 'operation') return 'Edit'
  if (kind === 'presence') return 'Cursor'
  return kind === 'join' ? 'Joined' : 'Left'
}

const FORMAT_INTRO: Record<CollabFormat, { body: string; steps: [string, string, string]; real: string }> = {
  sheets: {
    body: 'Both panes are independent Univer instances. Remote selections carry the collaborator’s color and name; while a cell is open, its uncommitted draft appears inside that cell in the other pane, then Enter sends the authoritative edit through the ordered operation log.',
    steps: [
      'Select a cell in Mira’s editor.',
      'Double-click and watch Mira’s draft appear for Noah.',
      'Type live, then press Enter to commit the value.',
    ],
    real: 'Both Univer editors, captured sheet mutations, remote replay, selection presence, conflict handling, and the public transport contract.',
  },
  docs: {
    body: 'Both panes are independent ProseMirror editors. Typing and caret {from,to} go through the ordered operation log. This is not a native DOCX editor.',
    steps: [
      'Type in Mira’s document.',
      'Watch the same characters appear for Noah.',
      'Move the caret and confirm {from,to} in Noah’s footer.',
    ],
    real: 'Both ProseMirror editors, document steps, remote replay, caret presence, and the public transport contract.',
  },
  slides: {
    body: 'Both panes share one DeckSpec. Title, theme, one-slide structure, and shape overrides sync. Bulk outline rebuilds do not.',
    steps: [
      'Edit the title in Mira’s deck.',
      'Watch Noah’s title update.',
      'Change slides to publish presence.',
    ],
    real: 'Both DeckSpec canvases, last-writer-wins field operations, slide presence, and the public transport contract.',
  },
  pdf: {
    body: 'Both panes annotate the same sample PDF. Highlights, notes, form values, and deletes sync. Page, text, and image edits are not collaborated.',
    steps: [
      'Click Highlight in Mira’s pane.',
      'Watch the markup appear for Noah.',
      'Toggle Agree or edit the memo, then confirm the form value on Noah’s side.',
    ],
    real: 'Both PDF annotators, the annotation/form collab codec, remote replay, page presence, and the public transport contract.',
  },
}

export function CollabSimulator() {
  const [format, setFormat] = useState<CollabFormat>(() => parseCollabQuery(window.location.href).format)
  const [session, setSession] = useState(() => ({ id: 1, hub: createBrowserCollabHub() }))
  const snapshot = useSyncExternalStore(session.hub.subscribe, session.hub.getSnapshot, session.hub.getSnapshot)
  const intro = FORMAT_INTRO[format]

  const resetSession = () => setSession(({ id }) => ({ id: id + 1, hub: createBrowserCollabHub() }))
  const selectFormat = (next: CollabFormat) => {
    if (next === format) return
    if (snapshot.head > 0 && !window.confirm('Switch format? Unsaved edits in both simulated editors will be lost.')) return
    const { artifact } = parseCollabQuery(window.location.href)
    writeCollabQuery({ artifact, format: next })
    setFormat(next)
    resetSession()
  }

  return (
    <div className="collab-simulator workbench-surface ds">
      <div className="collab-sim-toolbar workbench-toolbar ds-workstrip" role="toolbar" aria-label="Collaboration simulation controls">
        <div>
          <strong>Two editors, one browser room</strong>
          <span>Runs entirely in this page</span>
        </div>
        <span className="workbench-badge workbench-badge--live" role="status"><i aria-hidden="true" />Simulation live</span>
        <button type="button" className="workbench-button" onClick={() => {
          if (snapshot.head > 0 && !window.confirm('Reset both editors? Unsaved demo edits will be lost.')) return
          resetSession()
        }}>
          Reset both editors
        </button>
      </div>
      <SimFormatTabs format={format} onFormat={selectFormat} hint={SIM_FORMAT_HINT[format]} />

      <section className="collab-sim-intro" aria-labelledby="collab-sim-title">
        <div>
          <h2 id="collab-sim-title">Edit on either side. See exactly who is there.</h2>
          <p>{intro.body}</p>
        </div>
        <ol aria-label="How to try the collaboration simulator">
          {intro.steps.map((step) => <li key={step}>{step}</li>)}
        </ol>
      </section>

      <div className="collab-sim-editors ds-dual" key={`${session.id}-${format}`}>
        {format === 'sheets' ? EDITORS.map((profile) => <SimulatedEditor key={profile.id} hub={session.hub} profile={profile} />) : null}
        {format === 'docs' ? <CollabSimulatorDocs hub={session.hub} /> : null}
        {format === 'slides' ? <CollabSimulatorSlides hub={session.hub} /> : null}
        {format === 'pdf' ? <CollabSimulatorPdf hub={session.hub} /> : null}
      </div>

      <section className="collab-sim-ledger" aria-labelledby="collab-ledger-title">
        <header>
          <div>
            <h3 id="collab-ledger-title">Shared operation ledger</h3>
            <p>This is the browser hub’s ordered view of both editors.</p>
          </div>
          <dl>
            <div><dt>Editors</dt><dd>{snapshot.peers}</dd></div>
            <div><dt>Head sequence</dt><dd>{snapshot.head}</dd></div>
            <div><dt>Persistence</dt><dd>Memory only</dd></div>
          </dl>
        </header>
        <div className="collab-sim-ledger__body">
          {snapshot.activity.length ? (
            <ol className="collab-sim-events" aria-live="polite">
              {snapshot.activity.map((event) => (
                <li key={event.id}>
                  <span>{activityLabel(event.kind)}</span>
                  <strong>{event.actor}</strong>
                  <p>{event.detail}</p>
                  <code>{event.seq === undefined ? 'ephemeral' : `seq ${event.seq}`}</code>
                </li>
              ))}
            </ol>
          ) : <p className="collab-sim-empty">Connecting both editors…</p>}
          <aside className="collab-sim-truth" role="note">
            <strong>What is simulated</strong>
            <p>Room membership, server ordering, replay, and event delivery live only in this browser tab. No file is uploaded and refreshing clears the session.</p>
            <strong>What is real</strong>
            <p>{intro.real}</p>
          </aside>
        </div>
      </section>
    </div>
  )
}

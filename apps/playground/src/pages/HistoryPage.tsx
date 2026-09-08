import { useEffect, useId, useMemo, useRef, useState } from 'react'
import { diffGrids, diffText } from '@injoffice/history'
import type { HistoryCommandController } from '../../../../packages/history/src/commands'
import type { HistoryManager } from '../../../../packages/history/src/manager'
import { HistoryTimeline } from '../../../../packages/history/src/react'
import type { HistoryListPage, HistoryVersionInfo } from '../../../../packages/history/src/types'
import { DsButton, DsChip, DsSegment } from '../design-system/primitives'
import '../design-system/live-tools.css'
import { createPlaygroundHistory, playgroundHistoryAuthors, type HistorySnapshot } from '../historyLifecycle'
import { createPlaygroundHistoryController } from '../historyTimeline'

const BEFORE_GRID = 'Name,Plan,Actual\nNorth,120,118\nSouth,95,101\nWest,140,139'
const AFTER_GRID = 'Name,Plan,Actual\nNorth,120,126\nSouth,105,101\nWest,140,151\nEast,80,84'
const BEFORE_TEXT = 'Quarterly plan\nRevenue held close to target.\nThe West team added two accounts.\nNext review: Friday.'
const AFTER_TEXT = 'Quarterly plan\nRevenue finished above target.\nThe West team added two accounts.\nThe East team opened a new region.\nNext review: Friday.'

function parseGrid(value: string): string[][] {
  return value.split('\n').map((row) => row.split(',').map((cell) => cell.trim()))
}

export default function HistoryPage({ fixedFormat }: { fixedFormat?: 'sheets' | 'docs' } = {}) {
  const id = useId()
  const [selectedMode, setMode] = useState<'grid' | 'text'>('grid')
  const mode = fixedFormat === 'docs' ? 'text' : fixedFormat === 'sheets' ? 'grid' : selectedMode
  const [beforeGrid, setBeforeGrid] = useState(BEFORE_GRID)
  const [afterGrid, setAfterGrid] = useState(AFTER_GRID)
  const [beforeText, setBeforeText] = useState(BEFORE_TEXT)
  const [afterText, setAfterText] = useState(AFTER_TEXT)
  const [versions, setVersions] = useState<HistoryVersionInfo[]>([])
  const [selectedId, setSelectedId] = useState<string | null>(null)
  const [preview, setPreview] = useState<HistorySnapshot>('')
  const [status, setStatus] = useState('Host-backed versions stay isolated from the live editors.')
  const [error, setError] = useState<string | null>(null)
  const managerRef = useRef<HistoryManager<HistorySnapshot> | null>(null)
  const liveRef = useRef({ snapshot: AFTER_GRID, contentType: 'text/csv' })
  const [timeline, setTimeline] = useState<{
    controller: HistoryCommandController<HistorySnapshot>
    page: HistoryListPage
  } | null>(null)
  const authors = playgroundHistoryAuthors()
  const gridDiff = useMemo(() => diffGrids(parseGrid(beforeGrid), parseGrid(afterGrid)), [afterGrid, beforeGrid])
  const textDiff = useMemo(() => diffText(beforeText, afterText), [afterText, beforeText])
  const live = mode === 'grid' ? afterGrid : afterText
  const contentType = mode === 'grid' ? 'text/csv' : 'text/plain'
  liveRef.current = { snapshot: live, contentType }

  const refresh = async (manager: HistoryManager<HistorySnapshot>, preferred?: string) => {
    const page = await manager.listVersions()
    setVersions(page.versions)
    setTimeline((current) => current ? { ...current, page } : current)
    const nextId = preferred && page.versions.some((version) => version.id === preferred) ? preferred : page.versions[0]?.id ?? null
    setSelectedId(nextId)
    if (nextId) {
      const loaded = await manager.loadPreview(nextId)
      setPreview(loaded.snapshot)
    } else {
      setPreview('')
    }
  }

  useEffect(() => {
    const { manager } = createPlaygroundHistory(mode === 'text' ? [
      { snapshot: BEFORE_TEXT, contentType: 'text/plain', description: 'Seeded before document' },
      { snapshot: AFTER_TEXT, contentType: 'text/plain', description: 'Seeded after document' },
    ] : [
      { snapshot: BEFORE_GRID, contentType: 'text/csv', description: 'Seeded before grid' },
      { snapshot: AFTER_GRID, contentType: 'text/csv', description: 'Seeded after grid' },
    ])
    managerRef.current = manager
    const controller = createPlaygroundHistoryController(
      manager,
      () => liveRef.current,
      (loaded) => {
        setSelectedId(loaded.version.id)
        setPreview(loaded.snapshot)
        setStatus(`Loaded isolated preview of ${loaded.version.id}.`)
      },
      (snapshot) => {
        if (liveRef.current.contentType === 'text/csv') setAfterGrid(snapshot)
        else setAfterText(snapshot)
      },
    )
    void manager.listVersions().then((page) => {
      setTimeline({ controller, page })
      return refresh(manager)
    }).catch((reason: unknown) => setError(reason instanceof Error ? reason.message : String(reason)))
  }, [])

  const run = async (work: (manager: HistoryManager<HistorySnapshot>) => Promise<string>) => {
    const manager = managerRef.current
    if (!manager) return
    setError(null)
    try {
      const message = await work(manager)
      setStatus(message)
    } catch (reason: unknown) {
      setError(reason instanceof Error ? reason.message : String(reason))
    }
  }

  return (
    <div className="ds">
    <section className="tool-page" data-demo-surface="history" aria-label="Document history workbench">
      <div className="tool-page__controls ds-workstrip" role="toolbar" aria-label="History controls">
        {fixedFormat === undefined && <DsSegment
          label="Diff mode"
          value={mode}
          onChange={(id) => setMode(id as 'grid' | 'text')}
          options={[{ id: 'grid', label: 'Spreadsheet' }, { id: 'text', label: 'Document' }]}
        />}
        <DsButton variant="outlined" className="workbench-button" onClick={() => {
          if (mode === 'grid') { setBeforeGrid(BEFORE_GRID); setAfterGrid(AFTER_GRID) } else { setBeforeText(BEFORE_TEXT); setAfterText(AFTER_TEXT) }
        }}>Reset example</DsButton>
        <DsButton variant="outlined" className="workbench-button" onClick={() => void run(async (manager) => {
          const created = await manager.capture({ snapshot: live, author: authors.user, contentType, description: 'User capture' })
          await refresh(manager, created.id)
          return `Captured ${created.id} as a user save.`
        })}>Capture after</DsButton>
        <DsButton variant="outlined" className="workbench-button" onClick={() => void run(async (manager) => {
          const created = await manager.capture({ snapshot: live, author: authors.agent, reason: 'agent-delivery', contentType, description: 'Agent capture' })
          await refresh(manager, created.id)
          return `Captured ${created.id} as an agent delivery.`
        })}>Capture as agent</DsButton>
        <DsButton variant="outlined" className="workbench-button" disabled={!selectedId} onClick={() => void run(async (manager) => {
          const restored = await manager.restore({ sourceVersionId: selectedId!, author: authors.user })
          const loaded = await manager.loadPreview(restored.id)
          if (mode === 'grid') setAfterGrid(loaded.snapshot)
          else setAfterText(loaded.snapshot)
          await refresh(manager, restored.id)
          return `Restored ${selectedId} as new version ${restored.id}.`
        })}>Restore selected</DsButton>
        <span className="tool-page__status ds-muted">{status}</span>
      </div>
      {error && <p className="tool-error" role="alert">{error}</p>}

      <div className="tool-page__grid tool-page__grid--three ds-split ds-split--three">
        <section className="tool-card ds-split-main" aria-labelledby={`${id}-history-before-title`}>
          <span className="ds-eyebrow tool-eyebrow">Before</span>
          <h2 id={`${id}-history-before-title`}>Before</h2>
          <textarea className="tool-editor ds-outline" value={mode === 'grid' ? beforeGrid : beforeText} onChange={(event) => mode === 'grid' ? setBeforeGrid(event.target.value) : setBeforeText(event.target.value)} aria-label="Before content" />
        </section>
        <section className="tool-card ds-split-main" aria-labelledby={`${id}-history-after-title`}>
          <span className="ds-eyebrow tool-eyebrow">After</span>
          <h2 id={`${id}-history-after-title`}>After</h2>
          <textarea className="tool-editor ds-outline" value={mode === 'grid' ? afterGrid : afterText} onChange={(event) => mode === 'grid' ? setAfterGrid(event.target.value) : setAfterText(event.target.value)} aria-label="After content" />
        </section>
        <section className="tool-card tool-card--hero ds-split-main" aria-labelledby={`${id}-history-diff-title`}>
          <span className="ds-eyebrow tool-eyebrow">Changes</span>
          <div className="ds-row">
            <h2 id={`${id}-history-diff-title`}>Changes</h2>
            <DsChip>{mode === 'grid' ? gridDiff.changeCount : textDiff.added + textDiff.removed}</DsChip>
          </div>
          {mode === 'grid' ? (
            <div className="diff-list">
              {gridDiff.changes.map((change) => <div className="diff-row ds-diff-row" key={change.address}><strong>{change.address}</strong><del>{change.from || 'blank'}</del><span aria-hidden="true">→</span><ins>{change.to || 'blank'}</ins></div>)}
              {gridDiff.changes.length === 0 && <p className="tool-empty">The grids are identical.</p>}
              <p className="tool-note ds-muted">Rows {gridDiff.rowDelta >= 0 ? '+' : ''}{gridDiff.rowDelta} · columns {gridDiff.colDelta >= 0 ? '+' : ''}{gridDiff.colDelta}</p>
            </div>
          ) : (
            <div className="diff-text">
              {textDiff.ops.map((op, index) => <div className={`diff-text__line diff-text__line--${op.kind}`} key={`${index}-${op.line}`}><span>{op.kind === 'add' ? '+' : op.kind === 'remove' ? '−' : ' '}</span>{op.line || '\u00a0'}</div>)}
            </div>
          )}
        </section>
      </div>

      <div className="tool-page__grid ds-split ds-split--wide">
        <section className="tool-card ds-split-main" aria-labelledby={`${id}-history-versions-title`}>
          <span className="ds-eyebrow tool-eyebrow">Version timeline</span>
          <div className="ds-row">
            <h2 id={`${id}-history-versions-title`}>Version timeline</h2>
            <DsChip>{versions.length}</DsChip>
          </div>
          {timeline ? (
            <HistoryTimeline
              key={timeline.page.versions[0]?.id ?? 'empty'}
              controller={timeline.controller}
              initialPage={timeline.page}
              author={authors.user}
              title="Version history"
            />
          ) : (
            <p className="tool-empty">Loading saved versions…</p>
          )}
        </section>
        <section className="tool-card ds-split-side" aria-labelledby={`${id}-history-preview-title`}>
          <span className="ds-eyebrow tool-eyebrow">Isolated preview</span>
          <h2 id={`${id}-history-preview-title`}>{selectedId ?? 'No version'}</h2>
          <textarea className="tool-editor ds-outline" readOnly value={preview} aria-label="Isolated history preview" />
          <p className="tool-note ds-muted">The package timeline mounts here. Previews stay cloned; restore writes a new immutable version.</p>
        </section>
      </div>
    </section>
    </div>
  )
}

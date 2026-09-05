import { useCallback, useEffect, useMemo, useRef, useState, type CSSProperties, type KeyboardEvent, type ReactElement } from 'react'
import type { HistoryCommandController } from './commands'
import type {
  HistoryAuthor,
  HistoryAuthorKind,
  HistoryChangeReason,
  HistoryListPage,
  HistoryVersionFilter,
  HistoryVersionInfo,
} from './types'

export interface HistoryTimelineProps<TSnapshot> {
  controller: HistoryCommandController<TSnapshot>
  /** Supplying a page avoids an empty first render and supports server rendering. */
  initialPage?: HistoryListPage
  author?: HistoryAuthor
  canCapture?: boolean
  canRestore?: boolean
  pageSize?: number
  initialFilter?: HistoryVersionFilter
  title?: string
  formatTime?: (createdAt: number) => string
  formatSize?: (bytes: number) => string
  onError?: (error: unknown) => void
}

type BusyAction = 'list' | 'preview' | 'capture' | 'restore' | null

const REASONS: Array<{ value: '' | HistoryChangeReason; label: string }> = [
  { value: '', label: 'All changes' },
  { value: 'save', label: 'Saved' },
  { value: 'agent-delivery', label: 'Agent delivery' },
  { value: 'import', label: 'Imported' },
  { value: 'restore', label: 'Restored' },
]
const AUTHORS: Array<{ value: '' | HistoryAuthorKind; label: string }> = [
  { value: '', label: 'Everyone' },
  { value: 'user', label: 'People' },
  { value: 'agent', label: 'Agents' },
  { value: 'system', label: 'System' },
]
const EMPTY_FILTER: HistoryVersionFilter = {}

export function HistoryTimeline<TSnapshot>({
  controller,
  initialPage,
  author,
  canCapture = true,
  canRestore = true,
  pageSize = 30,
  initialFilter = EMPTY_FILTER,
  title = 'Version history',
  formatTime = defaultTime,
  formatSize = defaultSize,
  onError,
}: HistoryTimelineProps<TSnapshot>) {
  const [versions, setVersions] = useState(() => initialPage?.versions ?? [])
  const [nextCursor, setNextCursor] = useState(initialPage?.nextCursor)
  const [reason, setReason] = useState<'' | HistoryChangeReason>(initialFilter.reasons?.length === 1 ? initialFilter.reasons[0]! : '')
  const [authorKind, setAuthorKind] = useState<'' | HistoryAuthorKind>(initialFilter.authorKinds?.length === 1 ? initialFilter.authorKinds[0]! : '')
  const [busy, setBusy] = useState<BusyAction>(null)
  const [error, setError] = useState<string | null>(null)
  const [previewId, setPreviewId] = useState<string | null>(null)
  const [restoreTarget, setRestoreTarget] = useState<HistoryVersionInfo | null>(null)
  const confirmRef = useRef<HTMLDivElement>(null)

  const filter = useMemo<HistoryVersionFilter>(() => ({
    ...initialFilter,
    ...(reason ? { reasons: [reason] } : { reasons: undefined }),
    ...(authorKind ? { authorKinds: [authorKind] } : { authorKinds: undefined }),
  }), [authorKind, initialFilter, reason])

  const report = useCallback((caught: unknown) => {
    setError(caught instanceof Error ? caught.message : String(caught))
    onError?.(caught)
  }, [onError])

  const load = useCallback(async (cursor?: string) => {
    setBusy('list'); setError(null)
    try {
      const page = await controller.list({ limit: pageSize, filter, ...(cursor ? { cursor } : {}) })
      setVersions((current) => cursor ? mergeVersions(current, page.versions) : page.versions)
      setNextCursor(page.nextCursor)
    } catch (caught) { report(caught) } finally { setBusy(null) }
  }, [controller, filter, pageSize, report])

  useEffect(() => {
    if (initialPage && reason === (initialFilter.reasons?.length === 1 ? initialFilter.reasons[0] : '')
      && authorKind === (initialFilter.authorKinds?.length === 1 ? initialFilter.authorKinds[0] : '')) return
    const abort = new AbortController()
    setBusy('list'); setError(null)
    void controller.list({ limit: pageSize, filter, signal: abort.signal }).then((page) => {
      setVersions(page.versions); setNextCursor(page.nextCursor)
    }).catch((caught) => { if (!abort.signal.aborted) report(caught) }).finally(() => { if (!abort.signal.aborted) setBusy(null) })
    return () => abort.abort()
  }, [authorKind, controller, filter, initialFilter, initialPage, pageSize, reason, report])

  const preview = async (versionId: string) => {
    setBusy('preview'); setError(null)
    try { await controller.preview(versionId); setPreviewId(versionId) } catch (caught) { report(caught) } finally { setBusy(null) }
  }
  const capture = async () => {
    if (!author) return
    setBusy('capture'); setError(null)
    try { await controller.capture({ author }); await load() } catch (caught) { report(caught); setBusy(null) }
  }
  const restore = async () => {
    if (!author || !restoreTarget) return
    setBusy('restore'); setError(null)
    try {
      await controller.restore({ sourceVersionId: restoreTarget.id, author })
      setRestoreTarget(null)
      await load()
    } catch (caught) { report(caught); setBusy(null) }
  }

  const disabled = busy !== null
  return (
    <aside className="ioc-history" style={styles.shell} aria-labelledby="ioc-history-title" aria-busy={disabled}>
      <header style={styles.header}>
        <div>
          <h2 id="ioc-history-title" style={styles.title}>{title}</h2>
          <p style={styles.subtitle}>Browse saved states without changing the open workbook.</p>
        </div>
        {canCapture && author ? <button type="button" style={styles.primaryButton} disabled={disabled} onClick={() => void capture()}>Save version</button> : null}
      </header>

      <div style={styles.filters} role="group" aria-label="Filter version history">
        <label style={styles.field}>Change
          <select aria-label="Change type" value={reason} disabled={disabled} onChange={(event) => setReason(event.target.value as '' | HistoryChangeReason)} style={styles.select}>
            {REASONS.map((option) => <option key={option.value} value={option.value}>{option.label}</option>)}
          </select>
        </label>
        <label style={styles.field}>Author
          <select aria-label="Author type" value={authorKind} disabled={disabled} onChange={(event) => setAuthorKind(event.target.value as '' | HistoryAuthorKind)} style={styles.select}>
            {AUTHORS.map((option) => <option key={option.value} value={option.value}>{option.label}</option>)}
          </select>
        </label>
      </div>

      {error ? <div role="alert" style={styles.error}><strong>History could not be updated.</strong><span>{error}</span><button type="button" onClick={() => void load()} style={styles.textButton}>Try again</button></div> : null}
      {busy === 'list' && versions.length === 0 ? <p role="status" style={styles.empty}>Loading saved versions…</p> : null}
      {!busy && versions.length === 0 && !error ? <p style={styles.empty}>No saved versions match these filters.</p> : null}

      <ol style={styles.timeline} aria-label="Saved versions">
        {versions.map((version, index) => (
          <li key={version.id} style={styles.item} aria-current={previewId === version.id ? 'true' : undefined}>
            <span style={{ ...styles.railDot, ...(previewId === version.id ? styles.railDotActive : {}) }} aria-hidden="true" />
            <div style={styles.itemTop}>
              <div>
                <strong style={styles.author}>{version.author.displayName || version.author.id}</strong>
                <span style={styles.reason}>{reasonLabel(version.reason)}{index === 0 ? ' · Latest' : ''}</span>
              </div>
              <time dateTime={new Date(version.createdAt).toISOString()} style={styles.time}>{formatTime(version.createdAt)}</time>
            </div>
            <p style={styles.meta}>{formatSize(version.size)}{version.description ? ` · ${version.description}` : ''}</p>
            {version.sourceVersionId ? <p style={styles.lineage}>Restored from {version.sourceVersionId}</p> : null}
            <div style={styles.actions}>
              <button type="button" disabled={disabled} onClick={() => void preview(version.id)} style={styles.secondaryButton}>{previewId === version.id ? 'Previewing' : 'Preview'}</button>
              {canRestore && author && index !== 0 ? <button type="button" disabled={disabled} onClick={() => setRestoreTarget(version)} style={styles.textButton}>Restore</button> : null}
            </div>
          </li>
        ))}
      </ol>

      {nextCursor ? <button type="button" disabled={disabled} onClick={() => void load(nextCursor)} style={styles.loadButton}>Load older versions</button> : null}

      {restoreTarget ? (
        <div
          ref={confirmRef}
          role="alertdialog"
          aria-modal="true"
          aria-labelledby="ioc-history-confirm-title"
          tabIndex={-1}
          style={styles.confirm}
          onKeyDown={(event) => trapHistoryDialogFocus(event, confirmRef.current, () => setRestoreTarget(null))}
        >
          <strong id="ioc-history-confirm-title">Restore this version?</strong>
          <p style={styles.confirmCopy}>A new saved version will be created. Existing history stays intact.</p>
          <div style={styles.actions}>
            <button type="button" disabled={disabled} onClick={() => void restore()} style={styles.warningButton}>Restore version</button>
            <button type="button" disabled={disabled} onClick={() => setRestoreTarget(null)} style={styles.textButton}>Cancel</button>
          </div>
        </div>
      ) : null}
    </aside>
  )
}

export interface HistoryMountRoot {
  render(node: ReactElement): void
  unmount(): void
}

export function createHistorySidebarElement(doc: Pick<Document, 'createElement'>): HTMLElement {
  const aside = doc.createElement('aside')
  aside.className = 'ioc-history-sidebar'
  aside.setAttribute('data-ioc-history-sidebar', 'true')
  return aside
}

/** Mount the durable timeline into a host sidebar node. Callers pass a React
 * root factory so this module stays testable without a browser document. */
export function mountHistoryTimeline<TSnapshot>(
  target: Element,
  props: HistoryTimelineProps<TSnapshot>,
  createRoot: (container: Element) => HistoryMountRoot,
): { unmount(): void } {
  const root = createRoot(target)
  root.render(<HistoryTimeline {...props} />)
  return { unmount() { root.unmount() } }
}

export function attachHistorySidebar<TSnapshot>(
  parent: Element,
  props: HistoryTimelineProps<TSnapshot>,
  createRoot: (container: Element) => HistoryMountRoot,
  doc: Pick<Document, 'createElement'>,
): { element: HTMLElement; unmount(): void } {
  const element = createHistorySidebarElement(doc)
  parent.appendChild(element)
  const mounted = mountHistoryTimeline(element, props, createRoot)
  return {
    element,
    unmount() {
      mounted.unmount()
      parent.removeChild(element)
    },
  }
}

export function trapHistoryDialogFocus(
  event: Pick<KeyboardEvent<HTMLElement>, 'key' | 'shiftKey' | 'preventDefault'>,
  root: { querySelectorAll: (selector: string) => ArrayLike<HTMLElement> } | null,
  onCancel: () => void,
  active: HTMLElement | null = typeof document === 'undefined' ? null : document.activeElement as HTMLElement | null,
): void {
  if (event.key === 'Escape') {
    event.preventDefault()
    onCancel()
    return
  }
  if (event.key !== 'Tab' || !root) return
  const focusable = Array.from(root.querySelectorAll('button:not([disabled])'))
  if (focusable.length === 0) return
  const first = focusable[0]!
  const last = focusable.at(-1)!
  if (event.shiftKey && (active === first || active === null)) {
    event.preventDefault()
    last.focus()
  } else if (!event.shiftKey && active === last) {
    event.preventDefault()
    first.focus()
  }
}

export function defaultTime(createdAt: number): string {
  return new Intl.DateTimeFormat(undefined, { dateStyle: 'medium', timeStyle: 'short' }).format(new Date(createdAt))
}

export function defaultSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`
  if (bytes < 1024 ** 2) return `${(bytes / 1024).toFixed(bytes < 10 * 1024 ? 1 : 0)} KB`
  return `${(bytes / 1024 ** 2).toFixed(bytes < 10 * 1024 ** 2 ? 1 : 0)} MB`
}

function mergeVersions(current: HistoryVersionInfo[], older: HistoryVersionInfo[]): HistoryVersionInfo[] {
  const ids = new Set(current.map(({ id }) => id))
  return [...current, ...older.filter(({ id }) => !ids.has(id))]
}

function reasonLabel(reason: HistoryChangeReason): string {
  return REASONS.find((entry) => entry.value === reason)?.label ?? reason
}

const styles: Record<string, CSSProperties> = {
  shell: { '--ioc-history-ink': '#172033', '--ioc-history-muted': '#64748b', '--ioc-history-rule': '#d6deea', '--ioc-history-blue': '#3157c8', background: '#fbfcfe', color: 'var(--ioc-history-ink)', boxSizing: 'border-box', fontFamily: 'inherit', maxWidth: 420, minWidth: 280, padding: '18px 18px 22px', position: 'relative', width: '100%' } as CSSProperties,
  header: { alignItems: 'flex-start', display: 'flex', gap: 16, justifyContent: 'space-between' },
  title: { fontSize: 19, fontWeight: 680, letterSpacing: '-0.015em', lineHeight: 1.2, margin: 0 },
  subtitle: { color: 'var(--ioc-history-muted)', fontSize: 12, lineHeight: 1.45, margin: '5px 0 0', maxWidth: 240 },
  filters: { borderBottom: '1px solid var(--ioc-history-rule)', display: 'grid', gap: 10, gridTemplateColumns: '1fr 1fr', marginTop: 18, paddingBottom: 14 },
  field: { color: 'var(--ioc-history-muted)', display: 'grid', fontSize: 11, fontWeight: 650, gap: 5 },
  select: { background: '#fff', border: '1px solid var(--ioc-history-rule)', borderRadius: 5, color: 'var(--ioc-history-ink)', font: 'inherit', fontSize: 12, minHeight: 31, padding: '4px 7px' },
  timeline: { listStyle: 'none', margin: 0, padding: '3px 0 0 19px' },
  item: { borderLeft: '1px solid var(--ioc-history-rule)', padding: '15px 0 17px 19px', position: 'relative' },
  railDot: { background: '#fbfcfe', border: '2px solid #91a0b5', borderRadius: '50%', boxSizing: 'border-box', height: 11, left: -6, position: 'absolute', top: 19, width: 11 },
  railDotActive: { background: 'var(--ioc-history-blue)', borderColor: 'var(--ioc-history-blue)', boxShadow: '0 0 0 3px #dce5ff' },
  itemTop: { alignItems: 'flex-start', display: 'flex', gap: 10, justifyContent: 'space-between' },
  author: { display: 'block', fontSize: 13, fontWeight: 680, lineHeight: 1.35 },
  reason: { color: 'var(--ioc-history-muted)', display: 'block', fontSize: 11, lineHeight: 1.4, marginTop: 1 },
  time: { color: 'var(--ioc-history-muted)', flex: '0 0 auto', fontSize: 10, lineHeight: 1.4, maxWidth: 100, textAlign: 'right' },
  meta: { color: 'var(--ioc-history-muted)', fontSize: 11, lineHeight: 1.4, margin: '7px 0 0' },
  lineage: { background: '#eef3ff', borderLeft: '2px solid var(--ioc-history-blue)', color: '#294494', fontSize: 11, margin: '8px 0 0', padding: '5px 7px' },
  actions: { alignItems: 'center', display: 'flex', gap: 8, marginTop: 10 },
  primaryButton: { background: 'var(--ioc-history-blue)', border: 0, borderRadius: 5, color: '#fff', cursor: 'pointer', flex: '0 0 auto', font: 'inherit', fontSize: 12, fontWeight: 680, minHeight: 32, padding: '6px 10px' },
  secondaryButton: { background: '#fff', border: '1px solid var(--ioc-history-rule)', borderRadius: 5, color: 'var(--ioc-history-ink)', cursor: 'pointer', font: 'inherit', fontSize: 11, fontWeight: 650, minHeight: 28, padding: '4px 9px' },
  textButton: { background: 'transparent', border: 0, color: 'var(--ioc-history-blue)', cursor: 'pointer', font: 'inherit', fontSize: 11, fontWeight: 680, padding: '5px 3px' },
  loadButton: { background: '#fff', border: '1px solid var(--ioc-history-rule)', borderRadius: 5, color: 'var(--ioc-history-ink)', cursor: 'pointer', font: 'inherit', fontSize: 12, fontWeight: 650, minHeight: 34, width: '100%' },
  empty: { color: 'var(--ioc-history-muted)', fontSize: 12, lineHeight: 1.5, padding: '24px 4px', textAlign: 'center' },
  error: { background: '#fff2ed', borderLeft: '3px solid #bc3d17', color: '#70240f', display: 'grid', fontSize: 12, gap: 4, lineHeight: 1.4, marginTop: 14, padding: '10px 12px' },
  confirm: { background: '#fff8e8', border: '1px solid #e4bd69', bottom: 12, boxShadow: '0 12px 32px rgba(30, 41, 59, .18)', left: 12, padding: 14, position: 'sticky', right: 12, zIndex: 1 },
  confirmCopy: { color: '#6c4b17', fontSize: 12, lineHeight: 1.45, margin: '5px 0 0' },
  warningButton: { background: '#9a4d00', border: 0, borderRadius: 5, color: '#fff', cursor: 'pointer', font: 'inherit', fontSize: 11, fontWeight: 680, minHeight: 30, padding: '5px 10px' },
}

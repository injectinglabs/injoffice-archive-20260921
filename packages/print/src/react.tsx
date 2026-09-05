import {
  useCallback,
  useEffect,
  useId,
  useRef,
  useState,
  type CSSProperties,
  type FocusEvent,
  type KeyboardEvent,
  type ReactNode,
} from 'react'
import type { PrintConfigurationCommandController } from './commands'
import type {
  PrintPreviewDocument,
  PrintPreviewManager,
  PrintPreviewPage,
  PrintPreviewSession,
  PrintPageSelection,
} from './preview'
import type {
  PrintAlign,
  PrintArea,
  PrintHeaderFooter,
  PrintLayoutConfig,
  PrintPaperMargin,
  PrintPaperSize,
  PrintRenderConfig,
  PrintWatermark,
} from './types'

export interface PrintUIControlVisibility {
  area?: boolean
  paper?: boolean
  orientation?: boolean
  scale?: boolean
  margins?: boolean
  repeatedTitles?: boolean
  content?: boolean
  alignment?: boolean
  headerFooter?: boolean
  watermark?: boolean
  pageSelection?: boolean
}

export interface PrintUIConfig {
  controls?: PrintUIControlVisibility
  paperSizes?: ReadonlyArray<PrintPaperSize>
  areas?: ReadonlyArray<PrintArea>
  showPreview?: boolean
  showPrintAction?: boolean
  showCloseAction?: boolean
  autoPreview?: boolean
  modal?: boolean
  title?: string
}

export interface PrintWorkspaceProps<TPayload> {
  controller: PrintConfigurationCommandController
  previewManager: PrintPreviewManager<TPayload>
  renderPage(page: Readonly<PrintPreviewPage<TPayload>>): ReactNode
  config?: PrintUIConfig
  initialDocument?: PrintPreviewDocument<TPayload>
  createSessionId?: () => string
  className?: string
  style?: CSSProperties
  onClose?: () => void
  onPrinted?: () => void
  onError?: (error: unknown) => void
}

type BusyState = 'saving' | 'previewing' | 'printing' | null

const PAPERS: ReadonlyArray<PrintPaperSize> = ['Letter', 'Tabloid', 'Legal', 'Statement', 'Executive', 'Folio', 'A3', 'A4', 'A5', 'B4', 'B5', 'Custom']
const AREAS: ReadonlyArray<PrintArea> = ['CurrentSheet', 'Workbook', 'CurrentSelection', 'AllSelection']
const MARGINS: ReadonlyArray<PrintPaperMargin> = ['Normal', 'Narrow', 'Wide', 'None', 'Custom']
const ALIGNMENTS: ReadonlyArray<PrintAlign> = ['Start', 'Middle', 'End']
const TOKENS: ReadonlyArray<{ value: PrintHeaderFooter; label: string }> = [
  { value: 'PageSize', label: 'Page number and count' },
  { value: 'WorkbookTitle', label: 'Workbook title' },
  { value: 'WorksheetTitle', label: 'Worksheet title' },
  { value: 'Date', label: 'Date' },
  { value: 'Time', label: 'Time' },
]

let sessionSequence = 0
const nextSessionId = () => `print-ui-${Date.now()}-${++sessionSequence}`
const titleCase = (value: string) => value.replace(/([a-z])([A-Z])/g, '$1 $2')

export function canonicalPageRange(start: number, end: number, totalPages?: number): PrintPageSelection | null {
  if (!Number.isSafeInteger(start) || !Number.isSafeInteger(end) || start < 1 || end < start) return null
  if (totalPages !== undefined && end > totalPages) return null
  return { kind: 'range', startPage: start, endPage: end }
}

export function PrintWorkspace<TPayload>({
  controller,
  previewManager,
  renderPage,
  config = {},
  initialDocument,
  createSessionId = nextSessionId,
  className,
  style,
  onClose,
  onPrinted,
  onError,
}: PrintWorkspaceProps<TPayload>) {
  const headingId = `ioc-print-${useId().replace(/:/g, '')}`
  const root = useRef<HTMLElement>(null)
  const activeSession = useRef<PrintPreviewSession<TPayload> | null>(null)
  const [snapshot, setSnapshot] = useState(() => controller.snapshot())
  const [document, setDocument] = useState<PrintPreviewDocument<TPayload> | null>(initialDocument ?? null)
  const [selectionKind, setSelectionKind] = useState<'all' | 'range'>(() => initialDocument?.selection.kind === 'range' ? 'range' : 'all')
  const [rangeStart, setRangeStart] = useState(() => initialDocument?.selection.kind === 'range' ? initialDocument.selection.startPage : 1)
  const [rangeEnd, setRangeEnd] = useState(() => initialDocument?.selection.kind === 'range' ? initialDocument.selection.endPage : 1)
  const [busy, setBusy] = useState<BusyState>(null)
  const [error, setError] = useState<string | null>(null)
  const controls = config.controls ?? {}
  const disabled = busy !== null

  const report = useCallback((caught: unknown) => {
    setError(caught instanceof Error ? caught.message : String(caught))
    onError?.(caught)
  }, [onError])

  const currentSelection = useCallback((): PrintPageSelection | null => {
    if (selectionKind === 'all') return { kind: 'all' }
    return canonicalPageRange(rangeStart, rangeEnd, document?.totalPages)
  }, [document?.totalPages, rangeEnd, rangeStart, selectionKind])

  const refreshPreview = useCallback(async () => {
    const selection = currentSelection()
    if (!selection) {
      report(new RangeError('Enter a page range within the available page count.'))
      return false
    }
    activeSession.current?.cancel('superseded by a newer preview')
    setBusy('previewing'); setError(null)
    let session: PrintPreviewSession<TPayload> | null = null
    try {
      session = previewManager.start(createSessionId(), controller.manager.snapshot(), { selection })
      activeSession.current = session
      const result = await session.result
      if (activeSession.current !== session) return false
      setDocument(result)
      return true
    } catch (caught) {
      if ((!session || activeSession.current === session) && !(caught instanceof Error && caught.name === 'PrintPreviewError' && 'code' in caught && caught.code === 'ABORTED')) report(caught)
      return false
    } finally {
      if (!session || activeSession.current === session) {
        activeSession.current = null
        setBusy(null)
      }
    }
  }, [controller.manager, createSessionId, currentSelection, previewManager, report])

  useEffect(() => {
    controller.manager.openPrintDialog()
    setSnapshot(controller.snapshot())
    const stop = controller.manager.onEvent((event) => {
      if (event.name === 'changed') setSnapshot(controller.snapshot())
    })
    if (config.autoPreview !== false && config.showPreview !== false && !initialDocument) void refreshPreview()
    root.current?.focus()
    return () => {
      activeSession.current?.cancel('print workspace unmounted')
      stop()
    }
    // Preview identity intentionally comes from the first mounted workspace;
    // later setting changes explicitly start replacement sessions.
  }, [config.autoPreview, config.showPreview, controller, initialDocument])

  const commitLayout = async (patch: Partial<PrintLayoutConfig>) => {
    setBusy('saving'); setError(null)
    try {
      if (!await controller.updateLayout(patch)) throw new Error('The print layout could not be saved.')
      setSnapshot(controller.snapshot())
      if (config.showPreview !== false) await refreshPreview()
    } catch (caught) { report(caught) } finally { setBusy(null) }
  }
  const commitRender = async (patch: Partial<PrintRenderConfig>) => {
    setBusy('saving'); setError(null)
    try {
      if (!await controller.updateRender(patch)) throw new Error('The print appearance could not be saved.')
      setSnapshot(controller.snapshot())
      if (config.showPreview !== false) await refreshPreview()
    } catch (caught) { report(caught) } finally { setBusy(null) }
  }
  const close = () => {
    if (controller.manager.closePrintDialog()) onClose?.()
  }
  const print = async () => {
    setBusy('printing'); setError(null)
    try {
      if (!await controller.manager.print()) throw new Error('Printing was canceled.')
      onPrinted?.()
    } catch (caught) { report(caught) } finally { setBusy(null) }
  }
  const trapFocus = (event: KeyboardEvent<HTMLElement>) => {
    if (event.key === 'Escape' && config.showCloseAction !== false && !disabled) { event.preventDefault(); close(); return }
    if (event.key !== 'Tab' || config.modal === false || !root.current) return
    const focusable = [...root.current.querySelectorAll<HTMLElement>('button:not([disabled]), input:not([disabled]), select:not([disabled]), [tabindex]:not([tabindex="-1"])')]
    if (!focusable.length) return
    const first = focusable[0]!
    const last = focusable.at(-1)!
    if (event.shiftKey && (documentActiveElement() === first || documentActiveElement() === root.current)) { event.preventDefault(); last.focus() }
    else if (!event.shiftKey && documentActiveElement() === last) { event.preventDefault(); first.focus() }
  }

  const layout = snapshot.layout
  const render = snapshot.render
  return (
    <section
      ref={root}
      role={config.modal === false ? 'region' : 'dialog'}
      aria-modal={config.modal === false ? undefined : true}
      aria-labelledby={headingId}
      aria-busy={disabled}
      tabIndex={-1}
      onKeyDown={trapFocus}
      className={`ioc-print-workspace${className ? ` ${className}` : ''}`}
      style={style}
    >
      <style>{PRINT_WORKSPACE_CSS}</style>
      <header className="ioc-print-header">
        <div>
          <h2 id={headingId}>{config.title ?? 'Print'}</h2>
          <p>Choose what goes on paper, then check every page before printing.</p>
        </div>
        <div className="ioc-print-actions">
          {config.showCloseAction !== false ? <button type="button" className="ioc-print-button secondary" disabled={disabled} onClick={close}>Close</button> : null}
          {config.showPrintAction !== false ? <button type="button" className="ioc-print-button primary" disabled={disabled} onClick={() => void print()}>Print</button> : null}
        </div>
      </header>

      {config.showPreview === false && error ? <div className="ioc-print-error" role="alert"><strong>Print settings could not be updated.</strong><span>{error}</span></div> : null}

      <div className="ioc-print-body">
        <aside className="ioc-print-controls" aria-label="Print settings">
          {controls.area !== false ? <Field label="Print">
            <select aria-label="Print area" disabled={disabled} value={layout.area} onChange={(event) => void commitLayout({ area: event.target.value as PrintArea })}>
              {(config.areas ?? AREAS).map((area) => <option key={area} value={area}>{titleCase(area)}</option>)}
            </select>
          </Field> : null}

          <div className="ioc-print-control-grid">
            {controls.paper !== false ? <Field label="Paper size">
              <select aria-label="Paper size" disabled={disabled} value={layout.paperSize} onChange={(event) => {
                const paperSize = event.target.value as PrintPaperSize
                void commitLayout({ paperSize, ...(paperSize === 'Custom' ? { pageSizeCustom: layout.pageSizeCustom ?? { width: 612, height: 792 } } : {}) })
              }}>
                {(config.paperSizes ?? PAPERS).map((paper) => <option key={paper} value={paper}>{paper}</option>)}
              </select>
            </Field> : null}
            {controls.orientation !== false ? <Field label="Orientation">
              <select aria-label="Orientation" disabled={disabled} value={layout.direction} onChange={(event) => void commitLayout({ direction: event.target.value as PrintLayoutConfig['direction'] })}>
                <option value="Portrait">Portrait</option><option value="Landscape">Landscape</option>
              </select>
            </Field> : null}
          </div>

          {controls.paper !== false && layout.paperSize === 'Custom' ? <fieldset className="ioc-print-inset"><legend>Custom paper, points</legend><div className="ioc-print-control-grid">
            <NumberField label="Width" value={layout.pageSizeCustom?.width ?? 612} min={1} disabled={disabled} onCommit={(width) => width !== null && void commitLayout({ pageSizeCustom: { width, height: layout.pageSizeCustom?.height ?? 792 } })} />
            <NumberField label="Height" value={layout.pageSizeCustom?.height ?? 792} min={1} disabled={disabled} onCommit={(height) => height !== null && void commitLayout({ pageSizeCustom: { width: layout.pageSizeCustom?.width ?? 612, height } })} />
          </div></fieldset> : null}

          {controls.scale !== false ? <Field label="Scale">
            <select aria-label="Page scaling" disabled={disabled} value={layout.scale} onChange={(event) => void commitLayout({ scale: event.target.value as PrintLayoutConfig['scale'] })}>
              <option value="Origin">Actual size</option><option value="FitWidth">Fit to width</option><option value="FitHeight">Fit to height</option><option value="FitPage">Fit to page</option><option value="Custom">Custom</option>
            </select>
          </Field> : null}
          {controls.scale !== false && layout.scale === 'Custom' ? <NumberField label="Custom scale, percent" value={layout.customScale} min={10} max={400} disabled={disabled} onCommit={(customScale) => customScale !== null && void commitLayout({ customScale })} /> : null}

          {controls.margins !== false ? <Field label="Margins">
            <select aria-label="Page margins" disabled={disabled} value={layout.margin} onChange={(event) => {
              const margin = event.target.value as PrintPaperMargin
              void commitLayout({ margin, ...(margin === 'Custom' ? { customMargins: layout.customMargins ?? { top: 0.5, right: 0.5, bottom: 0.5, left: 0.5, header: 0.3, footer: 0.3 } } : {}) })
            }}>
              {MARGINS.map((margin) => <option key={margin} value={margin}>{margin}</option>)}
            </select>
          </Field> : null}
          {controls.margins !== false && layout.margin === 'Custom' ? <fieldset className="ioc-print-inset"><legend>Custom margins, inches</legend><div className="ioc-print-margin-grid">
            {(['top', 'right', 'bottom', 'left', 'header', 'footer'] as const).map((edge) => <NumberField key={edge} label={titleCase(edge)} value={layout.customMargins?.[edge] ?? 0.5} min={0} max={20} step={0.1} disabled={disabled} onCommit={(value) => value !== null && void commitLayout({ customMargins: { top: 0.5, right: 0.5, bottom: 0.5, left: 0.5, header: 0.3, footer: 0.3, ...layout.customMargins, [edge]: value } })} />)}
          </div></fieldset> : null}

          {controls.repeatedTitles !== false ? <fieldset><legend>Repeat on every page</legend><div className="ioc-print-control-grid">
            <NumberField label="First row" value={layout.repeatRows ? layout.repeatRows.startRow + 1 : undefined} min={1} allowBlank disabled={disabled} onCommit={(value) => void commitLayout({ repeatRows: value === null ? undefined : { startRow: value - 1, endRow: Math.max(value - 1, layout.repeatRows?.endRow ?? value - 1) } })} />
            <NumberField label="Last row" value={layout.repeatRows ? layout.repeatRows.endRow + 1 : undefined} min={1} allowBlank disabled={disabled} onCommit={(value) => void commitLayout({ repeatRows: value === null ? undefined : { startRow: Math.min(layout.repeatRows?.startRow ?? value - 1, value - 1), endRow: value - 1 } })} />
            <NumberField label="First column" value={layout.repeatColumns ? layout.repeatColumns.startColumn + 1 : undefined} min={1} allowBlank disabled={disabled} onCommit={(value) => void commitLayout({ repeatColumns: value === null ? undefined : { startColumn: value - 1, endColumn: Math.max(value - 1, layout.repeatColumns?.endColumn ?? value - 1) } })} />
            <NumberField label="Last column" value={layout.repeatColumns ? layout.repeatColumns.endColumn + 1 : undefined} min={1} allowBlank disabled={disabled} onCommit={(value) => void commitLayout({ repeatColumns: value === null ? undefined : { startColumn: Math.min(layout.repeatColumns?.startColumn ?? value - 1, value - 1), endColumn: value - 1 } })} />
          </div></fieldset> : null}

          {controls.content !== false ? <fieldset><legend>Page content</legend><div className="ioc-print-checks">
            <Check label="Show gridlines" checked={render.gridlines} disabled={disabled} onChange={(gridlines) => void commitRender({ gridlines })} />
            <Check label="Show row and column headings" checked={render.headings} disabled={disabled} onChange={(headings) => void commitRender({ headings })} />
          </div></fieldset> : null}

          {controls.alignment !== false ? <fieldset><legend>Center on page</legend><div className="ioc-print-control-grid">
            <Field label="Horizontal"><select aria-label="Horizontal page alignment" disabled={disabled} value={render.hAlign} onChange={(event) => void commitRender({ hAlign: event.target.value as PrintAlign })}>{ALIGNMENTS.map((value) => <option key={value} value={value}>{value}</option>)}</select></Field>
            <Field label="Vertical"><select aria-label="Vertical page alignment" disabled={disabled} value={render.vAlign} onChange={(event) => void commitRender({ vAlign: event.target.value as PrintAlign })}>{ALIGNMENTS.map((value) => <option key={value} value={value}>{value}</option>)}</select></Field>
          </div></fieldset> : null}

          {controls.headerFooter !== false ? <fieldset><legend>Headers and footers</legend><div className="ioc-print-checks">
            {TOKENS.map((token) => <Check key={token.value} label={token.label} checked={render.headerFooter.includes(token.value)} disabled={disabled} onChange={(checked) => void commitRender({ headerFooter: checked ? [...render.headerFooter, token.value] : render.headerFooter.filter((value) => value !== token.value) })} />)}
          </div></fieldset> : null}

          {controls.watermark !== false ? <fieldset><legend>Watermark</legend><div className="ioc-print-checks">
            <Check label="Add text watermark" checked={render.watermark?.kind === 'text'} disabled={disabled} onChange={(checked) => void commitRender({ watermark: checked ? { kind: 'text', text: 'DRAFT', opacity: 0.18, rotation: -35, fontSize: 54 } : undefined })} />
            {render.watermark?.kind === 'text' ? <Field label="Watermark text"><input key={render.watermark.text} aria-label="Watermark text" disabled={disabled} defaultValue={render.watermark.text} maxLength={512} onBlur={(event: FocusEvent<HTMLInputElement>) => void commitRender({ watermark: { ...render.watermark as Extract<PrintWatermark, { kind: 'text' }>, text: event.currentTarget.value || 'DRAFT' } })} /></Field> : null}
          </div></fieldset> : null}
        </aside>

        {config.showPreview !== false ? <main className="ioc-print-preview" aria-label="Print preview">
          <div className="ioc-print-preview-toolbar">
            <div aria-live="polite" role="status">{busy === 'previewing' ? 'Rendering preview…' : document ? `${document.totalPages} ${document.totalPages === 1 ? 'page' : 'pages'}` : 'Preview not rendered'}</div>
            <div className="ioc-print-preview-options">
              {controls.pageSelection !== false ? <><label>Pages <select aria-label="Preview pages" disabled={disabled} value={selectionKind} onChange={(event) => setSelectionKind(event.target.value as 'all' | 'range')}><option value="all">All</option><option value="range">Range</option></select></label>
                {selectionKind === 'range' ? <><NumberField label="From" value={rangeStart} min={1} max={document?.totalPages} disabled={disabled} onCommit={(value) => value !== null && setRangeStart(value)} /><NumberField label="To" value={rangeEnd} min={1} max={document?.totalPages} disabled={disabled} onCommit={(value) => value !== null && setRangeEnd(value)} /></> : null}</> : null}
              <button type="button" className="ioc-print-button secondary" disabled={disabled} onClick={() => void refreshPreview()}>Refresh preview</button>
            </div>
          </div>
          {error ? <div className="ioc-print-error" role="alert"><strong>Print could not continue.</strong><span>{error}</span></div> : null}
          {busy === 'previewing' && !document ? <div className="ioc-print-empty" role="status">Preparing pages…</div> : null}
          {!document && busy !== 'previewing' && !error ? <div className="ioc-print-empty">Select Refresh preview to prepare the pages.</div> : null}
          {document?.totalPages === 0 ? <div className="ioc-print-empty">There is nothing to print in this selection.</div> : null}
          <ol className="ioc-print-pages" aria-label="Preview pages">
            {document?.pages.map((page) => <li key={page.number} className="ioc-print-page-item">
              <div className="ioc-print-paper" style={{ aspectRatio: `${page.widthPoints} / ${page.heightPoints}` }}>{renderPage(page)}</div>
              <span>Page {page.number}</span>
            </li>)}
          </ol>
        </main> : null}
      </div>
    </section>
  )
}

function Field({ label, children }: { label: string; children: ReactNode }) {
  return <label className="ioc-print-field"><span>{label}</span>{children}</label>
}

function NumberField({ label, value, min, max, step = 1, allowBlank = false, disabled = false, onCommit }: { label: string; value?: number; min?: number; max?: number; step?: number; allowBlank?: boolean; disabled?: boolean; onCommit(value: number | null): void }) {
  const [draft, setDraft] = useState(value === undefined ? '' : String(value))
  useEffect(() => setDraft(value === undefined ? '' : String(value)), [value])
  return <label className="ioc-print-field"><span>{label}</span><input type="number" value={draft} min={min} max={max} step={step} disabled={disabled} onChange={(event) => setDraft(event.currentTarget.value)} onBlur={() => {
    if (allowBlank && draft === '') { onCommit(null); return }
    const next = Number(draft)
    if (draft !== '' && Number.isFinite(next) && (min === undefined || next >= min) && (max === undefined || next <= max)) onCommit(next)
    else setDraft(value === undefined ? '' : String(value))
  }} /></label>
}

function Check({ label, checked, disabled, onChange }: { label: string; checked: boolean; disabled: boolean; onChange(value: boolean): void }) {
  return <label className="ioc-print-check"><input type="checkbox" checked={checked} disabled={disabled} onChange={(event) => onChange(event.target.checked)} /><span>{label}</span></label>
}

function documentActiveElement(): Element | null {
  return typeof document === 'undefined' ? null : document.activeElement
}

const PRINT_WORKSPACE_CSS = `
.ioc-print-workspace{--ioc-print-ink:#16243a;--ioc-print-muted:#66758a;--ioc-print-rule:#ccd5e0;--ioc-print-panel:#f6f8fb;--ioc-print-paper:#fff;--ioc-print-blue:#1d4f7a;--ioc-print-coral:#d45b48;background:var(--ioc-print-panel);color:var(--ioc-print-ink);display:flex;flex-direction:column;font-family:ui-sans-serif,system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;height:min(860px,100vh);min-height:520px;outline:none;width:100%}
.ioc-print-workspace *{box-sizing:border-box}.ioc-print-header{align-items:center;background:#fff;border-bottom:1px solid var(--ioc-print-rule);display:flex;gap:24px;justify-content:space-between;padding:16px 20px}.ioc-print-header h2{font-size:20px;letter-spacing:-.015em;line-height:1.2;margin:0}.ioc-print-header p{color:var(--ioc-print-muted);font-size:12px;line-height:1.4;margin:3px 0 0}.ioc-print-actions,.ioc-print-preview-options{align-items:center;display:flex;flex-wrap:wrap;gap:8px}.ioc-print-button{border-radius:4px;cursor:pointer;font:inherit;font-size:12px;font-weight:680;min-height:34px;padding:6px 13px}.ioc-print-button:focus-visible,.ioc-print-workspace input:focus-visible,.ioc-print-workspace select:focus-visible{outline:3px solid #9cc9ed;outline-offset:1px}.ioc-print-button:disabled,.ioc-print-workspace input:disabled,.ioc-print-workspace select:disabled{cursor:not-allowed;opacity:.58}.ioc-print-button.primary{background:var(--ioc-print-blue);border:1px solid var(--ioc-print-blue);color:#fff}.ioc-print-button.secondary{background:#fff;border:1px solid #aeb9c7;color:var(--ioc-print-ink)}
.ioc-print-body{display:grid;flex:1;grid-template-columns:minmax(260px,340px) minmax(320px,1fr);min-height:0}.ioc-print-controls{background:#fff;border-right:1px solid var(--ioc-print-rule);display:flex;flex-direction:column;gap:14px;overflow:auto;padding:18px 20px}.ioc-print-control-grid{display:grid;gap:10px;grid-template-columns:1fr 1fr}.ioc-print-margin-grid{display:grid;gap:8px;grid-template-columns:repeat(3,1fr)}.ioc-print-field{color:var(--ioc-print-muted);display:grid;font-size:11px;font-weight:650;gap:5px}.ioc-print-field input,.ioc-print-field select,.ioc-print-preview-options select{background:#fff;border:1px solid #aeb9c7;border-radius:4px;color:var(--ioc-print-ink);font:inherit;font-size:12px;min-height:34px;padding:5px 8px;width:100%}.ioc-print-controls fieldset{border:0;border-top:1px solid #e2e7ee;margin:1px 0 0;padding:14px 0 0}.ioc-print-controls fieldset.ioc-print-inset{background:#f5f7fa;border:0;padding:10px}.ioc-print-controls legend{font-size:12px;font-weight:700;margin:0 0 9px;padding:0}.ioc-print-checks{display:grid;gap:8px}.ioc-print-check{align-items:flex-start;display:flex;font-size:12px;gap:8px;line-height:1.35}.ioc-print-check input{accent-color:var(--ioc-print-blue);margin:1px 0 0}
.ioc-print-preview{background:#dbe2ea;display:flex;flex-direction:column;min-width:0;overflow:hidden}.ioc-print-preview-toolbar{align-items:center;background:#edf1f5;border-bottom:1px solid #b9c5d2;color:var(--ioc-print-muted);display:flex;font-size:12px;justify-content:space-between;min-height:54px;padding:9px 16px}.ioc-print-preview-options label{align-items:center;display:flex;font-size:11px;font-weight:650;gap:5px}.ioc-print-preview-options .ioc-print-field{align-items:center;display:flex}.ioc-print-preview-options .ioc-print-field input{width:64px}.ioc-print-pages{display:grid;gap:30px;grid-template-columns:repeat(auto-fit,minmax(260px,500px));list-style:none;margin:0;overflow:auto;padding:30px}.ioc-print-page-item{align-items:center;display:flex;flex-direction:column;gap:9px}.ioc-print-page-item>span{background:var(--ioc-print-coral);border-radius:12px;color:#fff;font-size:10px;font-weight:700;padding:3px 9px}.ioc-print-paper{background:var(--ioc-print-paper);box-shadow:0 7px 22px rgba(28,44,65,.17);max-height:660px;overflow:hidden;width:100%}.ioc-print-empty{align-items:center;color:#53647a;display:flex;flex:1;font-size:13px;justify-content:center;padding:30px;text-align:center}.ioc-print-error{background:#fff0ed;border-left:4px solid #b84432;color:#73291f;display:grid;font-size:12px;gap:3px;margin:16px;padding:11px 13px}
@media(max-width:720px){.ioc-print-workspace{height:auto;min-height:100vh}.ioc-print-header{align-items:flex-start;flex-direction:column}.ioc-print-body{display:flex;flex-direction:column}.ioc-print-controls{border-bottom:1px solid var(--ioc-print-rule);border-right:0;max-height:none}.ioc-print-preview{min-height:500px}.ioc-print-preview-toolbar{align-items:flex-start;flex-direction:column;gap:9px}.ioc-print-pages{grid-template-columns:minmax(220px,1fr);padding:20px}}
@media(prefers-reduced-motion:reduce){.ioc-print-workspace *{scroll-behavior:auto!important;transition:none!important}}
`

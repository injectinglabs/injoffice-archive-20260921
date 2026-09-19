import { useEffect, useId, useRef, useState, type CSSProperties, type PointerEvent } from 'react'
import { applyPdfCommand, inspectPdf, PdfHistory, findPdfTextMatches, importPdfPages, exportPdfPages, parsePdfPageRange, type PdfCommand, type PdfAnnotationTarget, type PdfSummary } from './pdf-commands'
import { parsePdfRecoveryDraft, type PdfRecoveryDraft } from './pdf-recovery'
import {searchPdfDocument,type PdfSearchResult} from './pdf-search'
import type {PDFDocumentProxy} from 'pdfjs-dist'
import './pdf-editor.css'

type Tool = 'edit-note' | 'replace' | 'view' | 'text' | 'note' | 'highlight' | 'underline' | 'strikeout' | 'rectangle' | 'ellipse' | 'line' | 'arrow' | 'form'
type Placement = { at: [number, number]; end?: [number, number] }
export type PdfEditorProps = {
  registerHistory?: (commands: { undo(): void; redo(): void }) => void
  registerCommit?: (commit: () => Promise<boolean>) => void
  initialRecoveryDraft?: unknown
  onRecoveryDraftChange?: (draft: unknown | null) => void
  name: string
  bytes: Uint8Array
  onChange: (bytes: Uint8Array) => void
  /** Actual loading/native operation state. Pending text is reported independently via onDraftChange. */
  onBusyChange?: (busy: boolean) => void
  onDraftChange?: (dirty: boolean) => void
  viewOptions?: { zoom: number; navigation: boolean; focus: boolean }
}
type PdfViewport = { width: number; height: number; convertToPdfPoint(x: number, y: number): number[]; convertToViewportPoint(x: number, y: number): number[] }
let fontPromise: Promise<Uint8Array> | undefined
function localPdfFont() {
  fontPromise ??= fetch(new URL('./pdf-assets/standard_fonts/LiberationSans-Regular.ttf', document.baseURI)).then(async response => { if (!response.ok) throw new Error('The bundled PDF font could not be loaded.'); return new Uint8Array(await response.arrayBuffer()) }).catch(error => { fontPromise = undefined; throw error })
  return fontPromise
}
const message = (error: unknown) => error instanceof Error ? error.message : String(error)
export default function PdfEditor({ name, bytes, onChange, onBusyChange, onDraftChange, viewOptions, initialRecoveryDraft, onRecoveryDraftChange, registerCommit, registerHistory }: PdfEditorProps) {
  const arrowMarkerId = useId()
  const callbacks = useRef({ onChange, onBusyChange, onDraftChange, onRecoveryDraftChange }); callbacks.current = { onChange, onBusyChange, onDraftChange, onRecoveryDraftChange }
  const recovered = useRef(parsePdfRecoveryDraft(initialRecoveryDraft))
  const recoveryWarning = useRef(initialRecoveryDraft != null && !recovered.current)
  const history = useRef<PdfHistory | null>(null); history.current ??= new PdfHistory(bytes)
  const [current, setCurrent] = useState(history.current.bytes)
  const [summary, setSummary] = useState<PdfSummary>()
  const [page, setPage] = useState(recovered.current?.page ?? 1)
  const [fieldName, setFieldName] = useState(recovered.current?.fieldName ?? '')
  const [formDirty, setFormDirty] = useState(recovered.current?.tool === 'form')
  const [tool, setTool] = useState<Tool>(recovered.current?.tool ?? 'view')
  const [noteTarget,setNoteTarget]=useState<{ref:string;signature:string}|undefined>(recovered.current?.tool==='edit-note'?{ref:recovered.current.annotationRef!,signature:recovered.current.annotationSignature!}:undefined)
  const [oldText, setOldText] = useState(recovered.current?.oldText ?? '')
  const [text, setText] = useState(recovered.current?.text ?? '')
  const [size, setSize] = useState(recovered.current?.size ?? 16)
  const [color, setColor] = useState(recovered.current?.color ?? '#224f9e')
  const [placement, setPlacement] = useState<Placement | undefined>(recovered.current?.placement)
  const [working, setWorking] = useState(false)
  const [rendering, setRendering] = useState(true)
  const [error, setError] = useState('')
  const [notice, setNotice] = useState('')
  const [notes, setNotes] = useState<Array<{ id: string; rect: number[]; text: string }>>([])
  const [openNote, setOpenNote] = useState<string>()
  const [viewport, setViewport] = useState<PdfViewport>()
  const [exportRange, setExportRange] = useState('')
  const [query, setQuery] = useState('')
  const [loaded,setLoaded] = useState<{bytes:Uint8Array;pdf:PDFDocumentProxy;pdfjs:typeof import('pdfjs-dist/legacy/build/pdf.mjs')}>()
  const [searchResult,setSearchResult] = useState<PdfSearchResult>()
  const [searching,setSearching] = useState(false)
  const [searchPage,setSearchPage] = useState(0)
  const searchAbort = useRef<AbortController | undefined>(undefined)
  const pendingSearchHit = useRef<{page:number;match:number} | undefined>(undefined)
  const [pageText, setPageText] = useState<string[]>([])
  const [matchIndex, setMatchIndex] = useState(0)
  const textContainer = useRef<HTMLDivElement>(null)
  const textDivs = useRef<HTMLElement[]>([])
  const searchInput = useRef<HTMLInputElement>(null)
  const matches = findPdfTextMatches(pageText, query)
  const canvas = useRef<HTMLCanvasElement>(null)
  const draft = useRef(!!recovered.current)
  const locked = useRef(false)
  const live = useRef(true)
  const composing = useRef(false)
  const start = useRef<Placement | undefined>(undefined)
  const zoom = Math.max(50, Math.min(200, viewOptions?.zoom ?? 100)) / 100
  const busy = working || rendering
  useEffect(() => { callbacks.current.onBusyChange?.(busy) }, [busy])
  const hasDraft = tool === 'form' ? formDirty : !!placement || text.length > 0
  const selectedField = summary?.fields.find(field => field.name === fieldName)
  function notifyDraft(value: boolean) { draft.current = value; callbacks.current.onDraftChange?.(value); callbacks.current.onBusyChange?.(busy) }
  function persistDraft(overrides: Partial<PdfRecoveryDraft> = {}) {
    if (tool === 'view' && !overrides.tool) return
    const effectiveTool=overrides.tool??tool
    const candidate = { version: 1, format: 'pdf', page, tool:effectiveTool, text, size, color, placement, ...(effectiveTool === 'form' ? { fieldName } : {}), ...(['replace','edit-note'].includes(effectiveTool) ? { oldText } : {}), ...(effectiveTool==='edit-note'&&noteTarget?{annotationRef:noteTarget.ref,annotationSignature:noteTarget.signature}:{}), ...overrides }
    callbacks.current.onRecoveryDraftChange?.(parsePdfRecoveryDraft(candidate))
  }
  function cancel() { if (working) return; if(tool==='edit-note'){setNoteTarget(undefined);setTool('view')} setFormDirty(false); setText(tool === 'form' ? selectedField?.value ?? '' : ''); setOldText(''); setPlacement(undefined); start.current = undefined; notifyDraft(false); callbacks.current.onRecoveryDraftChange?.(null); if (page > (summary?.pages.length ?? page)) setPage(1); setError('') }
  useEffect(() => { live.current = true; if (recovered.current) { notifyDraft(true) }; return () => { live.current = false; callbacks.current.onBusyChange?.(false) } }, [])
  useEffect(() => {
    let cancelled=false, destroy:(()=>void)|undefined
    setLoaded(undefined); setSummary(undefined); setRendering(true); callbacks.current.onBusyChange?.(true); setError('')
    searchAbort.current?.abort(); setSearchResult(undefined); setSearching(false)
    void (async()=>{
      const [pdfjs,worker,info]=await Promise.all([import('pdfjs-dist/legacy/build/pdf.mjs'),import('pdfjs-dist/legacy/build/pdf.worker.min.mjs?url'),inspectPdf(current)])
      if (cancelled) return
      pdfjs.GlobalWorkerOptions.workerSrc=worker.default
      const assets=new URL('./pdf-assets/',document.baseURI).href
      const task=pdfjs.getDocument({data:current.slice(),cMapUrl:`${assets}cmaps/`,cMapPacked:true,standardFontDataUrl:`${assets}standard_fonts/`,wasmUrl:`${assets}wasm/`})
      destroy=()=>{void task.destroy()}
      const pdf=await task.promise
      if (!cancelled) {setSummary(info);setLoaded({bytes:current,pdf,pdfjs})}
    })().catch(error=>{if(!cancelled){setError(message(error));setRendering(false);locked.current=false}})
    return ()=>{cancelled=true;searchAbort.current?.abort();destroy?.()}
  },[current])
  useEffect(() => {
    if (!loaded || loaded.bytes !== current) return
    let cancelled=false, cancelRender:(()=>void)|undefined,cancelText:(()=>void)|undefined
    const {pdf,pdfjs}=loaded
    locked.current=true;setRendering(true);callbacks.current.onBusyChange?.(true);setViewport(undefined);setNotes([]);setPageText([]);textDivs.current=[];if(textContainer.current)textContainer.current.replaceChildren();setError('')
    void (async()=>{
      if (draft.current && page>pdf.numPages) throw new Error('The recovered draft refers to a missing page. Cancel the draft to continue.')
      const index = Math.min(page, pdf.numPages)
      if (cancelled) return
      if (index !== page) setPage(index)
      const selected = await pdf.getPage(index)
      if (cancelled || !canvas.current) return
      const view = selected.getViewport({ scale: zoom * 96 / 72 })
      if (view.width * view.height > 16000000 || view.width > 16000 || view.height > 16000) throw new Error('This page is too large to render at this zoom. Reduce the zoom.')
      const annotations = await selected.getAnnotations()
      if (cancelled) return
      setNotes(annotations.filter(value => value.subtype === 'Text').map(value => ({ id: value.id, rect: value.rect, text: value.contentsObj?.str ?? '' })))
      setOpenNote(undefined)
      setViewport(view)
      const ratio = Math.min(window.devicePixelRatio || 1, 2)
      canvas.current.width = Math.ceil(view.width * ratio); canvas.current.height = Math.ceil(view.height * ratio)
      const context = canvas.current.getContext('2d'); if (!context) throw new Error('The PDF page cannot be rendered.')
      const render = selected.render({ canvas: canvas.current, canvasContext: context, viewport: view, transform: ratio === 1 ? undefined : [ratio, 0, 0, ratio, 0, 0] })
      cancelRender = () => render.cancel()
      await render.promise
      if (!cancelled && textContainer.current) {
        const layer = new pdfjs.TextLayer({ textContentSource: await selected.getTextContent(), container: textContainer.current, viewport: view })
        cancelText = () => layer.cancel()
        await layer.render()
        if (!cancelled) { textDivs.current = layer.textDivs; setPageText(layer.textContentItemsStr); setMatchIndex(pendingSearchHit.current?.page === index ? pendingSearchHit.current.match : 0); pendingSearchHit.current = undefined }
      }
      if (!cancelled && recoveryWarning.current) { recoveryWarning.current = false; setError('The recovered PDF draft is invalid and could not be restored.') }
    })().catch(error=>{if(!cancelled)setError(message(error))}).finally(()=>{if(!cancelled){locked.current=false;setRendering(false)}})
    return ()=>{cancelled=true;cancelRender?.();cancelText?.()}
  },[loaded,current,page,zoom])
  async function searchAllPages() {
    if (!loaded || loaded.bytes!==current || searching || !query.trim()) return
    const controller=new AbortController();searchAbort.current=controller;setSearching(true);setSearchPage(0);setSearchResult(undefined)
    try {
      const result=await searchPdfDocument(loaded.pdf.numPages,async page=>{
        const content=await (await loaded.pdf.getPage(page)).getTextContent()
        return content.items.flatMap(item=>'str' in item ? [item.str] : [])
      },query,controller.signal,setSearchPage)
      if(!controller.signal.aborted && live.current)setSearchResult(result)
    } catch(error) {if(!controller.signal.aborted && live.current)setError(message(error))}
    finally {if(searchAbort.current===controller && live.current)setSearching(false)}
  }
  function jumpToSearchHit(hit:{page:number;match:number}) {
    if(busy||hasDraft)return
    if(hit.page===page)setMatchIndex(hit.match)
    else {pendingSearchHit.current=hit;setPage(hit.page)}
  }
  useEffect(() => {
    textDivs.current.forEach((element, index) => {
      element.classList.toggle('pdf-search-hit', matches.some(match => match.spans.includes(index)))
      element.classList.toggle('pdf-search-active', !!matches[matchIndex]?.spans.includes(index))
    })
    const currentMatch = matches[matchIndex]
    if (currentMatch) textDivs.current[currentMatch.spans[0]]?.scrollIntoView({ block: 'nearest', inline: 'nearest' })
  }, [query, pageText, matchIndex])
  async function command(value: PdfCommand, nextPage = page) {
    if (locked.current || !summary?.editable) return false
    locked.current = true; setWorking(true); callbacks.current.onBusyChange?.(true); setError('')
    try {
      const qualified = value.kind === 'text' || value.kind === 'form-value' ? { ...value, fontBytes: await localPdfFont() } : value
      const saved = await applyPdfCommand(history.current!.bytes, qualified)
      if (!live.current) return false
      const next = history.current!.push(saved)
      if(value.kind==='note.edit'){setTool('view');setNoteTarget(undefined);setOldText('')}
      setFormDirty(false); setText(value.kind === 'form-value' ? value.value : ''); setPlacement(undefined); draft.current = false; callbacks.current.onDraftChange?.(false)
      setPage(nextPage); setCurrent(next); callbacks.current.onChange(next.slice()); setNotice('Change applied. Save to keep it in your file.')
      return true
    } catch (error) { if (live.current) setError(message(error)); return false }
    finally { if (live.current) { locked.current = false; setWorking(false); callbacks.current.onBusyChange?.(rendering) } }
  }
  const bridge = window.injDesktop as (NonNullable<Window['injDesktop']> & {
    pickAsset?: (kind: 'pdf' | 'image') => Promise<{ name: string; bytes: Uint8Array } | null>
    exportBytes?: (input: { name: string; bytes: Uint8Array }) => Promise<{ name: string } | null>
  }) | undefined
  async function importAsset(kind: 'pdf' | 'image') {
    if (locked.current || draft.current || !bridge?.pickAsset || !summary?.editable) return
    setError('')
    try {
      const asset = await bridge.pickAsset(kind)
      if (!asset || !live.current) return
      if (kind === 'image') { await command({ kind: 'image', page, bytes: asset.bytes }); return }
      locked.current = true; setWorking(true); callbacks.current.onBusyChange?.(true)
      const saved = await importPdfPages(history.current!.bytes, asset.bytes, page)
      if (!live.current) return
      const next = history.current!.push(saved); setCurrent(next); setPage(page + 1); callbacks.current.onChange(next.slice()); setNotice(`Pages imported from ${asset.name}.`)
    } catch (error) { if (live.current) setError(message(error)) }
    finally { if (live.current) { locked.current = false; setWorking(false); callbacks.current.onBusyChange?.(rendering) } }
  }
  async function exportPages() {
    if (locked.current || draft.current || !bridge?.exportBytes || !summary) return
    setError('')
    try {
      const range = parsePdfPageRange(exportRange || String(page), summary.pages.length)
      locked.current = true; setWorking(true)
      const saved = await exportPdfPages(history.current!.bytes, range)
      const result = await bridge.exportBytes({ name: name.replace(/\.pdf$/i, '') + '-pages.pdf', bytes: saved })
      if (live.current && result) setNotice(`Exported ${range.length} page${range.length === 1 ? '' : 's'} to ${result.name}.`)
    } catch (error) { if (live.current) setError(message(error)) }
    finally { if (live.current) { locked.current = false; setWorking(false) } }
  }
  function restore(direction: 'undo' | 'redo') {
    if (busy || draft.current) return
    const next = history.current![direction](); setCurrent(next); callbacks.current.onChange(next.slice()); setNotice(direction === 'undo' ? 'Change undone.' : 'Change restored.')
  }
  function editNote(target:PdfAnnotationTarget) {
    if(busy||hasDraft||target.subtype!=='Text'||target.contents.length>10000)return
    const placement={at:target.rect.slice(0,2) as [number,number],end:target.rect.slice(2) as [number,number]}
    setNoteTarget({ref:target.ref,signature:target.signature});setTool('edit-note');setOldText(target.contents);setText(target.contents);setPlacement(placement);notifyDraft(true);setError('')
    persistDraft({tool:'edit-note',text:target.contents,oldText:target.contents,annotationRef:target.ref,annotationSignature:target.signature,placement})
  }
  function pickTool(next: Tool) { if (hasDraft || busy) return; setTool(next); setText(''); setFieldName(''); setError(''); setNotice(''); if (next === 'highlight') setColor('#ffcd38') }
  function point(event: PointerEvent<HTMLDivElement>): Placement | undefined {
    if (!viewport) return
    const rect = event.currentTarget.getBoundingClientRect()
    const x = Math.max(0, Math.min(viewport.width, event.clientX - rect.left))
    const y = Math.max(0, Math.min(viewport.height, event.clientY - rect.top))
    return { at: viewport.convertToPdfPoint(x, y) as [number, number] }
  }
  function down(event: PointerEvent<HTMLDivElement>) {
    if ((tool === 'view' || tool === 'form' || tool === 'replace' || tool === 'edit-note') || busy || !summary?.editable || event.button !== 0) return
    const value = point(event); if (!value) return
    event.currentTarget.setPointerCapture(event.pointerId); start.current = value; setPlacement(value); notifyDraft(true); persistDraft({ placement: value })
  }
  function move(event: PointerEvent<HTMLDivElement>) {
    if (!start.current || !['highlight','underline','strikeout','rectangle','ellipse','line','arrow'].includes(tool)) return
    const value = point(event); if (value) { const next = { ...start.current, end: value.at }; setPlacement(next); persistDraft({ placement: next }) }
  }
  async function apply(): Promise<boolean> {
    if (composing.current) return false
    if(tool==='edit-note'){if(text===oldText){cancel();return true}return noteTarget?command({kind:'note.edit',page,target:noteTarget,text}):false}
    if (tool === 'replace') { if (oldText === text) { cancel(); return true } return replaceExistingText() }
    if (tool === 'form') { if (selectedField && formDirty) return command({ kind: 'form-value', page, name: fieldName, value: text }); return false }
    if (!placement) return false
    if (tool === 'text') return command({ kind: 'text', page, at: placement.at, text, size, color })
    else if (tool === 'note') return command({ kind: 'note', page, at: placement.at, text, color })
    else if ((tool === 'line' || tool === 'arrow') && placement.end) return command({kind:tool,page,from:placement.at,to:placement.end,color})
    else if ((tool === 'ellipse' || tool === 'rectangle' || tool === 'highlight' || tool === 'underline' || tool === 'strikeout') && placement.end) {
      const [x, y] = placement.at, [ex, ey] = placement.end
      return command({ kind: tool, page, color, rect: [Math.min(x, ex), Math.min(y, ey), Math.max(x, ex), Math.max(y, ey)] })
    }
    return false
  }
  async function replaceExistingText() {
    if (locked.current || !summary?.editable || !placement?.end || !oldText || !bridge?.replacePdfText) return false
    locked.current = true; setWorking(true); callbacks.current.onBusyChange?.(true); setError('')
    try {
      const source = history.current!.bytes.slice()
      const digest = await crypto.subtle.digest('SHA-256', source)
      const revision = Array.from(new Uint8Array(digest), byte => byte.toString(16).padStart(2, '0')).join('')
      const saved = await bridge.replacePdfText({ bytes: source, revision, page, rect: [placement.at[0], placement.at[1], placement.end[0], placement.end[1]], oldText, newText: text })
      if (!live.current) return false
      const next = history.current!.push(new Uint8Array(saved)); setCurrent(next); setPlacement(undefined); setText(''); setOldText(''); draft.current = false; callbacks.current.onDraftChange?.(false); callbacks.current.onChange(next.slice()); setNotice('Existing text replaced. Save to keep the change.')
      return true
    } catch (error) { if (live.current) setError(message(error)); return false }
    finally { if (live.current) { locked.current = false; setWorking(false); callbacks.current.onBusyChange?.(rendering) } }
  }
  function selectExistingSpan(target: EventTarget) {
    if (tool !== 'replace' || busy || hasDraft || !viewport || !textContainer.current || !summary?.editable) return
    const element = textDivs.current.find(span => span === target || (target instanceof Node && span.contains(target)))
    if (!element?.textContent?.trim()) return
    const container = textContainer.current.getBoundingClientRect(), box = element.getBoundingClientRect()
    const corners = [[box.left, box.top], [box.right, box.top], [box.left, box.bottom], [box.right, box.bottom]].map(([x, y]) => viewport.convertToPdfPoint(x - container.left, y - container.top))
    const at: [number, number] = [Math.min(...corners.map(p => p[0])) - 1, Math.min(...corners.map(p => p[1])) - 1]
    const end: [number, number] = [Math.max(...corners.map(p => p[0])) + 1, Math.max(...corners.map(p => p[1])) + 1]
    const selectedText = element.textContent
    setOldText(selectedText); setText(selectedText); setPlacement({at, end}); notifyDraft(true); persistDraft({oldText:selectedText, text:selectedText, placement:{at,end}})
  }
  const commitRef = useRef<() => Promise<boolean>>(async () => false)
  commitRef.current = async () => (locked.current || composing.current) ? false : !draft.current ? true : apply()
  useEffect(() => { registerCommit?.(() => commitRef.current()) }, [registerCommit])
  const historyCommands = useRef({ undo() {}, redo() {} })
  historyCommands.current = { undo: () => restore('undo'), redo: () => restore('redo') }
  useEffect(() => { registerHistory?.({ undo: () => historyCommands.current.undo(), redo: () => historyCommands.current.redo() }) }, [registerHistory])
  const disabled = busy || hasDraft
  const count = summary?.pages.length ?? 0
  const displayStart = placement && viewport?.convertToViewportPoint(...placement.at)
  const displayEnd = placement?.end && viewport?.convertToViewportPoint(...placement.end)
  const drawStyle = displayStart && { left: Math.min(displayStart[0], displayEnd?.[0] ?? displayStart[0]), top: Math.min(displayStart[1], displayEnd?.[1] ?? displayStart[1]), width: Math.max(3, Math.abs((displayEnd?.[0] ?? displayStart[0]) - displayStart[0])), height: Math.max(3, Math.abs((displayEnd?.[1] ?? displayStart[1]) - displayStart[1])) }

  return <section className="pdf-workspace" aria-label={`PDF editor: ${name}`} onKeyDown={event => {
    if (event.key.toLowerCase() === 'f' && (event.ctrlKey || event.metaKey)) { event.preventDefault(); searchInput.current?.focus(); searchInput.current?.select() }
    if (event.key === 'Escape' && hasDraft) { event.preventDefault(); cancel() }
    if (event.key === 'Enter' && (event.ctrlKey || event.metaKey) && !event.nativeEvent.isComposing && hasDraft) { event.preventDefault(); apply() }
  }}>
    <div className="pdf-tools" role="toolbar" aria-label="PDF tools">
      {(['view', 'replace', 'text', 'note', 'highlight', 'underline', 'strikeout', 'rectangle', 'ellipse', 'line', 'arrow', 'form'] as Exclude<Tool,'edit-note'>[]).map(value => <button key={value} aria-pressed={tool === value} disabled={busy || hasDraft || (value !== 'view' && !summary?.editable)} onClick={() => pickTool(value)}>{({ view: 'Select', replace: 'Edit existing text', text: 'Add text', note: 'Note', highlight: 'Highlight area', underline:'Underline area', strikeout:'Strike through', rectangle: 'Rectangle', ellipse:'Ellipse', line:'Line', arrow:'Arrow', form: 'Fill forms' })[value]}</button>)}
      <button disabled={disabled || !summary?.editable || !bridge?.pickAsset} onClick={() => void importAsset('image')}>Insert image…</button>
      <span className="pdf-toolbar-spacer" />
      <div className="pdf-search"><input ref={searchInput} aria-label="Find PDF text" placeholder="Find text" value={query} maxLength={500} onChange={event => { searchAbort.current?.abort(); setSearching(false); setSearchResult(undefined); setQuery(event.target.value); setMatchIndex(0) }} /><span aria-live="polite">{query ? `${matches.length ? matchIndex + 1 : 0} / ${matches.length}` : ''}</span><button disabled={!query.trim() || searching || !loaded} onClick={() => void searchAllPages()}>All pages</button><button aria-label="Previous match" disabled={!matches.length} onClick={() => setMatchIndex((matchIndex - 1 + matches.length) % matches.length)}>‹</button><button aria-label="Next match" disabled={!matches.length} onClick={() => setMatchIndex((matchIndex + 1) % matches.length)}>›</button></div>
      <button disabled={disabled || !history.current?.canUndo} onClick={() => restore('undo')}>Undo</button><button disabled={disabled || !history.current?.canRedo} onClick={() => restore('redo')}>Redo</button>
    </div>
    {searching && <div className="pdf-search-results" role="status">Searching page {searchPage + 1}… <button onClick={() => {searchAbort.current?.abort();setSearching(false)}}>Cancel search</button></div>}
    {searchResult && <div className="pdf-search-results" aria-label="Document search results"><span role="status">{searchResult.hits.length} matches across {searchResult.pagesScanned} of {searchResult.totalPages} pages{searchResult.limited ? ' · Search limit reached; narrow your query' : ''}</span>{[...new Set(searchResult.hits.map(hit=>hit.page))].slice(0,100).map(number => <button key={number} disabled={busy||hasDraft} onClick={()=>jumpToSearchHit(searchResult.hits.find(hit=>hit.page===number)!)}>Page {number} · {searchResult.hits.filter(hit=>hit.page===number).length}</button>)}{new Set(searchResult.hits.map(hit=>hit.page)).size>100 && <span>Showing the first 100 matching pages.</span>}</div>}
    {error && <p className="pdf-message pdf-error" role="alert">{error}</p>}
    {summary?.refusal && <p className="pdf-message">{summary.refusal}</p>}
    <div className="pdf-body">
      {(viewOptions?.navigation ?? true) && !viewOptions?.focus && <aside className="pdf-pages" aria-label="PDF pages">
        <div className="pdf-pages-heading">Pages <span>{count}</span></div>
        <div className="pdf-page-list">{summary?.pages.map((info, index) => <button key={index} className="pdf-page-button" disabled={disabled} aria-current={page === index + 1 ? 'page' : undefined} onClick={() => { setPage(index + 1); setNotice('') }}><span className="pdf-page-symbol" aria-hidden="true">{index + 1}</span><span>Page {index + 1}<small>{Math.round(info.width)} × {Math.round(info.height)} pt</small></span></button>)}</div>
      </aside>}
      <div className="pdf-main">
        <div className="pdf-page-actions" role="toolbar" aria-label="Page controls">
          <button disabled={disabled || page <= 1} onClick={() => setPage(page - 1)} aria-label="Previous PDF page">‹</button><span>Page {page} of {count || '…'}</span><button disabled={disabled || page >= count} onClick={() => setPage(page + 1)} aria-label="Next PDF page">›</button>
          <span className="pdf-action-divider" />
          <button disabled={disabled || !summary?.editable} onClick={() => void command({ kind: 'add-page', page }, page + 1)}>Add page</button>
          <button disabled={disabled || !summary?.editable || !bridge?.pickAsset} onClick={() => void importAsset('pdf')}>Import pages…</button>
          <button disabled={disabled || !summary?.editable || count < 2} onClick={() => void command({ kind: 'delete-page', page }, Math.max(1, Math.min(page, count - 1)))}>Delete page</button>
          <button disabled={disabled || !summary?.editable} onClick={() => void command({ kind: 'rotate', page })}>Rotate 90°</button>
          <button disabled={disabled || !summary?.editable || page === 1} onClick={() => void command({ kind: 'move-page', page, to: page - 1 }, page - 1)}>Move earlier</button>
          <button disabled={disabled || !summary?.editable || page === count} onClick={() => void command({ kind: 'move-page', page, to: page + 1 }, page + 1)}>Move later</button>
        </div>
        {!!summary?.pages[page-1]?.annotations.length && <details className="pdf-annotation-list"><summary>Page annotations ({summary.pages[page-1].annotations.length})</summary>{summary.pages[page-1].annotations.map(annotation=><div key={annotation.ref}><span>{annotation.subtype==='Text'?'Note':annotation.subtype}{annotation.contents?`: ${annotation.contents.slice(0,140)}`:''}</span>{annotation.subtype==='Text'&&<button disabled={disabled||!summary.editable||annotation.contents.length>10000} onClick={()=>editNote(annotation)}>Edit note</button>}<button disabled={disabled||!summary.editable} aria-label={`Delete ${annotation.subtype} annotation ${annotation.ref}`} onClick={()=>void command({kind:'annotation.delete',page,target:annotation})}>Delete</button></div>)}</details>}
        <div className="pdf-export-row"><label>Export pages<input aria-label="PDF pages to export" value={exportRange} placeholder={String(page)} disabled={disabled} onChange={event => setExportRange(event.target.value)} maxLength={1000} /></label><button disabled={disabled || !bridge?.exportBytes} onClick={() => void exportPages()}>Export PDF…</button><span>For example: 1-3, 5</span></div>
        <div className="pdf-canvas-scroll" aria-busy={rendering}>
          <div className={`pdf-page-surface pdf-tool-${tool}`} style={{ width: viewport?.width, height: viewport?.height }} onPointerDown={down} onPointerMove={move} onPointerUp={() => { start.current = undefined }} onPointerCancel={() => { start.current = undefined }}>
            <canvas ref={canvas} style={{ width: viewport?.width, height: viewport?.height, visibility: viewport ? 'visible' : 'hidden' }} aria-label={`PDF page ${page}`} />
            <div ref={textContainer} className="pdf-text-layer" style={{ '--total-scale-factor': zoom * 96 / 72, pointerEvents: tool === 'view' || tool === 'replace' ? 'auto' : 'none' } as CSSProperties} onClick={event => selectExistingSpan(event.target)} />
            {notes.map(note => {
              const anchor = viewport?.convertToViewportPoint(note.rect[0], note.rect[3]) ?? [0, 0]
              return <div key={note.id} className="pdf-note-marker" style={{ left: Math.max(0, Math.min(anchor[0], (viewport?.width ?? 20) - 24)), top: Math.max(0, Math.min(anchor[1], (viewport?.height ?? 20) - 24)) }} onPointerDown={event => event.stopPropagation()}>
                <button aria-label={`Read note: ${note.text.slice(0, 80)}`} title={note.text} onClick={() => setOpenNote(openNote === note.id ? undefined : note.id)}>▤</button>
                {openNote === note.id && <div className="pdf-note-popup" role="note">{note.text}</div>}
              </div>
            })}
            {placement?.end && viewport && (tool === 'line' || tool === 'arrow') && (()=>{
              const from=viewport.convertToViewportPoint(...placement.at),to=viewport.convertToViewportPoint(...placement.end)
              return <svg className="pdf-line-preview" width={viewport.width} height={viewport.height} aria-hidden="true"><defs><marker id={arrowMarkerId} viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse"><path d="M 0 0 L 10 5 L 0 10" fill="none" stroke={color}/></marker></defs><line x1={from[0]} y1={from[1]} x2={to[0]} y2={to[1]} stroke={color} strokeWidth={1.5*zoom*96/72} markerEnd={tool==='arrow'?`url(#${arrowMarkerId})`:undefined}/></svg>
            })()}
            {placement && tool !== 'line' && tool !== 'arrow' && <div className={`pdf-placement pdf-placement-${tool}`} style={drawStyle ? {...drawStyle,color} : undefined}>{(tool === 'text' || tool === 'note') && <span style={{ color, fontSize: tool === 'text' ? size * zoom * 96 / 72 : 16 }}>{tool === 'text' ? text || 'Text' : 'Note'}</span>}</div>}
          </div>
        </div>
      </div>
      {tool !== 'view' && <aside className="pdf-inspector" aria-label="Tool settings">
        <h2>{({ 'edit-note':'Edit note', replace: 'Edit existing text', text: 'Add text', note: 'Add a note', highlight: 'Highlight an area', underline:'Underline an area', strikeout:'Strike through an area', rectangle: 'Draw a rectangle', ellipse:'Draw an ellipse', line:'Draw a line', arrow:'Draw an arrow', form: 'Fill form fields' })[tool]}</h2>
        <p>{tool==='edit-note'?'Edit the saved note text, then apply.':tool === 'replace' ? 'Click an existing text span, then enter its replacement. The original font and layout must support the change.' : tool === 'form' ? 'Choose a field, enter its value, and apply.' : tool === 'text' || tool === 'note' ? 'Click the page to place it.' : 'Drag across the page to choose an area.'}</p>
        {tool === 'form' && <>
          <label>Field<select aria-label="PDF form field" value={fieldName} disabled={busy || formDirty} onChange={event => { setFieldName(event.target.value); setText(summary?.fields.find(field => field.name === event.target.value)?.value ?? '') }}><option value="">Choose a field</option>{summary?.fields.map(field => <option key={field.name} value={field.name} disabled={field.readOnly}>{field.name}{field.readOnly ? ' (read only)' : ''}</option>)}</select></label>
          {!summary?.fields.length && <p>This PDF has no supported form fields.</p>}
          {selectedField && <label>Value{selectedField.kind === 'checkbox' ? <input aria-label="PDF checkbox value" type="checkbox" checked={text === 'true'} disabled={busy || selectedField.readOnly} onChange={event => { const value = String(event.target.checked); setText(value); setFormDirty(true); notifyDraft(true); persistDraft({ text: value }) }} /> : selectedField.kind === 'text' ? <textarea aria-label="PDF form value" value={text} disabled={working || selectedField.readOnly} maxLength={Math.min(selectedField.maxLength ?? 10000, 10000)} onCompositionStart={() => { composing.current = true }} onCompositionEnd={() => { composing.current = false }} onChange={event => { setText(event.target.value); setFormDirty(true); notifyDraft(true); persistDraft({ text: event.target.value }) }} /> : <select aria-label="PDF form option" value={text} disabled={busy || selectedField.readOnly} onChange={event => { setText(event.target.value); setFormDirty(true); notifyDraft(true); persistDraft({ text: event.target.value }) }}><option value="">No selection</option>{selectedField.options.map(option => <option key={option}>{option}</option>)}</select>}</label>}
        </>}
        {tool === 'replace' && <><label>Selected text<textarea readOnly value={oldText} aria-label="Original PDF text" /></label><label>Replacement<textarea value={text} aria-label="Replace existing PDF text" disabled={busy || !placement} maxLength={10000} onCompositionStart={() => { composing.current = true }} onCompositionEnd={() => { composing.current = false }} onChange={event => { setText(event.target.value); notifyDraft(true); persistDraft({text:event.target.value}) }} /></label></>}
        {(tool === 'text' || tool === 'note' || tool==='edit-note') && <label>{tool === 'text' ? 'Text' : 'Note'}<textarea aria-label={tool === 'text' ? 'New PDF text' : 'PDF note'} value={text} disabled={working} maxLength={10000} rows={4} onCompositionStart={() => { composing.current = true }} onCompositionEnd={() => { composing.current = false }} onChange={event => { setText(event.target.value); notifyDraft(!!placement || event.target.value.length > 0); persistDraft({ text: event.target.value }) }} /></label>}
        {tool === 'text' && <label>Size (pt)<input type="number" min={6} max={144} value={size} disabled={working} onChange={event => { const next = Math.max(6, Math.min(144, Number(event.target.value) || 6)); setSize(next); if (hasDraft) persistDraft({ size: next }) }} /></label>}
        {tool !== 'form' && tool !== 'replace' && tool !== 'edit-note' && <label>Color<input type="color" value={color} disabled={working} onChange={event => { setColor(event.target.value); if (hasDraft) persistDraft({ color: event.target.value }) }} /></label>}
        <div className="pdf-draft-actions"><button className="pdf-primary" disabled={busy || (tool === 'form' ? !formDirty || !selectedField || selectedField.readOnly : !placement || ((tool === 'text' || tool === 'note' || tool==='edit-note') ? !text.trim() : !placement.end))} onClick={apply}>Apply</button><button disabled={working || !hasDraft} onClick={cancel}>Cancel</button></div>
        {hasDraft && <p className="pdf-draft-hint">Apply or cancel before changing pages. Save also applies the draft.</p>}
      </aside>}
    </div>
    <div className="pdf-status"><span role="status">{working ? 'Applying change…' : rendering ? 'Rendering page…' : notice || (tool === 'view' ? 'Choose a tool to add content or arrange pages.' : 'Press Esc to cancel. Ctrl / ⌘ + Enter applies.')}</span><details><summary>Editing support</summary><p>Adds text, annotations, and form values. Existing text replacement uses the original font and refuses unsupported or ambiguous content. Select text to copy it, or search across document pages. Document search is bounded to 2,000 pages, 10 million characters and 10,000 matches. Imported pages retain page content; form PDFs cannot be imported or split. Images are inserted at the center of the page. Added text embeds Liberation Sans and supports available Latin, Greek, and Cyrillic characters. Complex scripts are not yet supported. Text stays on one line. Page changes and additions can be undone until the file is closed.</p></details></div>
  </section>
}

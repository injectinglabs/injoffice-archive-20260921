import { useEffect, useId, useMemo, useRef, useState } from 'react'
import { createDocxWasmClient } from '@injoffice/docx-wasm'
import type { NativeDocxDocumentV1 } from '../../../packages/docs/src/nativeContract'
import { buildDocxRunMutation, editableDocxRuns, verifyDocxRoundTrip } from '../../playground/src/docxRoundTrip'
import DocumentPreview from './DocumentPreview'
import {loadDocumentImages,type DocumentImageCache} from './document-media'
import { replaceParagraphLines, insertDocumentImage, deleteDocumentImage, replaceDocumentImage, replaceEditableDocumentText, mergeWithPreviousParagraph, insertDocumentTable, changeDocumentTable, changeDocumentTableGrid, type TableGridOperation } from './document-authoring'
import {runAppearance} from './document-style'
import {paragraphTextOffset, type DocumentTextRange} from './document-range'
import {createHiddenApplyScheduler} from './hidden-apply'
import HyperlinkControl from './HyperlinkControl'
import PageLayoutControl,{type PagePatch} from './PageLayoutControl'
import InsertTableControl from './InsertTableControl'
import ParagraphToolbar, { type ParagraphPatch } from './ParagraphToolbar'
import FormattingToolbar, { type FormattingPatch } from './FormattingToolbar'
import SelectionToolbar from './SelectionToolbar'
import ContextMenu, { activateRunAt, documentContextMenu, useContextMenu } from './ContextMenu'
import { documentRangeFormattingValues, formattingValues, documentFormatting, documentParagraphFormatting, docxSelection, documentTableSelection, paragraphOperations, documentStructure, type ParagraphOperation } from './formatting'
import {documentStatistics,statisticsWithDraft} from './document-statistics'
import './office-editor.css'

export interface OfficeEditorProps {
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

interface TextTarget { key: string; label: string; value: string }
type Preview = { kind: 'docx'; document: NativeDocxDocumentV1 }
/**
 * Preview media is held BESIDE the snapshot (WeakMap in the engine), never as a property of it. The
 * snapshot is the return value of engine calls that take find/replace text, so static analysis treats
 * every property read off it as DOM-derived (CodeQL js/xss-through-dom #48-#50). Image URLs must not be
 * read through that object; they are minted from media bytes and looked up by snapshot identity.
 */
type PreviewMedia = { images: Record<string, string>; notice: string }
const EMPTY_MEDIA: PreviewMedia = Object.freeze({ images: Object.freeze({}), notice: '' })
type TextRange = DocumentTextRange
interface Snapshot { bytes: Uint8Array; preview: Preview; targets: TextTarget[]; preferredSelection?:{key:string;range:TextRange} }
interface LocalEngine {
  media(snapshot: Snapshot): PreviewMedia
  hyperlink(snapshot:Snapshot,key:string,url:string|null):Promise<Snapshot>
  page(snapshot:Snapshot,patch:PagePatch):Promise<Snapshot>
  replaceImage(snapshot:Snapshot,id:string,bytes:Uint8Array,name:string):Promise<Snapshot>
  deleteImage(snapshot:Snapshot,id:string):Promise<{snapshot:Snapshot;key:string;text:string}>
  image(snapshot:Snapshot,key:string,bytes:Uint8Array,name:string):Promise<{snapshot:Snapshot;key:string;text:string}>
  read(bytes: Uint8Array): Promise<Snapshot>
  edit(snapshot: Snapshot, key: string, value: string): Promise<Snapshot>
  replaceAll?(snapshot: Snapshot, search: string, replacement: string): Promise<Snapshot>
  tableGrid?(snapshot:Snapshot,key:string,operation:TableGridOperation):Promise<{snapshot:Snapshot;key:string;text:string}>
  tableOperation?(snapshot:Snapshot,key:string,operation:ParagraphOperation):Promise<{snapshot:Snapshot;key:string;text:string}>
  table?(snapshot: Snapshot, key: string, rows: number, columns: number): Promise<{snapshot: Snapshot; key: string; text: string}>
  paragraphFormat?(snapshot: Snapshot, key: string, patch: ParagraphPatch): Promise<Snapshot>
  format?(snapshot: Snapshot, key: string, patch: FormattingPatch, range?:TextRange): Promise<Snapshot>
  join?(snapshot: Snapshot, key: string, text: string): Promise<{snapshot: Snapshot; key: string; text: string; caret: number} | null>
  lines?(snapshot: Snapshot, key: string, text: string): Promise<{ snapshot: Snapshot; key: string; text: string }>
  paragraph?(snapshot: Snapshot, key: string, operation: ParagraphOperation): Promise<Snapshot>
  terminate(): void
}

function operationId() { return `desktop-${crypto.randomUUID()}` }
function errorMessage(error: unknown) { return error instanceof Error ? error.message : String(error) }
function captureDraftCaret(text: string) {
  const fallback = paragraphTextOffset([text.length], 0, text.length) ?? 0
  const selection = typeof window === 'undefined' ? undefined : window.getSelection?.()
  if (!selection?.rangeCount) return fallback
  const range = selection.getRangeAt(0)
  const start = range.startContainer as { closest?: (selector: string) => Element | null; parentElement?: { closest?: (selector: string) => Element | null } | null }
  const run = (typeof start.closest === 'function' ? start : start.parentElement)?.closest?.('[data-docx-run]')
  if (!run || typeof range.cloneRange !== 'function') return fallback
  try {
    const prefix = range.cloneRange()
    prefix.selectNodeContents(run)
    prefix.setEnd(range.startContainer, range.startOffset)
    return paragraphTextOffset([text.length], 0, prefix.toString().length) ?? fallback
  } catch {
    return fallback
  }
}

function createEngine(extension: string): LocalEngine {
  if (extension !== 'docx') throw new Error('Choose a DOCX document.')
    const client = createDocxWasmClient()
    const imageCache:DocumentImageCache=new Map()
    const mediaFor = new WeakMap<Snapshot, PreviewMedia>()
    const read = async (bytes: Uint8Array): Promise<Snapshot> => {
      const document = await client.extract(bytes)
      const media=await loadDocumentImages(client,bytes,document,imageCache)
      const value: Snapshot = { bytes, preview: { kind: 'docx', document }, targets: editableDocxRuns(document).map(target => ({ key: target.key, label: target.label, value: target.text })) }
      mediaFor.set(value, { images: media.images, notice: media.notice })
      return value
    }
    return { read, media(snapshot){ return mediaFor.get(snapshot) ?? EMPTY_MEDIA }, async hyperlink(snapshot,key,url){
      const document=snapshot.preview.document,run=docxSelection(document,key)?.run
      if(!run?.can_edit_hyperlink)throw new Error('This text segment cannot be linked safely.')
      return read(await client.apply(snapshot.bytes,document,{protocol:'injoffice.office.mutations',version:1,format:'docx',mutation_id:operationId(),expected_revision:document.source.package_sha256,payload:{mutations:[{target_kind:'run',target_id:run.id,expected_xml_sha256:run.anchor.xml_sha256,operation:'hyperlink.set',hyperlink:{url,...(run.hyperlink?{expected_xml_sha256:run.hyperlink.anchor.xml_sha256}:{})}}]}}))
    }, async page(snapshot,patch){
      const document=snapshot.preview.document,section=document.sections[0]
      if(document.sections.length!==1||!section?.edit_policy?.allowed_operations.includes('section.page.patch'))throw new Error('Page settings cannot be changed safely in this document.')
      return read(await client.apply(snapshot.bytes,document,{protocol:'injoffice.office.mutations',version:1,format:'docx',mutation_id:operationId(),expected_revision:document.source.package_sha256,payload:{mutations:[{target_kind:'section',target_id:section.id,expected_xml_sha256:section.anchor.xml_sha256,operation:'section.page.patch',page:patch}]}}))
    }, async replaceImage(snapshot,id,bytes,name){
      const result=await replaceDocumentImage(client,snapshot.bytes,snapshot.preview.document,id,bytes,name,operationId)
      return read(result.bytes)
    }, async deleteImage(snapshot,id){
      const result=await deleteDocumentImage(client,snapshot.bytes,snapshot.preview.document,id,operationId)
      return {snapshot:await read(result.bytes),key:result.key,text:result.text}
    }, async image(snapshot,key,bytes,name){
      const result=await insertDocumentImage(client,snapshot.bytes,snapshot.preview.document,key,bytes,name,operationId)
      return {snapshot:await read(result.bytes),key:result.key,text:result.text}
    }, async tableGrid(snapshot,key,operation) {
      const result=await changeDocumentTableGrid(client,snapshot.bytes,snapshot.preview.document,key,operation,operationId)
      return {snapshot:await read(result.bytes),key:result.key,text:result.text}
    }, async tableOperation(snapshot,key,operation) {
      const result=await changeDocumentTable(client,snapshot.bytes,snapshot.preview.document,key,operation,operationId)
      return {snapshot:await read(result.bytes),key:result.key,text:result.text}
    }, async table(snapshot,key,rows,columns) {
      const result=await insertDocumentTable(client,snapshot.bytes,snapshot.preview.document,key,rows,columns,operationId)
      return {snapshot:await read(result.bytes),key:result.key,text:result.text}
    }, async replaceAll(snapshot, search, replacement) {
      const result = await replaceEditableDocumentText(client, snapshot.bytes, snapshot.preview.document, search, replacement, operationId)
      return read(result.bytes)
    }, async join(snapshot, key, text) {
      const result = await mergeWithPreviousParagraph(client, snapshot.bytes, snapshot.preview.document, key, text, operationId)
      return result ? {snapshot:await read(result.bytes),key:result.key,text:result.text,caret:result.caret} : null
    }, terminate: () => client.terminate(), async lines(snapshot, key, text) {
      if (snapshot.preview.kind !== 'docx') throw new Error('Invalid document session.')
      const result = await replaceParagraphLines(client, snapshot.bytes, snapshot.preview.document, key, text, operationId)
      return { snapshot: await read(result.bytes), key: result.key, text: result.text }
    }, async paragraph(snapshot, key, operation) {
      if (snapshot.preview.kind !== 'docx') throw new Error('Invalid document session.')
      return read(await client.apply(snapshot.bytes, snapshot.preview.document, documentStructure(snapshot.preview.document, key, operation, operationId())))
    }, async paragraphFormat(snapshot, key, patch) {
      return read(await client.apply(snapshot.bytes, snapshot.preview.document, documentParagraphFormatting(snapshot.preview.document, key, patch, operationId())))
    }, async format(snapshot, key, patch, range) {
      if (snapshot.preview.kind !== 'docx') throw new Error('Invalid document session.')
      const document = snapshot.preview.document
      if(range?.unsupported && patch.alignment===undefined)throw new Error('This selection includes content that cannot be formatted safely.')
      const selection=docxSelection(document,key)!
      const envelope = documentFormatting(document, key, patch, operationId(),range)
      const next=await read(await client.apply(snapshot.bytes, document, envelope))
      if(range?.paragraph_id && patch.alignment===undefined){
        const before=selection.paragraph,paragraph=next.preview.document.body.blocks.find(block=>block.paragraph?.anchor.path===before.anchor.path)?.paragraph
        if(!paragraph||paragraph.runs.map(run=>run.text??'').join('')!==before.runs.map(run=>run.text??'').join(''))throw new Error('Formatted paragraph text did not pass readback.')
        let offset=0;const first=paragraph.runs.find(run=>{offset+=(run.text??'').length;return offset>range.start_utf16})
        const focus=first&&next.targets.find(target=>target.key===`${encodeURIComponent(first.anchor.part_name)}:${first.id}`)
        if(!focus)throw new Error('Formatted paragraph selection could not be restored.')
        next.preferredSelection={key:focus.key,range:{...range,paragraph_id:paragraph.id}}
      }else if(range && patch.alignment===undefined){
        const source=selection.paragraph.anchor,model=next.preview.document
        const paragraphs=[model.body,...model.headers,...model.footers,...model.notes,...model.comment_stories].flatMap(story=>story.blocks.flatMap(block=>block.paragraph?[block.paragraph]:block.table?.rows.flatMap(row=>row.cells.flatMap(cell=>cell.paragraphs))??[]))
        const paragraph=paragraphs.find(value=>value.anchor.part_name===source.part_name&&value.anchor.path===source.path)
        const selectedIndex=selection.paragraph.runs.findIndex(run=>run.id===selection.run.id)+(range.start_utf16>0?1:0)
        const run=paragraph?.runs[selectedIndex],focus=run&&editableDocxRuns(model).find(target=>target.runId===run.id&&target.partName===run.anchor.part_name)
        if(!focus||focus.text!==selection.run.text?.slice(range.start_utf16,range.end_utf16))throw new Error('Formatted selection did not pass readback.')
        next.preferredSelection={key:focus.key,range:{start_utf16:0,end_utf16:focus.text.length}}
      }
      return next
    }, async edit(snapshot, key, value) {
      if (snapshot.preview.kind !== 'docx') throw new Error('Invalid document session.')
      const document = snapshot.preview.document
      const target = editableDocxRuns(document).find(candidate => candidate.key === key)
      if (!target) throw new Error('This text is not editable.')
      const next = await read(await client.apply(snapshot.bytes, document, buildDocxRunMutation(document, target, value, operationId())))
      if (next.preview.kind !== 'docx') throw new Error('Invalid document readback.')
      verifyDocxRoundTrip(document, next.preview.document, target, value, undefined, undefined, '', undefined, true)
      return next
    } }
}

export default function OfficeEditor({ name, bytes, onChange, onBusyChange, onDraftChange, viewOptions, initialRecoveryDraft, onRecoveryDraftChange, registerCommit, registerHistory }: OfficeEditorProps) {
  const ribbonId=useId()
  const menu=useContextMenu()
  const [ribbonTab,setRibbonTab]=useState<'Home'|'Insert'|'Layout'|'Table'>('Home')
  const [snapshot, setSnapshot] = useState<Snapshot>()
  const [busy, setBusy] = useState(true)
  const [error, setError] = useState('')
  const [selected, setSelected] = useState('')
  const [textRange,setTextRange] = useState<TextRange>()
  const [draft, setDraft] = useState('')
  const [search, setSearch] = useState('')
  const [replacement, setReplacement] = useState('')
  const searchInput = useRef<HTMLInputElement>(null)
  const [searchOpen, setSearchOpen] = useState(false)
  const [caretOffset, setCaretOffset] = useState<number>()
  const [undo, setUndo] = useState<Snapshot[]>([])
  const [redo, setRedo] = useState<Snapshot[]>([])
  const [composing, setComposing] = useState(false)
  const composingRef = useRef(false)
  const applyHiddenRef = useRef<() => Promise<void>>(async () => {})
  const hiddenApplyRef = useRef<ReturnType<typeof createHiddenApplyScheduler> | undefined>(undefined)
  if (!hiddenApplyRef.current) hiddenApplyRef.current = createHiddenApplyScheduler({
    delayMs: 80,
    composing: () => composingRef.current,
    apply: () => applyHiddenRef.current(),
  })
  const exportRequest=useRef<string|undefined>(undefined)
  const [exportStage,setExportStage]=useState<string|undefined>(undefined)
  const [exportNotice,setExportNotice]=useState('')
  const [exportMode,setExportMode]=useState<'original'|'preview'>('preview')
  useEffect(()=>typeof window!=='undefined'?window.injDesktop?.onPdfExportProgress?.(value=>{if(value.requestId===exportRequest.current)setExportStage(value.stage)}):undefined,[])
  useEffect(()=>()=>{if(exportRequest.current)void window.injDesktop?.cancelDocxPdf(exportRequest.current).catch(()=>{})},[])
  const draftPending = useRef(false)
  const engine = useRef<LocalEngine | undefined>(undefined)
  const mounted = useRef(false)
  const callbacks = useRef({ onChange, onBusyChange, onDraftChange, onRecoveryDraftChange })
  callbacks.current = { onChange, onBusyChange, onDraftChange, onRecoveryDraftChange }
  const target = snapshot?.targets.find(candidate => candidate.key === selected)
  const hasDraft = !!target && draft !== target.value
  useEffect(() => { callbacks.current.onBusyChange?.(busy||composing) }, [busy,composing])
  useEffect(() => { callbacks.current.onDraftChange?.(hasDraft) }, [hasDraft])
  useEffect(() => {
    let cancelled = false
    mounted.current = true
    let local: LocalEngine | undefined
    setBusy(true)
    try {
      local = createEngine(name.split('.').pop()?.toLowerCase() ?? '')
      engine.current = local
      void local.read(bytes).then(value => {
        if (cancelled) return
        setSnapshot(value)
        setBusy(false)
        const recovery = initialRecoveryDraft as { version?: number; format?: string; target?: string; text?: string } | null;
        if (recovery?.version === 1 && recovery.format === 'docx' && typeof recovery.target === 'string' && typeof recovery.text === 'string' && recovery.text.length <= 256 * 1024 && value.targets.some(target => target.key === recovery.target && target.value !== recovery.text)) {
          setSelected(recovery.target); setDraft(recovery.text); draftPending.current = true;
          return;
        }
        if (value.targets.length === 1 && value.targets[0]?.value === '' && value.preview.kind === 'docx') {
          setSelected(value.targets[0].key); setDraft('')
        }
      })
        .catch(reason => { if (!cancelled) setError(errorMessage(reason)) })
        .finally(() => { if (!cancelled) setBusy(false) })
    } catch (reason) { setError(errorMessage(reason)); setBusy(false) }
    return () => { cancelled = true; mounted.current = false; hiddenApplyRef.current?.cancel(); local?.terminate(); callbacks.current.onBusyChange?.(false) }
    // A new open session must mount a fresh editor; save paths do not reinitialize it.
  }, [])

  async function choose(key: string,range?:TextRange) {
    hiddenApplyRef.current?.cancel()
    if (busy || composing || !snapshot) return
    let source = snapshot
    if (draftPending.current) {
      if (!engine.current) return
      setBusy(true); callbacks.current.onBusyChange?.(true); setError('')
      try { source = await engine.current.edit(snapshot, selected, draft); if (!mounted.current) return; accept(source) }
      catch (reason) { if (mounted.current) setError(errorMessage(reason)); return }
      finally { if (mounted.current) setBusy(false) }
    }
    const next = source.targets.find(candidate => candidate.key === key)
    setCaretOffset(undefined);setTextRange(range)
    setSelected(next?.key ?? '')
    setDraft(next?.value ?? '')
  }
  function updateDraft(value: string) {
    if (!target || busy) return
    draftPending.current = value !== target.value
    setDraft(value);setTextRange(undefined)
    callbacks.current.onRecoveryDraftChange?.(draftPending.current ? { version: 1, format: 'docx', target: selected, text: value } : null)
    callbacks.current.onBusyChange?.(busy||composing)
    callbacks.current.onDraftChange?.(draftPending.current)
    if (draftPending.current) hiddenApplyRef.current?.schedule()
    else hiddenApplyRef.current?.cancel()
  }
  function cancelDraft() {
    if (busy) return
    hiddenApplyRef.current?.cancel()
    composingRef.current = false
    draftPending.current = false
    callbacks.current.onRecoveryDraftChange?.(null)
    setDraft(target?.value ?? ''); setTextRange(undefined); setSelected(''); setError(''); setComposing(false)
    callbacks.current.onBusyChange?.(false); callbacks.current.onDraftChange?.(false)
  }
  function accept(next: Snapshot) {
    if (!snapshot) return
    // Keep history bounded by bytes as well as count for large local files.
    let history = [...undo, snapshot].slice(-20)
    while (history.length > 1 && history.reduce((sum, item) => sum + item.bytes.byteLength, 0) > 128 * 1024 * 1024) history = history.slice(1)
    setUndo(history); setRedo([]); setSnapshot(next);setTextRange(undefined)
    draftPending.current = false
    setDraft(next.targets.find(candidate => candidate.key === selected)?.value ?? '')
    callbacks.current.onChange(next.bytes)
  }
  async function apply(restoreCaret = false) {
    hiddenApplyRef.current?.cancel()
    if (!snapshot || !engine.current || !target || !hasDraft || busy || composingRef.current) return
    const caret = captureDraftCaret(draft)
    setBusy(true); callbacks.current.onBusyChange?.(true); setError('')
    try {
      const next = await engine.current.edit(snapshot, selected, draft)
      if (mounted.current) {
        accept(next)
        if (restoreCaret) setCaretOffset(caret)
        else if (snapshot.preview.kind === 'docx') setSelected('')
      }
    } catch (reason) { if (mounted.current) setError(errorMessage(reason)) }
    finally { if (mounted.current) setBusy(false) }
  }
  applyHiddenRef.current = () => apply(true)
  async function changeFormatting(patch: FormattingPatch) {
    if (!snapshot || !engine.current?.format || !target || busy || composing) return
    hiddenApplyRef.current?.cancel()
    setBusy(true); callbacks.current.onBusyChange?.(true); setError('')
    try {
      const source = draftPending.current ? await engine.current.edit(snapshot, selected, draft) : snapshot
      const next = await engine.current.format(source, selected, patch, textRange)
      if (mounted.current) {accept(next);if(next.preferredSelection){setSelected(next.preferredSelection.key);setDraft(next.targets.find(target=>target.key===next.preferredSelection!.key)?.value??'');setTextRange(next.preferredSelection.range)}}
    } catch (reason) { if (mounted.current) setError(errorMessage(reason)) }
    finally { if (mounted.current) setBusy(false) }
  }
  async function changeParagraphFormatting(patch: ParagraphPatch) {
    if (!snapshot || !engine.current?.paragraphFormat || !target || busy || composing) return
    setBusy(true); callbacks.current.onBusyChange?.(true); setError('')
    try { const source = draftPending.current ? await engine.current.edit(snapshot, selected, draft) : snapshot; const next = await engine.current.paragraphFormat(source, selected, patch); if (mounted.current) accept(next) }
    catch (reason) { if (mounted.current) setError(errorMessage(reason)) }
    finally { if (mounted.current) setBusy(false) }
  }
  async function changeTableGrid(operation:TableGridOperation) {
    if(!snapshot||!engine.current?.tableGrid||!target||busy||composing)return
    setBusy(true);callbacks.current.onBusyChange?.(true);setError('')
    try {
      const source=draftPending.current?await engine.current.edit(snapshot,selected,draft):snapshot
      const result=await engine.current.tableGrid(source,selected,operation)
      if(mounted.current){accept(result.snapshot);setSelected(result.key);setDraft(result.text);setCaretOffset(0)}
    }catch(reason){if(mounted.current)setError(errorMessage(reason))}
    finally{if(mounted.current)setBusy(false)}
  }
  async function changeTable(operation:ParagraphOperation) {
    if(!snapshot||!engine.current?.tableOperation||!target||busy||composing)return
    setBusy(true);callbacks.current.onBusyChange?.(true);setError('')
    try {
      const source=draftPending.current?await engine.current.edit(snapshot,selected,draft):snapshot
      const result=await engine.current.tableOperation(source,selected,operation)
      if(mounted.current){accept(result.snapshot);setSelected(result.key);setDraft(result.text);setCaretOffset(0)}
    }catch(reason){if(mounted.current)setError(errorMessage(reason))}
    finally{if(mounted.current)setBusy(false)}
  }
  async function exportPdf(){
    const bridge=window.injDesktop
    if(exportRequest.current||!snapshot||!engine.current||busy||composing||!bridge?.exportDocxPdf||!bridge.exportBytes)return
    const requestId=crypto.randomUUID();exportRequest.current=requestId
    setBusy(true);callbacks.current.onBusyChange?.(true);setError('');setExportNotice((exportMode==='preview'?'Preview uses Liberation Sans instead of source fonts; line and page breaks can change. ':'')+'PDF text is exported as vector outlines and cannot be searched or selected.');setExportStage('reading')
    try{
      const source=draftPending.current?await engine.current.edit(snapshot,selected,draft):snapshot
      if(exportRequest.current!==requestId)return
      const bytes=await bridge.exportDocxPdf({requestId,bytes:source.bytes,revision:source.preview.document.source.package_sha256,mode:exportMode})
      if(!mounted.current||exportRequest.current!==requestId)return
      setExportStage('saving')
      const saved=await bridge.exportBytes({name:name.replace(/\.docx$/i,'')+'.pdf',bytes})
      if(mounted.current)setExportNotice(saved?`Saved ${saved.name}${exportMode==='preview'?' with substituted fonts':''}. PDF text uses vector outlines and cannot be selected.`:'PDF export canceled.')
    }catch(reason){if(mounted.current)setError(errorMessage(reason))}
    finally{exportRequest.current=undefined;if(mounted.current){setExportStage(undefined);setBusy(false)}}
  }
  async function cancelPdfExport(){
    const request=exportRequest.current;if(!request||exportStage==='saving')return
    exportRequest.current=undefined
    try{await window.injDesktop?.cancelDocxPdf(request)}catch(reason){if(mounted.current)setError(errorMessage(reason))}
  }
  async function changeLink(url:string|null){
    if(!snapshot||!engine.current||!target||busy||composing||textRange?.unsupported||(textRange&&textRange.start_utf16!==textRange.end_utf16))return
    setBusy(true);callbacks.current.onBusyChange?.(true);setError('')
    try{const source=draftPending.current?await engine.current.edit(snapshot,selected,draft):snapshot;const next=await engine.current.hyperlink(source,selected,url);if(mounted.current)accept(next)}
    catch(reason){if(mounted.current)setError(errorMessage(reason))}
    finally{if(mounted.current)setBusy(false)}
  }
  async function changePage(patch:PagePatch){
    if(!snapshot||!engine.current||busy||composing)return
    setBusy(true);callbacks.current.onBusyChange?.(true);setError('')
    try{const source=draftPending.current?await engine.current.edit(snapshot,selected,draft):snapshot;const next=await engine.current.page(source,patch);if(mounted.current)accept(next)}
    catch(reason){if(mounted.current)setError(errorMessage(reason))}
    finally{if(mounted.current)setBusy(false)}
  }
  async function replaceImage(id:string){
    const pick=typeof window!=='undefined'?window.injDesktop?.pickAsset:undefined
    if(!snapshot||!engine.current||busy||composing||!pick)return
    // Queue the native dialog before reporting the editing operation as busy.
    const pending=pick('image')
    setBusy(true);callbacks.current.onBusyChange?.(true);setError('')
    try{
      const asset=await pending;if(!asset||!mounted.current)return
      const source=draftPending.current?await engine.current.edit(snapshot,selected,draft):snapshot
      const next=await engine.current.replaceImage(source,id,asset.bytes,asset.name)
      if(mounted.current)accept(next)
    }catch(reason){if(mounted.current)setError(errorMessage(reason))}
    finally{if(mounted.current)setBusy(false)}
  }
  async function deleteImage(id:string) {
    if(!snapshot||!engine.current||busy||composing)return
    setBusy(true);callbacks.current.onBusyChange?.(true);setError('')
    try {
      const source=draftPending.current?await engine.current.edit(snapshot,selected,draft):snapshot
      const result=await engine.current.deleteImage(source,id)
      if(mounted.current){accept(result.snapshot);setSelected(result.key);setDraft(result.text);setCaretOffset(0)}
    }catch(reason){if(mounted.current)setError(errorMessage(reason))}
    finally{if(mounted.current)setBusy(false)}
  }
  async function insertImage() {
    const pick=typeof window!=='undefined'?window.injDesktop?.pickAsset:undefined
    if(!snapshot||!engine.current||!target||busy||composing||!pick)return
    setBusy(true);callbacks.current.onBusyChange?.(true);setError('')
    try {
      const asset=await pick('image');if(!asset||!mounted.current)return
      const source=draftPending.current?await engine.current.edit(snapshot,selected,draft):snapshot
      const result=await engine.current.image(source,selected,asset.bytes,asset.name)
      if(mounted.current){accept(result.snapshot);setSelected(result.key);setDraft(result.text);setCaretOffset(0)}
    }catch(reason){if(mounted.current)setError(errorMessage(reason))}
    finally{if(mounted.current)setBusy(false)}
  }
  async function insertTable(rows:number,columns:number) {
    if(!snapshot||!engine.current?.table||!target||busy||composing)return
    setBusy(true);callbacks.current.onBusyChange?.(true);setError('')
    try {
      const source=draftPending.current?await engine.current.edit(snapshot,selected,draft):snapshot
      const result=await engine.current.table(source,selected,rows,columns)
      if(mounted.current){accept(result.snapshot);setSelected(result.key);setDraft(result.text);setCaretOffset(0)}
    }catch(reason){if(mounted.current)setError(errorMessage(reason))}
    finally{if(mounted.current)setBusy(false)}
  }
  async function changeParagraph(operation: ParagraphOperation) {
    if (!snapshot || !engine.current?.paragraph || !target || busy || draftPending.current) return
    setBusy(true); callbacks.current.onBusyChange?.(true); setError('')
    try {
      const next = await engine.current.paragraph(snapshot, selected, operation)
      if (mounted.current) { accept(next); setSelected(''); setDraft('') }
    } catch (reason) { if (mounted.current) setError(errorMessage(reason)) }
    finally { if (mounted.current) setBusy(false) }
  }
  async function joinPrevious() {
    if (!snapshot || !engine.current?.join || !target || busy || composing) return
    setBusy(true); callbacks.current.onBusyChange?.(true); setError('')
    try { const result = await engine.current.join(snapshot, selected, draft); if (mounted.current && result) { accept(result.snapshot); setSelected(result.key); setDraft(result.text); setCaretOffset(result.caret) } }
    catch (reason) { if (mounted.current) setError(errorMessage(reason)) }
    finally { if (mounted.current) setBusy(false) }
  }
  async function insertLines(text: string, caret: number) {
    if (!snapshot || !engine.current?.lines || !target || busy || composing) return
    setBusy(true); callbacks.current.onBusyChange?.(true); setError('')
    try {
      const result = await engine.current.lines(snapshot, selected, text)
      if (mounted.current) { accept(result.snapshot); setSelected(result.key); setDraft(result.text); setCaretOffset(caret) }
    } catch (reason) { if (mounted.current) setError(errorMessage(reason)) }
    finally { if (mounted.current) setBusy(false) }
  }
  const commitLatest = useRef<() => Promise<boolean>>(async () => false)
  commitLatest.current = async () => {
    if (busy || composingRef.current) return false
    if (!draftPending.current) return true
    await apply()
    if (draftPending.current) return false
    callbacks.current.onDraftChange?.(false); callbacks.current.onBusyChange?.(false)
    return true
  }
  useEffect(() => { registerCommit?.(() => commitLatest.current()) }, [registerCommit])
  function travel(direction: 'undo' | 'redo') {
    if (!snapshot || busy || draftPending.current) return
    const stack = direction === 'undo' ? undo : redo
    const next = stack.at(-1)
    if (!next) return
    if (direction === 'undo') { setUndo(undo.slice(0, -1)); setRedo([...redo, snapshot]) }
    else { setRedo(redo.slice(0, -1)); setUndo([...undo, snapshot]) }
    draftPending.current = false
    setSnapshot(next); setTextRange(undefined); setSelected(''); setDraft(''); setError(''); callbacks.current.onChange(next.bytes)
  }
  async function replaceAll() {
    if (!snapshot || !engine.current?.replaceAll || busy || hasDraft || composing) return
    setBusy(true); callbacks.current.onBusyChange?.(true); setError('')
    try { const next = await engine.current.replaceAll(snapshot, search, replacement); if (mounted.current) { accept(next); setSelected(''); setDraft('') } }
    catch (reason) { if (mounted.current) setError(errorMessage(reason)) }
    finally { if (mounted.current) setBusy(false) }
  }
  const matches = search ? snapshot?.targets.filter(target => target.value.includes(search)) ?? [] : []
  const historyCommands = useRef({ undo() {}, redo() {} })
  historyCommands.current = { undo: () => travel('undo'), redo: () => travel('redo') }
  useEffect(() => { registerHistory?.({ undo: () => historyCommands.current.undo(), redo: () => historyCommands.current.redo() }) }, [registerHistory])
  const zoom = Math.max(50, Math.min(200, viewOptions?.zoom ?? 100)) / 100
  const toolbarValues=snapshot && (textRange?.paragraph_id?documentRangeFormattingValues(snapshot.preview.document,selected,textRange,hasDraft?draft:undefined):formattingValues(snapshot.preview,selected))
  if(toolbarValues && textRange && (textRange.unsupported || (textRange.paragraph_id&&docxSelection(snapshot!.preview.document,selected)?.paragraph.runs.some(run=>runAppearance(snapshot!.preview.document,docxSelection(snapshot!.preview.document,selected)!.paragraph,run).hidden)) || (textRange.paragraph_id?!docxSelection(snapshot!.preview.document,selected)?.paragraph.can_format_range:!docxSelection(snapshot!.preview.document,selected)?.run.can_format_range)))toolbarValues.characterEditable=false
  const selectedTable=snapshot && documentTableSelection(snapshot.preview.document,selected)
  const baseStatistics=useMemo(()=>snapshot?documentStatistics(snapshot.preview.document):undefined,[snapshot])
  const statisticsTarget=snapshot&&hasDraft?docxSelection(snapshot.preview.document,selected):undefined
  const statistics=baseStatistics&&statisticsWithDraft(baseStatistics,statisticsTarget?{paragraphId:statisticsTarget.paragraph.id,runId:statisticsTarget.run.id,text:draft}:undefined)
  const tabs:typeof ribbonTab[]=['Home','Insert','Layout',...(selectedTable?['Table' as const]:[])]
  const activeRibbonTab=tabs.includes(ribbonTab)?ribbonTab:'Home'
  const isDocument = snapshot?.preview.kind === 'docx'
  return <div className={`office-editor ${isDocument ? 'office-editor-document' : ''}`} onKeyDown={event => { if ((event.ctrlKey || event.metaKey) && !event.altKey && ['b','i','u'].includes(event.key.toLowerCase()) && !(event.target as HTMLElement).closest('input,select,textarea')) {event.preventDefault();const key=event.key.toLowerCase(),property=key==='b'?'bold':key==='i'?'italic':'underline';void changeFormatting({[property]:!toolbarValues?.[property]});return} if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'f') { event.preventDefault(); setSearchOpen(true); requestAnimationFrame(() => searchInput.current?.focus()) } }}>
    <div className="office-toolbar">
      <button aria-expanded={searchOpen} onClick={() => setSearchOpen(value => !value)}>Find / replace</button>
      <button onClick={() => travel('undo')} disabled={busy || hasDraft || !undo.length}>Undo</button>
      <button onClick={() => travel('redo')} disabled={busy || hasDraft || !redo.length}>Redo</button>
      {snapshot&&typeof window!=='undefined'&&window.injDesktop?.exportDocxPdf&&<><select aria-label="PDF font handling" disabled={busy||composing} value={exportMode} onChange={event=>setExportMode(event.target.value as 'original'|'preview')}><option value="preview">Bundled fonts · preview</option><option value="original">Original embedded fonts</option></select><button disabled={busy||composing} title="Export vector PDF; text will be outlined rather than selectable" onClick={()=>void exportPdf()}>Export PDF…</button></>}
      {isDocument && target && <><button className="office-apply" disabled={busy || !hasDraft || composing} onClick={() => void apply()}>Apply change</button><button disabled={busy} onClick={cancelDraft}>Cancel</button></>}
      {textRange && <span>{textRange.unsupported?'This selection includes unsupported content.':`${textRange.end_utf16-textRange.start_utf16} selected characters`}</span>}
      {(busy || hasDraft || isDocument) && <span>{busy ? exportStage ? 'Exporting PDF…' : snapshot ? 'Applying change…' : 'Opening…' : hasDraft ? 'Save applies your pending text.' : target ? 'Ctrl / ⌘ + Enter to apply; Esc to cancel.' : 'Click text to edit.'}</span>}
    </div>
    {exportNotice&&<p className="office-document-note" role="status">{exportNotice}</p>}
    {exportStage&&<div className="office-document-note" role="status">{exportStage==='saving'?'Choose where to save the PDF…':`Exporting PDF · ${exportStage}…`}{exportStage!=='saving'&&<button onClick={()=>void cancelPdfExport()}>Cancel PDF export</button>}</div>}
    {searchOpen && <section className="document-search" aria-label="Find and replace editable text"><input ref={searchInput} type="search" aria-label="Find editable document text" placeholder="Find editable text" maxLength={1000} value={search} onChange={event => setSearch(event.target.value)} /><span role="status">{matches.length} matching segments</span><button disabled={!matches.length || busy || hasDraft} onClick={() => { const index = matches.findIndex(match => match.key === selected); choose(matches[(index + 1) % matches.length].key) }}>Next match</button><input aria-label="Replacement text" placeholder="Replace with" maxLength={10000} value={replacement} onChange={event => setReplacement(event.target.value)} /><button disabled={!matches.length || busy || hasDraft || search === replacement} onClick={() => void replaceAll()}>Replace all</button><button onClick={() => setSearchOpen(false)} aria-label="Close document search">×</button><small>Case-sensitive; searches each editable text segment independently.</small></section>}
    {error && <div className="office-error" role="alert">{error}</div>}
    {snapshot && <>
      <div className="office-ribbon">
        <div className="office-ribbon-tabs" role="tablist" aria-label="Document tools">{tabs.map(tab=><button key={tab} role="tab" id={`${ribbonId}-tab-${tab}`} aria-selected={activeRibbonTab===tab} aria-controls={`${ribbonId}-panel-${tab}`} tabIndex={activeRibbonTab===tab?0:-1} onClick={()=>setRibbonTab(tab)} onKeyDown={event=>{
          if(!['ArrowLeft','ArrowRight','Home','End'].includes(event.key))return
          event.preventDefault();const index=event.key==='Home'?0:event.key==='End'?tabs.length-1:(tabs.indexOf(tab)+(event.key==='ArrowRight'?1:-1)+tabs.length)%tabs.length
          setRibbonTab(tabs[index]);(event.currentTarget.parentElement?.children[index] as HTMLButtonElement)?.focus()
        }}>{tab}</button>)}</div>
        <div className="office-ribbon-panel" role="tabpanel" id={`${ribbonId}-panel-Home`} aria-labelledby={`${ribbonId}-tab-Home`} hidden={activeRibbonTab!=='Home'}>
      {<FormattingToolbar scopeLabel={textRange?'Selected text range':undefined} kind={snapshot.preview.kind} values={toolbarValues} disabled={busy || composing} onChange={patch => void changeFormatting(patch)} />}
          {docxSelection(snapshot.preview.document,selected)?.paragraph.can_format_range&&!docxSelection(snapshot.preview.document,selected)!.paragraph.runs.some(run=>runAppearance(snapshot.preview.document,docxSelection(snapshot.preview.document,selected)!.paragraph,run).hidden)&&<button disabled={busy||composing||hasDraft||!docxSelection(snapshot.preview.document,selected)!.paragraph.runs.some(run=>run.text?.length)} onClick={()=>{const paragraph=docxSelection(snapshot.preview.document,selected)!.paragraph;setTextRange({paragraph_id:paragraph.id,start_utf16:0,end_utf16:paragraph.runs.reduce((length,run)=>length+(run.text??'').length,0)})}}>Select paragraph text</button>}
          <ParagraphToolbar section="home" properties={docxSelection(snapshot.preview.document, selected)?.paragraph.properties} styles={snapshot.preview.document.paragraph_styles ?? []} numbering={snapshot.preview.document.numbering_definitions} disabled={busy || composing || !docxSelection(snapshot.preview.document, selected)?.paragraph.edit_policy.allowed_operations.includes('properties.patch')} onChange={patch => void changeParagraphFormatting(patch)} />
        </div>
        <div className="office-ribbon-panel office-ribbon-insert" role="tabpanel" id={`${ribbonId}-panel-Insert`} aria-labelledby={`${ribbonId}-tab-Insert`} hidden={activeRibbonTab!=='Insert'}>
      {snapshot?.preview.kind === 'docx' && paragraphOperations(snapshot.preview.document, selected).map(operation => <button key={operation} disabled={busy || hasDraft || composing} onClick={() => void changeParagraph(operation)}>{operation === 'block.insert_after' ? 'Insert paragraph below' : 'Delete paragraph'}</button>)}
      {snapshot && paragraphOperations(snapshot.preview.document,selected).includes('block.insert_after') && <InsertTableControl disabled={busy||composing} onInsert={(rows,columns)=>void insertTable(rows,columns)} />}
      {snapshot && paragraphOperations(snapshot.preview.document,selected).includes('block.insert_after') && typeof window!=='undefined' && window.injDesktop?.pickAsset && <button disabled={busy||composing} onClick={()=>void insertImage()}>Insert image…</button>}
          {docxSelection(snapshot.preview.document,selected)?.run.can_edit_hyperlink&&<HyperlinkControl key={selected} url={docxSelection(snapshot.preview.document,selected)?.run.hyperlink?.url} disabled={busy||composing||!!textRange?.unsupported||!!(textRange&&textRange.start_utf16!==textRange.end_utf16)} onChange={url=>void changeLink(url)} />}
          {!target&&<p>Select a body paragraph to insert content.</p>}
        </div>
        <div className="office-ribbon-panel" role="tabpanel" id={`${ribbonId}-panel-Layout`} aria-labelledby={`${ribbonId}-tab-Layout`} hidden={activeRibbonTab!=='Layout'}>
          <PageLayoutControl section={snapshot.preview.document.sections.length===1?snapshot.preview.document.sections[0]:undefined} disabled={busy||composing} onChange={patch=>void changePage(patch)} />
          <ParagraphToolbar section="layout" properties={docxSelection(snapshot.preview.document, selected)?.paragraph.properties} styles={snapshot.preview.document.paragraph_styles ?? []} numbering={snapshot.preview.document.numbering_definitions} disabled={busy || composing || !docxSelection(snapshot.preview.document, selected)?.paragraph.edit_policy.allowed_operations.includes('properties.patch')} onChange={patch => void changeParagraphFormatting(patch)} />
        </div>
        {selectedTable&&<div className="office-ribbon-panel office-ribbon-insert" role="tabpanel" id={`${ribbonId}-panel-Table`} aria-labelledby={`${ribbonId}-tab-Table`} hidden={activeRibbonTab!=='Table'}>
      {selectedTable && (['block.insert_after','block.delete'] as const).filter(operation=>selectedTable.edit_policy.allowed_operations.includes(operation)).map(operation=><button key={operation} disabled={busy||composing} onClick={()=>void changeTable(operation)}>{operation==='block.insert_after'?'Continue after table':'Delete table'}</button>)}
      {selectedTable && (['table.row.insert_after','table.row.delete','table.column.insert_after','table.column.delete'] as const).filter(operation=>selectedTable.edit_policy.allowed_operations.includes(operation)).map(operation=><button key={operation} disabled={busy||composing} onClick={()=>void changeTableGrid(operation)}>{({'table.row.insert_after':'Add row below','table.row.delete':'Delete row','table.column.insert_after':'Add column right','table.column.delete':'Delete column'})[operation]}</button>)}
        </div>}
      </div>
      <div className={`office-preview ${isDocument ? 'office-document-preview' : ''}`} onContextMenu={event=>{activateRunAt(event);menu.open(event)}}>
        <div className="office-preview-scale" style={isDocument ? undefined : { zoom }}>
        {snapshot.preview.kind === 'docx' && <DocumentPreview replaceImage={typeof window!=='undefined'&&window.injDesktop?.pickAsset?id=>void replaceImage(id):undefined} deleteImage={id=>void deleteImage(id)} images={(engine.current?.media(snapshot) ?? EMPTY_MEDIA).images} imageNotice={(engine.current?.media(snapshot) ?? EMPTY_MEDIA).notice} document={snapshot.preview.document} choose={choose} selected={selected} draft={draft} textRange={textRange} onTextRangeChange={setTextRange} caretOffset={caretOffset} joinPrevious={() => void joinPrevious()} insertLines={(text, caret) => void insertLines(text, caret)} updateDraft={updateDraft} apply={() => void apply()} cancel={cancelDraft} busy={busy} hasDraft={hasDraft} onCompositionChange={value=>{composingRef.current=value;setComposing(value);callbacks.current.onBusyChange?.(busy||value);if(!value&&draftPending.current)hiddenApplyRef.current?.schedule()}} zoom={zoom} navigation={(viewOptions?.navigation ?? true) && !viewOptions?.focus} />}
        </div>
      </div>
      <SelectionToolbar values={toolbarValues} disabled={busy||composing} onChange={patch=>void changeFormatting(patch)} />
      {menu.anchor&&<ContextMenu anchor={menu.anchor} label="Document" onClose={menu.close} items={documentContextMenu({anchor:menu.anchor,target:!!target,values:toolbarValues,disabled:busy||composing,link:!!docxSelection(snapshot.preview.document,selected)?.run.can_edit_hyperlink&&!textRange?.unsupported&&!(textRange&&textRange.start_utf16!==textRange.end_utf16),table:paragraphOperations(snapshot.preview.document,selected).includes('block.insert_after'),onFormat:patch=>void changeFormatting(patch),onFind:()=>{setSearchOpen(true);requestAnimationFrame(()=>searchInput.current?.focus())},onRibbonTab:setRibbonTab})} />}
    </>}
    {statistics&&<footer className="office-document-status" aria-label="Document statistics" title="Body text including tables and pending edits. Headers, footers, notes and hidden text are excluded."><span>{statistics.words.toLocaleString()} words</span><span>{statistics.characters.toLocaleString()} characters</span><span>Body text{hasDraft?' · includes pending edits':''}</span></footer>}
    {!snapshot && !busy && <div className="office-empty">This file could not be opened. Choose another file to continue.</div>}
  </div>
}

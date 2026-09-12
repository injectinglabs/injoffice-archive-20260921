import { createContext, useContext, useEffect, useMemo, useRef, useState, type CSSProperties, type ReactNode } from 'react'
import { DOCX_WASM_NATIVE_MAX_TEXT_CODE_UNITS } from '@injoffice/docx-wasm'
import {createNativeDocxPartialContentPreviewV1} from '@injoffice/docs/native-docx'
import {
  DsButton,
  DsCallout,
  DsChip,
  DsField,
  DsSelect,
  DsTextarea,
} from '../design-system/primitives'
import '../design-system/live-create-edit.css'
import './docs-workspace.css'
import { extractDocxPreviewImages } from '../docxPreviewImages'
import { NativeDocxPages } from '../components/NativeDocxPages'
import {NativeDocxPartialCoverage} from '../components/NativeDocxPartialCoverage'
import {
  DOCX_MEDIA_TYPE,
  nativeDocxHighlight,
  nativeDocxTableRows,
  nativeDocxParagraphText,
  nativeDocxPreviewStats,
  nativeDocxStoryText,
  visibleNativeDocxBlocks,
} from '../docsNativePreview'
import {
  buildDocxMutationEvidence,
  buildDocxRunMutation,
  docxDownloadName,
  editableDocxRuns,
  findEditableDocxRun,
  verifyDocxRoundTrip,
  type DocxRoundTripProof,
  type EditableDocxRun,
} from '../docxRoundTrip'
import {
  createBrowserDocxRoundTripRuntime,
  createServerDocxRoundTripRuntime,
  type DocxRoundTripMode,
  type DocxRoundTripRuntime,
} from '../docxRoundTripRuntime'
import type {
  NativeDocxBlockV1,
  NativeDocxDocumentV1,
  NativeDocxParagraphV1,
  NativeDocxRunV1,
  NativeDocxTableV1,
  NativeDocxTableBorderV1,
} from '../../../../packages/docs/src/nativeContract'

const API_BASE = (import.meta.env.VITE_INJOFFICE_API_BASE ?? '').trim().replace(/\/$/, '')
const SERVER_FALLBACK_CONFIGURED = API_BASE.length > 0
const SAMPLE_PATH = `${import.meta.env.BASE_URL}native-docx/northstar-launch-brief.docx.b64`
const RunSelection = createContext<{ targets: EditableDocxRun[]; selected: string; busy: boolean; choose: (target: EditableDocxRun) => void }>({ targets: [], selected: '', busy: false, choose: () => {} })
const PreviewImages = createContext<ReadonlyMap<string, string>>(new Map())

function shortDigest(value: string): string {
  const digest = value.split(':')[1] ?? value
  return digest.length > 22 ? `${digest.slice(0, 12)}…${digest.slice(-8)}` : digest
}

function copyArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  const copy = new Uint8Array(bytes.byteLength)
  copy.set(bytes)
  return copy.buffer
}

function runStyle(run: NativeDocxRunV1): CSSProperties {
  const properties = run.properties
  if (!properties) return {}
  return {
    color: properties.color && properties.color !== 'auto' ? `#${properties.color}` : undefined,
    fontFamily: properties.font_family ? `"${properties.font_family}", sans-serif` : undefined,
    fontSize: properties.font_size_half_points ? `${properties.font_size_half_points / 2}pt` : undefined,
    fontStyle: properties.italic ? 'italic' : undefined,
    fontWeight: properties.bold ? 700 : undefined,
    textDecoration: properties.underline && properties.underline !== 'none' ? 'underline' : undefined,
    textDecorationStyle: properties.underline === 'double' ? 'double' : undefined,
    backgroundColor: nativeDocxHighlight(properties.highlight),
    direction: properties.rtl ? 'rtl' : undefined,
    unicodeBidi: properties.rtl ? 'isolate' : undefined,
  }
}

function RunView({ run }: { run: NativeDocxRunV1 }) {
  const selection = useContext(RunSelection)
  const images = useContext(PreviewImages)
  const [failedImage, setFailedImage] = useState('')
  if (run.kind === 'text') {
    if (run.properties?.hidden) return null
    if (run.page_field) return <span className="docx-control" title="Resolved from pagination in the native page preview">[{run.page_field === 'PAGE' ? 'page number' : 'total pages'}]</span>
    const target = selection.targets.find((candidate) => candidate.runId === run.id && candidate.partName === run.anchor.part_name)
    if (target) return <button type="button" className="docx-editable-run" data-run-key={target.key} style={runStyle(run)} aria-pressed={selection.selected === target.key} aria-label={`Edit ${target.label}: ${run.text || 'empty text'}`} aria-controls="docx-replacement-text" disabled={selection.busy} onClick={() => selection.choose(target)}>{run.text || '\u00a0'}</button>
    return <span style={runStyle(run)}>{run.text}</span>
  }
  if (run.kind === 'control') {
    if (run.control === 'tab') return <span>{'\t'}</span>
    if (run.control === 'soft-hyphen') return <span aria-hidden="true">&shy;</span>
    if (run.control === 'line-break') return <br />
    return <><br /><span className="docx-control">[{run.control}]</span></>
  }
  if (run.kind === 'drawing') {
    const drawing = run.drawing
    const image = drawing?.media_part ? images.get(drawing.media_part) : undefined
    if (image && drawing && failedImage !== image) return <span className="docx-preview-image" title={drawing.placement === 'floating' ? 'Floating image shown inline in this approximate preview' : 'Embedded image'}>
      <img src={image} alt={drawing.alt_text || drawing.name || 'Embedded document image'} loading="lazy" onError={() => setFailedImage(image)} style={{ width: `${Math.min(drawing.width_emu / 9525, 1200)}px`, height: 'auto', maxWidth: '100%' }} />
      {drawing.placement === 'floating' && <small>Floating image · shown inline</small>}
    </span>
    return <span className="docx-object" title={drawing?.media_part ?? 'Native drawing'}>▧ {drawing?.alt_text || drawing?.name || 'drawing'} · {drawing?.placement}</span>
  }
  return <sup className="docx-reference" title={`${run.reference?.kind ?? 'reference'} ${run.reference?.target_id ?? ''}`}>[{run.reference?.kind ?? 'ref'}]</sup>
}

export function ParagraphView({ paragraph }: { paragraph: NativeDocxParagraphV1 }) {
  const properties = paragraph.properties
  const style: CSSProperties = { textAlign: properties.alignment === 'both' || properties.alignment === 'distribute' ? 'justify' : properties.alignment }
  const empty = nativeDocxParagraphText(paragraph).length === 0 && !paragraph.runs.some((run) => run.kind === 'drawing' || run.kind === 'reference')
  const className = properties.paragraph_style_id?.toLowerCase().startsWith('heading') ? 'docx-paragraph docx-heading' : 'docx-paragraph'
  return (
    <p className={className} style={style} data-edit-mode={paragraph.edit_policy.mode} title={`${paragraph.id} · ${paragraph.edit_policy.mode}`}>
      {properties.numbering && <span className="docx-unresolved-numbering" title={`List ${properties.numbering.num_id}, level ${properties.numbering.level}. The extracted content contract does not resolve the list marker.`}>[list]</span>}
      {paragraph.runs.map((run) => <RunView key={run.id} run={run} />)}
      {empty && <span aria-hidden="true">&nbsp;</span>}
    </p>
  )
}

export function TableView({ table }: { table: NativeDocxTableV1 }) {
  const widths = table.grid_widths_twips
  const totalWidth = widths?.reduce((total, width) => total + width, 0) ?? 0
  const rows = nativeDocxTableRows(table)
  const border = (value?: NativeDocxTableBorderV1) => value ? value.style === 'none' ? 'none' : `${value.size_eighth_points / 8}pt solid ${value.color_rgb ? `#${value.color_rgb}` : 'currentColor'}` : undefined
  return (
    <div className="docx-table-wrap">
      <table className="docx-table" data-edit-mode={table.edit_policy.mode} title={`${table.id} · ${table.edit_policy.mode}`} style={{ width: table.width_twips ? `${table.width_twips / 20}pt` : undefined, maxWidth: '100%', marginInlineStart: table.indent_twips ? `${table.indent_twips / 20}pt` : undefined }}>
        {widths && totalWidth > 0 && <colgroup>{widths.map((width, index) => <col key={index} style={{ width: `${width / totalWidth * 100}%` }} />)}</colgroup>}
        <tbody>
          {table.rows.map((row, rowIndex) => (
            <tr key={row.id} style={{ height: row.height_twips ? `${row.height_twips / 20}pt` : undefined }}>
              {rows[rowIndex].map(({ cell, column, rowSpan, orphanContinuation }) => {
                const Cell = row.repeat_header ? 'th' : 'td'
                const margins = table.cell_margins
                const missingBackground = !cell.shading_rgb && cell.paragraphs.some((paragraph) => paragraph.runs.some((run) => run.properties?.color?.toUpperCase() === 'FFFFFF' && !run.properties.hidden && !nativeDocxHighlight(run.properties.highlight)))
                return <Cell key={cell.id} className={missingBackground ? 'docx-missing-background' : undefined} scope={row.repeat_header ? 'col' : undefined} colSpan={cell.grid_span} rowSpan={rowSpan} style={{
                  background: cell.shading_rgb ? `#${cell.shading_rgb}` : undefined,
                  padding: margins ? `${margins.top_twips / 20}pt ${margins.right_twips / 20}pt ${margins.bottom_twips / 20}pt ${margins.left_twips / 20}pt` : undefined,
                  borderTop: border(cell.borders?.top ?? (rowIndex === 0 ? table.borders?.top : table.borders?.inside_horizontal)),
                  borderBottom: border(cell.borders?.bottom ?? (rowIndex + rowSpan === table.rows.length ? table.borders?.bottom : table.borders?.inside_horizontal)),
                  borderLeft: border(cell.borders?.left ?? (column === 0 ? table.borders?.left : table.borders?.inside_vertical)),
                  borderRight: border(cell.borders?.right ?? (column + cell.grid_span === row.cells.reduce((sum, item) => sum + item.grid_span, 0) ? table.borders?.right : table.borders?.inside_vertical)),
                }}>
                  {orphanContinuation && <span className="docx-control" title="The merge could not be projected safely; cell content is shown separately.">[unresolved merge]</span>}
                  {missingBackground && <small className="docx-background-notice">Background unavailable · text outlined for readability</small>}
                  {cell.paragraphs.map((paragraph) => <ParagraphView key={paragraph.id} paragraph={paragraph} />)}
                </Cell>
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}

function BlockView({ block }: { block: NativeDocxBlockV1 }) {
  if (block.paragraph) return <ParagraphView paragraph={block.paragraph} />
  if (block.table) return <TableView table={block.table} />
  return null
}

function MetaRow({ label, children }: { label: string; children: ReactNode }) {
  return <div className="docx-meta-row"><dt>{label}</dt><dd>{children}</dd></div>
}

export default function DocsPage() {
  const uploadRef = useRef<HTMLInputElement | null>(null)
  const browserRuntimeRef = useRef<DocxRoundTripRuntime | null>(null)
  const serverRuntimeRef = useRef<DocxRoundTripRuntime | null>(null)
  const sourceDigestRef = useRef('')
  const previewRef = useRef<HTMLElement | null>(null)
  const [mode, setMode] = useState<DocxRoundTripMode>('browser')
  const [document, setDocument] = useState<NativeDocxDocumentV1 | null>(null)
  const [authoritativeBytes, setAuthoritativeBytes] = useState<Uint8Array | null>(null)
  const [artifactId, setArtifactId] = useState('')
  const [sourceName, setSourceName] = useState('document.docx')
  const [selection, setSelection] = useState('')
  const [draft, setDraft] = useState('')
  const [status, setStatus] = useState('Choose the bundled document or upload your own .docx file.')
  const [error, setError] = useState('')
  const [busy, setBusy] = useState(false)
  const [output, setOutput] = useState<Blob | null>(null)
  const [downloadURL, setDownloadURL] = useState('')
  const [proof, setProof] = useState<DocxRoundTripProof | null>(null)
  const [undoBytes, setUndoBytes] = useState<Uint8Array[]>([])
  const [changed, setChanged] = useState(false)
  const [previewImages, setPreviewImages] = useState<ReadonlyMap<string, string>>(new Map())

  useEffect(() => {
    let cancelled = false
    const controller = new AbortController()
    const urls: string[] = []
    setPreviewImages(new Map())
    if (authoritativeBytes && document) void extractDocxPreviewImages(authoritativeBytes, document, controller.signal).then((images) => {
      if (cancelled) return
      const next = new Map<string, string>()
      for (const [part, image] of images) {
        const url = URL.createObjectURL(new Blob([Uint8Array.from(image.bytes).buffer], { type: image.mime }))
        urls.push(url)
        next.set(part, url)
      }
      setPreviewImages(next)
    }).catch(() => { /* Preserve document preview when optional media is unavailable. */ })
    return () => { cancelled = true; controller.abort(); for (const url of urls) URL.revokeObjectURL(url) }
  }, [authoritativeBytes, document])

  const stats = useMemo(() => document ? nativeDocxPreviewStats(document) : null, [document])
  const partialCoverage=useMemo(()=>{if(!document)return null;try{return {value:createNativeDocxPartialContentPreviewV1(document,{policy:'source-text-with-omissions-v1',read_only:true})}}catch{return {error:true as const}}},[document])
  const preview = useMemo(() => document ? visibleNativeDocxBlocks(document) : null, [document])
  const targets = useMemo(() => document ? editableDocxRuns(document) : [], [document])
  const target = targets.find((candidate) => candidate.key === selection) ?? targets[0]
  const hasDraft = Boolean(target && draft !== target.text)
  const selectedOutsidePreview = document && preview && target
    ? document.body.blocks.slice(preview.blocks.length).find((block) => block.paragraph?.id === target.paragraphId || block.table?.rows.some((row) => row.cells.some((cell) => cell.paragraphs.some((paragraph) => paragraph.id === target.paragraphId))))
    : undefined
  const canReplace = () => !changed && !hasDraft || window.confirm('Replace this document? Download any changes you want to keep first.')
  const mutationEvidence = useMemo(() => document && target
    ? JSON.stringify(buildDocxMutationEvidence(document, target, draft), null, 2)
    : '', [document, target, draft])

  const runtimeFor = (selectedMode: DocxRoundTripMode): DocxRoundTripRuntime => {
    if (selectedMode === 'browser') {
      browserRuntimeRef.current ??= createBrowserDocxRoundTripRuntime()
      return browserRuntimeRef.current
    }
    serverRuntimeRef.current ??= createServerDocxRoundTripRuntime(API_BASE)
    return serverRuntimeRef.current
  }

  const resetBrowserRuntime = () => {
    browserRuntimeRef.current?.terminate()
    browserRuntimeRef.current = null
  }

  useEffect(() => () => {
    browserRuntimeRef.current?.terminate()
    serverRuntimeRef.current?.terminate()
  }, [])

  useEffect(() => {
    if (!output) {
      setDownloadURL('')
      return
    }
    const url = URL.createObjectURL(output)
    setDownloadURL(url)
    return () => URL.revokeObjectURL(url)
  }, [output])

  const adoptDocument = (next: NativeDocxDocumentV1, preferred?: Pick<EditableDocxRun, 'runId' | 'partName'>) => {
    const nextTargets = editableDocxRuns(next)
    const nextTarget = preferred ? findEditableDocxRun(next, preferred) ?? nextTargets[0] : nextTargets[0]
    setDocument(next)
    setSelection(nextTarget?.key ?? '')
    setDraft(nextTarget?.text ?? '')
  }

  const clearSession = () => {
    setDocument(null)
    setAuthoritativeBytes(null)
    setArtifactId('')
    setSelection('')
    setDraft('')
    setProof(null)
    setOutput(null)
    setError('')
    setUndoBytes([])
    setChanged(false)
  }

  const chooseMode = (nextMode: DocxRoundTripMode) => {
    if (nextMode === mode) return
    if (!canReplace()) return
    if (nextMode === 'server' && !SERVER_FALLBACK_CONFIGURED) {
      setError('Set VITE_INJOFFICE_API_BASE to enable the server fallback.')
      return
    }
    if (nextMode === 'server') resetBrowserRuntime()
    clearSession()
    setMode(nextMode)
    setStatus(nextMode === 'browser'
      ? 'Browser-local mode selected. Choose a document; its bytes will not be uploaded.'
      : 'Server fallback selected. Opening a document will upload it to the configured DOCX API.')
  }

  const extractFile = async (blob: Blob, name: string) => {
    setBusy(true)
    setError('')
    setStatus(`Opening ${name}…`)
    try {
      const bytes = new Uint8Array(await blob.arrayBuffer())
      const extracted = await runtimeFor(mode).extract(bytes)
      const next = extracted.document
      setAuthoritativeBytes(bytes)
      setArtifactId(extracted.artifactId)
      setSourceName(name)
      adoptDocument(next)
      sourceDigestRef.current = next.source.package_sha256
      setProof(null)
      setOutput(null)
      setUndoBytes([])
      setChanged(false)
      const count = editableDocxRuns(next).length
      setStatus(`Opened ${name}. Select text in the preview to edit it. ${count} editable passage${count === 1 ? '' : 's'}.`)
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason))
      if (mode === 'browser') {
        resetBrowserRuntime()
        setStatus('Browser extraction did not complete. No document bytes were uploaded; choose Server fallback explicitly if you want to retry remotely.')
      } else {
        setStatus('Server extraction did not complete. Check the configured DOCX API and retry.')
      }
    } finally {
      setBusy(false)
    }
  }

  const loadSample = async () => {
    if (!canReplace()) return
    setBusy(true)
    setError('')
    setStatus('Loading the bundled native DOCX sample…')
    try {
      const response = await fetch(SAMPLE_PATH)
      if (!response.ok) throw new Error(`Sample request failed with HTTP ${response.status}.`)
      const encoded = (await response.text()).trim()
      const binary = atob(encoded)
      const bytes = Uint8Array.from(binary, (character) => character.charCodeAt(0))
      await extractFile(new Blob([copyArrayBuffer(bytes)], { type: DOCX_MEDIA_TYPE }), 'northstar-launch-brief.docx')
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason))
      setStatus('The bundled sample could not be loaded.')
      setBusy(false)
    }
  }

  const chooseTarget = (next: EditableDocxRun) => {
    if (busy || next.key === target?.key) return
    if (hasDraft && !window.confirm('Discard the unapplied text edit and select another passage?')) return
    setSelection(next.key)
    setDraft(next.text)
    setError('')
    requestAnimationFrame(() => {
      const element = [...(previewRef.current?.querySelectorAll<HTMLElement>('[data-run-key]') ?? [])].find((node) => node.dataset.runKey === next.key)
      element?.scrollIntoView({ block: 'nearest', inline: 'nearest' })
    })
  }

  const undo = async () => {
    const bytes = undoBytes.at(-1)
    if (!bytes || busy) return
    if (hasDraft && !window.confirm('Discard the unapplied text edit and undo the last change?')) return
    setBusy(true)
    setError('')
    try {
      const extracted = await runtimeFor(mode).extract(bytes)
      adoptDocument(extracted.document, target)
      setAuthoritativeBytes(bytes)
      setArtifactId(extracted.artifactId)
      setOutput(new Blob([copyArrayBuffer(bytes)], { type: DOCX_MEDIA_TYPE }))
      setProof(null)
      setUndoBytes((history) => history.slice(0, -1))
      setChanged(extracted.document.source.package_sha256 !== sourceDigestRef.current)
      setStatus('Last change undone. The restored document is ready to edit or download.')
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason))
      setStatus('Could not restore the previous document. Your current document is still available.')
    } finally { setBusy(false) }
  }

  const copyMutationEvidence = async () => {
    try {
      if (!navigator.clipboard) throw new Error('Clipboard access is unavailable in this browser context.')
      await navigator.clipboard.writeText(mutationEvidence)
      setError('')
      setStatus('Copied the current revision- and anchor-bound mutation evidence.')
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : 'Clipboard access was refused.')
    }
  }

  const mutate = async () => {
    if (!document || !target || !authoritativeBytes) return
    let savedBytes: Blob | null = null
    setBusy(true)
    setError('')
    setStatus(`Applying one fail-closed text mutation to ${target.label}…`)
    try {
      const envelope = buildDocxRunMutation(document, target, draft, `demo-${Date.now().toString(36)}`)
      const runtime = runtimeFor(mode)
      const applied = await runtime.apply({ original: authoritativeBytes, document, envelope, artifactId, sourceName })
      const mutatedBytes = applied.bytes
      const mutatedBlob = new Blob([copyArrayBuffer(mutatedBytes)], { type: DOCX_MEDIA_TYPE })
      savedBytes = mutatedBlob

      // Always reopen the exact returned bytes. Server mode intentionally does
      // not follow its mutable artifact pointer for this verification step.
      const extracted = await runtime.extract(mutatedBytes)
      const next = extracted.document
      const nextProof = verifyDocxRoundTrip(
        document, next, target, draft, applied.revision, applied.packageSHA256,
        artifactId, applied.artifactId, mode === 'browser',
      )

      setAuthoritativeBytes(mutatedBytes)
      setUndoBytes((history) => [...history.slice(-9), authoritativeBytes])
      setChanged(true)
      setArtifactId(extracted.artifactId)
      setOutput(mutatedBlob)
      setProof(nextProof)
      adoptDocument(next, target)
      setStatus(`Saved, reopened, and verified ${target.label}. The download is the mutated OOXML package.`)
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason))
      if (mode === 'browser') resetBrowserRuntime()
      // A remote request may have advanced its artifact before a verification
      // error or lost response. Retry from our last verified local bytes.
      setArtifactId('')
      if (savedBytes) {
        setStatus('The edit could not be verified. Your previous document and download are unchanged; you can retry or discard the text edit.')
      } else {
        setStatus(mode === 'browser'
          ? 'The browser mutation was refused; no download was produced and no document bytes were uploaded.'
          : 'The server mutation was refused or the fallback API could not be reached; no download was produced.')
      }
    } finally {
      setBusy(false)
    }
  }

  return (
    <div className="platen-fill native-demo docx-demo workbench-surface ds" data-demo-surface="docs" data-demo-dirty={changed || hasDraft} data-demo-busy={busy}>
      <div className="view-switcher ds-workstrip" role="group" aria-label="Document actions">
        <DsButton variant="filled" className="workbench-button workbench-button--primary" disabled={busy} onClick={() => void loadSample()}>
          Open sample
        </DsButton>
        <DsButton variant="outlined" className="workbench-button" disabled={busy} onClick={() => uploadRef.current?.click()}>
          Open .docx
        </DsButton>
        <input
          ref={uploadRef}
          className="visually-hidden"
          type="file"
          accept={`.docx,${DOCX_MEDIA_TYPE}`}
          disabled={busy}
          aria-label="Open a DOCX file"
          onChange={(event) => {
            const file = event.currentTarget.files?.[0]
            if (file && canReplace()) void extractFile(file, file.name)
            event.currentTarget.value = ''
          }}
        />
        <DsChip>{mode === 'browser' ? 'Browser-local · no upload' : 'Explicit server fallback'}</DsChip>
        <DsButton variant="outlined" disabled={busy || !undoBytes.length} onClick={() => void undo()}>Undo change</DsButton>
        {downloadURL && <a className="native-download ds-btn ds-btn--filled" href={downloadURL} download={docxDownloadName(sourceName)}>Download .docx</a>}
      </div>
      <p className="native-status workbench-status ds-status" role="status" aria-live="polite" aria-atomic="true" data-state={error ? 'error' : busy ? 'busy' : document ? 'ready' : 'idle'}>
        {busy ? 'Working · ' : ''}{status}
      </p>
      <DsCallout
        tone="note"
        title="Approximate content preview · not Word pagination"
      >
        {mode === 'browser'
          ? 'Edit supported text and download a real Word file, entirely in your browser. This preview shows document structure; page layout and drawings may look different in Word.'
          : 'Files are uploaded to the configured server. Select supported text to edit it. This preview shows document structure; page layout and drawings may look different in Word.'}
      </DsCallout>
      {error && <DsCallout tone="refuse" title={mode === 'browser' ? 'Browser engine' : 'Server response'}>{error}</DsCallout>}

      <RunSelection.Provider value={{ targets, selected: target?.key ?? '', busy, choose: chooseTarget }}>
      <PreviewImages.Provider value={previewImages}>
      <div className="native-workspace docx-workspace ds-split">
        <section ref={previewRef} className="native-main docx-main ds-split-main" aria-label="Document preview">
          {SERVER_FALLBACK_CONFIGURED && authoritativeBytes && document && <NativeDocxPages bytes={authoritativeBytes} packageDigest={document.source.package_sha256} apiBase={API_BASE} />}
          {!document || !preview ? (
            <div className="native-empty">
              <div>
                <span className="native-empty-mark">D</span>
                <h2>Open a real DOCX</h2>
                <p>Try the sample or open your own document. Select a passage, change its text, and download the edited Word file.</p>
                <DsButton variant="filled" disabled={busy} onClick={() => void loadSample()}>Open sample</DsButton>
              </div>
            </div>
          ) : (
            <article className="docx-contract-sheet ds-page" data-rendering-mode="approximate-content">
              <h4>{sourceName}</h4>
              <p className="native-muted ds-muted">Select a passage to edit. The selected passage is highlighted.</p>
              <p className="docx-preview-boundary">Continuous content view. Fonts and wrapping may differ; supported embedded PNG/JPEG images appear inline, other drawings use placeholders. List markers are unresolved, and headers/footers appear below the body. Unsupported content remains in the original file.</p>
              {partialCoverage&&(partialCoverage.value?<NativeDocxPartialCoverage coverage={partialCoverage.value}/>:<p role="status">Preview coverage is unavailable. The existing editor and source file are unchanged.</p>)}
              {preview.blocks.map((block) => <BlockView key={block.id} block={block} />)}
              {preview.omitted > 0 && <p className="docx-omitted">Preview stopped after 200 body blocks; {preview.omitted} remain in the validated contract.</p>}
              {selectedOutsidePreview && <section className="docx-preview-story" aria-label="Selected passage outside the preview"><h5>Selected passage</h5><BlockView block={selectedOutsidePreview} /></section>}
              {[...document.headers, ...document.footers, ...document.notes, ...document.comment_stories].map((story) => <section key={story.id} className="docx-preview-story" aria-label={story.kind}><h5>{story.kind}</h5>{story.blocks.map((block) => <BlockView key={block.id} block={block} />)}</section>)}
            </article>
          )}
        </section>

        <aside className="native-side docx-side workbench-inspector ds-split-side" aria-label="Edit document">
          <section className="native-panel ds-panel">
            <h3>Edit selected text</h3>
            {target ? (
              <>
                <DsField className="native-field" label="Selected passage">
                  <DsSelect disabled={busy} value={target.key} onChange={(event) => {
                    const next = targets.find((candidate) => candidate.key === event.target.value)
                    if (next) chooseTarget(next)
                  }}>
                    {targets.map((candidate) => <option key={candidate.key} value={candidate.key}>{candidate.text.slice(0, 60) || '(empty text)'} — {candidate.label}</option>)}
                  </DsSelect>
                </DsField>
                <DsField className="native-field" label="Replacement text">
                  <DsTextarea id="docx-replacement-text" disabled={busy} value={draft} maxLength={DOCX_WASM_NATIVE_MAX_TEXT_CODE_UNITS} rows={4} onChange={(event) => {
                    setDraft(event.target.value)
                  }} />
                </DsField>
                <div className="native-actions">
                  <DsButton variant="green" disabled={busy || draft === target.text} onClick={() => void mutate()}>Apply text change</DsButton>
                  <DsButton variant="outlined" disabled={busy || !hasDraft} onClick={() => setDraft(target.text)}>Discard text edit</DsButton>
                </div>
                <p className="native-muted ds-muted">{hasDraft ? 'Apply this edit to include it in the download.' : 'Select text in the preview or choose a passage above.'}</p>
              </>
            ) : <p className="native-muted ds-muted">{document ? 'This file has no supported editable text. You can open another document.' : 'Open a document to start editing.'}</p>}
          </section>

          <details className="docx-technical-details">
            <summary>Technical details</summary>
            <div className="tool-segment ds-segment" role="group" aria-label="DOCX processing runtime">
              <button type="button" aria-pressed={mode === 'browser'} disabled={busy} onClick={() => chooseMode('browser')}>In browser</button>
              <button type="button" aria-pressed={mode === 'server'} disabled={busy || !SERVER_FALLBACK_CONFIGURED} title={SERVER_FALLBACK_CONFIGURED ? 'Upload to the configured DOCX API' : 'Set VITE_INJOFFICE_API_BASE to enable'} onClick={() => chooseMode('server')}>Server fallback</button>
            </div>
            <p className="native-muted ds-muted">The browser loads a 5.3 MiB Go engine in a worker. Edits are checked against the source revision and verified by reopening the returned bytes.</p>
            <dl className="native-proof-list ds-proof">
              <div className="native-proof-row"><dt>Runtime</dt><dd>{mode === 'browser' ? 'browser-local' : 'server fallback'}</dd></div>
              <div className="native-proof-row"><dt>Artifact</dt><dd>{mode === 'browser' ? 'browser memory' : artifactId ? 'stored' : 'request-local'}</dd></div>
            </dl>
            {mutationEvidence && <div className="native-contract-evidence"><h4>Current mutation evidence</h4><DsButton variant="outlined" disabled={busy} onClick={() => void copyMutationEvidence()}>Copy evidence</DsButton><pre>{mutationEvidence}</pre></div>}
          <section className="native-panel ds-panel">
            <h3>Last verified change</h3>
            {proof ? (
              <>
                <dl className="native-proof-list ds-proof">
                  <div className="native-proof-row"><dt>Target</dt><dd>{proof.target}</dd></div>
                  <div className="native-proof-row"><dt>Before</dt><dd>{proof.before || 'blank'}</dd></div>
                  <div className="native-proof-row"><dt>After</dt><dd>{proof.after || 'blank'}</dd></div>
                  <div className="native-proof-row"><dt>Artifact identity</dt><dd>{proof.artifactIdentity}</dd></div>
                  <div className="native-proof-row"><dt>Native ID</dt><dd>{proof.documentIdentity}</dd></div>
                  <div className="native-proof-row"><dt>Preserved parts</dt><dd>{proof.preservedParts} verified</dd></div>
                  <div className="native-proof-row"><dt>CAS moved</dt><dd>{shortDigest(proof.previousRevision)} → {shortDigest(proof.revision)}</dd></div>
                </dl>
                <DsChip tone="green">Applied</DsChip>
              </>
            ) : (
              <p className="native-muted ds-muted">After save, the page re-extracts the exact response bytes and verifies text, main-part identity, package revision, and preserve-verbatim inventory. Download stays disabled unless every check passes.</p>
            )}
          </section>

          <section className="native-panel ds-panel">
            <p className="native-kicker ds-eyebrow">Fidelity boundary</p>
            <h3>What was actually parsed</h3>
            {document && stats ? (
              <dl className="docx-meta ds-proof">
                <MetaRow label="Revision"><code>{shortDigest(document.revision)}</code></MetaRow>
                <MetaRow label="Package"><code>{shortDigest(document.source.package_sha256)}</code></MetaRow>
                <MetaRow label="Main part"><code>{document.source.main_part}</code></MetaRow>
                <MetaRow label="Body"><span>{stats.blocks} blocks · {stats.paragraphs} paragraphs · {stats.tables} tables</span></MetaRow>
                <MetaRow label="Inline"><span>{stats.textRuns} text runs · {stats.drawings} drawings · {stats.references} references</span></MetaRow>
                <MetaRow label="Read only"><span>{stats.readOnlyBlocks} modeled block{stats.readOnlyBlocks === 1 ? '' : 's'}</span></MetaRow>
                <MetaRow label="Stories"><span>{document.headers.length} headers · {document.footers.length} footers · {document.notes.length} notes · {document.comments.length} comments</span></MetaRow>
                <MetaRow label="Sections"><span>{document.sections.length}</span></MetaRow>
              </dl>
            ) : <p className="native-muted ds-muted">No package has been extracted.</p>}
          </section>

          {document && <>
            <section className="native-panel ds-panel">
              <p className="native-kicker ds-eyebrow">Headers and footers</p>
              {[...document.headers, ...document.footers].length === 0
                ? <p className="native-muted ds-muted">None modeled.</p>
                : [...document.headers, ...document.footers].map((story) => <div className="docx-story" key={story.id}><strong>{story.kind}</strong><span>{nativeDocxStoryText(story) || 'Empty story'}</span></div>)}
            </section>
            <section className="native-panel ds-panel">
              <p className="native-kicker ds-eyebrow">Declared capabilities</p>
              <div className="docx-chips">{document.capabilities.map((capability) => <span key={capability.name} title={capability.detail}>{capability.name} · {capability.level}</span>)}</div>
            </section>
            <section className="native-panel ds-panel">
              <p className="native-kicker ds-eyebrow">Preserved or refused</p>
              <h3>{document.unsupported.length} explicit limitation{document.unsupported.length === 1 ? '' : 's'}</h3>
              {document.unsupported.length === 0
                ? <p className="native-muted ds-muted">The extractor reported no unsupported capabilities for this package.</p>
                : <ul className="docx-limitations">{document.unsupported.slice(0, 12).map((item) => <li key={item.id}><strong>{item.code}</strong><span>{item.message}</span><small>{item.preservation}</small></li>)}</ul>}
              {document.unsupported.length > 12 && <p className="native-muted ds-muted">{document.unsupported.length - 12} more remain in the native contract.</p>}
            </section>
          </>}
          </details>
        </aside>
      </div>
      </PreviewImages.Provider>
      </RunSelection.Provider>
    </div>
  )
}

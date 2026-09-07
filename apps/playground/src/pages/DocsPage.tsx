import { useEffect, useMemo, useRef, useState, type CSSProperties, type ReactNode } from 'react'
import { DOCX_WASM_NATIVE_MAX_TEXT_CODE_UNITS } from '@injoffice/docx-wasm'
import {
  DsButton,
  DsCallout,
  DsChip,
  DsField,
  DsSelect,
  DsTextarea,
} from '../design-system/primitives'
import '../design-system/live-create-edit.css'
import {
  DOCX_MEDIA_TYPE,
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
} from '../../../../packages/docs/src/nativeContract'

const API_BASE = (import.meta.env.VITE_INJOFFICE_API_BASE ?? '').trim().replace(/\/$/, '')
const SERVER_FALLBACK_CONFIGURED = API_BASE.length > 0
const SAMPLE_PATH = `${import.meta.env.BASE_URL}native-docx/northstar-launch-brief.docx.b64`

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
  }
}

function RunView({ run }: { run: NativeDocxRunV1 }) {
  if (run.kind === 'text') {
    if (run.properties?.hidden) return null
    return <span style={runStyle(run)}>{run.text}</span>
  }
  if (run.kind === 'control') {
    if (run.control === 'tab') return <span className="docx-control" title="Native tab">→</span>
    if (run.control === 'soft-hyphen') return <span aria-hidden="true">&shy;</span>
    return <><br /><span className="docx-control">[{run.control}]</span></>
  }
  if (run.kind === 'drawing') {
    const drawing = run.drawing
    return <span className="docx-object" title={drawing?.media_part ?? 'Native drawing'}>▧ {drawing?.alt_text || drawing?.name || 'drawing'} · {drawing?.placement}</span>
  }
  return <sup className="docx-reference" title={`${run.reference?.kind ?? 'reference'} ${run.reference?.target_id ?? ''}`}>[{run.reference?.kind ?? 'ref'}]</sup>
}

function ParagraphView({ paragraph }: { paragraph: NativeDocxParagraphV1 }) {
  const properties = paragraph.properties
  const style: CSSProperties = { textAlign: properties.alignment === 'both' || properties.alignment === 'distribute' ? 'justify' : properties.alignment }
  const empty = nativeDocxParagraphText(paragraph).length === 0 && !paragraph.runs.some((run) => run.kind === 'drawing' || run.kind === 'reference')
  const className = properties.paragraph_style_id?.toLowerCase().startsWith('heading') ? 'docx-paragraph docx-heading' : 'docx-paragraph'
  return (
    <p className={className} style={style} data-edit-mode={paragraph.edit_policy.mode} title={`${paragraph.id} · ${paragraph.edit_policy.mode}`}>
      {properties.numbering && <span className="docx-numbering" title={`List ${properties.numbering.num_id}, level ${properties.numbering.level}`}>•</span>}
      {paragraph.runs.map((run) => <RunView key={run.id} run={run} />)}
      {empty && <span aria-hidden="true">&nbsp;</span>}
    </p>
  )
}

function TableView({ table }: { table: NativeDocxTableV1 }) {
  const widths = table.grid_widths_twips
  const totalWidth = widths?.reduce((total, width) => total + width, 0) ?? 0
  return (
    <div className="docx-table-wrap">
      <table className="docx-table" data-edit-mode={table.edit_policy.mode} title={`${table.id} · ${table.edit_policy.mode}`}>
        {widths && totalWidth > 0 && <colgroup>{widths.map((width, index) => <col key={index} style={{ width: `${width / totalWidth * 100}%` }} />)}</colgroup>}
        <tbody>
          {table.rows.map((row) => (
            <tr key={row.id}>
              {row.cells.map((cell) => (
                <td key={cell.id} colSpan={cell.grid_span} style={{ background: cell.shading_rgb ? `#${cell.shading_rgb}` : undefined }}>
                  {cell.vertical_merge === 'continue'
                    ? <span className="docx-control">[continued merged cell]</span>
                    : cell.paragraphs.map((paragraph) => <ParagraphView key={paragraph.id} paragraph={paragraph} />)}
                </td>
              ))}
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

  const stats = useMemo(() => document ? nativeDocxPreviewStats(document) : null, [document])
  const preview = useMemo(() => document ? visibleNativeDocxBlocks(document) : null, [document])
  const targets = useMemo(() => document ? editableDocxRuns(document) : [], [document])
  const target = targets.find((candidate) => candidate.key === selection) ?? targets[0]
  const mutationEvidence = useMemo(() => document && target
    ? JSON.stringify(buildDocxMutationEvidence(document, target, draft), null, 2)
    : '', [document, target, draft])
  const firstSection = document?.sections[0]
  const pageRatio = firstSection ? `${firstSection.page.width_twips} / ${firstSection.page.height_twips}` : '8.5 / 11'

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
  }

  const chooseMode = (nextMode: DocxRoundTripMode) => {
    if (nextMode === mode) return
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
    setStatus(`Extracting ${name} into the native DOCX contract…`)
    setProof(null)
    setOutput(null)
    try {
      const bytes = new Uint8Array(await blob.arrayBuffer())
      const extracted = await runtimeFor(mode).extract(bytes)
      const next = extracted.document
      setAuthoritativeBytes(bytes)
      setArtifactId(extracted.artifactId)
      setSourceName(name)
      adoptDocument(next)
      const count = editableDocxRuns(next).length
      setStatus(`Extracted ${next.body.blocks.length} body block${next.body.blocks.length === 1 ? '' : 's'} with ${count} safe text edit target${count === 1 ? '' : 's'}.`)
    } catch (reason) {
      setDocument(null)
      setAuthoritativeBytes(null)
      setArtifactId('')
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
    setSelection(next.key)
    setDraft(next.text)
    setProof(null)
    setOutput(null)
    setError('')
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
      setArtifactId(extracted.artifactId)
      setOutput(mutatedBlob)
      setProof(nextProof)
      adoptDocument(next, target)
      setStatus(`Saved, reopened, and verified ${target.label}. The download is the mutated OOXML package.`)
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason))
      if (mode === 'browser') resetBrowserRuntime()
      if (savedBytes) {
        setDocument(null)
        setAuthoritativeBytes(null)
        setArtifactId('')
        setSelection('')
        setDraft('')
        setOutput(null)
        setProof(null)
        setStatus('The engine returned DOCX bytes, but exact-byte readback verification failed. No download was enabled; re-open the source before another edit.')
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
    <div className="platen-fill native-demo docx-demo workbench-surface ds" data-demo-surface="docs">
      <div className="view-switcher ds-workstrip" role="group" aria-label="DOCX processing runtime">
        <div className="tool-segment ds-segment" role="group" aria-label="DOCX processing runtime">
          <button type="button" aria-pressed={mode === 'browser'} disabled={busy} onClick={() => chooseMode('browser')}>In browser (default)</button>
          <button
            type="button"
            aria-pressed={mode === 'server'}
            disabled={busy || !SERVER_FALLBACK_CONFIGURED}
            title={SERVER_FALLBACK_CONFIGURED ? 'Upload to the configured DOCX API' : 'Set VITE_INJOFFICE_API_BASE to enable'}
            onClick={() => chooseMode('server')}
          >Server fallback{SERVER_FALLBACK_CONFIGURED ? '' : ' (not configured)'}</button>
        </div>
        <DsButton variant="filled" className="workbench-button workbench-button--primary" disabled={busy} onClick={() => void loadSample()}>
          Use bundled .docx
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
            if (file) void extractFile(file, file.name)
            event.currentTarget.value = ''
          }}
        />
        <DsChip>{mode === 'browser' ? 'Browser-local · no upload' : 'Explicit server fallback'}</DsChip>
      </div>
      <p className="native-status workbench-status ds-status" role="status" aria-live="polite" aria-atomic="true" data-state={error ? 'error' : busy ? 'busy' : document ? 'ready' : 'idle'}>
        {busy ? 'Working · ' : ''}{status}
      </p>
      <DsCallout
        tone="note"
        title={mode === 'browser' ? 'Semantic contract, not Word paint' : 'Explicit server fallback'}
      >
        {mode === 'browser'
          ? 'The first operation loads the version-matched Go engine (about 5.3 MiB) in a Web Worker. This is real DOCX extraction and guarded text write-back, but not Word-compatible pagination or native page-paint. Server fallback stays disabled until VITE_INJOFFICE_API_BASE is set.'
          : 'This mode uploads the document. It uses the same fail-closed extraction and mutation contract as the browser engine; unsupported structures stay explicit instead of being approximated.'}
      </DsCallout>
      {error && <DsCallout tone="refuse" title={mode === 'browser' ? 'Browser engine' : 'Server response'}>{error}</DsCallout>}

      <div className="native-workspace docx-workspace ds-split">
        <main className="native-main docx-main ds-split-main" aria-label="Native DOCX semantic preview">
          {!document || !preview ? (
            <div className="native-empty">
              <div>
                <span className="native-empty-mark">D</span>
                <h2>Open a real DOCX</h2>
                <p>The Go engine runs inside a browser Worker by default. It extracts a revision-bound model, applies one supported text edit, reopens the saved package, and returns real <code>.docx</code> bytes without a server.</p>
                <DsButton variant="filled" className="workbench-button workbench-button--primary" disabled={busy} onClick={() => void loadSample()}>Run the bundled proof</DsButton>
              </div>
            </div>
          ) : (
            <article className="docx-contract-sheet ds-page" style={{ aspectRatio: pageRatio }}>
              <h4>{sourceName}</h4>
              <p className="native-kicker ds-eyebrow">Native source flow · not paginated · {document.protocol} v{document.version}</p>
              {preview.blocks.map((block) => <BlockView key={block.id} block={block} />)}
              {preview.omitted > 0 && <p className="docx-omitted">Preview stopped after 200 body blocks; {preview.omitted} remain in the validated contract.</p>}
            </article>
          )}
        </main>

        <aside className="native-side docx-side workbench-inspector ds-split-side" aria-label="Native DOCX evidence">
          <section className="native-panel ds-panel">
            <p className="native-kicker ds-eyebrow">01 · Extract</p>
            {document ? (
              <dl className="native-proof-list ds-proof">
                <div className="native-proof-row"><dt>Revision</dt><dd>{shortDigest(document.revision)}</dd></div>
                <div className="native-proof-row"><dt>Safe targets</dt><dd>{targets.length}</dd></div>
                <div className="native-proof-row"><dt>Runtime</dt><dd>{mode === 'browser' ? 'browser-local' : 'server fallback'}</dd></div>
                <div className="native-proof-row"><dt>Artifact</dt><dd>{mode === 'browser' ? 'browser memory' : artifactId ? 'stored' : 'request-local'}</dd></div>
              </dl>
            ) : <p className="native-muted ds-muted">No document extracted yet.</p>}
          </section>

          <section className="native-panel ds-panel">
            <p className="native-kicker ds-eyebrow">02 · Mutate</p>
            {target ? (
              <>
                <DsField className="native-field" label="Safe text run">
                  <DsSelect value={target.key} onChange={(event) => {
                    const next = targets.find((candidate) => candidate.key === event.target.value)
                    if (next) chooseTarget(next)
                  }}>
                    {targets.map((candidate) => <option key={candidate.key} value={candidate.key}>{candidate.label}</option>)}
                  </DsSelect>
                </DsField>
                <DsField className="native-field" label="Replacement text">
                  <DsTextarea value={draft} maxLength={DOCX_WASM_NATIVE_MAX_TEXT_CODE_UNITS} rows={4} onChange={(event) => {
                    setDraft(event.target.value)
                    setProof(null)
                    setOutput(null)
                  }} />
                </DsField>
                <div className="native-actions">
                  <DsButton variant="green" className="workbench-button workbench-button--primary" disabled={busy || draft === target.text} onClick={() => void mutate()}>Save to DOCX</DsButton>
                  <DsButton variant="outlined" className="workbench-button" disabled={busy || !mutationEvidence} onClick={() => void copyMutationEvidence()}>Copy evidence</DsButton>
                </div>
                {draft === target.text && <p className="native-muted ds-muted">Change the text to enable a non-empty transaction.</p>}
                <div className="native-contract-evidence" aria-label="Current DOCX mutation evidence">
                  <span className="native-kicker ds-eyebrow">Agent contract · bounded evidence</span>
                  <pre>{mutationEvidence}</pre>
                </div>
              </>
            ) : <p className="native-muted ds-muted">This document exposes no safely editable visible text runs. That refusal is intentional.</p>}
          </section>

          <section className="native-panel ds-panel">
            <p className="native-kicker ds-eyebrow">03 · Verify</p>
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
                {downloadURL && <a className="native-download workbench-button workbench-button--primary ds-btn ds-btn--filled" href={downloadURL} download={docxDownloadName(sourceName)}>Download verified .docx</a>}
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
        </aside>
      </div>
    </div>
  )
}

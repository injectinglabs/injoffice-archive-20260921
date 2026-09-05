import { useEffect, useMemo, useRef, useState } from 'react'
import type { NativePptxDeck } from '@injoffice/pptx-native'
import {
  buildPptxMutationEvidence,
  buildPptxShapeMutation,
  buildPptxTextMutation,
  editablePptxTargets,
  exactPptxShapePresets,
  findPptxTarget,
  firstPptxRunText,
  pptxDownloadName,
  pptxTargetKey,
  pptxTargetValue,
  verifyPptxRoundTrip,
  type EditablePptxTarget,
  type ExactPptxShapePreset,
  type PptxRoundTripProof,
} from '../pptxRoundTrip'
import {
  createBrowserPptxRoundTripRuntime,
  createServerPptxRoundTripRuntime,
  type PptxRoundTripMode,
  type PptxRoundTripRuntime,
} from '../pptxRoundTripRuntime'

const PPTX_MEDIA_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
const SAMPLE_PATH = new URL('../../../../go/pptxpatch/testdata/playground_northstar_review.pptx', import.meta.url).href
const SAMPLE_NAME = 'northstar-launch-review.pptx'
const API_BASE = (import.meta.env.VITE_INJOFFICE_API_BASE ?? '').trim().replace(/\/$/, '')
const SERVER_FALLBACK_CONFIGURED = API_BASE.length > 0

export default function PptxNativePage() {
  const uploadRef = useRef<HTMLInputElement | null>(null)
  const browserRuntimeRef = useRef<PptxRoundTripRuntime | null>(null)
  const serverRuntimeRef = useRef<PptxRoundTripRuntime | null>(null)
  const [mode, setMode] = useState<PptxRoundTripMode>('browser')
  const [deck, setDeck] = useState<NativePptxDeck | null>(null)
  const [authoritativeBytes, setAuthoritativeBytes] = useState<Uint8Array | null>(null)
  const [artifactId, setArtifactId] = useState('')
  const [sourceName, setSourceName] = useState('presentation.pptx')
  const [selection, setSelection] = useState('')
  const [draft, setDraft] = useState('')
  const [shapePreset, setShapePreset] = useState<ExactPptxShapePreset>('rect')
  const [shapeFill, setShapeFill] = useState('625BF6')
  const [status, setStatus] = useState('Choose the bundled presentation or upload your own .pptx file.')
  const [error, setError] = useState('')
  const [busy, setBusy] = useState(false)
  const [output, setOutput] = useState<Blob | null>(null)
  const [downloadURL, setDownloadURL] = useState('')
  const [proof, setProof] = useState<PptxRoundTripProof | null>(null)

  const targets = useMemo(() => deck ? editablePptxTargets(deck) : [], [deck])
  const target = targets.find((candidate) => pptxTargetKey(candidate) === selection) ?? targets[0]
  const targetValue = pptxTargetValue(target)
  const textInvalid = target?.operationKind === 'text.replace' && /[\t\r\n]/.test(draft)
  const shapeUnchanged = target?.operationKind === 'autoshape.update'
    && shapePreset === target.autoShape.preset && shapeFill === (target.autoShape.fill ?? '')
  const mutation = useMemo(() => {
    if (!deck || !target) return null
    try {
      return target.operationKind === 'text.replace'
        ? buildPptxTextMutation(deck, target, draft, 'demo-pptx-text-edit')
        : buildPptxShapeMutation(deck, target, shapePreset, shapeFill || undefined, 'demo-pptx-shape-edit')
    } catch {
      return null
    }
  }, [deck, draft, shapeFill, shapePreset, target])
  const mutationEvidence = useMemo(() => mutation && target
    ? JSON.stringify(buildPptxMutationEvidence(target, mutation), null, 2)
    : '', [mutation, target])

  const runtimeFor = (selectedMode: PptxRoundTripMode): PptxRoundTripRuntime => {
    if (selectedMode === 'browser') {
      browserRuntimeRef.current ??= createBrowserPptxRoundTripRuntime()
      return browserRuntimeRef.current
    }
    serverRuntimeRef.current ??= createServerPptxRoundTripRuntime(API_BASE)
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

  const adoptDeck = (next: NativePptxDeck, preferred?: EditablePptxTarget) => {
    const nextTarget = preferred ? findPptxTarget(next, preferred) : editablePptxTargets(next)[0]
    setDeck(next)
    setSelection(nextTarget ? pptxTargetKey(nextTarget) : '')
    setDraft(nextTarget?.operationKind === 'text.replace' ? firstPptxRunText(nextTarget) : '')
    if (nextTarget?.operationKind === 'autoshape.update') {
      setShapePreset(nextTarget.autoShape.preset)
      setShapeFill(nextTarget.autoShape.fill ?? '')
    }
  }

  const clearSession = () => {
    setDeck(null)
    setAuthoritativeBytes(null)
    setArtifactId('')
    setSourceName('presentation.pptx')
    setSelection('')
    setDraft('')
    setShapePreset('rect')
    setShapeFill('625BF6')
    setProof(null)
    setOutput(null)
    setError('')
  }

  const chooseMode = (nextMode: PptxRoundTripMode) => {
    if (nextMode === mode) return
    if (nextMode === 'server' && !SERVER_FALLBACK_CONFIGURED) {
      setError('Set VITE_INJOFFICE_API_BASE to enable the server fallback.')
      return
    }
    if (nextMode === 'server') resetBrowserRuntime()
    clearSession()
    setMode(nextMode)
    setStatus(nextMode === 'browser'
      ? 'Browser-local mode selected. Choose a presentation; its bytes will not be uploaded.'
      : 'Server fallback selected. Opening a presentation will upload it to the configured PPTX API.')
  }

  const extractFile = async (blob: Blob, name: string) => {
    clearSession()
    setBusy(true)
    setStatus(`Extracting ${name} into the native PPTX contract…`)
    try {
      const bytes = new Uint8Array(await blob.arrayBuffer())
      const extracted = await runtimeFor(mode).extract(bytes)
      setAuthoritativeBytes(bytes)
      setArtifactId(extracted.artifactId)
      setSourceName(name)
      adoptDeck(extracted.deck)
      const count = editablePptxTargets(extracted.deck).length
      setStatus(`Extracted ${extracted.deck.slides.length} slide${extracted.deck.slides.length === 1 ? '' : 's'} with ${count} exact text or AutoShape edit target${count === 1 ? '' : 's'}.`)
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason))
      if (mode === 'browser') {
        resetBrowserRuntime()
        setStatus('Browser extraction did not complete. No presentation bytes were uploaded; choose Server fallback explicitly if you want to retry remotely.')
      } else setStatus('Server extraction did not complete. Check the configured PPTX API and retry.')
    } finally {
      setBusy(false)
    }
  }

  const loadSample = async () => {
    clearSession()
    setBusy(true)
    setStatus('Loading the bundled native PPTX sample…')
    try {
      const response = await fetch(SAMPLE_PATH)
      if (!response.ok) throw await responseError(response)
      await extractFile(new Blob([await response.arrayBuffer()], { type: PPTX_MEDIA_TYPE }), SAMPLE_NAME)
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason))
      setStatus('The sample could not be loaded.')
      setBusy(false)
    }
  }

  const chooseTarget = (next: EditablePptxTarget) => {
    setSelection(pptxTargetKey(next))
    setDraft(next.operationKind === 'text.replace' ? firstPptxRunText(next) : '')
    if (next.operationKind === 'autoshape.update') {
      setShapePreset(next.autoShape.preset)
      setShapeFill(next.autoShape.fill ?? '')
    }
    setProof(null)
    setOutput(null)
    setError('')
  }

  const copyMutationEvidence = async () => {
    try {
      if (!navigator.clipboard) throw new Error('Clipboard access is unavailable in this browser context.')
      await navigator.clipboard.writeText(mutationEvidence)
      setError('')
      setStatus('Copied the current revision- and fingerprint-bound mutation evidence.')
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : 'Clipboard access was refused.')
    }
  }

  const mutate = async () => {
    if (!deck || !target || !authoritativeBytes || !mutation) return
    let savedBytes: Blob | null = null
    setBusy(true)
    setError('')
    setStatus(`Applying one fail-closed ${target.operationKind} mutation to ${target.elementName}…`)
    try {
      const runtime = runtimeFor(mode)
      const applied = await runtime.apply({ original: authoritativeBytes, deck, mutation, artifactId, sourceName })
      const mutatedBytes = applied.bytes
      const mutatedBlob = new Blob([copyArrayBuffer(mutatedBytes)], { type: PPTX_MEDIA_TYPE })
      savedBytes = mutatedBlob

      // Verify the exact bytes returned by apply. Server mode intentionally
      // avoids following the mutable artifact pointer at this boundary.
      const extracted = await runtime.extract(mutatedBytes)
      const next = extracted.deck
      const nextProof = verifyPptxRoundTrip(
        deck,
        next,
        target,
        target.operationKind === 'text.replace' ? { text: draft } : { preset: shapePreset, fill: shapeFill || undefined },
        applied.revision,
        applied.packageSHA256,
        artifactId,
        applied.artifactId,
        mode === 'browser',
      )

      setAuthoritativeBytes(mutatedBytes)
      setArtifactId(extracted.artifactId)
      setOutput(mutatedBlob)
      setProof(nextProof)
      adoptDeck(next, target)
      setStatus(`Saved, reopened, and verified ${target.elementName}. The download is the mutated OOXML package.`)
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason))
      if (mode === 'browser') resetBrowserRuntime()
      if (savedBytes) {
        setDeck(null)
        setAuthoritativeBytes(null)
        setArtifactId('')
        setSelection('')
        setDraft('')
        setOutput(null)
        setProof(null)
        setStatus('The engine returned PPTX bytes, but exact-byte readback verification failed. No download was enabled; re-open the source before another edit.')
      } else {
        setStatus(mode === 'browser'
          ? 'The browser mutation was refused; no download was produced and no presentation bytes were uploaded.'
          : 'The server mutation was refused or the fallback API could not be reached; no download was produced.')
      }
    } finally {
      setBusy(false)
    }
  }

  return (
    <div className="platen-fill native-demo workbench-surface" data-demo-surface="pptx-native">
      <div className="view-switcher" role="group" aria-label="PPTX processing runtime">
        <div className="tool-segment">
          <button type="button" aria-pressed={mode === 'browser'} disabled={busy} onClick={() => chooseMode('browser')}>In browser (default)</button>
          <button type="button" aria-pressed={mode === 'server'} disabled={busy || !SERVER_FALLBACK_CONFIGURED} title={SERVER_FALLBACK_CONFIGURED ? 'Upload to the configured PPTX API' : 'Set VITE_INJOFFICE_API_BASE to enable'} onClick={() => chooseMode('server')}>Server fallback{SERVER_FALLBACK_CONFIGURED ? '' : ' (not configured)'}</button>
        </div>
        <span>{mode === 'browser' ? 'Original bytes stay in this browser' : 'Uploads bytes to the configured API'}</span>
      </div>
      <div className="native-toolbar workbench-toolbar" role="group" aria-label="Native PPTX actions">
        <button type="button" className="workbench-button workbench-button--primary" disabled={busy} onClick={() => void loadSample()}>Use bundled .pptx</button>
        <button type="button" className="workbench-button" disabled={busy} onClick={() => uploadRef.current?.click()}>Open .pptx</button>
        <input ref={uploadRef} className="visually-hidden" type="file" accept={`.pptx,${PPTX_MEDIA_TYPE}`} disabled={busy} aria-label="Open a PPTX file" onChange={(event) => {
          const file = event.currentTarget.files?.[0]
          if (file) void extractFile(file, file.name)
          event.currentTarget.value = ''
        }} />
        <span className="native-badge workbench-badge">{mode === 'browser' ? 'Browser-local · no upload' : 'Explicit server fallback'}</span>
      </div>
      <p className="native-status workbench-status" role="status" aria-live="polite" aria-atomic="true" data-state={error ? 'error' : busy ? 'busy' : deck ? 'ready' : 'idle'}>{busy ? 'Working · ' : ''}{status}</p>
      <p className="native-help workbench-callout workbench-callout--warning">
        {mode === 'browser'
          ? <>The first operation loads the version-matched Go engine (about 6.2 MiB) in a Web Worker. This proof exposes exact text replacement and four exact AutoShape presets; unsupported presentation features remain preserved or refused. Server fallback stays disabled until <code>VITE_INJOFFICE_API_BASE</code> is set.</>
          : <>This mode uploads the presentation. A static build needs a secured compatible API configured with <code>VITE_INJOFFICE_API_BASE</code>.</>}
      </p>
      {error && <div className="native-error workbench-callout workbench-callout--error" role="alert"><strong>{mode === 'browser' ? 'Browser engine' : 'Server response'}</strong><span>{error}</span></div>}

      <div className="native-workspace">
        <section className="native-main" aria-label="Extracted presentation preview">
          {deck ? (
            <>
              <div className="native-sheet-heading"><div><span className="native-kicker">Exact native projection</span><h2>{sourceName}</h2></div><span className="native-muted">{deck.slides.length} slide{deck.slides.length === 1 ? '' : 's'} · {deck.origin}</span></div>
              {targets.length > 0 ? (
                <div className="native-grid-wrap"><table className="native-grid"><thead><tr><th>Slide</th><th>Exact target</th><th>Operation</th><th>Current value</th></tr></thead><tbody>{targets.map((candidate) => {
                  const active = target && pptxTargetKey(candidate) === pptxTargetKey(target)
                  return <tr key={pptxTargetKey(candidate)}><td><span>{candidate.slideIndex + 1}</span></td><td className={active ? 'native-cell-active' : undefined}><button type="button" onClick={() => chooseTarget(candidate)}>{candidate.elementName}</button></td><td><span>{candidate.operationKind}</span></td><td><span>{pptxTargetValue(candidate) || '\u00a0'}</span></td></tr>
                })}</tbody></table></div>
              ) : <p className="native-muted">This presentation exposes no exact text or AutoShape target. That refusal is intentional.</p>}
            </>
          ) : (
            <div className="native-empty"><span className="native-kicker">Native PPTX round trip</span><h2>Put real PPTX bytes through the engine.</h2><p>The bundled proof is a populated, reproducible three-slide launch review with styled metrics, plan content, and exact AutoShapes—not an empty conformance shell. The Go engine extracts a revision-bound model, applies one exact edit, reopens the saved package, and returns real <code>.pptx</code> bytes without a server.</p><button type="button" className="workbench-button workbench-button--primary" disabled={busy} onClick={() => void loadSample()}>Run the bundled proof</button></div>
          )}
        </section>

        <aside className="native-side workbench-inspector" aria-label="PPTX mutation controls">
          <div className="native-panel"><span className="native-kicker">01 · Extract</span>{deck ? (
            <dl className="native-proof-list">
              <div className="native-proof-row"><dt>Document</dt><dd>{deck.documentId.slice(0, 20)}…</dd></div>
              <div className="native-proof-row"><dt>Revision</dt><dd>{shortRevision(deck.sourceRevision || '')}</dd></div>
              <div className="native-proof-row"><dt>Warnings</dt><dd>{deck.compatibility.diagnostics.length}</dd></div>
              <div className="native-proof-row"><dt>Runtime</dt><dd>{mode === 'browser' ? 'browser-local' : 'server fallback'}</dd></div>
              <div className="native-proof-row"><dt>Artifact</dt><dd>{mode === 'browser' ? 'browser memory' : artifactId ? 'stored' : 'request-local'}</dd></div>
            </dl>
          ) : <p className="native-muted">No presentation extracted yet.</p>}</div>

          <div className="native-panel"><span className="native-kicker">02 · Mutate</span>{target ? (
            <>
              <label className="native-field">Exact mutation target<select value={pptxTargetKey(target)} onChange={(event) => {
                const next = targets.find((candidate) => pptxTargetKey(candidate) === event.target.value)
                if (next) chooseTarget(next)
              }}>{targets.map((candidate) => <option key={pptxTargetKey(candidate)} value={pptxTargetKey(candidate)}>Slide {candidate.slideIndex + 1} · {candidate.elementName} · {candidate.operationKind}</option>)}</select></label>
              {target.operationKind === 'text.replace' ? (
                <label className="native-field">New first-run text<input value={draft} maxLength={32_767} onChange={(event) => {
                  setDraft(event.target.value)
                  setProof(null)
                  setOutput(null)
                }} /></label>
              ) : <>
                <label className="native-field">Exact preset<select value={shapePreset} onChange={(event) => {
                  setShapePreset(event.target.value as ExactPptxShapePreset)
                  setProof(null)
                  setOutput(null)
                }}>{exactPptxShapePresets.map((preset) => <option key={preset} value={preset}>{preset}</option>)}</select></label>
                <label className="native-field">Fill (blank means no fill)<input value={shapeFill} maxLength={6} pattern="[0-9A-Fa-f]{6}" onChange={(event) => {
                  setShapeFill(event.target.value.toUpperCase().replace(/[^0-9A-F]/g, '').slice(0, 6))
                  setProof(null)
                  setOutput(null)
                }} /></label>
              </>}
              {textInvalid && <p className="native-muted">Tabs and line breaks are outside this exact mutation subset.</p>}
              <div className="native-actions">
                <button type="button" className="workbench-button workbench-button--primary" disabled={busy || !mutation || textInvalid || target.operationKind === 'autoshape.update' && shapeFill.length > 0 && shapeFill.length !== 6} onClick={() => void mutate()}>Save to PPTX</button>
                <button type="button" className="workbench-button" disabled={busy || !mutationEvidence} onClick={() => void copyMutationEvidence()}>Copy evidence</button>
              </div>
              {(target.operationKind === 'text.replace' ? draft === targetValue : shapeUnchanged) && <p className="native-muted">Change the selected value to enable a non-empty transaction.</p>}
              <div className="native-contract-evidence" aria-label="Current PPTX mutation evidence">
                <span className="native-kicker">Agent contract · bounded evidence</span>
                <pre>{mutationEvidence || 'Change the selected value to preview the guarded mutation.'}</pre>
              </div>
            </>
          ) : <p className="native-muted">No fully materialized text or AutoShape target can be edited safely.</p>}</div>

          <div className="native-panel"><span className="native-kicker">03 · Verify</span>{proof ? (
            <>
              <dl className="native-proof-list">
                <div className="native-proof-row"><dt>Target</dt><dd>{proof.target}</dd></div>
                <div className="native-proof-row"><dt>Before</dt><dd>{proof.before || 'blank'}</dd></div>
                <div className="native-proof-row"><dt>After</dt><dd>{proof.after || 'blank'}</dd></div>
                <div className="native-proof-row"><dt>Artifact identity</dt><dd>{proof.artifactIdentity}</dd></div>
                <div className="native-proof-row"><dt>Untouched anchors</dt><dd>{proof.preservedElements} verified</dd></div>
                <div className="native-proof-row"><dt>CAS moved</dt><dd>{shortRevision(proof.previousRevision)} → {shortRevision(proof.revision)}</dd></div>
              </dl>
              {downloadURL && <a className="native-download workbench-button workbench-button--primary" href={downloadURL} download={pptxDownloadName(sourceName)}>Download verified .pptx</a>}
            </>
          ) : <p className="native-muted">After save, the page re-extracts the exact produced bytes and verifies the requested edit, document identity, revision headers, and every untouched source anchor. Download stays disabled unless every check passes.</p>}</div>
        </aside>
      </div>
    </div>
  )
}

function shortRevision(revision: string): string {
  const digest = revision.replace(/^rev-/, '')
  return digest.length > 16 ? `${digest.slice(0, 10)}…${digest.slice(-6)}` : digest || 'unavailable'
}

function copyArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  const copy = new Uint8Array(bytes.byteLength)
  copy.set(bytes)
  return copy.buffer
}

async function responseError(response: Response): Promise<Error> {
  const detail = (await response.text()).trim()
  return new Error(detail || `Request failed with HTTP ${response.status}.`)
}

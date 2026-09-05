import { useEffect, useMemo, useRef, useState, type CSSProperties } from 'react'
import { adaptWorkbookMutationBatchV1 } from '@injoffice/xlsx-wasm'
import {
  buildCellMutation,
  displayCellValue,
  downloadName,
  editableTargets,
  targetKey,
  type EditableTarget,
  type NativeCell,
  type NativeSheet,
  type NativeWorkbook,
} from '../nativeRoundTrip'
import {
  createBrowserXlsxRoundTripRuntime,
  createServerXlsxRoundTripRuntime,
  type XlsxRoundTripMode,
  type XlsxRoundTripRuntime,
} from '../xlsxRoundTripRuntime'

const XLSX_MEDIA_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
const SAMPLE_PATH = `${import.meta.env.BASE_URL}native-fixture/launch-readiness-plan.xlsx`
const API_BASE = (import.meta.env.VITE_INJOFFICE_API_BASE ?? '').trim().replace(/\/$/, '')
const SERVER_FALLBACK_CONFIGURED = API_BASE.length > 0

type RoundTripProof = {
  cell: string
  before: string
  after: string
  previousRevision: string
  revision: string
}

function shortRevision(revision: string): string {
  const digest = revision.split(':')[1] ?? revision
  return `${digest.slice(0, 10)}…${digest.slice(-6)}`
}

function copyArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  const copy = new Uint8Array(bytes.byteLength)
  copy.set(bytes)
  return copy.buffer
}

async function responseError(response: Response): Promise<Error> {
  const detail = (await response.text()).trim()
  if (detail) {
    try {
      const parsed = JSON.parse(detail) as unknown
      if (typeof parsed === 'object' && parsed !== null && 'error' in parsed && typeof parsed.error === 'string') {
        return new Error(parsed.error)
      }
    } catch {
      // The handler may also return a plain-text refusal.
    }
  }
  return new Error(detail || `Request failed with HTTP ${response.status}.`)
}

function columnName(index: number): string {
  let value = index + 1
  let result = ''
  while (value > 0) {
    value -= 1
    result = String.fromCharCode(65 + (value % 26)) + result
    value = Math.floor(value / 26)
  }
  return result
}

function previewBounds(sheet: NativeSheet): { rows: number; columns: number } {
  const rows = Math.max(3, Math.min(8, 1 + Math.max(-1, ...sheet.cells.map((cell) => cell.row))))
  const columns = Math.max(3, Math.min(6, 1 + Math.max(-1, ...sheet.cells.map((cell) => cell.column))))
  return { rows, columns }
}

function previewCellValue(workbook: NativeWorkbook, cell: NativeCell | undefined): string {
  const value = displayCellValue(cell)
  const numberFormat = cell ? workbook.styles[cell.style_id]?.effective.number_format : undefined
  if (cell?.value?.kind === 'number' && numberFormat?.startsWith('$')) {
    const number = Number(value)
    if (Number.isFinite(number)) return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', maximumFractionDigits: 0 }).format(number)
  }
  return value
}

function previewCellStyle(workbook: NativeWorkbook, cell: NativeCell | undefined): CSSProperties | undefined {
  if (!cell) return undefined
  const style = workbook.styles[cell.style_id]?.effective
  if (!style) return undefined
  return {
    backgroundColor: style.fill_color,
    color: style.font_color,
    fontFamily: style.font_name ? `"${style.font_name}", sans-serif` : undefined,
    fontSize: style.font_size_points ? `${style.font_size_points}pt` : undefined,
    fontStyle: style.italic ? 'italic' : undefined,
    fontWeight: style.bold ? 700 : undefined,
    textAlign: style.horizontal_alignment === 'general' ? undefined : style.horizontal_alignment,
    verticalAlign: style.vertical_alignment === 'middle' ? 'middle' : style.vertical_alignment,
    whiteSpace: style.wrap_text ? 'normal' : undefined,
  }
}

export default function NativeRoundTripPage() {
  const uploadRef = useRef<HTMLInputElement | null>(null)
  const browserRuntimeRef = useRef<XlsxRoundTripRuntime | null>(null)
  const serverRuntimeRef = useRef<XlsxRoundTripRuntime | null>(null)
  const [mode, setMode] = useState<XlsxRoundTripMode>('browser')
  const [workbook, setWorkbook] = useState<NativeWorkbook | null>(null)
  const [authoritativeBytes, setAuthoritativeBytes] = useState<Uint8Array | null>(null)
  const [artifactId, setArtifactId] = useState('')
  const [sourceName, setSourceName] = useState('workbook.xlsx')
  const [selection, setSelection] = useState('')
  const [draft, setDraft] = useState('')
  const [status, setStatus] = useState('Choose the bundled workbook or upload your own .xlsx file.')
  const [error, setError] = useState('')
  const [busy, setBusy] = useState(false)
  const [output, setOutput] = useState<Blob | null>(null)
  const [downloadURL, setDownloadURL] = useState('')
  const [proof, setProof] = useState<RoundTripProof | null>(null)

  const targets = useMemo(() => workbook ? editableTargets(workbook) : [], [workbook])
  const target = targets.find((candidate) => targetKey(candidate) === selection) ?? targets[0]
  const activeSheet = workbook?.sheets.find((sheet) => sheet.id === target?.sheetId) ?? workbook?.sheets[0]
  const targetValue = displayCellValue(target)

  const runtimeFor = (selectedMode: XlsxRoundTripMode): XlsxRoundTripRuntime => {
    if (selectedMode === 'browser') {
      browserRuntimeRef.current ??= createBrowserXlsxRoundTripRuntime()
      return browserRuntimeRef.current
    }
    serverRuntimeRef.current ??= createServerXlsxRoundTripRuntime(API_BASE)
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

  const adoptWorkbook = (next: NativeWorkbook, preferred?: Pick<EditableTarget, 'sheetId' | 'row' | 'column'>) => {
    const nextTargets = editableTargets(next)
    const preferredKey = preferred ? targetKey(preferred) : ''
    const nextTarget = nextTargets.find((candidate) => targetKey(candidate) === preferredKey) ?? nextTargets[0]
    setWorkbook(next)
    setSelection(nextTarget ? targetKey(nextTarget) : '')
    setDraft(displayCellValue(nextTarget))
  }

  const clearSession = () => {
    setWorkbook(null)
    setAuthoritativeBytes(null)
    setArtifactId('')
    setSelection('')
    setDraft('')
    setProof(null)
    setOutput(null)
    setError('')
  }

  const chooseMode = (nextMode: XlsxRoundTripMode) => {
    if (nextMode === mode) return
    if (nextMode === 'server' && !SERVER_FALLBACK_CONFIGURED) {
      setError('Set VITE_INJOFFICE_API_BASE to enable the server fallback.')
      return
    }
    if (nextMode === 'server') resetBrowserRuntime()
    clearSession()
    setMode(nextMode)
    setStatus(nextMode === 'browser'
      ? 'Browser-local mode selected. Choose a workbook; its bytes will not be uploaded.'
      : 'Server fallback selected. Opening a workbook will upload it to the configured XLSX API.')
  }

  const extractFile = async (blob: Blob, name: string) => {
    setBusy(true)
    setError('')
    setStatus(`Extracting ${name} into the native XLSX contract…`)
    setProof(null)
    setOutput(null)
    try {
      const bytes = new Uint8Array(await blob.arrayBuffer())
      const extracted = await runtimeFor(mode).extract(bytes)
      const next = extracted.workbook
      setAuthoritativeBytes(bytes)
      setArtifactId(extracted.artifactId)
      setSourceName(name)
      adoptWorkbook(next)
      const count = editableTargets(next).length
      setStatus(`Extracted ${next.sheets.length} sheet${next.sheets.length === 1 ? '' : 's'} with ${count} safe literal edit target${count === 1 ? '' : 's'}.`)
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason))
      if (mode === 'browser') {
        resetBrowserRuntime()
        setStatus('Browser extraction did not complete. No workbook bytes were uploaded; choose Server fallback explicitly if you want to retry remotely.')
      } else {
        setStatus('Server extraction did not complete. Check the configured XLSX API and retry.')
      }
    } finally {
      setBusy(false)
    }
  }

  const loadSample = async () => {
    setBusy(true)
    setError('')
    setStatus('Loading the bundled Excel-authored sample…')
    try {
      const response = await fetch(SAMPLE_PATH)
      if (!response.ok) throw await responseError(response)
      const blob = new Blob([await response.arrayBuffer()], { type: XLSX_MEDIA_TYPE })
      await extractFile(blob, 'launch-readiness-plan.xlsx')
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason))
      setStatus('The sample could not be loaded.')
      setBusy(false)
    }
  }

  const chooseTarget = (next: EditableTarget) => {
    setSelection(targetKey(next))
    setDraft(displayCellValue(next))
    setProof(null)
    setError('')
  }

  const mutate = async () => {
    if (!workbook || !target || !authoritativeBytes) return
    let savedBytes: Blob | null = null
    setBusy(true)
    setError('')
    setStatus(`Applying one fail-closed mutation to ${target.sheetName}!${target.ref}…`)
    try {
      const operationID = `demo-${Date.now().toString(36)}`
      const batch = buildCellMutation(workbook, target, draft, operationID)
      const transaction = adaptWorkbookMutationBatchV1(workbook, batch)
      const runtime = runtimeFor(mode)
      const applied = await runtime.apply({
        original: authoritativeBytes,
        workbook,
        transaction,
        artifactId,
        sourceName,
      })
      const mutatedBytes = applied.bytes
      const mutatedBlob = new Blob([copyArrayBuffer(mutatedBytes)], { type: XLSX_MEDIA_TYPE })
      savedBytes = mutatedBlob
      setAuthoritativeBytes(mutatedBytes)
      setOutput(mutatedBlob)

      // Reopen the exact produced bytes, not a mutable artifact pointer. On
      // the server path this also mints a fresh snapshot id, keeping the next
      // CAS and download bound to the model displayed below.
      const extracted = await runtime.extract(mutatedBytes)
      const next = extracted.workbook
      const nextTarget = editableTargets(next).find((candidate) => targetKey(candidate) === targetKey(target))
      const before = targetValue
      const after = displayCellValue(nextTarget)
      if (!nextTarget || after !== draft) throw new Error(`Readback mismatch at ${target.sheetName}!${target.ref}.`)
      // Standalone browser extraction has no server-owned identity registry and
      // may derive a fresh document id after package bytes change. The exact
      // saved package is still verified by its revision/digest and cell value.
      if (mode === 'server' && next.document_id !== workbook.document_id) throw new Error('Readback changed the workbook identity.')
      if (applied.revision && applied.revision !== next.revision) throw new Error('Readback revision does not match the saved package.')
      if (applied.packageSHA256 && applied.packageSHA256 !== next.source.package_sha256) throw new Error('Readback digest does not match the saved package.')

      setArtifactId(extracted.artifactId)
      setProof({
        cell: `${target.sheetName}!${target.ref}`,
        before,
        after,
        previousRevision: workbook.revision,
        revision: next.revision,
      })
      adoptWorkbook(next, target)
      setStatus(`Saved, reopened, and verified ${target.sheetName}!${target.ref}. The download is the mutated OOXML package.`)
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason))
      if (mode === 'browser') resetBrowserRuntime()
      if (savedBytes) {
        setWorkbook(null)
        setArtifactId('')
        setSelection('')
        setProof(null)
        setStatus('The XLSX save completed and is available to download, but exact-byte readback verification failed. Re-upload it before another edit.')
      } else {
        setStatus(mode === 'browser'
          ? 'The browser mutation was refused; no download was produced and no workbook bytes were uploaded.'
          : 'The server mutation was refused or the fallback API could not be reached; no download was produced.')
      }
    } finally {
      setBusy(false)
    }
  }

  const bounds = activeSheet ? previewBounds(activeSheet) : null
  const cellMap = new Map(activeSheet?.cells.map((cell) => [`${cell.row}:${cell.column}`, cell]) ?? [])
  const editableMap = new Map(targets.map((candidate) => [targetKey(candidate), candidate]))

  return (
    <div className="platen-fill native-demo workbench-surface" data-demo-surface="native">
      <div className="view-switcher" role="group" aria-label="XLSX processing runtime">
        <div className="tool-segment">
          <button type="button" aria-pressed={mode === 'browser'} disabled={busy} onClick={() => chooseMode('browser')}>In browser (default)</button>
          <button
            type="button"
            aria-pressed={mode === 'server'}
            disabled={busy || !SERVER_FALLBACK_CONFIGURED}
            title={SERVER_FALLBACK_CONFIGURED ? 'Upload to the configured XLSX API' : 'Set VITE_INJOFFICE_API_BASE to enable'}
            onClick={() => chooseMode('server')}
          >Server fallback{SERVER_FALLBACK_CONFIGURED ? '' : ' (not configured)'}</button>
        </div>
        <span>{mode === 'browser' ? 'Original bytes stay in this browser' : 'Uploads bytes to the configured API'}</span>
      </div>
      <div className="native-toolbar workbench-toolbar" role="group" aria-label="Native XLSX actions">
        <button type="button" className="workbench-button workbench-button--primary" disabled={busy} onClick={() => void loadSample()}>
          Use bundled .xlsx
        </button>
        <button type="button" className="workbench-button" disabled={busy} onClick={() => uploadRef.current?.click()}>
          Open .xlsx
        </button>
        <input
          ref={uploadRef}
          className="visually-hidden"
          type="file"
          accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          disabled={busy}
          aria-label="Open an XLSX file"
          onChange={(event) => {
            const file = event.currentTarget.files?.[0]
            if (file) void extractFile(file, file.name)
            event.currentTarget.value = ''
          }}
        />
        <span className="native-badge workbench-badge">{mode === 'browser' ? 'Browser-local · no upload' : 'Explicit server fallback'}</span>
      </div>
      <p className="native-status workbench-status" role="status" aria-live="polite" aria-atomic="true" data-state={error ? 'error' : busy ? 'busy' : workbook ? 'ready' : 'idle'}>
        {busy ? 'Working · ' : ''}{status}
      </p>
      <p className="native-help workbench-callout workbench-callout--warning">
        {mode === 'browser'
          ? <>The first operation loads the version-matched Go engine (about 6.3 MiB) in a Web Worker. A native feature refusal remains a refusal on the server path. Server fallback stays disabled until <code>VITE_INJOFFICE_API_BASE</code> is set.</>
          : <>This mode uploads the workbook. Local Vite uses its proxy; a static build needs a secured compatible API configured with <code>VITE_INJOFFICE_API_BASE</code>.</>}
      </p>
      {error && <div className="native-error workbench-callout workbench-callout--error" role="alert"><strong>{mode === 'browser' ? 'Browser engine' : 'Server response'}</strong><span>{error}</span></div>}

      <div className="native-workspace">
        <section className="native-main" aria-label="Extracted workbook preview">
          {workbook && activeSheet && bounds ? (
            <>
              <div className="native-sheet-heading">
                <div>
                  <span className="native-kicker">Exact native projection</span>
                  <h2>{activeSheet.name}</h2>
                </div>
                <span className="native-muted">{sourceName} · {workbook.source.authority}</span>
              </div>
              <div className="native-grid-wrap">
                <table className="native-grid">
                  <thead><tr><th aria-label="Row" />{Array.from({ length: bounds.columns }, (_, column) => <th key={column}>{columnName(column)}</th>)}</tr></thead>
                  <tbody>
                    {Array.from({ length: bounds.rows }, (_, row) => (
                      <tr key={row}>
                        <th>{row + 1}</th>
                        {Array.from({ length: bounds.columns }, (__, column) => {
                          const cell = cellMap.get(`${row}:${column}`)
                          const candidate = activeSheet ? editableMap.get(targetKey({ sheetId: activeSheet.id, row, column })) : undefined
                          const active = candidate && targetKey(candidate) === targetKey(target ?? candidate)
                          return (
                            <td key={column} className={active ? 'native-cell-active' : undefined} style={previewCellStyle(workbook, cell)}>
                              {candidate ? (
                                <button type="button" onClick={() => chooseTarget(candidate)} aria-label={`Edit ${activeSheet.name} ${cell?.ref ?? `${columnName(column)}${row + 1}`}`}>
                                  {previewCellValue(workbook, cell) || '\u00a0'}
                                </button>
                              ) : <span>{previewCellValue(workbook, cell) || '\u00a0'}</span>}
                            </td>
                          )
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </>
          ) : (
            <div className="native-empty">
              <span className="native-kicker">Native XLSX round trip</span>
              <h2>Put real XLSX bytes through the engine.</h2>
              <p>The Go engine runs inside a browser Worker by default. It extracts a revision-bound model, applies one supported edit, reopens the saved package, and returns real <code>.xlsx</code> bytes without a server.</p>
              <button type="button" className="workbench-button workbench-button--primary" disabled={busy} onClick={() => void loadSample()}>Run the bundled proof</button>
            </div>
          )}
        </section>

        <aside className="native-side workbench-inspector" aria-label="Mutation controls">
          <div className="native-panel">
            <span className="native-kicker">01 · Extract</span>
            {workbook ? (
              <dl className="native-proof-list">
                <div className="native-proof-row"><dt>Document</dt><dd>{workbook.document_id.slice(0, 20)}…</dd></div>
                <div className="native-proof-row"><dt>Revision</dt><dd>{shortRevision(workbook.revision)}</dd></div>
                <div className="native-proof-row"><dt>Preserved warnings</dt><dd>{workbook.unsupported?.length ?? 0}</dd></div>
                <div className="native-proof-row"><dt>Runtime</dt><dd>{mode === 'browser' ? 'browser-local' : 'server fallback'}</dd></div>
                <div className="native-proof-row"><dt>Artifact</dt><dd>{mode === 'browser' ? 'browser memory' : artifactId ? 'stored' : 'request-local'}</dd></div>
              </dl>
            ) : <p className="native-muted">No workbook extracted yet.</p>}
          </div>

          <div className="native-panel">
            <span className="native-kicker">02 · Mutate</span>
            {target ? (
              <>
                <label className="native-field">
                  Safe cell
                  <select value={targetKey(target)} onChange={(event) => {
                    const next = targets.find((candidate) => targetKey(candidate) === event.target.value)
                    if (next) chooseTarget(next)
                  }}>
                    {targets.map((candidate) => <option key={targetKey(candidate)} value={targetKey(candidate)}>{candidate.sheetName}!{candidate.ref}</option>)}
                  </select>
                </label>
                <label className="native-field">
                  New literal value
                  <input value={draft} maxLength={32767} onChange={(event) => setDraft(event.target.value)} />
                </label>
                <div className="native-actions">
                  <button type="button" className="workbench-button workbench-button--primary" disabled={busy || draft === targetValue} onClick={() => void mutate()}>Save to XLSX</button>
                </div>
                {draft === targetValue && <p className="native-muted">Change the value to enable a non-empty transaction.</p>}
              </>
            ) : <p className="native-muted">This workbook exposes no safely editable literal cells. That refusal is intentional.</p>}
          </div>

          <div className="native-panel">
            <span className="native-kicker">03 · Verify</span>
            {proof ? (
              <>
                <dl className="native-proof-list">
                  <div className="native-proof-row"><dt>Cell</dt><dd>{proof.cell}</dd></div>
                  <div className="native-proof-row"><dt>Before</dt><dd>{proof.before || 'blank'}</dd></div>
                  <div className="native-proof-row"><dt>After</dt><dd>{proof.after || 'blank'}</dd></div>
                  <div className="native-proof-row"><dt>CAS moved</dt><dd>{shortRevision(proof.previousRevision)} → {shortRevision(proof.revision)}</dd></div>
                </dl>
                {downloadURL && <a className="native-download workbench-button workbench-button--primary" href={downloadURL} download={downloadName(sourceName)}>Download verified .xlsx</a>}
              </>
            ) : (
              <>
                <p className="native-muted">After save, the page re-extracts the exact response bytes and shows the revision change here.</p>
                {downloadURL && <a className="native-download workbench-button" href={downloadURL} download={downloadName(sourceName)}>Download saved .xlsx</a>}
              </>
            )}
          </div>
        </aside>
      </div>
    </div>
  )
}

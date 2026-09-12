import { useEffect, useMemo, useRef, useState, type CSSProperties } from 'react'
import { adaptWorkbookMutationBatchV1 } from '@injoffice/xlsx-wasm'
import {
  DsButton,
  DsCallout,
  DsChip,
  DsField,
  DsInput,
  DsSelect,
} from '../design-system/primitives'
import '../design-system/live-create-edit.css'
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
import { nativeCellPreview } from '../nativeCellPreview'
import { NativeWorkbookObjects } from '../components/NativeWorkbookObjects'
import { NativeSheetPages } from '../components/NativeSheetPages'
import {nativeStoredRowPreviewV1,nativeTableFillPreview,nativeTableHeaderTextPreview,nativeTableTotalsTextPreview,nativeTableBorderPreview,type NativeWorkbookObjectsV1,type NativeTableBorderSideV1} from '@injoffice/sheets/browser'

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

function previewBounds(sheet: NativeSheet,objects?:NativeWorkbookObjectsV1|null): { rows: number; columns: number } {
  let lastRow=sheet.cells.reduce((max,cell)=>Math.max(max,cell.row),-1),lastColumn=sheet.cells.reduce((max,cell)=>Math.max(max,cell.column),-1)
  for(const table of objects?.tables??[]){if(table.sheet_part!==sheet.part_name)continue;const match=/^[A-Z]{1,3}[1-9][0-9]*:([A-Z]{1,3})([1-9][0-9]*)$/.exec(table.ref);if(match){lastRow=Math.max(lastRow,Number(match[2])-1);lastColumn=Math.max(lastColumn,[...match[1]!].reduce((n,c)=>n*26+c.charCodeAt(0)-64,0)-1)}}
  const rows = Math.max(3, Math.min(32, 1 + lastRow))
  const columns = Math.max(3, Math.min(12, 1 + lastColumn))
  return { rows, columns }
}

function previewCellStyle(workbook: NativeWorkbook, cell: NativeCell | undefined,implicitStyleID?:number): CSSProperties | undefined {
  const id=cell?.style_id??implicitStyleID
  if (id===undefined) return undefined
  const style = workbook.styles[id]?.effective
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
  const [objects,setObjects]=useState<NativeWorkbookObjectsV1|null>(null)
  useEffect(()=>setObjects(null),[workbook,mode])
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

  const bounds = activeSheet ? previewBounds(activeSheet,objects?.package_sha256===workbook?.source.package_sha256?objects:undefined) : null
  const cellMap = new Map(activeSheet?.cells.map((cell) => [`${cell.row}:${cell.column}`, cell]) ?? [])
  const rowStyles=new Map(activeSheet?.rows.filter(r=>r.style_id!==undefined).map(r=>[r.row,r.style_id])??[])
  const styleAt=(row:number,column:number):number|null|undefined=>{
    if(row<0||column<0||row>=1048576||column>=16384)return null
    const cell=cellMap.get(`${row}:${column}`);if(cell)return cell.style_id
    const rowStyle=rowStyles.get(row),columnStyles=activeSheet?.columns.filter(c=>c.column<=column&&column<=c.end_column&&c.style_id!==undefined).map(c=>c.style_id!)??[]
    if(new Set(columnStyles).size>1||rowStyle!==undefined&&columnStyles.length>0&&columnStyles[0]!==rowStyle)return undefined
    return rowStyle??columnStyles[0]??0
  }
  const edgeCSS=(edge:NativeTableBorderSideV1)=>`${edge.widthPoints}pt ${edge.style} ${edge.color}`
  const editableMap = new Map(targets.map((candidate) => [targetKey(candidate), candidate]))
  const previewWarnings = workbook && activeSheet && bounds
    ? activeSheet.cells.filter((cell) => cell.row < bounds.rows && cell.column < bounds.columns)
      .flatMap((cell) => {
        const warning = nativeCellPreview(workbook, cell,objects,activeSheet.part_name).warning
        return warning ? [{ ref: cell.ref, warning }] : []
      })
    : []

  return (
    <div className="platen-fill native-demo workbench-surface ds" data-demo-surface="native" data-demo-busy={busy}>
      <div className="view-switcher ds-workstrip" role="group" aria-label="XLSX processing runtime">
        <div className="tool-segment ds-segment" role="group" aria-label="XLSX processing runtime">
          <button type="button" aria-pressed={mode === 'browser'} disabled={busy} onClick={() => chooseMode('browser')}>In browser (default)</button>
          <button
            type="button"
            aria-pressed={mode === 'server'}
            disabled={busy || !SERVER_FALLBACK_CONFIGURED}
            title={SERVER_FALLBACK_CONFIGURED ? 'Upload to the configured XLSX API' : 'Set VITE_INJOFFICE_API_BASE to enable'}
            onClick={() => chooseMode('server')}
          >Server fallback{SERVER_FALLBACK_CONFIGURED ? '' : ' (not configured)'}</button>
        </div>
        <span className="ds-muted">{mode === 'browser' ? 'Original bytes stay in this browser' : 'Uploads bytes to the configured API'}</span>
      </div>
      <div className="native-toolbar workbench-toolbar ds-workstrip" role="group" aria-label="Native XLSX actions">
        <DsButton variant="filled" className="workbench-button workbench-button--primary" disabled={busy} onClick={() => void loadSample()}>
          Use bundled .xlsx
        </DsButton>
        <DsButton variant="outlined" className="workbench-button" disabled={busy} onClick={() => uploadRef.current?.click()}>
          Open .xlsx
        </DsButton>
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
        <DsChip>{mode === 'browser' ? 'Browser-local · no upload' : 'Explicit server fallback'}</DsChip>
      </div>
      <p className="native-status workbench-status ds-status" role="status" aria-live="polite" aria-atomic="true" data-state={error ? 'error' : busy ? 'busy' : workbook ? 'ready' : 'idle'}>
        {busy ? 'Working · ' : ''}{status}
      </p>
      <details className="ds-panel" data-xlsx-technical>
        <summary>Technical details</summary>
        <DsCallout
          tone="note"
          title={mode === 'browser' ? 'Browser-local engine' : 'Explicit server fallback'}
        >
          {mode === 'browser'
            ? 'The first operation downloads the version-matched Go engine and runs it in a browser Worker. Unsupported features remain unsupported on the server path.'
            : 'This mode uploads the workbook to the configured API. Unsupported features remain unsupported on this path.'}
        </DsCallout>
        <p className="ds-muted">{SERVER_FALLBACK_CONFIGURED
          ? 'A server API is configured. Selecting Server fallback uploads the workbook; it does not happen automatically.'
          : 'Server fallback is unavailable in this build. Developers can configure it with VITE_INJOFFICE_API_BASE.'}</p>
      </details>
      {error && <DsCallout tone="refuse" title={mode === 'browser' ? 'Browser engine' : 'Server response'}>{error}</DsCallout>}

      <div className="native-workspace ds-split">
        <section className={`native-main ds-split-main${workbook && activeSheet && bounds ? ' ds-split-main--flush' : ''}`} aria-label="Extracted workbook preview">
          {workbook && activeSheet && bounds ? (
            <>
              <div className="native-sheet-heading">
                <div>
                  <span className="native-kicker ds-eyebrow">Spreadsheet preview · limited layout</span>
                  <h2>{activeSheet.name}</h2>
                </div>
                <span className="native-muted ds-muted">{sourceName} · {workbook.source.authority}</span>
              </div>
              <div className="native-grid-wrap">
                <table className="native-grid">
                  <thead><tr><th aria-label="Row" />{Array.from({ length: bounds.columns }, (_, column) => <th key={column}>{columnName(column)}</th>)}</tr></thead>
                  <tbody>
                    {Array.from({ length: bounds.rows }, (_, row) => {
                      const stored=objects?nativeStoredRowPreviewV1(objects,workbook.source.package_sha256,activeSheet.part_name,row):undefined
                      return (
                      <tr key={row} className={stored?'native-stored-row':undefined} style={stored?{'--native-stored-row-height':`${stored.height_points}pt`,height:`${stored.height_points}pt`,...(stored.hidden?{display:'none'}:{})} as CSSProperties:undefined}>
                        <th><span>{row + 1}</span></th>
                        {Array.from({ length: bounds.columns }, (__, column) => {
                          const cell = cellMap.get(`${row}:${column}`)
                          const candidate = activeSheet ? editableMap.get(targetKey({ sheetId: activeSheet.id, row, column })) : undefined
                          const active = candidate && targetKey(candidate) === targetKey(target ?? candidate)
                          const preview = nativeCellPreview(workbook, cell,objects,activeSheet.part_name)
                          const styleID=styleAt(row,column),fill=styleID===null||styleID===undefined?undefined:workbook.styles[styleID]?.effective.fill
                          const tableFill=objects?nativeTableFillPreview(objects,workbook.source.package_sha256,activeSheet.part_name,row,column,fill,styleID??-1):undefined
                          const tableTotals=objects?nativeTableTotalsTextPreview(objects,workbook.source.package_sha256,activeSheet.part_name,row,column,styleID??-1):false
                          const tableHeader=objects?nativeTableHeaderTextPreview(objects,workbook.source.package_sha256,activeSheet.part_name,row,column,fill,styleID??-1):false
                          const tableEdges=objects?nativeTableBorderPreview(objects,workbook.source.package_sha256,activeSheet.part_name,row,column,styleID??-1,{top:styleAt(row-1,column),right:styleAt(row,column+1),bottom:styleAt(row+1,column),left:styleAt(row,column-1)}):undefined
                          return (
                            <td key={column} className={active ? 'native-cell-active' : undefined} style={{...previewCellStyle(workbook, cell,styleID??undefined),...(tableFill?{backgroundColor:tableFill}:{}),...(tableHeader?{color:'#FFFFFF',fontWeight:700}:{}),...(tableTotals?{fontWeight:700}:{}),...(tableEdges?.top?{borderTop:edgeCSS(tableEdges.top)}:{}),...(tableEdges?.right?{borderRight:edgeCSS(tableEdges.right)}:{}),...(tableEdges?.bottom?{borderBottom:edgeCSS(tableEdges.bottom)}:{}),...(tableEdges?.left?{borderLeft:edgeCSS(tableEdges.left)}:{})}} title={preview.warning ?? (preview.cached ? 'Saved formula result; not recalculated.' : undefined)}>
                              {candidate ? (
                                <button type="button" style={tableFill?{background:'transparent'}:undefined} onClick={() => chooseTarget(candidate)} aria-label={`Edit ${activeSheet.name} ${cell?.ref ?? `${columnName(column)}${row + 1}`}`}>
                                  {preview.text || '\u00a0'}
                                  {preview.warning && <sup role="img" aria-label={preview.warning}> ⚠</sup>}
                                </button>
                              ) : <span>{preview.text || '\u00a0'}{preview.warning && <sup role="img" aria-label={preview.warning}> ⚠</sup>}</span>}
                            </td>
                          )
                        })}
                      </tr>
                    )})}
                  </tbody>
                </table>
              </div>
              <p className="ds-muted">Formula cells show saved results, which may be stale; this preview does not recalculate. Warning markers identify raw values or missing saved results.</p>
              <p className="ds-muted">The grid preview is limited to the first 32 rows and 12 columns. Original content outside this window remains in the file.</p>
              {objects?.package_sha256===workbook.source.package_sha256&&<p className="ds-muted">{objects.row_geometry?.find(sheet=>sheet.sheet_part===activeSheet.part_name)?.warnings.join(' ')??'Stored row geometry is unavailable for this sheet; host preview sizes remain in use.'} Text may be clipped to stored heights; no font metrics or automatic fitting are inferred.</p>}
              {authoritativeBytes&&<NativeWorkbookObjects bytes={authoritativeBytes} revision={workbook.source.package_sha256} mode={mode} onInspection={setObjects} inspect={(bytes,revision)=>{const inspect=runtimeFor(mode).inspectObjects;if(!inspect)return Promise.reject(new Error('This runtime does not support object inspection'));return inspect(bytes,revision)}}/>}
              {objects?.package_sha256 === workbook.source.package_sha256 && <NativeSheetPages workbook={workbook} sheet={activeSheet} objects={objects} rows={bounds.rows} columns={bounds.columns}/>}
              {previewWarnings.length > 0 && <details className="ds-muted">
                <summary>{previewWarnings.length} preview cell warnings</summary>
                <ul>{previewWarnings.slice(0, 12).map(({ ref, warning }) => <li key={ref}><strong>{ref}:</strong> {warning}</li>)}</ul>
                {previewWarnings.length > 12 && <p>{previewWarnings.length - 12} more cells are marked in the preview; hover their markers for details.</p>}
              </details>}
            </>
          ) : (
            <div className="native-empty">
              <span className="native-kicker ds-eyebrow">XLSX files</span>
              <h2>Open a spreadsheet</h2>
              <p>Open an <code>.xlsx</code> file using the toolbar, or try the sample below. Preview its cells, make a supported edit, and download the updated file.</p>
              <DsButton variant="filled" className="workbench-button workbench-button--primary" disabled={busy} onClick={() => void loadSample()}>Try sample spreadsheet</DsButton>
            </div>
          )}
        </section>

        <aside className="native-side workbench-inspector ds-split-side" aria-label="Mutation controls">
          <div className="native-panel ds-panel">
            <span className="native-kicker ds-eyebrow">01 · Extract</span>
            {workbook ? (
              <dl className="native-proof-list ds-proof">
                <div className="native-proof-row"><dt>Document</dt><dd>{workbook.document_id.slice(0, 20)}…</dd></div>
                <div className="native-proof-row"><dt>Revision</dt><dd>{shortRevision(workbook.revision)}</dd></div>
                <div className="native-proof-row"><dt>Warnings</dt><dd>{workbook.unsupported?.length ?? 0}</dd></div>
                <div className="native-proof-row"><dt>Runtime</dt><dd>{mode === 'browser' ? 'browser-local' : 'server fallback'}</dd></div>
                <div className="native-proof-row"><dt>Artifact</dt><dd>{mode === 'browser' ? 'browser memory' : artifactId ? 'stored' : 'request-local'}</dd></div>
              </dl>
            ) : <p className="native-muted ds-muted">No workbook extracted yet.</p>}
          </div>

          <div className="native-panel ds-panel">
            <span className="native-kicker ds-eyebrow">02 · Mutate</span>
            {target ? (
              <>
                <DsField className="native-field" label="Safe cell">
                  <DsSelect value={targetKey(target)} onChange={(event) => {
                    const next = targets.find((candidate) => targetKey(candidate) === event.target.value)
                    if (next) chooseTarget(next)
                  }}>
                    {targets.map((candidate) => <option key={targetKey(candidate)} value={targetKey(candidate)}>{candidate.sheetName}!{candidate.ref}</option>)}
                  </DsSelect>
                </DsField>
                <DsField className="native-field" label="New literal value">
                  <DsInput value={draft} maxLength={32767} onChange={(event) => setDraft(event.target.value)} />
                </DsField>
                <div className="native-actions">
                  <DsButton variant="green" className="workbench-button workbench-button--primary" disabled={busy || draft === targetValue} onClick={() => void mutate()}>Save to XLSX</DsButton>
                </div>
                {draft === targetValue && <p className="native-muted ds-muted">Change the value to enable a non-empty transaction.</p>}
              </>
            ) : <p className="native-muted ds-muted">This workbook exposes no safely editable literal cells. That refusal is intentional.</p>}
          </div>

          <div className="native-panel ds-panel">
            <span className="native-kicker ds-eyebrow">03 · Verify</span>
            {proof ? (
              <>
                <dl className="native-proof-list ds-proof">
                  <div className="native-proof-row"><dt>Cell</dt><dd>{proof.cell}</dd></div>
                  <div className="native-proof-row"><dt>Before</dt><dd>{proof.before || 'blank'}</dd></div>
                  <div className="native-proof-row"><dt>After</dt><dd>{proof.after || 'blank'}</dd></div>
                  <div className="native-proof-row"><dt>CAS moved</dt><dd>{shortRevision(proof.previousRevision)} → {shortRevision(proof.revision)}</dd></div>
                </dl>
                <DsChip tone="green">Applied</DsChip>
                {downloadURL && <a className="native-download workbench-button workbench-button--primary ds-btn ds-btn--filled" href={downloadURL} download={downloadName(sourceName)}>Download verified .xlsx</a>}
              </>
            ) : (
              <>
                <p className="native-muted ds-muted">After save, the page re-extracts the exact response bytes and shows the revision change here.</p>
                {downloadURL && <a className="native-download workbench-button ds-btn ds-btn--outlined" href={downloadURL} download={downloadName(sourceName)}>Download saved .xlsx</a>}
              </>
            )}
          </div>
        </aside>
      </div>
    </div>
  )
}

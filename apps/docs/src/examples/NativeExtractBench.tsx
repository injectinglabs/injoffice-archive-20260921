import { useEffect, useMemo, useState } from 'react'
import { adaptWorkbookMutationBatchV1, createXlsxWasmClient } from '@injoffice/xlsx-wasm'
import type { NativeWorkbookV2 } from '@injoffice/sheets/browser'
import { LiveBench } from '../components/LiveBench'

const SAMPLE = `${import.meta.env.BASE_URL}native-fixture/launch-readiness-plan.xlsx`

function cellText(workbook: NativeWorkbookV2, sheetIndex = 0, cellIndex = 0): string {
  const cell = workbook.sheets[sheetIndex]?.cells[cellIndex]
  return cell?.value?.text ?? cell?.value?.lexical ?? cell?.formula?.text ?? ''
}

export function NativeExtractBench() {
  const [status, setStatus] = useState('Loading the bundled workbook in a browser Worker…')
  const [error, setError] = useState('')
  const [busy, setBusy] = useState(false)
  const [original, setOriginal] = useState<Uint8Array | null>(null)
  const [workbook, setWorkbook] = useState<NativeWorkbookV2 | null>(null)
  const [draft, setDraft] = useState('')
  const [download, setDownload] = useState<string | null>(null)
  const client = useMemo(() => createXlsxWasmClient(), [])

  useEffect(() => () => client.terminate(), [client])
  useEffect(() => () => { if (download) URL.revokeObjectURL(download) }, [download])

  useEffect(() => {
    let cancelled = false
    void (async () => {
      try {
        const response = await fetch(SAMPLE)
        if (!response.ok) throw new Error(`Could not load ${SAMPLE} (${response.status})`)
        const bytes = new Uint8Array(await response.arrayBuffer())
        const extracted = await client.extract(bytes)
        if (cancelled) return
        setOriginal(bytes)
        setWorkbook(extracted)
        setDraft(cellText(extracted))
        setStatus(`Extracted ${extracted.sheets.length} sheet${extracted.sheets.length === 1 ? '' : 's'} · ${extracted.source.package_sha256.slice(0, 19)}…`)
      } catch (reason) {
        if (!cancelled) {
          setError(reason instanceof Error ? reason.message : String(reason))
          setStatus('Browser extract did not complete. No sidecar was contacted.')
        }
      }
    })()
    return () => { cancelled = true }
  }, [client])

  const apply = async () => {
    if (!original || !workbook) return
    const sheet = workbook.sheets[0]
    const cell = sheet?.cells[0]
    if (!sheet || !cell) {
      setError('The extracted workbook has no editable sample cell.')
      return
    }
    setBusy(true)
    setError('')
    setStatus(`Applying cell.set_value to ${sheet.name}!${cell.ref}…`)
    try {
      const numeric = Number(draft)
      const value = draft.trim() !== '' && Number.isFinite(numeric) ? numeric : draft
      const transaction = adaptWorkbookMutationBatchV1(workbook, {
        protocol: 'injoffice.xlsx.mutations',
        version: 1,
        batch_id: 'docs-wasm-1',
        expected_revision: workbook.source.package_sha256,
        operations: [{
          operation_id: 'docs-edit-1',
          kind: 'cell.set_value',
          sheet_id: sheet.id,
          cell: { row: cell.row, column: cell.column },
          value,
        }],
      })
      const saved = await client.apply(original, workbook, transaction)
      const reopened = await client.extract(saved)
      setWorkbook(reopened)
      setOriginal(saved)
      setDraft(cellText(reopened))
      if (download) URL.revokeObjectURL(download)
      const copy = new Uint8Array(saved.byteLength)
      copy.set(saved)
      setDownload(URL.createObjectURL(new Blob([copy.buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })))
      setStatus(`Applied and re-extracted · ${reopened.source.package_sha256.slice(0, 19)}…`)
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason))
      setStatus('Apply refused. The original bytes were not replaced.')
    } finally {
      setBusy(false)
    }
  }

  const sheet = workbook?.sheets[0]
  const preview = sheet?.cells.slice(0, 8) ?? []

  return (
    <LiveBench title="Live example" hint="@injoffice/xlsx-wasm · extract + apply in a Worker">
      <div className="bench-controls">
        <label className="field" style={{ flex: 1 }}>
          First cell value
          <input value={draft} onChange={(event) => setDraft(event.target.value)} disabled={!workbook || busy} />
        </label>
        <button type="button" className="bench-button" onClick={() => void apply()} disabled={!workbook || busy}>Apply in the browser</button>
        {download ? <a className="bench-button" href={download} download="injoffice-docs.xlsx">Download .xlsx</a> : null}
      </div>
      <p>{status}</p>
      {error ? <p><span className="badge badge--patch">refused</span> {error}</p> : null}
      {sheet ? (
        <table className="grid-table">
          <thead><tr><th>Ref</th><th>Value</th></tr></thead>
          <tbody>
            {preview.map((cell) => (
              <tr key={cell.ref}><td>{cell.ref}</td><td>{cell.value?.text ?? cell.value?.lexical ?? cell.formula?.text ?? ''}</td></tr>
            ))}
          </tbody>
        </table>
      ) : null}
    </LiveBench>
  )
}

import { useEffect, useRef, useState, type CSSProperties } from 'react'
import { createXlsxSourceStylePreviewClient, type XlsxSourceStylePreviewClient, type XlsxSourceStylePreviewV1, type XlsxSourceStyleV1, type XlsxSourceStyleCellV1 } from '@injoffice/xlsx-wasm'
import { formatNativeSheetCellDisplayV2 } from '@injoffice/sheets/browser'
import { DsButton } from '../design-system/primitives'

export function sourceStyleCellText(cell: XlsxSourceStyleCellV1, style: XlsxSourceStyleV1, date1904: boolean): { text: string; warning?: string } {
  if (cell.kind === 'blank') return { text: '' }
  if (cell.kind === 'string') return { text: cell.text }
  if (cell.kind === 'boolean') return { text: cell.lexical === '1' ? 'TRUE' : cell.lexical === '0' ? 'FALSE' : cell.lexical }
  if (cell.kind === 'number' || cell.kind === 'date') {
    // This exact source-declared literal suffix is a browser display addition;
    // strict glyph formatting and workbook bytes are not modified.
    const euro = cell.kind === 'number' && style.number_format === '#,##0.00" €"'
    const result = formatNativeSheetCellDisplayV2(cell.kind, cell.lexical, euro ? '#,##0.00' : style.number_format, date1904)
    if (result.status === 'ready') return { text: result.text + (euro ? ' €' : '') }
    return { text: cell.lexical, warning: `${cell.ref}: showing the saved value without its unsupported number/date format.` }
  }
  return { text: cell.lexical || cell.text }
}

export function XlsxSourceStylePreview({ bytes, name }: { bytes: Uint8Array; name: string }) {
  const [preview, setPreview] = useState<XlsxSourceStylePreviewV1 | null>(null)
  const [message, setMessage] = useState('')
  const [busy, setBusy] = useState(false)
  const generation = useRef(0)
  const client = useRef<XlsxSourceStylePreviewClient | null>(null)
  useEffect(() => {
    generation.current++
    setPreview(null); setMessage(''); setBusy(false)
    return () => { generation.current++; client.current?.terminate(); client.current = null }
  }, [bytes])
  function cancel() {
    generation.current++; client.current?.terminate(); client.current = null
    setPreview(null); setBusy(false); setMessage('Preview cancelled.')
  }
  async function run() {
    const token = ++generation.current
    client.current?.terminate()
    const worker = createXlsxSourceStylePreviewClient()
    client.current = worker
    setBusy(true); setPreview(null); setMessage('Reading source styles locally…')
    try {
      const result = await worker.preview(bytes)
      if (token !== generation.current) return
      setPreview(result); setMessage('Read-only source grid ready. The original file is unchanged.')
    } catch (error) {
      if (token === generation.current) setMessage(error instanceof Error ? error.message : 'Source preview unavailable.')
    } finally {
      worker.terminate()
      if (client.current === worker) client.current = null
      if (token === generation.current) setBusy(false)
    }
  }
  return <section className="ds-panel" aria-label="Read-only source-style recovery">
    <h3>Read-only source preview</h3>
    <p>The file could not open for native editing. You can try a limited source grid for conflicting fill and number-format records. This runs locally in your browser.</p>
    <p>Fonts, widths, wrapping and borders are approximate. Saved formula results may be stale. Print settings are not applied.</p>
    <DsButton disabled={busy} onClick={() => void run()}>{busy ? 'Reading source…' : 'Preview source styles'}</DsButton>
    {busy && <DsButton onClick={cancel}>Cancel preview</DsButton>}
    <p role="status">{message}</p>
    {preview && <SourceStyleGrid preview={preview} name={name} />}
  </section>
}

export function SourceStyleGrid({ preview, name }: { preview: XlsxSourceStylePreviewV1; name: string }) {
  const sheet = preview.sheets[0]!
  const styles = new Map(preview.styles.map(style => [style.id, style]))
  const cells = new Map(sheet.cells.map(cell => [`${cell.row}:${cell.column}`, cell]))
  const merges = new Map(sheet.merges.map(merge => [`${merge.row}:${merge.column}`, merge]))
  const covered = new Set<string>()
  for (const merge of sheet.merges) for (let row = merge.row; row <= merge.end_row; row++) for (let col = merge.column; col <= merge.end_column; col++) if (row !== merge.row || col !== merge.column) covered.add(`${row}:${col}`)
  const text = new Map(sheet.cells.map(cell => [cell.ref, sourceStyleCellText(cell, styles.get(cell.style_id)!, preview.date1904)]))
  const warnings = [...new Set([...preview.warnings, ...preview.styles.flatMap(style => style.warnings), ...[...text.values()].flatMap(value => value.warning ? [value.warning] : [])])]
  const widths = sheet.column_widths.map(width => width * 7 + 5)
  return <>
    <h4>{name} · {sheet.name}</h4>
    <ul>{warnings.map(warning => <li key={warning}>{warning}</li>)}</ul>
    <div style={{ overflow: 'auto', maxHeight: 600 }}>
      <table aria-label="Approximate read-only source grid" style={{ tableLayout: 'fixed', borderCollapse: 'collapse', width: widths.reduce((a, b) => a + b, 0), background: '#fff', color: '#000' }}>
        <colgroup>{widths.map((width, index) => <col key={index} style={{ width }} />)}</colgroup>
        <tbody>{sheet.row_heights.map((height, row) => <tr key={row} style={{ height: height * 96 / 72 }}>
          {widths.map((_, column) => {
            const key = `${row}:${column}`
            if (covered.has(key)) return null
            const cell = cells.get(key), merge = merges.get(key), style = styles.get(cell?.style_id ?? 0)!
            const css: CSSProperties = {
              fontFamily: `"${style.font_name}", sans-serif`, fontSize: style.font_size_points * 96 / 72,
              color: style.font_color, background: style.fill_color || '#fff', fontWeight: style.bold ? 700 : 400, fontStyle: style.italic ? 'italic' : 'normal',
              textAlign: style.horizontal === 'general' ? cell?.kind === 'number' || cell?.kind === 'date' ? 'right' : 'left' : style.horizontal,
              verticalAlign: style.vertical === 'center' ? 'middle' : style.vertical,
              whiteSpace: style.wrap ? 'pre-wrap' : 'pre', padding: '1px 3px', overflow: 'hidden', boxSizing: 'border-box',
            }
            for (const side of ['left', 'right', 'top', 'bottom'] as const) {
              const color = style.borders[side]
              css[`border${side[0]!.toUpperCase()}${side.slice(1)}` as 'borderLeft'] = color ? `1px solid ${color}` : 'none'
            }
            return <td key={column} data-source-cell={cell?.ref} colSpan={merge ? merge.end_column - column + 1 : undefined} rowSpan={merge ? merge.end_row - row + 1 : undefined} style={css} title={cell ? `${cell.ref}${cell.cached ? ` · saved formula result · ${cell.formula}` : ''}` : undefined}>
              <div style={{ overflow: 'hidden', maxWidth: '100%' }}>{cell ? text.get(cell.ref)!.text : ''}</div>
            </td>
          })}
        </tr>)}</tbody>
      </table>
    </div>
    <details><summary>Source conflicts and saved values</summary>
      <p style={{ overflowWrap: 'anywhere' }}>Original file: <code>{preview.package_sha256}</code></p>
      <p>{preview.strict_error}</p>
      <ul>{preview.conflicts.map(conflict => <li key={`${conflict.style_id}:${conflict.component}`}>Style {conflict.style_id}: absent {conflict.apply_flag}; displaying direct {conflict.component} {conflict.direct_component_id} instead of parent {conflict.parent_component_id}.</li>)}</ul>
      <table className="ds-table"><thead><tr><th>Cell</th><th>Saved value</th><th>Formula (not recalculated)</th><th>Source format</th></tr></thead><tbody>{sheet.cells.map(cell => <tr key={cell.ref}><td>{cell.ref}</td><td>{cell.kind === 'string' ? cell.text : cell.lexical}</td><td>{cell.formula}</td><td>{styles.get(cell.style_id)!.number_format}</td></tr>)}</tbody></table>
    </details>
  </>
}

import { useEffect, useRef, useState, type CSSProperties } from 'react'
import { createXlsxSourceStylePreviewClient, createXlsxSourceStylePreviewV2Client, type XlsxSourceStylePreviewV2Client, type XlsxSourceStylePreviewV2, type XlsxSourceStylePreviewClient, type XlsxSourceStylePreviewV1, type XlsxSourceStyleV1, type XlsxSourceStyleCellV1 } from '@injoffice/xlsx-wasm'
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
  const [preview, setPreview] = useState<XlsxSourceStylePreviewV1 | XlsxSourceStylePreviewV2 | null>(null)
  const [message, setMessage] = useState('')
  const [busy, setBusy] = useState(false)
  const generation = useRef(0)
  const client = useRef<XlsxSourceStylePreviewClient | XlsxSourceStylePreviewV2Client | null>(null)
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
    setBusy(true); setPreview(null); setMessage('Reading source styles locally…')
    try {
      const factories = [createXlsxSourceStylePreviewClient, createXlsxSourceStylePreviewV2Client]
      for (let i = 0; i < factories.length; i++) {
        if (token !== generation.current) return
        const worker = factories[i]!()
        client.current = worker
        try {
          const result = await worker.preview(bytes)
          if (token !== generation.current) return
          setPreview(result); setMessage('Read-only source grid ready. The original file is unchanged.')
          return
        } catch (error) {
          if (token !== generation.current) return
          if (i === factories.length - 1) setMessage(error instanceof Error ? error.message : 'Source preview unavailable.')
        } finally {
          worker.terminate()
          if (client.current === worker) client.current = null
        }
      }
    } finally {
      if (token === generation.current) setBusy(false)
    }
  }
  return <section className="ds-panel" aria-label="Read-only source-style recovery">
    <h3>Read-only source preview</h3>
    <p>The file could not open for native editing. You can try a limited source grid for conflicting fill and number-format records, including qualified status colors and data bars. This runs locally in your browser.</p>
    <p>Fonts, widths, wrapping and borders are approximate. Saved formula results may be stale. Print settings are not applied.</p>
    <DsButton disabled={busy} onClick={() => void run()}>{busy ? 'Reading source…' : 'Preview source styles'}</DsButton>
    {busy && <DsButton onClick={cancel}>Cancel preview</DsButton>}
    <p role="status">{message}</p>
    {preview && <SourceStyleGrid preview={preview.version === 2 ? preview.grid : preview} conditional={preview.version === 2 ? preview : undefined} name={name} />}
  </section>
}

export function SourceStyleGrid({ preview, conditional, name }: { preview: XlsxSourceStylePreviewV1; conditional?: XlsxSourceStylePreviewV2; name: string }) {
  const sheet = preview.sheets[0]!
  const matches = new Set(conditional?.text_rule?.cells.filter(cell => cell.matched).map(cell => cell.ref) ?? [])
  const bars = new Map(conditional?.data_bar?.cells.map(cell => [cell.ref, cell]) ?? [])
  const styles = new Map(preview.styles.map(style => [style.id, style]))
  const cells = new Map(sheet.cells.map(cell => [`${cell.row}:${cell.column}`, cell]))
  const merges = new Map(sheet.merges.map(merge => [`${merge.row}:${merge.column}`, merge]))
  const covered = new Set<string>()
  for (const merge of sheet.merges) for (let row = merge.row; row <= merge.end_row; row++) for (let col = merge.column; col <= merge.end_column; col++) if (row !== merge.row || col !== merge.column) covered.add(`${row}:${col}`)
  const text = new Map(sheet.cells.map(cell => [cell.ref, sourceStyleCellText(cell, styles.get(cell.style_id)!, preview.date1904)]))
  const warnings = [...new Set([...preview.warnings, ...(conditional?.warnings ?? []), ...preview.styles.flatMap(style => style.warnings), ...[...text.values()].flatMap(value => value.warning ? [value.warning] : [])])]
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
            const matched = cell && matches.has(cell.ref), bar = cell && bars.get(cell.ref)
            const css: CSSProperties = {
              fontFamily: `"${style.font_name}", sans-serif`, fontSize: style.font_size_points * 96 / 72,
              color: matched ? conditional!.text_rule!.font_color : style.font_color, background: matched ? conditional!.text_rule!.background_color : style.fill_color || '#fff', fontWeight: matched || style.bold ? 700 : 400, fontStyle: style.italic ? 'italic' : 'normal',
              textAlign: style.horizontal === 'general' ? cell?.kind === 'number' || cell?.kind === 'date' ? 'right' : 'left' : style.horizontal,
              verticalAlign: style.vertical === 'center' ? 'middle' : style.vertical,
              position: 'relative', whiteSpace: style.wrap ? 'pre-wrap' : 'pre', padding: '1px 3px', overflow: 'hidden', boxSizing: 'border-box',
            }
            for (const side of ['left', 'right', 'top', 'bottom'] as const) {
              const color = style.borders[side]
              css[`border${side[0]!.toUpperCase()}${side.slice(1)}` as 'borderLeft'] = color ? `1px solid ${color}` : 'none'
            }
            return <td key={column} data-source-cell={cell?.ref} data-source-conditional-match={matched ? "true" : undefined} colSpan={merge ? merge.end_column - column + 1 : undefined} rowSpan={merge ? merge.end_row - row + 1 : undefined} style={css} title={cell ? `${cell.ref}${cell.cached ? ` · saved formula result · ${cell.formula}` : ''}` : undefined}>
              {bar && <span aria-hidden="true" data-source-bar={cell!.ref} data-source-bar-percent={bar.length_percent} style={{ position: 'absolute', left: 0, top: 0, bottom: 0, width: `${bar.length_percent}%`, background: `linear-gradient(to right, ${conditional!.data_bar!.color}, #fff)`, pointerEvents: 'none' }} />}
              <div style={{ position: 'relative', overflow: 'hidden', maxWidth: '100%' }}>{cell ? text.get(cell.ref)!.text : ''}</div>
            </td>
          })}
        </tr>)}</tbody>
      </table>
    </div>
    {conditional && <details><summary>Source conditional rules</summary>
      <p style={{ overflowWrap: 'anywhere' }}>Styles part: <code>{conditional.styles_sha256}</code><br />Worksheet part: <code>{conditional.worksheet_sha256}</code></p>
      {conditional.text_rule && <p>Range {conditional.text_rule.range}: saved strings equal to “{conditional.text_rule.text}” receive bold text and the declared font/background colors. {matches.size} of {conditional.text_rule.cells.length} stored strings match.</p>}
      {conditional.data_bar && <p>Range {conditional.data_bar.range}: {conditional.data_bar.cells.length} stored integer values use bounds {conditional.data_bar.minimum}–{conditional.data_bar.maximum} and lengths {conditional.data_bar.min_length}–{conditional.data_bar.max_length}%. The gradient width is approximate.</p>}
      {conditional.frozen_view && <p>The source freezes {conditional.frozen_view.frozen_rows} rows. This grid scrolls normally.</p>}
    </details>}
    <details><summary>Source conflicts and saved values</summary>
      <p style={{ overflowWrap: 'anywhere' }}>Original file: <code>{preview.package_sha256}</code></p>
      <p>{preview.strict_error}</p>
      <ul>{preview.conflicts.map(conflict => <li key={`${conflict.style_id}:${conflict.component}`}>Style {conflict.style_id}: absent {conflict.apply_flag}; displaying direct {conflict.component} {conflict.direct_component_id} instead of parent {conflict.parent_component_id}.</li>)}</ul>
      <table className="ds-table"><thead><tr><th>Cell</th><th>Saved value</th><th>Formula (not recalculated)</th><th>Source format</th></tr></thead><tbody>{sheet.cells.map(cell => <tr key={cell.ref}><td>{cell.ref}</td><td>{cell.kind === 'string' ? cell.text : cell.lexical}</td><td>{cell.formula}</td><td>{styles.get(cell.style_id)!.number_format}</td></tr>)}</tbody></table>
    </details>
  </>
}

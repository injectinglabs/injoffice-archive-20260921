import { useEffect, useId, useRef, useState } from 'react'
import {
  projectNativeWorkbookV2, createNativeMaximumDigitWidthAuthorityV2,
  compileNativeSheetGeometryV2, compileNativeStoredRowSheetGeometryV1, compileNativeSheetPagePreviewV1,
  nativeTableFillPreview, nativeTableHeaderTextPreview, nativeTableTotalsTextPreview,
  type NativeWorkbookObjectsV1, type NativeSheetGeometryV2,
  type NativeSheetPagePreviewV1, type NativeSheetHostPagePolicyV1,
} from '@injoffice/sheets/browser'
import type { NativeWorkbook, NativeSheet } from '../nativeRoundTrip'
import { nativeCellPreview } from '../nativeCellPreview'
import { DsButton, DsSelect, DsInput } from '../design-system/primitives'
import './native-sheet-pages.css'

const EMU_PER_PIXEL = 9525
const MAX_FONT_BYTES = 32 * 1024 * 1024
type Props = { workbook: NativeWorkbook; sheet: NativeSheet; objects: NativeWorkbookObjectsV1; rows: number; columns: number }
type Result = { geometry: NativeSheetGeometryV2; plan: NativeSheetPagePreviewV1; fontFamily: string }

// Key the inner view by source identity: pending font reads and page choices never
// carry over to a different sheet/revision, even before effect cleanup runs.
export function NativeSheetPages(props: Props) {
  return <SheetPagesSession key={`${props.workbook.source.package_sha256}:${props.sheet.id}`} {...props}/>
}

function SheetPagesSession({ workbook, sheet, objects, rows, columns }: Props) {
  const [font, setFont] = useState<File | null>(null)
  const [paper, setPaper] = useState<'A4' | 'Letter'>('A4')
  const [orientation, setOrientation] = useState<'portrait' | 'landscape'>('portrait')
  const [scale, setScale] = useState('100')
  const [useSource, setUseSource] = useState(true)
  const [useStoredRows, setUseStoredRows] = useState(false)
  const [result, setResult] = useState<Result | null>(null)
  const [message, setMessage] = useState('')
  const [busy, setBusy] = useState(false)
  const generation = useRef(0)
  const installed = useRef<FontFace | null>(null)
  const instance = useId().replace(/[^a-zA-Z0-9]/g, '')
  useEffect(() => () => {
    generation.current++
    if (installed.current) document.fonts.delete(installed.current)
  }, [])
  function invalidate() { generation.current++; setResult(null); setMessage(''); setBusy(false) }
  async function renderPages() {
    const token = ++generation.current
    setResult(null); setBusy(true); setMessage('Reading font metrics and laying out the selected range…')
    let loaded: FontFace | undefined
    try {
      if (!font || font.size < 12 || font.size > MAX_FONT_BYTES) throw new Error('Choose the workbook’s Normal font as a standalone TrueType file, up to 32 MiB.')
      const bytes = new Uint8Array(await font.arrayBuffer())
      if (generation.current !== token) return
      const model = projectNativeWorkbookV2(workbook)
      const authority = createNativeMaximumDigitWidthAuthorityV2(model, bytes)
      const viewport = { row: 0, column: 0, end_row: rows - 1, end_column: columns - 1 }
      const geometry = useStoredRows
        ? compileNativeStoredRowSheetGeometryV1(model, sheet.id, viewport, authority, objects)
        : compileNativeSheetGeometryV2(model, sheet.id, viewport, authority)
      const host: NativeSheetHostPagePolicyV1 | undefined = useSource ? undefined : {
        kind: 'explicit-host-page-policy-v1', paper, orientation, scale: Number(scale),
        left_inches: 0.5, right_inches: 0.5, top_inches: 0.5, bottom_inches: 0.5,
      }
      const plan = compileNativeSheetPagePreviewV1(geometry, objects, host)
      const fontFamily = `injoffice-sheet-${instance}-${token}`
      loaded = await new FontFace(fontFamily, bytes.buffer, {
        weight: model.normal_style?.font_bold ? '700' : '400',
        style: model.normal_style?.font_italic ? 'italic' : 'normal',
      }).load()
      if (generation.current !== token) return
      if (installed.current) document.fonts.delete(installed.current)
      document.fonts.add(loaded); installed.current = loaded
      setResult({ geometry, plan, fontFamily })
      setMessage(`${plan.pages.length} preview pages. Read-only; the workbook is unchanged.`)
    } catch (error) {
      if (generation.current === token) setMessage(error instanceof Error ? error.message : 'Page preview unavailable.')
    } finally {
      if (generation.current === token) setBusy(false)
    }
  }
  return <section className="ds-panel native-sheet-pages" aria-label="Spreadsheet page preview">
    <h3>Page preview</h3>
    <p className="ds-muted">Preview the first {rows} rows and {columns} columns on paper. Column widths use the exact Normal font; text uses approximate browser layout. This is not Excel print fidelity.</p>
    <div className="native-sheet-page-controls">
      <label>Normal font: {workbook.normal_style?.font_name || 'not available'}<input type="file" accept=".ttf,font/ttf" onChange={event => { invalidate(); setFont(event.target.files?.[0] ?? null) }}/></label>
      <label><input type="checkbox" checked={useStoredRows} onChange={event => { invalidate(); setUseStoredRows(event.target.checked) }}/> Allow approximate stored-row layout</label>
      <label><input type="checkbox" checked={useSource} onChange={event => { invalidate(); setUseSource(event.target.checked) }}/> Use saved page settings</label>
      {!useSource && <>
        <label>Paper<DsSelect value={paper} onChange={event => { invalidate(); setPaper(event.target.value as 'A4' | 'Letter') }}><option>A4</option><option>Letter</option></DsSelect></label>
        <label>Orientation<DsSelect value={orientation} onChange={event => { invalidate(); setOrientation(event.target.value as 'portrait' | 'landscape') }}><option value="portrait">Portrait</option><option value="landscape">Landscape</option></DsSelect></label>
        <label>Scale (%)<DsInput type="number" min={10} max={400} step={1} value={scale} onChange={event => { invalidate(); setScale(event.target.value) }}/></label>
        <p className="ds-muted">Your page choices use 0.5-inch margins. These are preview choices, not workbook defaults.</p>
      </>}
      <DsButton disabled={busy || !font} onClick={() => void renderPages()}>{busy ? 'Preparing pages…' : 'Preview pages'}</DsButton>
    </div>
    <p className="ds-muted">The font stays in this browser and is not saved in the workbook. Other fonts may be substituted by the browser. Charts, drawings, headers, print titles and print areas are not drawn here. Formula values are saved caches, not recalculated results.</p>
    <p role="status">{message}</p>
    {result && <>
      <p>{result.plan.settings_origin === 'source' ? 'Saved paper, margins and scale' : 'Your paper, margins and scale'} · approximate selected-range preview</p>
      <details><summary>Page preview limitations</summary><ul>{result.plan.warnings.map((warning, index) => <li key={index}>{warning}</li>)}<li>Text is single-line and clipped to cells; wrapping, rotation and text overflow are not reproduced. Unsupported styles and rich runs may differ. Cell text longer than 2,048 characters is truncated in this view.</li></ul></details>
      <NativeSheetPageImages {...{workbook, sheet, objects}} {...result}/>
    </>}
  </section>
}

export function NativeSheetPageImages({ workbook, sheet, objects, geometry, plan, fontFamily }: Pick<Props, 'workbook' | 'sheet' | 'objects'> & Result) {
  const uid = useId().replace(/[^a-zA-Z0-9]/g, '')
  if (plan.source_package_sha256 !== workbook.source.package_sha256 || geometry.source_package_sha256 !== workbook.source.package_sha256 || objects.package_sha256 !== workbook.source.package_sha256 || plan.sheet_id !== sheet.id || geometry.sheet_id !== sheet.id || plan.geometry_sha256 !== geometry.geometry_sha256) return <p role="alert">Page preview no longer matches this workbook.</p>
  const cellMap = new Map(sheet.cells.map(cell => [`${cell.row}:${cell.column}`, cell]))
  const styleAt = (row: number, column: number) => cellMap.get(`${row}:${column}`)?.style_id ?? sheet.rows.find(r => r.row === row)?.style_id ?? sheet.columns.find(c => column >= c.column && column <= c.end_column)?.style_id ?? 0
  const cells = geometry.rows.flatMap(row => geometry.columns.flatMap(column => {
    if (row.hidden || column.hidden) return []
    const merge = geometry.merged_ranges.find(m => row.row >= m.row && row.row <= m.end_row && column.column >= m.column && column.column <= m.end_column)
    if (merge && (row.row !== merge.row || column.column !== merge.column)) return []
    const cell = cellMap.get(`${row.row}:${column.column}`), styleId = styleAt(row.row, column.column)
    const style = workbook.styles[styleId]?.effective
    const display = nativeCellPreview(workbook, cell, objects, sheet.part_name)
    const fill = nativeTableFillPreview(objects, workbook.source.package_sha256, sheet.part_name, row.row, column.column, style?.fill, styleId) ?? style?.fill_color ?? '#FFFFFF'
    const header = nativeTableHeaderTextPreview(objects, workbook.source.package_sha256, sheet.part_name, row.row, column.column, style?.fill, styleId)
    const totals = nativeTableTotalsTextPreview(objects, workbook.source.package_sha256, sheet.part_name, row.row, column.column, styleId)
    const rect = merge?.rect ?? { x_emu: column.x_emu, y_emu: row.y_emu, width_emu: column.width_emu, height_emu: row.height_emu }
    return [{ key: `${row.row}-${column.column}`, row: row.row, column: column.column, rect, style, display, fill, header, totals }]
  }))
  return <div className="native-sheet-page-list">{plan.pages.map(page => {
    const clip = `${uid}-page-${page.number}`
    return <figure key={page.number}>
      <figcaption>Page {page.number} · rows {page.rows.start + 1}–{page.rows.end + 1}, columns {page.columns.start + 1}–{page.columns.end + 1}</figcaption>
      <svg role="img" aria-label={`Approximate spreadsheet page ${page.number}`} viewBox={`0 0 ${page.width_emu / EMU_PER_PIXEL} ${page.height_emu / EMU_PER_PIXEL}`}>
        <rect width="100%" height="100%" fill="#FFFFFF"/>
        <defs><clipPath id={clip}><rect x={page.source_clip.x_emu / EMU_PER_PIXEL} y={page.source_clip.y_emu / EMU_PER_PIXEL} width={page.source_clip.width_emu / EMU_PER_PIXEL} height={page.source_clip.height_emu / EMU_PER_PIXEL}/></clipPath></defs>
        <g transform={`translate(${page.translate_x_emu / EMU_PER_PIXEL} ${page.translate_y_emu / EMU_PER_PIXEL}) scale(${page.scale})`} clipPath={`url(#${clip})`}>
          {cells.filter(cell => cell.row >= page.rows.start && cell.row <= page.rows.end && cell.column >= page.columns.start && cell.column <= page.columns.end).map(cell => {
            const x = cell.rect.x_emu / EMU_PER_PIXEL, y = cell.rect.y_emu / EMU_PER_PIXEL, w = cell.rect.width_emu / EMU_PER_PIXEL, h = cell.rect.height_emu / EMU_PER_PIXEL
            const id = `${clip}-${cell.key}`, size = (cell.style?.font_size_points ?? 11) * 96 / 72
            const right = cell.style?.horizontal_alignment === 'right', center = cell.style?.horizontal_alignment === 'center'
            const exactNormal = cell.style?.font_name === workbook.normal_style?.font_name && Boolean(cell.style?.bold) === Boolean(workbook.normal_style?.font_bold) && Boolean(cell.style?.italic) === Boolean(workbook.normal_style?.font_italic)
            return <g key={cell.key}><title>{cell.display.text}{cell.display.warning ? ` — ${cell.display.warning}` : ''}</title>
              <rect x={x} y={y} width={w} height={h} fill={cell.fill}/>
              <clipPath id={id}><rect x={x} y={y} width={w} height={h}/></clipPath>
              <text clipPath={`url(#${id})`} x={right ? x + w - 2 : center ? x + w / 2 : x + 2} y={cell.style?.vertical_alignment === 'top' ? y + size : cell.style?.vertical_alignment === 'middle' ? y + (h + size) / 2 - 2 : y + h - 2} textAnchor={right ? 'end' : center ? 'middle' : 'start'} fontFamily={exactNormal ? fontFamily : cell.style?.font_name || 'sans-serif'} fontSize={size} fontWeight={cell.header || cell.totals || cell.style?.bold ? 700 : 400} fontStyle={cell.style?.italic ? 'italic' : 'normal'} fill={cell.header ? '#FFFFFF' : cell.style?.font_color || '#000000'}>{cell.display.text.slice(0, 2048)}</text>
            </g>
          })}
        </g>
      </svg>
    </figure>
  })}</div>
}

import { NativeRichTextSpans, NativeRichTextDetails } from './NativeRichTextSpans'
import { Fragment, useEffect, useId, useRef, useState } from 'react'
import {
  projectNativeWorkbookV2, createNativeMaximumDigitWidthAuthorityV2,
  compileNativeSheetGeometryV2, compileNativeStoredRowSheetGeometryV1, compileNativeSheetPagePreviewV1,
  layoutNativeDrawingObjectsV1, layoutNativeCachedChartV1,
  selectNativeSheetPrintAreaSetV1, compileNativeSheetPrintAreaSetPreviewV1, selectNativeSheetPrintTitleViewportV1,
  compileNativeSheetPrintPagePreviewV1, type NativeSheetPrintPagePreviewV1,
  nativeTableFillPreview, nativeTableHeaderTextPreview, nativeTableTotalsTextPreview,
  selectNativeConditionalFillPreviewV1, type NativeConditionalFillPreviewV1,
  selectNativeRichTextPreviewV1, type NativeRichTextPreviewV1,
  type NativeWorkbookObjectsV1, type NativeSheetGeometryV2, type NativeSheetViewportV2,
  type NativeSheetPagePreviewV1, type NativeSheetHostPagePolicyV1,
  type NativePositionedDrawingV1, type NativeChartPreviewV1,
} from '@injoffice/sheets/browser'
import type { NativeWorkbook, NativeSheet } from '../nativeRoundTrip'
import { nativeSheetPageCellPreview } from '../nativeSheetPageCellPreview'
import { nativeSheetPageFitChoice } from '../nativeSheetPageFitChoice'
import { DsButton, DsSelect, DsInput } from '../design-system/primitives'
import './native-sheet-pages.css'

const EMU_PER_PIXEL = 9525
const MAX_FONT_BYTES = 32 * 1024 * 1024
// Demo responsiveness bound, not a library one: packages/sheets paginates any
// range the source authors, and neither it nor the Go tier carries a row or
// column cap. This one only keeps a single browser layout pass small. The
// largest saved print area in the hard-v2 corpus is 34 rows (A1:H34) and the
// widest is 28 columns, so 64 x 40 clears the corpus with room to spare while
// staying an order of magnitude below a full worksheet.
const MAX_PREVIEW_ROWS = 64
const MAX_PREVIEW_COLUMNS = 40
type Props = { workbook: NativeWorkbook; sheet: NativeSheet; objects: NativeWorkbookObjectsV1; rows: number; columns: number }
type Result = { selectedViewport?: NativeSheetViewportV2; geometry: NativeSheetGeometryV2; plan: NativeSheetPagePreviewV1; fontFamily: string; drawings?: NativePositionedDrawingV1[]; compactGeneral?: boolean; rangeOrigin?: 'source-print-area' | 'explicit-host'; areaIndex?: number; conditionalFills?: boolean; richRuns?: boolean; printPage?: NativeSheetPrintPagePreviewV1 }

function cellAddress(row: number, column: number) {
  let letters = '', index = column + 1
  while (index > 0) { letters = String.fromCharCode(65 + (index - 1) % 26) + letters; index = Math.floor((index - 1) / 26) }
  return `${letters}${row + 1}`
}

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
  const [scaling, setScaling] = useState<'percent' | 'fit'>('percent')
  const [fitWidth, setFitWidth] = useState('1')
  const [fitHeight, setFitHeight] = useState('1')
  const [pageOrder, setPageOrder] = useState<'downThenOver' | 'overThenDown'>('downThenOver')
  const [margins, setMargins] = useState({ left: '0.5', right: '0.5', top: '0.5', bottom: '0.5' })
  const [useSource, setUseSource] = useState(true)
  const [useStoredRows, setUseStoredRows] = useState(false)
  const [compactGeneral, setCompactGeneral] = useState(false)
  const [conditionalFills, setConditionalFills] = useState(false)
  const [richRuns, setRichRuns] = useState(false)
  let rich: NativeRichTextPreviewV1 | undefined, richError = ''
  if (richRuns) { try { rich = selectNativeRichTextPreviewV1(workbook, sheet.id, objects) } catch (error) { richError = error instanceof Error ? error.message : 'Rich source unavailable.' } }
  const [usePrintArea, setUsePrintArea] = useState(false)
  const [printPagePreview, setPrintPagePreview] = useState(false)
  const [repeatHeadings, setRepeatHeadings] = useState(false)
  const [rangeRows, setRangeRows] = useState(String(rows))
  const [rangeColumns, setRangeColumns] = useState(String(columns))
  const [results, setResult] = useState<Result[] | null>(null)
  const [message, setMessage] = useState('')
  const [busy, setBusy] = useState(false)
  const generation = useRef(0)
  const installed = useRef<FontFace | null>(null)
  const instance = useId().replace(/[^a-zA-Z0-9]/g, '')
  const conditionalSource = objects.conditional_fills?.find(entry => entry.sheet_id === sheet.id && entry.sheet_part === sheet.part_name)
  const savedArea = objects.print_areas?.find(area => area.sheet_id === sheet.id && area.sheet_part === sheet.part_name)
  const savedSet = objects.print_area_sets?.find(area => area.sheet_id === sheet.id && area.sheet_part === sheet.part_name)
  const savedRanges = savedSet?.status === 'available' ? savedSet.areas : !objects.print_area_sets && savedArea?.status === 'available' ? [savedArea.area] : []
  const neutrality = objects.dimension_neutrality?.filter(entry => entry.sheet_part === sheet.part_name) ?? []
  const neutralCodes = neutrality.length === 1 ? neutrality[0]!.codes : []
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
      const sourcePrintPage = printPagePreview
      const viewports = usePrintArea || sourcePrintPage ? selectNativeSheetPrintAreaSetV1(model, sheet.id, objects) : [{
        row: 0, column: 0, end_row: Number(rangeRows) - 1, end_column: Number(rangeColumns) - 1,
      }]
      const geometries = viewports.map((viewport, index) => {
      const selectedRows = viewport.end_row - viewport.row + 1, selectedColumns = viewport.end_column - viewport.column + 1
      if (!Number.isInteger(selectedRows) || selectedRows < 1 || selectedRows > MAX_PREVIEW_ROWS || !Number.isInteger(selectedColumns) || selectedColumns < 1 || selectedColumns > MAX_PREVIEW_COLUMNS) throw new Error(usePrintArea || sourcePrintPage
        ? `The saved print area${viewports.length > 1 ? ` ${index + 1}` : ''} exceeds this demo’s ${MAX_PREVIEW_ROWS}-row or ${MAX_PREVIEW_COLUMNS}-column limit. No areas were previewed. Choose a range from A1 instead; saved areas are not changed.`
        : `Choose 1–${MAX_PREVIEW_ROWS} rows and 1–${MAX_PREVIEW_COLUMNS} columns for this bounded preview.`)
      const geometryViewport = repeatHeadings ? selectNativeSheetPrintTitleViewportV1(model, sheet.id, viewport, objects) : viewport
      if (geometryViewport.end_row - geometryViewport.row + 1 > MAX_PREVIEW_ROWS || geometryViewport.end_column - geometryViewport.column + 1 > MAX_PREVIEW_COLUMNS) throw new Error(`The selected range and saved headings together exceed this demo’s ${MAX_PREVIEW_ROWS}-row or ${MAX_PREVIEW_COLUMNS}-column geometry limit. No areas were previewed.`)
      return useStoredRows
        ? compileNativeStoredRowSheetGeometryV1(model, sheet.id, geometryViewport, authority, objects)
        : compileNativeSheetGeometryV2(model, sheet.id, geometryViewport, authority, objects)
      })
      if (!sourcePrintPage && !useSource && Object.values(margins).some(value => !value.trim())) throw new Error('Enter all four preview margins. Use 0 for a zero margin.')
      const host: NativeSheetHostPagePolicyV1 | undefined = useSource || sourcePrintPage ? undefined : {
        kind: 'explicit-host-page-policy-v1', paper, orientation, scale: scaling === 'fit' ? 100 : Number(scale),
        ...(scaling === 'fit' ? { fit_to_page: nativeSheetPageFitChoice(fitWidth, fitHeight) } : {}),
        page_order: pageOrder,
        left_inches: Number(margins.left), right_inches: Number(margins.right), top_inches: Number(margins.top), bottom_inches: Number(margins.bottom),
      }
      const options = repeatHeadings ? { repeat_print_titles: true as const } : undefined
      let printPage: NativeSheetPrintPagePreviewV1 | undefined
      const plans = (() => {
        if (sourcePrintPage) {
          printPage = compileNativeSheetPrintPagePreviewV1(geometries, objects, options)
          if (printPage.status !== 'available') throw new Error(printPage.reason)
          return printPage.source_plan.areas.map(area => area.plan)
        }
        return usePrintArea
          ? compileNativeSheetPrintAreaSetPreviewV1(geometries, objects, host, options).areas.map(area => area.plan)
          : [compileNativeSheetPagePreviewV1(geometries[0]!, objects, host, options ? { ...options, body_viewport: viewports[0]! } : undefined)]
      })()
      const layouts = geometries.map((geometry, index) => {
        const plan = plans[index]!, drawings = layoutNativeDrawingObjectsV1(geometry, objects)
        if (repeatHeadings) assertNativeSheetHeadingDrawings(plan, drawings)
        return { geometry, plan, drawings, selectedViewport: viewports[index]! }
      })
      const fontFamily = `injoffice-sheet-${instance}-${token}`
      loaded = await new FontFace(fontFamily, bytes.buffer, {
        weight: model.normal_style?.font_bold ? '700' : '400',
        style: model.normal_style?.font_italic ? 'italic' : 'normal',
      }).load()
      if (generation.current !== token) return
      if (installed.current) document.fonts.delete(installed.current)
      document.fonts.add(loaded); installed.current = loaded
      setResult(layouts.map((layout, index) => ({ ...layout, fontFamily, compactGeneral, conditionalFills, richRuns, rangeOrigin: usePrintArea || sourcePrintPage ? 'source-print-area' : 'explicit-host', ...(usePrintArea || sourcePrintPage ? { areaIndex: index } : {}), ...(printPage ? { printPage } : {}) })))
      setMessage(`${plans.reduce((total, plan) => total + plan.pages.length, 0)} preview pages${sourcePrintPage ? ' at 96 DPI' : ''}. Read-only; the workbook is unchanged.`)
    } catch (error) {
      if (generation.current === token) setMessage(error instanceof Error ? error.message : 'Page preview unavailable.')
    } finally {
      if (generation.current === token) setBusy(false)
    }
  }
  return <section className="ds-panel native-sheet-pages" aria-label="Spreadsheet page preview">
    <h3>Page preview</h3>
    <p className="ds-muted">Preview an A1 range or supported saved print areas on paper. Saved areas keep their source order and start on separate pages. Column widths use the exact Normal font; text and cached charts use approximate browser layout. This is not Excel print fidelity.</p>
    <label><input type="checkbox" checked={printPagePreview} onChange={event => { invalidate(); setPrintPagePreview(event.target.checked); if (event.target.checked) { setUsePrintArea(true); setUseSource(true) } }}/> Source print-page preview (96 DPI)</label>
    <p className="ds-muted">Opt-in source print-page preview uses saved print areas and authored page setup only. Printer-dependent DPI, missing paper/margins/scale, and fit-to-page are refused with an explicit reason; defaults are not invented. This is not Excel printer calibration or LibreOffice PDF parity. The grid preview stays available.</p>
    <div className="native-sheet-page-controls">
      <label>Preview range<DsSelect aria-label="Preview range" value={usePrintArea || printPagePreview ? 'saved' : 'a1'} onChange={event => { invalidate(); setUsePrintArea(event.target.value === 'saved') }}><option value="a1">Choose a range from A1</option><option value="saved">Use saved print area</option></DsSelect></label>
      {usePrintArea || printPagePreview ? <p className="ds-muted">{savedRanges.length
        ? `Saved print area${savedRanges.length > 1 ? 's' : ''}: ${savedRanges.map(area => `${cellAddress(area.row, area.column)}:${cellAddress(area.end_row, area.end_column)}`).join(', ')}. Each area supports at most ${MAX_PREVIEW_ROWS} rows and ${MAX_PREVIEW_COLUMNS} columns; none are cropped automatically. Fit limits apply separately to each area, with at most 100 pages in total.`
        : `No supported saved print area is available. ${(savedSet ?? savedArea)?.warnings.join(' ') ?? ''} Choose a range from A1 to continue.`}</p> : <>
        <label>Rows<DsInput type="number" min={1} max={MAX_PREVIEW_ROWS} step={1} value={rangeRows} onChange={event => { invalidate(); setRangeRows(event.target.value) }}/></label>
        <label>Columns<DsInput type="number" min={1} max={MAX_PREVIEW_COLUMNS} step={1} value={rangeColumns} onChange={event => { invalidate(); setRangeColumns(event.target.value) }}/></label>
      </>}
      <label>Normal font: {workbook.normal_style?.font_name || 'not available'}<input type="file" accept=".ttf,font/ttf" onChange={event => { invalidate(); setFont(event.target.files?.[0] ?? null) }}/></label>
      <label><input type="checkbox" checked={useStoredRows} onChange={event => { invalidate(); setUseStoredRows(event.target.checked) }}/> Allow approximate stored-row layout</label>
      <label><input type="checkbox" checked={conditionalFills} onChange={event => { invalidate(); setConditionalFills(event.target.checked) }}/> Preview supported conditional fills</label>
      <p aria-label="Conditional fill availability" className="ds-muted">{conditionalSource ? conditionalSource.warnings.join(" ") : "No conditional fill rules were qualified in this source."} Conditional fills are not included unless you select this option.</p>
      <label><input type="checkbox" checked={compactGeneral} onChange={event => { invalidate(); setCompactGeneral(event.target.checked) }}/> Compact General numbers (host preview)</label>
      {compactGeneral && <p className="ds-muted">Your display choice rounds General numbers to seven significant digits, with scientific notation below 0.000001 or at 10000000 and above. This is not Excel General formatting. Stored values and formula caches are unchanged.</p>}
      <label><input type="checkbox" checked={useSource} onChange={event => { invalidate(); setUseSource(event.target.checked) }}/> Use saved page settings</label>
      <label><input type="checkbox" checked={repeatHeadings} onChange={event => { invalidate(); setRepeatHeadings(event.target.checked) }}/> Repeat saved print headings</label>
      <p className="ds-muted">{repeatHeadings ? `Saved heading rows and columns may lie inside or outside the selected range. The combined geometry must fit ${MAX_PREVIEW_ROWS} rows and ${MAX_PREVIEW_COLUMNS} columns. Only the selected body and heading bands are printed; intervening cells are omitted. Source-positioned chart previews and drawing placeholders repeat with the heading regions. Drawings crossing a region boundary are clipped into separate pieces; their plots remain approximate.` : 'Saved print headings are not repeated unless you select this option.'}</p>
      {!useSource && !printPagePreview && <>
        <label>Paper<DsSelect value={paper} onChange={event => { invalidate(); setPaper(event.target.value as 'A4' | 'Letter') }}><option>A4</option><option>Letter</option></DsSelect></label>
        <label>Orientation<DsSelect value={orientation} onChange={event => { invalidate(); setOrientation(event.target.value as 'portrait' | 'landscape') }}><option value="portrait">Portrait</option><option value="landscape">Landscape</option></DsSelect></label>
        <label>Scaling<DsSelect aria-label="Scaling" value={scaling} onChange={event => { invalidate(); setScaling(event.target.value as 'percent' | 'fit') }}><option value="percent">Scale percentage</option><option value="fit">Fit to page limits</option></DsSelect></label>
        {scaling === 'percent' ? <label>Scale (%)<DsInput type="number" min={10} max={400} step={1} value={scale} onChange={event => { invalidate(); setScale(event.target.value) }}/></label> : <>
          <label>Pages wide<DsInput aria-label="Pages wide" type="number" min={0} max={100} step={1} value={fitWidth} onChange={event => { invalidate(); setFitWidth(event.target.value) }}/></label>
          <label>Pages tall<DsInput aria-label="Pages tall" type="number" min={0} max={100} step={1} value={fitHeight} onChange={event => { invalidate(); setFitHeight(event.target.value) }}/></label>
          <p className="ds-muted">Use 0 for an unlimited dimension, with at least one positive limit. Fit reduces the preview in whole-percent steps from 100% to 10%; it never enlarges content. Limits apply only to the selected range, not the whole workbook.</p>
        </>}
        <label>Page order<DsSelect value={pageOrder} onChange={event => { invalidate(); setPageOrder(event.target.value as 'downThenOver' | 'overThenDown') }}><option value="downThenOver">Down, then across</option><option value="overThenDown">Across, then down</option></DsSelect></label>
        {(['left', 'right', 'top', 'bottom'] as const).map(side => <label key={side}>{side[0]!.toUpperCase() + side.slice(1)} margin (inches)<DsInput type="number" min={0} max={20} step="0.05" value={margins[side]} onChange={event => { invalidate(); setMargins(current => ({ ...current, [side]: event.target.value })) }}/></label>)}
        <p className="ds-muted">Paper, margins, scale and page order are your preview choices, not workbook defaults.</p>
      </>}
      <DsButton disabled={busy || !font} onClick={() => void renderPages()}>{busy ? 'Preparing pages…' : 'Preview pages'}</DsButton>
    </div>
    <p className="ds-muted">The font stays in this browser and is not saved in the workbook. Other fonts may be substituted by the browser. Supported chart caches use saved drawing anchors; plot colors and axes are approximate. Unknown drawings get placeholders when their position is known. Page headers and footers are not drawn here. Saved print areas support up to 16 non-overlapping rectangles. Formula values are saved caches, not recalculated results.</p>
    {(usePrintArea || printPagePreview) && savedRanges.length > 0 && <p className="ds-muted" aria-label="Saved print area provenance">{(savedSet ?? savedArea)?.warnings.join(' ')}</p>}
    <p className="ds-muted" aria-label="Unmodeled worksheet markup">{neutralCodes.length
      ? `Unmodeled worksheet markup cleared as non-dimensional for this sheet: ${neutralCodes.join(', ')}. ${neutrality[0]!.warnings.join(' ')}`
      : 'No unmodeled worksheet markup was cleared as non-dimensional for this sheet. Dimension refusals stand.'}</p>
    <label><input type="checkbox" checked={richRuns} onChange={event => { invalidate(); setRichRuns(event.target.checked) }}/> Preview supported rich-text runs</label>
    {richRuns && <>
      {richError && <p role="alert">{richError} Run styling omitted; plain source text retained.</p>}
      <NativeRichTextDetails entries={rich?.cells ?? []} styles={workbook.styles} warnings={[...(rich?.warnings ?? []), ...sheet.cells.filter(c => c.value?.rich && !rich?.cells.some(e => e.ref === c.ref)).slice(0, 256).map(c => `${c.ref}: source run evidence unavailable or outside preview bounds; plain source text retained.`)]}/>
    </>}
    <p role="status">{message}</p>
    {results?.map((result, index) => <div key={index} data-print-area={result.areaIndex}>
      {result.areaIndex !== undefined && <h4>Saved print area {result.areaIndex + 1}: {cellAddress((result.selectedViewport ?? result.geometry.viewport).row, (result.selectedViewport ?? result.geometry.viewport).column)}:{cellAddress((result.selectedViewport ?? result.geometry.viewport).end_row, (result.selectedViewport ?? result.geometry.viewport).end_column)}</h4>}
      <p>{result.plan.settings_origin === 'source' ? 'Saved paper, margins and scaling' : 'Your paper, margins and scaling'} · {result.printPage ? 'source print-page preview at 96 DPI' : 'approximate selected-range preview'}</p>
      {result.printPage?.status === 'available' && <p>Isolated capture rectangles: {result.printPage.pages.map(page => `${page.width_css_px}×${page.height_css_px}`).join(', ')} CSS pixels at 96 DPI. Fractional A4 sizes are not rounded.</p>}
      {result.plan.settings.fit_to_page && <p>Fit limits: {result.plan.settings.fit_to_page.width || 'unlimited'} wide × {result.plan.settings.fit_to_page.height || 'unlimited'} tall. Effective preview scale: {result.plan.pages.length ? `${Math.round(result.plan.pages[0]!.scale * 100)}%` : 'no visible cells'}.</p>}
      <p>{result.rangeOrigin === 'source-print-area' ? 'Saved print area' : 'Your preview range'}: {cellAddress((result.selectedViewport ?? result.geometry.viewport).row, (result.selectedViewport ?? result.geometry.viewport).column)}:{cellAddress((result.selectedViewport ?? result.geometry.viewport).end_row, (result.selectedViewport ?? result.geometry.viewport).end_column)}. Range selection is separate from paper settings.</p>
      <p className="ds-muted">Page order: {result.plan.settings.page_order === 'overThenDown' ? 'across, then down' : 'down, then across'}.</p>
      {result.plan.pages.some(page => page.regions) && <p>Saved print headings and source-positioned drawing fragments repeat on each page. Page captions list body rows and columns; heading cells are shown separately.</p>}
      <details><summary>Page preview limitations</summary><ul>{result.plan.warnings.map((warning, index) => <li key={index}>{warning}</li>)}<li>Text is single-line and clipped to cells; wrapping, rotation and text overflow are not reproduced. Unsupported styles and rich runs may differ. Cell text longer than 2,048 characters is truncated in this view.</li></ul></details>
      {!!result.drawings?.length && <details open><summary>Drawing coverage ({result.drawings.length})</summary><ul>{result.drawings.map((drawing, index) => <li key={index}>{drawing.source.kind === 'chart' ? `Chart ${index + 1}` : `Drawing ${index + 1}`}: {drawing.status === 'positioned' ? 'saved position available' : drawing.status}. {drawing.warning} {drawing.source.warnings.join(' ')}</li>)}</ul></details>}
      <NativeSheetPageImages {...{workbook, sheet, objects}} {...result}/>
    </div>)}
  </section>
}

/** Each qualified drawing is clipped independently into the page's disjoint
 * regions. Unknown positions still refuse the complete repeated-heading preview. */
export function assertNativeSheetHeadingDrawings(plan: NativeSheetPagePreviewV1, drawings: NativePositionedDrawingV1[]) {
  if (!plan.pages.some(page => page.regions)) return
  if (drawings.some(drawing => drawing.status === 'unavailable' || drawing.status === 'positioned' && (!drawing.rect || !drawing.clip))) {
    throw new Error('Drawings have unavailable positions. Turn off repeated headings to preview this range without repeating them.')
  }
}

export function NativeSheetPageImages({ workbook, sheet, objects, geometry, plan, fontFamily, drawings = [], compactGeneral = false, conditionalFills = false, richRuns = false, printPage }: Pick<Props, 'workbook' | 'sheet' | 'objects'> & Result) {
  const uid = useId().replace(/[^a-zA-Z0-9]/g, '')
  if (plan.source_package_sha256 !== workbook.source.package_sha256 || geometry.source_package_sha256 !== workbook.source.package_sha256 || objects.package_sha256 !== workbook.source.package_sha256 || plan.sheet_id !== sheet.id || geometry.sheet_id !== sheet.id || plan.geometry_sha256 !== geometry.geometry_sha256) return <p role="alert">Page preview no longer matches this workbook.</p>
  let conditional: NativeConditionalFillPreviewV1 | undefined
  let conditionalError = ''
  if (conditionalFills) {
    try { conditional = selectNativeConditionalFillPreviewV1(workbook, sheet.id, objects) }
    catch (error) { conditionalError = error instanceof Error ? error.message : 'Conditional source evidence is unavailable.' }
  }
  let rich: NativeRichTextPreviewV1 | undefined
  let richError = ''
  if (richRuns) { try { rich = selectNativeRichTextPreviewV1(workbook, sheet.id, objects) } catch (error) { richError = error instanceof Error ? error.message : 'Rich source unavailable.' } }
  const richCells = new Map(rich?.cells.map(entry => [entry.ref, entry]) ?? [])
  const conditionalMatches = new Set(conditional?.status === 'available' ? conditional.cells.filter(cell => cell.matches).map(cell => `${cell.row}:${cell.column}`) : [])
  const paintRegions = plan.pages.some(page => page.regions) ? plan.pages.flatMap(page => page.regions ?? [{ ...page, kind: 'body' as const }]) : undefined
  const cellMap = new Map(sheet.cells.map(cell => [`${cell.row}:${cell.column}`, cell]))
  const styleAt = (row: number, column: number) => cellMap.get(`${row}:${column}`)?.style_id ?? sheet.rows.find(r => r.row === row)?.style_id ?? sheet.columns.find(c => column >= c.column && column <= c.end_column)?.style_id ?? 0
  const cells = geometry.rows.flatMap(row => geometry.columns.flatMap(column => {
    if (row.hidden || column.hidden || paintRegions && !paintRegions.some(region => row.row >= region.rows.start && row.row <= region.rows.end && column.column >= region.columns.start && column.column <= region.columns.end)) return []
    const merge = geometry.merged_ranges.find(m => row.row >= m.row && row.row <= m.end_row && column.column >= m.column && column.column <= m.end_column)
    if (merge && (row.row !== merge.row || column.column !== merge.column)) return []
    const cell = cellMap.get(`${row.row}:${column.column}`), styleId = styleAt(row.row, column.column)
    const style = workbook.styles[styleId]?.effective
    const display = nativeSheetPageCellPreview(workbook, cell, objects, sheet.part_name, compactGeneral)
    const conditionalMatch = conditionalMatches.has(`${row.row}:${column.column}`)
    const fill = conditionalMatch && conditional?.status === 'available' ? conditional.rule.fill : nativeTableFillPreview(objects, workbook.source.package_sha256, sheet.part_name, row.row, column.column, style?.fill, styleId) ?? style?.fill_color ?? '#FFFFFF'
    const header = nativeTableHeaderTextPreview(objects, workbook.source.package_sha256, sheet.part_name, row.row, column.column, style?.fill, styleId)
    const totals = nativeTableTotalsTextPreview(objects, workbook.source.package_sha256, sheet.part_name, row.row, column.column, styleId)
    const rect = merge?.rect ?? { x_emu: column.x_emu, y_emu: row.y_emu, width_emu: column.width_emu, height_emu: row.height_emu }
    return [{ key: `${row.row}-${column.column}`, row: row.row, column: column.column, rect, style, display, fill, header, totals, conditionalMatch }]
  }))
  const disclosures = cells.filter(cell => cell.display.cached || cell.display.warnings.length || cell.display.truncated || cell.display.compacted)
  const address = cellAddress
  return <div className="native-sheet-page-list">
    {conditionalFills && <section aria-label="Conditional fill preview details">
      <p role="status">{conditionalError || conditional?.warnings.join(' ') || 'No source-qualified conditional fills are available. No conditional fills were applied.'}</p>
      {conditional?.status === 'available' && <p>Source rule: {conditional.rule.ref} {conditional.rule.operator} {conditional.rule.operand}; priority {conditional.rule.priority}; stopIfTrue {String(conditional.rule.stop_if_true)}; differential style {conditional.rule.dxf_id}; fill {conditional.rule.fill}. {conditional.cells.filter(cell => cell.matches).length} matching source cells, including {conditional.cells.filter(cell => cell.matches && cell.cached).length} saved formula results. This overlay does not change source styles.</p>}
    </section>}
    <section aria-label="Cell display details">
      <p role="status">{cells.filter(c => c.display.cached).length} saved formula results; freshness is unknown and formulas are not recalculated. {cells.filter(c => c.display.warnings.length).length} cells have display warnings. {cells.filter(c => c.display.truncated).length} cells exceed the 2,048-character display limit. {cells.filter(c => c.display.compacted).length} General values use your compact display choice.</p>
      {!!disclosures.length && <details><summary>Cell display details ({disclosures.length})</summary><ul>{disclosures.slice(0,100).map(cell => <li key={cell.key}>{address(cell.row,cell.column)}: {cell.display.cached && 'Saved formula result; freshness unknown. '}{cell.display.warnings.join(' ')}{cell.display.truncated && ' Text is truncated in this preview. '}{cell.display.compacted && ' Host rounding applied; not Excel General. '}Stored value: {cell.display.stored.slice(0,256)}{cell.display.stored.length > 256 && '… (detail shortened)'}</li>)}</ul>{disclosures.length > 100 && <p>Showing the first 100 of {disclosures.length} cell details.</p>}</details>}
    </section>
    {richRuns && <p>{richError || 'Rich-run styling is approximate; unavailable cell evidence retains plain text.'}</p>}
    {plan.pages.map(page => {
    return <figure key={page.number}>
      <figcaption>Page {page.number} · {page.regions ? 'body ' : ''}rows {page.rows.start + 1}–{page.rows.end + 1}, columns {page.columns.start + 1}–{page.columns.end + 1}</figcaption>
      <svg role="img" aria-label={`Approximate spreadsheet page ${page.number}`} viewBox={`0 0 ${page.width_emu / EMU_PER_PIXEL} ${page.height_emu / EMU_PER_PIXEL}`} {...(printPage?.status === 'available' ? { width: page.width_emu / EMU_PER_PIXEL, height: page.height_emu / EMU_PER_PIXEL, 'data-capture-dpi': 96 } : {})}>
        <rect width="100%" height="100%" fill="#FFFFFF"/>
        {(page.regions ?? [{ ...page, kind: 'body' as const }]).map(region => {
        const clip = `${uid}-page-${page.number}${page.regions ? `-${region.kind}` : ''}`
        return <Fragment key={region.kind}>
        <defs><clipPath id={clip}><rect x={region.source_clip.x_emu / EMU_PER_PIXEL} y={region.source_clip.y_emu / EMU_PER_PIXEL} width={region.source_clip.width_emu / EMU_PER_PIXEL} height={region.source_clip.height_emu / EMU_PER_PIXEL}/></clipPath></defs>
        <g data-page-region={page.regions ? region.kind : undefined} transform={`translate(${region.translate_x_emu / EMU_PER_PIXEL} ${region.translate_y_emu / EMU_PER_PIXEL}) scale(${page.scale})`} clipPath={`url(#${clip})`}>
          {cells.filter(cell => cell.row >= region.rows.start && cell.row <= region.rows.end && cell.column >= region.columns.start && cell.column <= region.columns.end).map(cell => {
            const x = cell.rect.x_emu / EMU_PER_PIXEL, y = cell.rect.y_emu / EMU_PER_PIXEL, w = cell.rect.width_emu / EMU_PER_PIXEL, h = cell.rect.height_emu / EMU_PER_PIXEL
            const id = `${clip}-${cell.key}`, size = (cell.style?.font_size_points ?? 11) * 96 / 72
            const right = cell.display.horizontal === 'right', center = cell.display.horizontal === 'center'
            const exactNormal = cell.style?.font_name === workbook.normal_style?.font_name && Boolean(cell.style?.bold) === Boolean(workbook.normal_style?.font_bold) && Boolean(cell.style?.italic) === Boolean(workbook.normal_style?.font_italic)
            return <g key={cell.key}><title>{`${address(cell.row,cell.column)}: ${cell.display.text.slice(0,2048)}${cell.display.warnings.length ? ` — ${cell.display.warnings.join(' ')}` : ''}`}</title>
              <rect x={x} y={y} width={w} height={h} fill={cell.fill} data-conditional-fill={cell.conditionalMatch ? "true" : undefined}/>
              <clipPath id={id}><rect x={x} y={y} width={w} height={h}/></clipPath>
              <text clipPath={`url(#${id})`} x={right ? x + w - 2 : center ? x + w / 2 : x + 2} y={cell.style?.vertical_alignment === 'top' ? y + size : cell.style?.vertical_alignment === 'middle' ? y + (h + size) / 2 - 2 : y + h - 2} textAnchor={right ? 'end' : center ? 'middle' : 'start'} fontFamily={exactNormal ? fontFamily : cell.style?.font_name || 'sans-serif'} fontSize={size} fontWeight={cell.header || cell.totals || cell.style?.bold ? 700 : 400} fontStyle={cell.style?.italic ? 'italic' : 'normal'} fill={cell.header ? '#FFFFFF' : cell.style?.font_color || '#000000'}>{richRuns && richCells.get(address(cell.row, cell.column))?.status === 'available' ? <NativeRichTextSpans entry={richCells.get(address(cell.row, cell.column))!} base={cell.style ?? {}} normal={workbook.normal_style} loadedFont={fontFamily}/> : cell.display.text.slice(0, 2048)}</text>
            </g>
          })}
          {drawings.filter(d => d.status === 'positioned' && d.rect && d.clip).map((drawing, index) => {
            const r = drawing.rect!, c = drawing.clip!, p = region.source_clip
            const left = Math.max(c.x_emu, p.x_emu), top = Math.max(c.y_emu, p.y_emu), right = Math.min(c.x_emu + c.width_emu, p.x_emu + p.width_emu), bottom = Math.min(c.y_emu + c.height_emu, p.y_emu + p.height_emu)
            if (right <= left || bottom <= top) return null
            const id = `${clip}-drawing-${index}`, chart = objects.charts.find(chart => chart.part === drawing.source.chart_part)
            return <g key={index} aria-label={`Source-positioned drawing ${index + 1}`}>
              <clipPath id={id}><rect x={left / EMU_PER_PIXEL} y={top / EMU_PER_PIXEL} width={(right - left) / EMU_PER_PIXEL} height={(bottom - top) / EMU_PER_PIXEL}/></clipPath>
              <g clipPath={`url(#${id})`}><g transform={`translate(${r.x_emu / EMU_PER_PIXEL} ${r.y_emu / EMU_PER_PIXEL}) scale(${r.width_emu / EMU_PER_PIXEL / 600} ${r.height_emu / EMU_PER_PIXEL / 260})`}>
                <NativePositionedChartPlot chart={chart} index={index}/>
              </g></g>
            </g>
          })}
        </g>
        </Fragment>
        })}
      </svg>
    </figure>
  })}</div>
}

/** Cached-data graphic only. Its outer placement is source-backed; this plot is
 * deliberately not an Office chart renderer or hyperlink surface. */
export function NativePositionedChartPlot({ chart, index }: { chart?: NativeChartPreviewV1; index: number }) {
  const layout = chart ? layoutNativeCachedChartV1(chart) : null
  return <g>
    <rect width={600} height={260} fill="#FFFFFF" stroke="#666666" strokeWidth={1}/>
    <text x={12} y={18} fontFamily="sans-serif" fontSize={12} fill="#333333">{layout ? `Chart ${index + 1}: saved data, approximate plot` : 'Drawing preview unavailable'}</text>
    {layout && chart && <g transform="translate(45 30)">
      {chart.type === 'col' ? <line x1={0} x2={520} y1={layout.baseline * 185} y2={layout.baseline * 185} stroke="#333333"/> : <line y1={0} y2={185} x1={layout.baseline * 520} x2={layout.baseline * 520} stroke="#333333"/>}
      {layout.marks.map((mark, i) => <rect key={i} x={mark.x * 520} y={mark.y * 185} width={mark.width * 520} height={mark.height * 185} fill={['#3366CC', '#CC6633', '#339966', '#993399'][mark.series % 4]}><title>{`${chart.series[mark.series]?.name || `Unnamed series ${mark.series + 1}`}: ${mark.value}`}</title></rect>)}
    </g>}
    {layout && <text x={12} y={248} fontFamily="sans-serif" fontSize={12} fill="#333333">Saved value range: {layout.minimum} to {layout.maximum}</text>}
  </g>
}

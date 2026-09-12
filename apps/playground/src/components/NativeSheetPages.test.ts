import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { describe, expect, it } from 'vitest'
import { NativeSheetPageImages, NativeSheetPages, NativePositionedChartPlot } from './NativeSheetPages'
import type { NativeWorkbook, NativeSheet } from '../nativeRoundTrip'
import type { NativeWorkbookObjectsV1, NativeSheetGeometryV2, NativeSheetPagePreviewV1, NativeChartPreviewV1, NativePositionedDrawingV1 } from '@injoffice/sheets/browser'

function fixture() {
  const revision = `sha256:${'a'.repeat(64)}`
  const sheet = { id: '1', name: 'Budget', part_name: 'xl/worksheets/sheet1.xml', rows: [], columns: [], cells: [
    { row: 0, column: 0, ref: 'A1', style_id: 0, value: { kind: 'string', text: 'Rent <script>bad</script>' } },
    { row: 1, column: 0, ref: 'A2', style_id: 0, value: { kind: 'string', text: 'Outside page one' } },
  ] } as unknown as NativeSheet
  const workbook = { source: { package_sha256: revision }, normal_style: { font_name: 'Exact Font', font_bold: false, font_italic: false }, styles: [{ effective: { font_name: 'Exact Font', font_size_points: 11, unsupported: [] } }], sheets: [sheet] } as unknown as NativeWorkbook
  const objects = { protocol: 'injoffice.xlsx.objects-preview', version: 1, package_sha256: revision, tables: [], charts: [] } as unknown as NativeWorkbookObjectsV1
  const geometry = { sheet_id: '1', source_package_sha256: revision, geometry_sha256: `sha256:${'b'.repeat(64)}`, rows: [{ row: 0, y_emu: 0, height_emu: 190500, hidden: false }, { row: 1, y_emu: 190500, height_emu: 190500, hidden: false }], columns: [{ column: 0, x_emu: 0, width_emu: 952500, hidden: false }], merged_ranges: [] } as unknown as NativeSheetGeometryV2
  const plan = { source_package_sha256: revision, sheet_id: '1', geometry_sha256: geometry.geometry_sha256, pages: [{ number: 1, width_emu: 7772400, height_emu: 10058400, source_clip: { x_emu: 0, y_emu: 0, width_emu: 952500, height_emu: 190500 }, content_clip: { x_emu: 457200, y_emu: 457200, width_emu: 6858000, height_emu: 9144000 }, scale: 1, translate_x_emu: 457200, translate_y_emu: 457200, rows: { start: 0, end: 0 }, columns: { start: 0, end: 0 } }] } as unknown as NativeSheetPagePreviewV1
  return { workbook, sheet, objects, geometry, plan, fontFamily: 'uploaded-exact-font' }
}
const render = (props = fixture()) => renderToStaticMarkup(createElement(NativeSheetPageImages, props))

describe('selected-range page presentation', () => {
  it('escapes text and renders only source cells on the selected page', () => {
    const html = render()
    expect(html).toContain('&lt;script&gt;bad&lt;/script&gt;')
    expect(html).not.toContain('<script>')
    expect(html).not.toContain('Outside page one')
    expect(html).toContain('Approximate spreadsheet page 1')
    expect(html).toContain('translate(48 48) scale(1)')
    expect(html).toContain('width="100" height="20"')
    expect(html).toContain('font-family="uploaded-exact-font"')
    expect(html).not.toContain('<button')
  })
  it.each(['objects', 'geometry', 'plan'] as const)('rejects stale %s source identity', field => {
    const props = fixture()
    if (field === 'objects') props.objects = { ...props.objects, package_sha256: 'other' }
    else props[field] = { ...props[field], source_package_sha256: 'other' } as never
    expect(render(props)).toContain('no longer matches')
    expect(render(props)).not.toContain('<svg')
  })
  it('rejects mismatched sheet and geometry-plan joins', () => {
    const props = fixture()
    props.plan = { ...props.plan, geometry_sha256: 'other' }
    expect(render(props)).not.toContain('<svg')
    props.plan = { ...fixture().plan, sheet_id: '2' }
    expect(render(props)).not.toContain('<svg')
  })
  it('does not map unrelated font families to the uploaded Normal face', () => {
    const props = fixture()
    props.workbook = { ...props.workbook, normal_style: { ...props.workbook.normal_style!, font_name: 'Different Font' } }
    expect(render(props)).toContain('font-family="Exact Font"')
    expect(render(props)).not.toContain('font-family="uploaded-exact-font"')
  })
  it('omits hidden rows without renumbering their source positions', () => {
    const props = fixture()
    props.geometry = { ...props.geometry, rows: props.geometry.rows.map(row => ({ ...row, hidden: true })) }
    expect(render(props)).not.toContain('Rent')
  })
  it('suppresses non-origin merged cells', () => {
    const props = fixture()
    props.geometry = { ...props.geometry, merged_ranges: [{ ref: 'A1:A2', row: 0, column: 0, end_row: 1, end_column: 0, rect: { x_emu: 0, y_emu: 0, width_emu: 952500, height_emu: 381000 } }] }
    props.plan.pages[0]!.rows.end = 1
    expect(render(props)).not.toContain('Outside page one')
    expect(render(props)).toContain('height="40"')
  })
  it('labels the setup as approximate and requires an explicit local font', () => {
    const props = fixture()
    const html = renderToStaticMarkup(createElement(NativeSheetPages, { ...props, rows: 2, columns: 1 }))
    expect(html).toContain('not Excel print fidelity')
    expect(html).toContain('type="file"')
    expect(html).toContain('Normal font: Exact Font')
    expect(html).toContain('Use saved page settings')
    expect(html).toContain('disabled=""')
    expect(html).toContain('not drawn here')
    expect(html).not.toContain('<svg')
    expect(html).toContain('max="26"')
    expect(html).toContain('max="32"')
  })
  it('draws cache graphics only inside qualified page intersections', () => {
    const props = fixture()
    const drawings: NativePositionedDrawingV1[] = [{ source: { sheet_id: '1', sheet_part: props.sheet.part_name, drawing_part: 'xl/drawings/drawing1.xml', ordinal: 1, kind: 'unsupported', warnings: ['Unmodeled picture'] }, status: 'positioned', rect: { x_emu: 0, y_emu: 0, width_emu: 952500, height_emu: 190500 }, clip: { x_emu: 0, y_emu: 0, width_emu: 952500, height_emu: 190500 } }]
    const html = renderToStaticMarkup(createElement(NativeSheetPageImages, { ...props, drawings }))
    expect(html).toContain('Source-positioned drawing 1')
    expect(html).toContain('Drawing preview unavailable')
    expect(html).toContain('scale(0.16666666666666666 0.07692307692307693)')
    drawings[0] = { ...drawings[0]!, clip: { ...drawings[0]!.clip!, y_emu: 381000 } }
    expect(renderToStaticMarkup(createElement(NativeSheetPageImages, { ...props, drawings }))).not.toContain('Source-positioned drawing')
  })
  it('labels cached plots as approximate and never activates source links', () => {
    const chart: NativeChartPreviewV1 = { part: 'xl/charts/chart1.xml', type: 'col', series: [{ name: '<a href="https://example.test">Revenue</a>', labels: ['Q1'], values: [10] }], warnings: [] }
    const html = renderToStaticMarkup(createElement('svg', null, createElement(NativePositionedChartPlot, { chart, index: 0 })))
    expect(html).toContain('saved data, approximate plot')
    expect(html).toContain('&lt;a href=')
    expect(html).not.toContain('<a ')
    expect(html).not.toContain('<image')
    expect(html).toContain('Saved value range:')
  })
})

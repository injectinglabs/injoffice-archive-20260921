import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { describe, expect, it } from 'vitest'
import { NativeSheetPageImages, NativeSheetPages, NativePositionedChartPlot, assertNativeSheetHeadingDrawings } from './NativeSheetPages'
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
  it('right-aligns General numeric values but retains left-aligned numeric-looking text', () => {
    const props = fixture()
    props.workbook.styles[0]!.effective.number_format = 'General'
    props.sheet.cells[0] = { ...props.sheet.cells[0]!, value: { kind: 'number', storage: 'number', lexical: '123', rich: false } }
    expect(render(props)).toMatch(/<text[^>]*x="98"[^>]*text-anchor="end"[^>]*>123<\/text>/)
    props.sheet.cells[0] = { ...props.sheet.cells[0]!, value: { kind: 'string', storage: 'inline-string', text: '123', rich: false } } as NativeSheet['cells'][number]
    expect(render(props)).toMatch(/<text[^>]*x="2"[^>]*text-anchor="start"[^>]*>123<\/text>/)
    props.workbook.styles[0]!.effective.horizontal_alignment = 'center'
    expect(render(props)).toMatch(/<text[^>]*x="50"[^>]*text-anchor="middle"[^>]*>123<\/text>/)
  })
  it('shows visible cache and compact-number disclosures without changing strict default output', () => {
    const props = fixture()
    props.workbook.styles[0]!.effective.number_format = 'General'
    props.sheet.cells[0] = { ...props.sheet.cells[0]!, value: undefined, formula: { type: 'normal', text: '16/3', cached: { kind: 'number', storage: 'number', lexical: '5.3333333333333304', rich: false } } }
    expect(render(props)).toContain('5.3333333333333304')
    const html = renderToStaticMarkup(createElement(NativeSheetPageImages, { ...props, compactGeneral: true }))
    expect(html).toContain('>5.333333</text>')
    expect(html).toContain('1 saved formula results; freshness is unknown')
    expect(html).toContain('A1: Saved formula result; freshness unknown.')
    expect(html).toContain('Host rounding applied; not Excel General.')
    expect(html).toContain('Stored value: 5.3333333333333304')
  })
  it('makes unsupported display and truncation visible outside the SVG tooltip', () => {
    const props = fixture()
    props.sheet.cells[0] = { ...props.sheet.cells[0]!, value: { kind: 'number', storage: 'number', lexical: '123.45', rich: false } }
    props.sheet.cells[1] = { ...props.sheet.cells[1]!, value: { kind: 'string', storage: 'shared-string', text: 'x'.repeat(2049), rich: false } } as never
    const html = render(props)
    expect(html).toContain('1 cells have display warnings. 1 cells exceed')
    expect(html).toContain('<li>A1: Number format unavailable; showing the stored value.')
    expect(html).toContain('<li>A2:  Text is truncated in this preview.')
  })
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
    expect(html).toContain('Repeat saved print headings')
    expect(html).not.toContain('checked=""/> Repeat saved print headings')
    expect(html).toContain('aria-label="Preview range"')
    expect(html).toContain('value="a1" selected=""')
    expect(html).toContain('Use saved print area')
    expect(html).toContain('disabled=""')
    expect(html).toContain('not drawn here')
    expect(html).not.toContain('<svg')
    expect(html).toContain('max="26"')
    expect(html).toContain('max="32"')
    expect(html).toContain('Compact General numbers (host preview)')
    expect(html).not.toContain('checked=""/> Compact General')
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
  it('keeps non-A1 source addresses while painting in viewport-local coordinates', () => {
    const props = fixture()
    props.sheet.cells.push({ row: 2, column: 1, ref: 'B3', style_id: 0, value: { kind: 'string', storage: 'inline-string', text: 'Inside saved area', rich: false } } as NativeSheet['cells'][number])
    props.geometry = { ...props.geometry, viewport: { row: 2, column: 1, end_row: 2, end_column: 1 }, rows: [{ row: 2, y_emu: 0, height_emu: 190500, hidden: false }], columns: [{ column: 1, x_emu: 0, width_emu: 952500, hidden: false }] } as NativeSheetGeometryV2
    props.plan.pages[0] = { ...props.plan.pages[0]!, rows: { start: 2, end: 2 }, columns: { start: 1, end: 1 } }
    const html = render(props)
    expect(html).toContain('rows 3–3, columns 2–2')
    expect(html).toContain('<title>B3: Inside saved area</title>')
    expect(html).toContain('translate(48 48) scale(1)')
    expect(html).toContain('x="2"')
    expect(html).not.toContain('Rent')
    expect(html).not.toContain('Outside page one')
  })
  it('applies the same effective fit transform to cells and positioned drawings', () => {
    const props = fixture()
    props.plan.pages[0]!.scale = 0.37
    const drawings: NativePositionedDrawingV1[] = [{ source: { sheet_id: '1', sheet_part: props.sheet.part_name, drawing_part: 'xl/drawings/drawing1.xml', ordinal: 1, kind: 'unsupported', warnings: ['Unmodeled picture'] }, status: 'positioned', rect: { x_emu: 0, y_emu: 0, width_emu: 952500, height_emu: 190500 }, clip: { x_emu: 0, y_emu: 0, width_emu: 952500, height_emu: 190500 } }]
    const html = renderToStaticMarkup(createElement(NativeSheetPageImages, { ...props, drawings }))
    expect(html).toContain('translate(48 48) scale(0.37)')
    expect(html).toContain('Source-positioned drawing 1')
    expect(html).toContain('Rent &lt;script&gt;bad&lt;/script&gt;')
    expect(html).not.toContain('Outside page one')
    expect(html.indexOf('scale(0.37)')).toBeLessThan(html.indexOf('Source-positioned drawing 1'))
  })
  it('keeps separate area pages clipped and gives their SVG definitions unique identities', () => {
    const first = fixture(), second = fixture()
    second.sheet.cells = [{ ...second.sheet.cells[0]!, ref: 'B3', row: 2, column: 1, value: { kind: 'string', text: 'Second saved area' } } as NativeSheet['cells'][number]]
    second.geometry = { ...second.geometry, rows: [{ ...second.geometry.rows[0]!, row: 2 }], columns: [{ ...second.geometry.columns[0]!, column: 1 }] }
    second.plan.pages[0] = { ...second.plan.pages[0]!, rows: { start: 2, end: 2 }, columns: { start: 1, end: 1 } }
    const html = renderToStaticMarkup(createElement('div', null,
      createElement(NativeSheetPageImages, first), createElement(NativeSheetPageImages, second)))
    expect(html.match(/aria-label="Approximate spreadsheet page 1"/g)).toHaveLength(2)
    expect(html.match(/<title>A1:/g)).toHaveLength(1)
    expect(html.match(/<title>B3:/g)).toHaveLength(1)
    expect(html).not.toContain('Outside page one')
    const ids = [...html.matchAll(/<clipPath id="([^"]+)"/g)].map(match => match[1])
    expect(ids.length).toBeGreaterThan(1)
    expect(new Set(ids).size).toBe(ids.length)
  })
  it('paints repeated heading regions separately without duplicating corner cells', () => {
    const props = fixture(), page = props.plan.pages[0]!
    props.geometry.columns.push({ ...props.geometry.columns[0]!, column: 1, x_emu: 952500 })
    props.sheet.cells.push({ ...props.sheet.cells[0]!, row: 0, column: 1, ref: 'B1', value: { kind: 'string', text: 'Column heading' } } as NativeSheet['cells'][number])
    props.sheet.cells.push({ ...props.sheet.cells[0]!, row: 1, column: 1, ref: 'B2', value: { kind: 'string', text: 'Body value' } } as NativeSheet['cells'][number])
    const region = (kind: 'body' | 'repeat-rows' | 'repeat-columns' | 'repeat-corner', row: number, column: number) => ({
      kind, rows: { start: row, end: row }, columns: { start: column, end: column },
      source_clip: { x_emu: column * 952500, y_emu: row * 190500, width_emu: 952500, height_emu: 190500 },
      translate_x_emu: 457200, translate_y_emu: 457200,
    })
    page.rows = { start: 1, end: 1 }; page.columns = { start: 1, end: 1 }
    page.regions = [region('repeat-corner', 0, 0), region('repeat-rows', 0, 1), region('repeat-columns', 1, 0), region('body', 1, 1)]
    const html = render(props)
    expect(html).toContain('body rows 2–2, columns 2–2')
    for (const kind of ['body', 'repeat-rows', 'repeat-columns', 'repeat-corner']) expect(html).toContain(`data-page-region="${kind}"`)
    for (const address of ['A1', 'B1', 'A2', 'B2']) expect(html.match(new RegExp(`<title>${address}:`, 'g'))).toHaveLength(1)
    expect(html).toContain('>Body value</text>')
    const drawing: NativePositionedDrawingV1 = { source: { sheet_id: '1', sheet_part: props.sheet.part_name, drawing_part: 'xl/drawings/drawing1.xml', ordinal: 1, kind: 'unsupported', warnings: [] }, status: 'positioned', rect: region('body', 1, 1).source_clip, clip: region('body', 1, 1).source_clip }
    expect(() => assertNativeSheetHeadingDrawings(props.plan, [drawing])).not.toThrow()
    expect(() => assertNativeSheetHeadingDrawings(props.plan, [{ ...drawing, rect: region('repeat-corner', 0, 0).source_clip, clip: region('repeat-corner', 0, 0).source_clip }])).not.toThrow()
    expect(() => assertNativeSheetHeadingDrawings(props.plan, [{ ...drawing, clip: undefined }])).toThrow('unavailable positions')
  })
  it('clips crossing drawings into all four heading regions with unique per-page identities', () => {
    const props = fixture(), page = props.plan.pages[0]!
    const region = (kind: 'body' | 'repeat-rows' | 'repeat-columns' | 'repeat-corner', row: number, column: number) => ({
      kind, rows: { start: row, end: row }, columns: { start: column, end: column },
      source_clip: { x_emu: column * 952500, y_emu: row * 190500, width_emu: 952500, height_emu: 190500 },
      translate_x_emu: 457200, translate_y_emu: 457200,
    })
    page.regions = [region('body', 1, 1), region('repeat-rows', 0, 1), region('repeat-columns', 1, 0), region('repeat-corner', 0, 0)]
    page.scale = 0.5
    props.plan.pages.push({ ...page, number: 2 })
    const rect = { x_emu: 476250, y_emu: 95250, width_emu: 952500, height_emu: 190500 }
    const drawings: NativePositionedDrawingV1[] = [{ source: { sheet_id: '1', sheet_part: props.sheet.part_name, drawing_part: 'xl/drawings/drawing1.xml', ordinal: 1, kind: 'unsupported', warnings: [] }, status: 'positioned', rect, clip: rect }]
    expect(() => assertNativeSheetHeadingDrawings(props.plan, drawings)).not.toThrow()
    const html = renderToStaticMarkup(createElement(NativeSheetPageImages, { ...props, drawings }))
    expect(html.match(/aria-label="Source-positioned drawing 1"/g)).toHaveLength(8)
    expect(html.match(/Drawing preview unavailable/g)).toHaveLength(8)
    for (const [x, y] of [[50, 10], [100, 10], [50, 20], [100, 20]]) {
      expect(html.match(new RegExp(`<rect x="${x}" y="${y}" width="50" height="10"`, 'g'))).toHaveLength(2)
    }
    const ids = [...html.matchAll(/<clipPath id="([^"]+)"/g)].map(match => match[1])
    expect(new Set(ids).size).toBe(ids.length)
    expect(html.match(/translate\(48 48\) scale\(0.5\)/g)).toHaveLength(8)
    drawings[0] = { ...drawings[0]!, rect: { x_emu: 0, y_emu: 0, width_emu: 952500, height_emu: 190500 }, clip: { x_emu: 0, y_emu: 0, width_emu: 952500, height_emu: 190500 } }
    const cornerOnly = renderToStaticMarkup(createElement(NativeSheetPageImages, { ...props, drawings }))
    expect(cornerOnly.match(/aria-label="Source-positioned drawing 1"/g)).toHaveLength(2)
    expect(() => assertNativeSheetHeadingDrawings(props.plan, [{ ...drawings[0]!, status: 'unavailable' }])).toThrow('unavailable positions')
    expect(() => assertNativeSheetHeadingDrawings(props.plan, [{ ...drawings[0]!, status: 'outside', rect: undefined, clip: undefined }])).not.toThrow()
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

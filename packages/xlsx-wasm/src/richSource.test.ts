import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { decodeXlsxRichSourcePreviewV1 } from './sourceStyles'
const source = JSON.parse(readFileSync(new URL('../testdata/rich-source-preview.json', import.meta.url), 'utf8'))
const decode = (value = source) => decodeXlsxRichSourcePreviewV1(JSON.stringify(value), source.package_sha256)
describe('rich source decoder', () => {
  it('joins exact Go source content, Unicode offsets and outside geometry', () => {
    const p = decode()
    expect(p.rich_cells[0]?.runs[1]).toMatchObject({ start: 4, end: 9, style: { text: ' 🧮 3', bold: false, font_color: '#FF0000' } })
    expect(p.sheet.cells[3]?.lexical).toBe('42250557.5799999')
    expect(p.sheet.cells[1]?.text).toBe('Budget ')
    expect(p.unused_style_ids).toEqual([0])
    expect(p.omitted_rows).toEqual([{ index: 2, size: 409 }])
    expect(Object.isFrozen(p.rich_cells[0]?.runs[0]?.style)).toBe(true)
    expect(Object.isFrozen(p.sheet.cells)).toBe(true)
    expect(p).not.toHaveProperty('revision')
  })
  it.each([
    ['authority', (p: any) => { p.editable = true }],
    ['source', (p: any) => { p.package_sha256 = 'sha256:' + '0'.repeat(64) }],
    ['metadata', (p: any) => { delete p.styles_sha256 }],
    ['present count', (p: any) => { p.parent_count.declaration = 'present' }],
    ['unbounded count', (p: any) => { p.parent_count.observed = 129 }],
    ['default style', (p: any) => { p.sheet.cells[0].style_id = 0 }],
    ['sparse', (p: any) => { p.sheet.cells.pop() }],
    ['duplicate coordinate', (p: any) => { p.sheet.cells[1] = p.sheet.cells[0] }],
    ['wrong coordinate', (p: any) => { p.sheet.cells[0].ref = 'B1' }],
    ['unused projected style', (p: any) => { p.unused_style_ids.push(1) }],
    ['unused gap', (p: any) => { p.unused_style_ids.push(3) }],
    ['duplicate unused', (p: any) => { p.unused_style_ids.push(0) }],
    ['parent join', (p: any) => { p.styles[0].parent_id = 1 }],
    ['unresolved font', (p: any) => { p.styles[0].warnings.push('fallback') }],
    ['format', (p: any) => { p.styles[0].number_format = '0.00' }],
    ['formula', (p: any) => { p.sheet.cells[2].formula = '1'; p.sheet.cells[2].cached = true }],
    ['numeric text', (p: any) => { p.sheet.cells[3].lexical = 'NaN' }],
    ['UTF8 text bound', (p: any) => { p.sheet.cells[1].text = 'é'.repeat(2500) }],
    ['string lexical', (p: any) => { p.sheet.cells[1].lexical = '1' }],
    ['rich style join', (p: any) => { p.rich_cells[0].style_id = 2 }],
    ['rich text join', (p: any) => { p.rich_cells[0].text = 'changed' }],
    ['run text join', (p: any) => { p.rich_cells[0].runs[0].style.text = 'Xear' }],
    ['UTF16 offset', (p: any) => { p.rich_cells[0].runs[1].end = 8 }],
    ['offset gap', (p: any) => { p.rich_cells[0].runs[1].start = 5 }],
    ['duplicate rich', (p: any) => { p.rich_cells.push(p.rich_cells[0]) }],
    ['unknown run effect', (p: any) => { p.rich_cells[0].runs[0].style.strike = true }],
    ['invalid inherited props', (p: any) => { p.rich_cells[0].runs[0].style.properties = 'cell-inherited' }],
    ['active underline', (p: any) => { p.rich_cells[0].runs[0].style.underline = 'single' }],
    ['underline origin', (p: any) => { delete p.rich_cells[0].runs[0].style.underline_origin }],
    ['unpaired surrogate', (p: any) => { p.rich_cells[0].runs[0].style.text = '\ud800' }],
    ['inside omitted row', (p: any) => { p.omitted_rows[0].index = 1 }],
    ['duplicate omitted band', (p: any) => { p.omitted_columns.push(p.omitted_columns[0]) }],
    ['outside row height', (p: any) => { p.omitted_rows[0].size = 410 }],
    ['missing disclosure', (p: any) => { p.warnings = [] }],
  ])('refuses %s', (_name, edit) => { const p = structuredClone(source); edit(p); expect(() => decode(p)).toThrow() })
  it('retains a fully inherited run with the qualified cell font', () => {
    const p = structuredClone(source)
    p.rich_cells[0].runs[1].style = { text: ' 🧮 3', properties: 'cell-inherited', omitted: [] }
    expect(decode(p).rich_cells[0]?.runs[1]?.style.properties).toBe('cell-inherited')
  })
})

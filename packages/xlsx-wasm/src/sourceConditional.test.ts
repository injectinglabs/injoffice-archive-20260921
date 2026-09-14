import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { decodeXlsxSourceStylePreviewV2 } from './sourceStyles'
const source = JSON.parse(readFileSync(new URL('../testdata/source-conditional-preview.json', import.meta.url), 'utf8'))
const decode = (value = source) => decodeXlsxSourceStylePreviewV2(JSON.stringify(value), source.grid.package_sha256)
describe('conditional source evidence decoder', () => {
  it('joins and freezes the actual Go synthetic output', () => {
    const preview = decode()
    expect(preview.data_bar?.cells[0]).toEqual({ ref: 'A2', value: 327, length_percent: 62.32 })
    expect(preview.text_rule?.cells[0]).toEqual({ ref: 'B2', value: 'REORDER', cached: true, matched: true })
    expect(Object.isFrozen(preview.text_rule?.cells[0])).toBe(true)
    expect(Object.isFrozen(preview.grid.sheets[0]?.cells)).toBe(true)
    expect(preview).not.toHaveProperty('revision')
  })
  it.each([
    ['authority', (p: any) => { p.editable = true }],
    ['base identity', (p: any) => { p.grid.package_sha256 = 'sha256:' + '0'.repeat(64) }],
    ['missing source identity', (p: any) => { delete p.worksheet_sha256 }],
    ['unqualified base', (p: any) => { p.grid.sheets[0].cells[0].style_id = 99 }],
    ['missing effect', (p: any) => { p.data_bar.cells = [] }],
    ['duplicate effect', (p: any) => { p.data_bar.cells.push(p.data_bar.cells[0]) }],
    ['wrong effect source', (p: any) => { p.text_rule.cells[0].ref = 'A2' }],
    ['wrong cache', (p: any) => { p.text_rule.cells[0].cached = false }],
    ['invented match', (p: any) => { p.text_rule.cells[0].matched = false }],
    ['invented value', (p: any) => { p.text_rule.cells[0].value = 'OK' }],
    ['bar length', (p: any) => { p.data_bar.cells[0].length_percent = 80 }],
    ['bar value', (p: any) => { p.data_bar.cells[0].value = 328 }],
    ['bar bounds', (p: any) => { p.data_bar.maximum = 100 }],
    ['bar formula', (p: any) => { p.grid.sheets[0].cells[2].formula = '300+27'; p.grid.sheets[0].cells[2].cached = true }],
    ['bar fill', (p: any) => { p.grid.sheets[0].cells[2].style_id = 1 }],
    ['range mismatch', (p: any) => { p.data_bar.range = 'B2' }],
    ['unknown gradient', (p: any) => { p.data_bar.gradient = false }],
    ['priority collision', (p: any) => { p.data_bar.priority = p.text_rule.priority }],
    ['rule absence', (p: any) => { p.data_bar = null; p.text_rule = null }],
    ['bad pane', (p: any) => { p.frozen_view = { frozen_rows: 1, top_left_cell: 'B2', active_pane: 'bottomLeft' }; p.warnings.push('viewport omitted') }],
    ['pane outside grid', (p: any) => { p.frozen_view = { frozen_rows: 2, top_left_cell: 'A3', active_pane: 'bottomLeft' }; p.warnings.push('viewport omitted') }],
    ['unwarned pane', (p: any) => { p.frozen_view = { frozen_rows: 1, top_left_cell: 'A2', active_pane: 'bottomLeft' } }],
  ])('refuses %s', (_name, edit) => {
    const value = structuredClone(source); edit(value)
    expect(() => decode(value)).toThrow()
  })
  it('keeps saved cache decisions and a disclosed frozen viewport', () => {
    const value = structuredClone(source)
    value.grid.sheets[0].cells[3].text = 'OK'
    value.text_rule.cells[0].value = 'OK'; value.text_rule.cells[0].matched = false
    value.frozen_view = { frozen_rows: 1, top_left_cell: 'A2', active_pane: 'bottomLeft' }; value.warnings.push('viewport omitted')
    expect(decode(value).text_rule?.cells[0].matched).toBe(false)
  })
})

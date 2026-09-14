import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { decodeXlsxSourceStylePreviewV1 } from './sourceStyles'
const source = JSON.parse(readFileSync(new URL('../testdata/source-style-preview.json', import.meta.url), 'utf8'))
const decode = (value = source) => decodeXlsxSourceStylePreviewV1(JSON.stringify(value), source.package_sha256)
describe('read-only source-style decoder', () => {
  it('decodes the actual synthetic Go output as an immutable separate model', () => {
    const preview = decode()
    expect(preview.read_only).toBe(true)
    expect(preview.styles[1]?.fill_color).toBe('#1E2761')
    expect(preview.styles[1]?.font_color).toBe('#FFFFFF')
    expect(preview.sheets[0]?.cells[2]?.formula).toBe('SUM(1,2)')
    expect(Object.isFrozen(preview.sheets[0]?.cells[2])).toBe(true)
    expect(preview).not.toHaveProperty('revision')
    expect(preview).not.toHaveProperty('editable')
  })
  it.each([
    ['native authority', (p: any) => { p.editable = true }],
    ['package identity', (p: any) => { p.package_sha256 = 'sha256:' + '0'.repeat(64) }],
    ['mutable authority marker', (p: any) => { p.read_only = false }],
    ['stale style join', (p: any) => { p.sheets[0].cells[0].style_id = 127 }],
    ['coordinate alias', (p: any) => { p.sheets[0].cells[0].ref = 'B1' }],
    ['duplicate cell', (p: any) => { p.sheets[0].cells.push(p.sheets[0].cells[0]) }],
    ['oversize geometry', (p: any) => { p.sheets[0].column_widths[0] = Infinity }],
    ['unpaired Unicode', (p: any) => { p.sheets[0].cells[0].text = '\ud800' }],
    ['invalid number', (p: any) => { p.sheets[0].cells[2].lexical = 'NaN' }],
    ['bad RGB', (p: any) => { p.styles[1].fill_color = 'url(evil)' }],
    ['unlabeled cache', (p: any) => { p.sheets[0].cells[2].cached = false }],
    ['covered merge content', (p: any) => { p.sheets[0].cells[1].kind = 'string'; p.sheets[0].cells[1].text = 'hidden' }],
    ['overlapping merge', (p: any) => { p.sheets[0].merges.push(p.sheets[0].merges[0]) }],
    ['conflict ID substitution', (p: any) => { p.conflicts[0].direct_component_id = 127 }],
    ['conflict kind substitution', (p: any) => { p.conflicts[0].apply_flag = 'applyNumberFormat' }],
  ])('refuses %s', (_name, edit) => {
    const value = structuredClone(source); edit(value)
    expect(() => decode(value)).toThrow()
  })
})

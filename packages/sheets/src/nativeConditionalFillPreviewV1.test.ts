import { describe, expect, it } from 'vitest'
import { decodeNativeConditionalFillPreviewsV1, selectNativeConditionalFillPreviewV1, type NativeConditionalFillPreviewV1 } from './nativeConditionalFillPreviewV1.js'
import { decodeNativeWorkbookObjectsV1, type NativeWorkbookObjectsV1 } from './nativeObjectsPreviewV1.js'
import type { NativeWorkbookV2 } from './nativeContractV2.generated.js'

type Mutable<T> = { -readonly [K in keyof T]: Mutable<T[K]> }

const hash = `sha256:${'a'.repeat(64)}`
function fixture() {
  const entry: NativeConditionalFillPreviewV1 = {
    sheet_id: '7', sheet_part: 'Sheets/s1.xml', status: 'available', warnings: ['Saved formula caches have unknown freshness: 0 / 1 values; source rules untouched.'],
    rule: { ref: 'A1:A3', operator: 'greaterThan', operand: '0', priority: 7, stop_if_true: true, dxf_id: 0, fill: '#12AB34' },
    cells: [{ row: 0, column: 0, lexical: '-2', cached: false, matches: false }, { row: 1, column: 0, lexical: '-0', cached: false, matches: false }, { row: 2, column: 0, lexical: '3', cached: true, matches: true }],
  }
  const cells = entry.cells.map((cell, index) => {
    const value = { kind: 'number', storage: 'number', lexical: cell.lexical, rich: false }
    return { row: index, column: 0, ref: `A${index + 1}`, style_id: 0, editable: !cell.cached, ...(cell.cached ? { formula: { type: 'normal', text: '1+2', cached: value } } : { value }) }
  })
  const workbook = { source: { package_sha256: hash }, unsupported: [], sheets: [{ id: '7', part_name: 'Sheets/s1.xml', cells, merged_ranges: [] }] } as unknown as Mutable<NativeWorkbookV2>
  const objects: NativeWorkbookObjectsV1 = { protocol: 'injoffice.xlsx.preview-objects', version: 1, package_sha256: hash, tables: [], charts: [], conditional_fills: [entry] }
  return { entry, workbook, objects }
}

describe('source-qualified conditional fill supplement', () => {
  it('decodes through the public object envelope and joins every stored/cache value without mutations', () => {
    const { workbook, objects, entry } = fixture()
    const before = JSON.stringify({ workbook, objects })
    const result = selectNativeConditionalFillPreviewV1(workbook, '7', decodeNativeWorkbookObjectsV1(objects, hash))
    expect(result).toEqual(entry)
    expect(JSON.stringify({ workbook, objects })).toBe(before)
    if (result?.status === 'available') result.cells[0]!.lexical = '99'
    expect(objects.conditional_fills![0]).toEqual(entry)
  })

  it('retains whole-overlay refusal and supports absent legacy additive metadata', () => {
    const { workbook, objects } = fixture()
    objects.conditional_fills = [{ sheet_id: '7', sheet_part: 'Sheets/s1.xml', status: 'unavailable', warnings: ['Multiple rules are not resolved.'] }]
    expect(selectNativeConditionalFillPreviewV1(workbook, '7', objects)).toEqual(objects.conditional_fills[0])
    delete objects.conditional_fills
    expect(selectNativeConditionalFillPreviewV1(workbook, '7', objects)).toBeUndefined()
  })

  it('refuses stale hashes, wrong owner, modified values, lost caches and merges atomically', () => {
    for (const mutate of [
      ({ objects }: ReturnType<typeof fixture>) => { objects.tables = [{ sheet_part: 'Sheets/s1.xml' }] as NativeWorkbookObjectsV1['tables'] },
      ({ objects }: ReturnType<typeof fixture>) => { objects.package_sha256 = `sha256:${'b'.repeat(64)}` },
      ({ entry }: ReturnType<typeof fixture>) => { entry.sheet_part = 'Sheets/other.xml' },
      ({ workbook }: ReturnType<typeof fixture>) => { workbook.sheets[0]!.cells[0]!.value!.lexical = '-1' },
      ({ workbook }: ReturnType<typeof fixture>) => { delete workbook.sheets[0]!.cells[2]!.formula!.cached },
      ({ workbook }: ReturnType<typeof fixture>) => { workbook.sheets[0]!.cells[2]!.formula!.type = 'array' },
      ({ workbook }: ReturnType<typeof fixture>) => { workbook.sheets[0]!.cells.push(workbook.sheets[0]!.cells[0]!) },
      ({ workbook }: ReturnType<typeof fixture>) => { workbook.sheets[0]!.cells[0]!.ref = 'B1' },
      ({ workbook }: ReturnType<typeof fixture>) => { workbook.unsupported = [{ part_name: 'Sheets/s1.xml', cell_ref: 'A1' }] as unknown as Mutable<NativeWorkbookV2>['unsupported'] },
      ({ workbook }: ReturnType<typeof fixture>) => { workbook.sheets[0]!.merged_ranges = [{ row: 0, column: 0, end_row: 0, end_column: 1, ref: 'A1:B1', editable: false }] },
    ]) {
      const f = fixture(); mutate(f)
      expect(() => selectNativeConditionalFillPreviewV1(f.workbook, '7', f.objects)).toThrow()
    }
  })

  it('validates all identities, exact evidence shape and independently checks integer comparison results', () => {
    const invalid = [
      (e: any) => { e.cells[0].matches = true },
      (e: any) => { e.cells.pop() },
      (e: any) => { e.cells[1].row = 0 },
      (e: any) => { e.cells[0].column = -0 },
      (e: any) => { e.cells[0].lexical = '01' },
      (e: any) => { e.cells[0].lexical = '1e1' },
      (e: any) => { e.rule.operand = '0.1' },
      (e: any) => { e.rule.operand = '1000000000000000' },
      (e: any) => { e.rule.operator = 'between' },
      (e: any) => { e.rule.fill = 'red' },
      (e: any) => { e.rule.priority = 0 },
      (e: any) => { e.rule.dxf_id = -0 },
      (e: any) => { e.rule.extra = true },
      (e: any) => { e.warnings = ['Actual control\ncharacter'] },
      (e: any) => { e.warnings = ['DEL\u007f'] },
      (e: any) => { e.rule.ref = 'A1:A4097' },
      (e: any) => { e.rule.ref = 'A1 A3' },
      (e: any) => { e.sheet_part = '../Sheets/s1.xml' },
      (e: any) => { e.status = 'unavailable' },
    ]
    for (const mutate of invalid) {
      const { entry } = fixture(); mutate(entry)
      expect(() => decodeNativeConditionalFillPreviewsV1([entry])).toThrow('Invalid conditional fill preview')
    }
    const { entry } = fixture()
    expect(() => decodeNativeConditionalFillPreviewsV1([entry, entry])).toThrow()
    expect(() => decodeNativeConditionalFillPreviewsV1([{ ...entry, sheet_id: '8' }, entry])).toThrow()
    expect(() => decodeNativeConditionalFillPreviewsV1([{ ...entry, sheet_part: 'Sheets/s2.xml' }, entry])).toThrow()
  })

  it('bounds aggregate evidence across worksheets', () => {
    const { entry } = fixture()
    if (entry.status !== 'available') throw new Error('fixture')
    const large = { ...entry, rule: { ...entry.rule, ref: 'A1:A4096' }, cells: Array.from({ length: 4096 }, (_, row) => ({ row, column: 0, lexical: '1', cached: false, matches: true })) }
    const entries = Array.from({ length: 4 }, (_, i) => ({ ...large, sheet_id: String(i + 1), sheet_part: `Sheets/s${i + 1}.xml` }))
    expect(decodeNativeConditionalFillPreviewsV1(entries)).toHaveLength(4)
    expect(() => decodeNativeConditionalFillPreviewsV1([...entries, { ...large, sheet_id: '5', sheet_part: 'Sheets/s5.xml' }])).toThrow()
  })
})

import { expect, it } from 'vitest'
import type { NativeWorkbookV2 } from './nativeContractV2.generated.js'
import { decodeNativeRichTextPreviewV1, selectNativeRichTextPreviewV1, nativeRichTextRunDisclosureV1, type NativeRichTextPreviewV1 } from './nativeRichTextPreviewV1.js'
function fixture(): NativeRichTextPreviewV1 { return { cells: [{ sheet_id: '1', sheet_part: 'xl/worksheets/sheet1.xml', row: 0, column: 0, ref: 'A1', style_id: 0, storage: 'shared', source_part: 'xl/sharedStrings.xml', shared_index: '0', text: 'bold plain', status: 'available', warnings: ['Missing run properties use approximate cell-font fallback.'], runs: [{ text: 'bold ', properties: 'direct', bold: true, omitted: [] }, { text: 'plain', properties: 'direct', font_color: '#FF0000', omitted: ['baseline', 'font-family-hint'] }] }], warnings: [] } }
function source(p = fixture()) {
  const e = p.cells[0]!
  const workbook = { source: { package_sha256: `sha256:${'a'.repeat(64)}` }, sheets: [{ id: '1', part_name: e.sheet_part, merged_ranges: [], cells: [{ row: 0, column: 0, ref: 'A1', style_id: 0, value: { kind: 'string', rich: true, storage: 'shared', lexical: '0', text: e.text, runs: e.runs!.map(({ omitted, properties, ...r }) => r) } }] }], styles: [{ id: 0, effective: { projection: 'full', unsupported: [], bold: false, italic: false } }] } as unknown as NativeWorkbookV2
  return { workbook, objects: { package_sha256: workbook.source.package_sha256, rich_text: p, tables: [] } }
}
it('decodes owned source evidence and distinguishes direct from inherited fallback', () => {
  const p = fixture(), d = decodeNativeRichTextPreviewV1(p); d.cells[0]!.runs![0]!.text = 'changed'; expect(p.cells[0]!.runs![0]!.text).toBe('bold ')
  expect(nativeRichTextRunDisclosureV1(p.cells[0]!.runs![1]!)).toContain('missing font_name, font_size_points, bold, italic')
  expect(nativeRichTextRunDisclosureV1({ text: 'x', properties: 'cell-inherited', omitted: [] })).toContain('No rPr: cell font inherited')
  const { workbook, objects } = source(); expect(selectNativeRichTextPreviewV1(workbook, '1', objects)).toEqual(fixture())
})
it('refuses forged, malformed, duplicated, unknown or unbounded evidence without getters', () => {
  for (const edit of [
    (p: any) => p.cells[0].extra = true, (p: any) => p.cells.push(p.cells[0]), (p: any) => p.cells[0].ref = 'A01',
    (p: any) => p.cells[0].runs[0].bold = '', (p: any) => p.cells[0].runs[0].font_color = 'red',
    (p: any) => p.cells[0].runs[0].font_name = 'url(evil)', (p: any) => p.cells[0].runs[0].text = 'wrong',
    (p: any) => p.cells[0].runs[0].omitted = ['unknown'], (p: any) => p.cells[0].warnings = ['bad\u0000'],
    (p: any) => p.cells[0].runs[0].properties = 'cell-inherited', (p: any) => p.cells[0].runs = Array(65).fill(p.cells[0].runs[0]),
    (p: any) => p.cells[0].text = 'x'.repeat(2049), (p: any) => p.cells = Array(257).fill(p.cells[0]),
  ]) { const p = fixture(); edit(p); expect(() => decodeNativeRichTextPreviewV1(p)).toThrow() }
  let called = false; const p = fixture(); Object.defineProperty(p, 'cells', { get() { called = true; return [] } }); expect(() => decodeNativeRichTextPreviewV1(p)).toThrow(); expect(called).toBe(false)
})
it('refuses stale cell, style, index, run, table and grouped-source joins', () => {
  for (const edit of [
    (s: any) => s.objects.package_sha256 = `sha256:${'b'.repeat(64)}`,
    (s: any) => s.workbook.sheets[0].cells[0].value.lexical = '1',
    (s: any) => s.workbook.sheets[0].cells[0].value.runs[0].bold = false,
    (s: any) => delete s.objects.rich_text.cells[0].runs[0].bold,
    (s: any) => { delete s.objects.rich_text.cells[0].runs[0].bold; s.objects.rich_text.cells[0].runs[0].properties = 'cell-inherited' },
    (s: any) => s.workbook.sheets[0].cells[0].style_id = 1,
    (s: any) => s.workbook.styles[0].effective.wrap_text = true,
    (s: any) => s.objects.tables.push({ sheet_part: 'xl/worksheets/sheet1.xml' }),
    (s: any) => s.workbook.sheets[0].cells.push({ ref: 'B1', row: 0, column: 1, formula: { type: 'array' } }),
  ]) { const s = source(); edit(s); expect(() => selectNativeRichTextPreviewV1(s.workbook, '1', s.objects)).toThrow() }
})
it('keeps omitted styling as readable plaintext evidence with no partial runs', () => {
  const p = fixture(), c = p.cells[0]!; c.status = 'omitted'; delete c.runs
  expect(decodeNativeRichTextPreviewV1(p).cells[0]!.text).toBe('bold plain')
  c.runs = []; expect(() => decodeNativeRichTextPreviewV1(p)).toThrow()
})
it('accepts combined maximum rich shape within its structural snapshot allowance', () => {
  const p = fixture(), seed = p.cells[0]!; p.cells = Array.from({ length: 256 }, (_, row) => ({ ...structuredClone(seed), row, ref: `A${row + 1}`, text: 'x'.repeat(128), runs: Array.from({ length: 4 }, () => ({ text: 'x'.repeat(32), properties: 'direct' as const, font_name: 'Calibri', font_color: '#112233', font_size_points: 11, bold: false, italic: false, omitted: ['baseline', 'font-family-hint', 'underline-none', 'strike-false'] })) }))
  expect(decodeNativeRichTextPreviewV1(p).cells).toHaveLength(256)
  p.cells[0]!.runs!.push({ text: 'y', properties: 'direct', omitted: [] }); p.cells[0]!.text += 'y'
  expect(() => decodeNativeRichTextPreviewV1(p)).toThrow()
})
it('enforces aggregate run and text budgets independently', () => {
  const seed = fixture().cells[0]!
  const runHeavy = { cells: Array.from({ length: 17 }, (_, row) => ({ ...structuredClone(seed), row, ref: `A${row + 1}`, text: 'x'.repeat(64), runs: Array.from({ length: 64 }, () => ({ text: 'x', properties: 'direct' as const, omitted: [] })) })), warnings: [] }
  expect(() => decodeNativeRichTextPreviewV1(runHeavy)).toThrow()
  runHeavy.cells.pop(); expect(decodeNativeRichTextPreviewV1(runHeavy).cells).toHaveLength(16)
  const textHeavy = { cells: Array.from({ length: 17 }, (_, row) => ({ ...structuredClone(seed), row, ref: `A${row + 1}`, text: 'x'.repeat(2048), runs: [{ text: 'x'.repeat(2048), properties: 'direct' as const, omitted: [] }] })), warnings: [] }
  expect(() => decodeNativeRichTextPreviewV1(textHeavy)).toThrow()
  textHeavy.cells.pop(); expect(decodeNativeRichTextPreviewV1(textHeavy).cells).toHaveLength(16)
})

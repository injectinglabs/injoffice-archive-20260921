import { expect, it } from 'vitest'
import type { NativeWorkbookV2 } from './nativeContractV2.generated.js'
import { decodeNativeRichTextPreviewV1, selectNativeRichTextPreviewV1, nativeRichTextRunDisclosureV1, type NativeRichTextPreviewV1 } from './nativeRichTextPreviewV1.js'
function fixture(): NativeRichTextPreviewV1 { return { cells: [{ sheet_id: '1', sheet_part: 'xl/worksheets/sheet1.xml', row: 0, column: 0, ref: 'A1', style_id: 0, storage: 'shared', source_part: 'xl/sharedStrings.xml', shared_index: '0', text: 'bold plain', status: 'available', warnings: ['Missing run properties use approximate cell-font fallback.'], runs: [{ text: 'bold ', properties: 'direct', bold: true, omitted: [] }, { text: 'plain', properties: 'direct', font_color: '#FF0000', omitted: ['baseline', 'font-family-hint'] }] }], warnings: [] } }
function source(p = fixture()) {
  const e = p.cells[0]!
  const workbook = { source: { package_sha256: `sha256:${'a'.repeat(64)}` }, sheets: [{ id: '1', part_name: e.sheet_part, merged_ranges: [], cells: [{ row: 0, column: 0, ref: 'A1', style_id: 0, value: { kind: 'string', rich: true, storage: 'shared', lexical: '0', text: e.text, runs: e.runs!.map(({ omitted, properties, ...r }) => r) } }] }], styles: [{ id: 0, effective: { projection: 'full', unsupported: [], bold: false, italic: false } }] } as unknown as NativeWorkbookV2
  return { workbook, objects: { protocol: 'injoffice.xlsx.preview-objects' as const, version: 1 as const, package_sha256: workbook.source.package_sha256, rich_text: p, tables: [], charts: [] } }
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
  const p = fixture(), seed = p.cells[0]!; p.cells = Array.from({ length: 256 }, (_, row) => ({ ...structuredClone(seed), row, ref: `A${row + 1}`, text: 'x'.repeat(128), runs: Array.from({ length: 4 }, () => ({ text: 'x'.repeat(32), properties: 'direct' as const, font_name: 'Calibri', font_color: '#112233', font_size_points: 11, bold: false, italic: false, underline: 'single' as const, underline_origin: 'explicit-val' as const, omitted: ['baseline', 'font-family-hint', 'strike-false'] })) }))
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
it('decodes maximum conditional-fill and rich-run evidence together in the public envelope', async () => {
  const { decodeNativeWorkbookObjectsV1 } = await import('./nativeObjectsPreviewV1.js')
  const rich = fixture(), seed = rich.cells[0]!
  rich.cells = Array.from({ length: 256 }, (_, row) => ({ ...structuredClone(seed), row, ref: `A${row + 1}`, text: 'x'.repeat(128), runs: Array.from({ length: 4 }, () => ({ text: 'x'.repeat(32), properties: 'direct' as const, font_name: 'Calibri', font_scheme: 'minor' as const, declared_font_name: 'Arial', theme_part: 'xl/theme/theme1.xml', theme_sha256: `sha256:${'b'.repeat(64)}`, font_size_points: 11, font_color: '#112233', bold: false, italic: false, underline: 'single' as const, underline_origin: 'explicit-val' as const, omitted: ['baseline', 'font-family-hint', 'strike-false'] })) }))
  const conditional_fills = Array.from({ length: 4 }, (_, index) => ({ sheet_id: String(index + 1), sheet_part: `xl/worksheets/sheet${index + 1}.xml`, status: 'available', warnings: ['Stored integers only.'], rule: { ref: 'A1:BL64', operator: 'equal', operand: '0', priority: 1, stop_if_true: false, dxf_id: 0, fill: '#112233' }, cells: Array.from({ length: 4096 }, (_, i) => ({ row: Math.floor(i / 64), column: i % 64, lexical: '0', cached: false, matches: true })) }))
  const hash = `sha256:${'a'.repeat(64)}`, envelope = { protocol: 'injoffice.xlsx.preview-objects', version: 1, package_sha256: hash, tables: [], charts: [], conditional_fills, rich_text: rich }
  const result = decodeNativeWorkbookObjectsV1(envelope, hash)
  expect(result.conditional_fills).toHaveLength(4); expect(result.rich_text?.cells).toHaveLength(256)
})
it('retains Unicode source paths and explicitly omitted noncanonical saved indices', () => {
  const p = fixture(), c = p.cells[0]!; c.sheet_part = 'xl/worksheets/预算.xml'; c.source_part = 'xl/字符串.xml'; c.shared_index = '00'; c.status = 'omitted'; delete c.runs
  expect(decodeNativeRichTextPreviewV1(p).cells[0]!.shared_index).toBe('00')
  c.status = 'available'; c.runs = fixture().cells[0]!.runs; expect(() => decodeNativeRichTextPreviewV1(p)).toThrow()
})

it('validates paired raw underline attestations without inventing an extracted-model property', () => {
  for (const [underline, underline_origin] of [['single', 'explicit-val'], ['single', 'default-val'], ['none', 'explicit-val']] as const) {
    const s = source(); Object.assign(s.objects.rich_text.cells[0]!.runs![0]!, { underline, underline_origin })
    expect(selectNativeRichTextPreviewV1(s.workbook, '1', s.objects).cells[0]!.runs![0]!.underline).toBe(underline)
    expect(nativeRichTextRunDisclosureV1(s.objects.rich_text.cells[0]!.runs![0]!)).toContain(underline_origin === 'default-val' ? 'schema-default single' : 'explicit source val')
  }
  for (const extra of [{underline:['single'],underline_origin:'explicit-val'}, {underline:'single'}, {underline_origin:'explicit-val'}, {underline:'none',underline_origin:'default-val'}, {underline:'double',underline_origin:'explicit-val'}, {underline:'singleAccounting',underline_origin:'explicit-val'}, {underline:'single',underline_origin:'unknown'}, {underline:null,underline_origin:'explicit-val'}, {underline:'single',underline_origin:'explicit-val',omitted:['underline-none']}, {underline:'single',underline_origin:'explicit-val',properties:'cell-inherited'}]) {
    const p = fixture(); Object.assign(p.cells[0]!.runs![0]!, extra); expect(() => decodeNativeRichTextPreviewV1(p)).toThrow()
  }
  expect(nativeRichTextRunDisclosureV1(fixture().cells[0]!.runs![0]!)).toContain('No direct underline declaration; undecorated host fallback')
  const p = fixture(); p.cells[0]!.runs![0]!.omitted = ['underline-none']; expect(decodeNativeRichTextPreviewV1(p).cells).toHaveLength(1)
})

it('joins resolved scheme fonts and retains raw declaration evidence', () => {
  const s = source(), run = s.objects.rich_text.cells[0]!.runs![0]!
  Object.assign(run, {font_name:'Cambria',font_scheme:'major',declared_font_name:'Arial',theme_part:'xl/theme/theme1.xml',theme_sha256:`sha256:${'b'.repeat(64)}`})
  Object.assign(s.workbook.sheets[0]!.cells[0]!.value!.runs![0]!, {font_name: 'Cambria'})
  const selected = selectNativeRichTextPreviewV1(s.workbook,'1',s.objects)
  expect(nativeRichTextRunDisclosureV1(selected.cells[0]!.runs![0]!)).toContain('declared Arial, resolved theme Latin face Cambria')
  Object.assign(s.workbook.sheets[0]!.cells[0]!.value!.runs![0]!, {font_name: 'Arial'})
  expect(() => selectNativeRichTextPreviewV1(s.workbook,'1',s.objects)).toThrow()
})
it('refuses partial, malformed or non-ASCII theme attestations', () => {
  const evidence = {font_name:'Cambria',font_scheme:'major',declared_font_name:'Arial',theme_part:'xl/theme/theme1.xml',theme_sha256:`sha256:${'b'.repeat(64)}`}
  for (const extra of [{font_scheme:'none'},{font_scheme:['major']},{theme_part:'../theme.xml'},{theme_sha256:'bad'},{declared_font_name:''},{properties:'cell-inherited'},{text:'世界'}]) {
    const p=fixture(); Object.assign(p.cells[0]!.runs![0]!,evidence,extra); expect(() => decodeNativeRichTextPreviewV1(p)).toThrow()
  }
  for(const k of Object.keys(evidence)) {
    const p=fixture();const r=p.cells[0]!.runs![0]!;Object.assign(r,evidence);delete (r as unknown as Record<string,unknown>)[k];expect(() => decodeNativeRichTextPreviewV1(p)).toThrow()
  }
})

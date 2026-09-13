import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { NativePptxDeck } from './types'
import { decodeNativePptxTableInspection, createNativePptxTableGeometryPreview, type NativePptxTableInspection } from './tableInspection'

const sha = 'a'.repeat(64)
function fixture() {
  const deck = JSON.parse(readFileSync(new URL('../../../go/pptxpatch/testdata/native-contract/valid/parsed-full.json', import.meta.url), 'utf8')) as NativePptxDeck
  deck.sourceRevision = `rev-${sha}`
  const slide = deck.slides[0]!, source = slide.source!, opaque = slide.passthrough[0]!
  const input: NativePptxTableInspection = {
    protocol: 'pptx-table-content-inspection-v1', package_sha256: sha, source_revision: `rev-${sha}`,
    tables: [{ slide_id: slide.id, slide_index: 0, part_name: source.partName, part_sha256: 'b'.repeat(64), slide_source_sha256: source.fingerprintSha256,
      object_id: 'cNvPr-123', source_sha256: opaque.fingerprintSha256,
      rect: { x: -1, y: 2, width: 100, height: 200 }, cells: [{ row: 0, column: 0, rect: { x: 0, y: 0, width: 100, height: 200 }, paragraphs: ['Plain <text> 😀', ''] }],
      warnings: ['Plain source text and stored cell geometry only; this is not a rendered Office table.', 'Table styles, borders and fills are omitted, including any authored dash patterns.', 'Source text styling, margins, alignment, spacing and inherited font metrics are not applied.'],
    }], omissions: [],
  }
  return { deck, input }
}

describe('read-only source-bound PPTX table decoder', () => {
  it('copies plain text and geometry without native tokens or authored styles', () => {
    const { deck, input } = fixture()
    const result = decodeNativePptxTableInspection(input, deck, sha)
    expect(result).toEqual(input)
    expect(result).not.toBe(input)
    expect(result.tables[0]!.cells).not.toBe(input.tables[0]!.cells)
    expect(JSON.stringify(result)).not.toContain('token')
  })
  it.each([
    ['hash', (v: NativePptxTableInspection) => { v.package_sha256 = 'b'.repeat(64) }],
    ['revision', (v: NativePptxTableInspection) => { v.source_revision = `rev-${'b'.repeat(64)}` }],
    ['slide', (v: NativePptxTableInspection) => { v.tables[0]!.slide_id = 'unknown' }],
    ['part', (v: NativePptxTableInspection) => { v.tables[0]!.part_name = 'ppt/slides/slide2.xml' }],
    ['root hash', (v: NativePptxTableInspection) => { v.tables[0]!.slide_source_sha256 = 'b'.repeat(64) }],
    ['source hash', (v: NativePptxTableInspection) => { v.tables[0]!.source_sha256 = 'b'.repeat(64) }],
    ['duplicate table', (v: NativePptxTableInspection) => { v.tables.push(structuredClone(v.tables[0]!)) }],
    ['negative zero', (v: NativePptxTableInspection) => { v.tables[0]!.rect.x = -0 }],
    ['offset', (v: NativePptxTableInspection) => { v.tables[0]!.cells[0]!.rect.x = 1 }],
    ['grid extent', (v: NativePptxTableInspection) => { v.tables[0]!.cells[0]!.rect.width = 99 }],
    ['grid order', (v: NativePptxTableInspection) => { v.tables[0]!.cells[0]!.column = 1 }],
    ['empty cells', (v: NativePptxTableInspection) => { v.tables[0]!.cells = [] }],
    ['empty paragraphs', (v: NativePptxTableInspection) => { v.tables[0]!.cells[0]!.paragraphs = [] }],
    ['paragraph budget', (v: NativePptxTableInspection) => { v.tables[0]!.cells[0]!.paragraphs = Array(257).fill('') }],
    ['text budget', (v: NativePptxTableInspection) => { v.tables[0]!.cells[0]!.paragraphs = ['x'.repeat(65536), 'x'] }],
    ['unpaired surrogate', (v: NativePptxTableInspection) => { v.tables[0]!.cells[0]!.paragraphs = ['\ud800'] }],
    ['missing warning', (v: NativePptxTableInspection) => { v.tables[0]!.warnings.pop() }],
    ['invalid omission', (v: NativePptxTableInspection) => { v.omissions = [{ slide_id: 'unknown', object_id: 'slide', reason: 'unsupported' }] }],
  ])('refuses %s', (_name, mutate) => {
    const { deck, input } = fixture(); mutate(input)
    expect(() => decodeNativePptxTableInspection(input, deck, sha)).toThrow()
  })
  it('rejects getters without executing them, sparse arrays and custom prototypes', () => {
    const { deck, input } = fixture()
    let invoked = false
    Object.defineProperty(input.tables[0], 'cells', { get() { invoked = true; return [] } })
    expect(() => decodeNativePptxTableInspection(input, deck, sha)).toThrow()
    expect(invoked).toBe(false)
    const sparse = fixture(); delete sparse.input.tables[0]
    expect(() => decodeNativePptxTableInspection(sparse.input, sparse.deck, sha)).toThrow()
    const prototype = fixture(); Object.setPrototypeOf(prototype.input.tables, Object.create(Array.prototype))
    expect(() => decodeNativePptxTableInspection(prototype.input, prototype.deck, sha)).toThrow()
    const symbol = fixture(); Object.defineProperty(symbol.input, Symbol('authority'), { value: true })
    expect(() => decodeNativePptxTableInspection(symbol.input, symbol.deck, sha)).toThrow()
  })
  it('requires unique source membership and rejects output capabilities', () => {
    const { deck, input } = fixture()
    deck.slides[0]!.passthrough.push({ ...deck.slides[0]!.passthrough[0]!, token: 'different-token' })
    expect(() => decodeNativePptxTableInspection(input, deck, sha)).toThrow()
    const other = fixture()
    Object.assign(other.input.tables[0]!, { passthrough: [] })
    expect(() => decodeNativePptxTableInspection(other.input, other.deck, sha)).toThrow()
  })
})


describe('source-positioned table geometry preview', () => {
  it('requires decoder-bound immutable source and explicit policy', () => {
    const {deck,input}=fixture()
    expect(()=>createNativePptxTableGeometryPreview(input,'host-sans-12pt-clipped-v1')).toThrow()
    const admitted=decodeNativePptxTableInspection(input,deck,sha)
    expect(()=>createNativePptxTableGeometryPreview(admitted,'other' as never)).toThrow()
    expect(()=>{admitted.tables[0]!.rect.x=20}).toThrow()
    expect(()=>createNativePptxTableGeometryPreview(structuredClone(admitted),'host-sans-12pt-clipped-v1')).toThrow()
    const preview=createNativePptxTableGeometryPreview(admitted,'host-sans-12pt-clipped-v1')
    expect(preview).toMatchObject({packageSHA256:sha,sourceRevision:`rev-${sha}`,fontSize:16,lineHeight:20,inset:2,cssPixelsPerInch:96})
    expect(preview.slides[0]).toMatchObject({sourceBounds:{x:-1,y:2,width:100,height:200},width:100/9525,height:200/9525})
    expect(preview.slides[0]!.tables[0]!.cells[0]!.paragraphs).toEqual(['Plain <text> 😀',''])
    expect(Object.isFrozen(preview.slides[0]!.tables[0]!.cells[0]!.paragraphs)).toBe(true)
    expect(JSON.stringify(preview)).not.toContain('passthrough')
  })
  it('retains source displacement and overlap rather than repositioning tables into a grid',()=>{
    const {deck,input}=fixture(), second=structuredClone(input.tables[0]!)
    second.object_id='cNvPr-124';second.source_sha256='e'.repeat(64);second.rect.x=49;second.rect.y=52
    deck.slides[0]!.passthrough.push({...deck.slides[0]!.passthrough[0]!,token:'other',fingerprintSha256:second.source_sha256})
    input.tables.push(second)
    const preview=createNativePptxTableGeometryPreview(decodeNativePptxTableInspection(input,deck,sha),'host-sans-12pt-clipped-v1')
    expect(preview.slides[0]!.tables[1]!.rect).toEqual({x:50/9525,y:50/9525,width:100/9525,height:200/9525})
    expect(preview.slides[0]!.sourceBounds).toEqual({x:-1,y:2,width:150,height:250})
  })
  it('omits an oversized arrangement as a whole with a visible reason',()=>{
    const {deck,input}=fixture();input.tables[0]!.rect.width=1_000_000_000;input.tables[0]!.cells[0]!.rect.width=1_000_000_000
    const preview=createNativePptxTableGeometryPreview(decodeNativePptxTableInspection(input,deck,sha),'host-sans-12pt-clipped-v1')
    expect(preview.slides).toEqual([])
    expect(preview.omissions[0]!.reason).toContain('16384 CSS pixel')
  })
})

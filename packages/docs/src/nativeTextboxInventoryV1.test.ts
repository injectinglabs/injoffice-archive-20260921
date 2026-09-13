import {readFileSync} from 'node:fs'
import {describe, expect, it} from 'vitest'
import type {NativeDocxDocumentV1} from './nativeContract.js'
import {createNativeDocxTextboxInventoryV1 as create, decodeNativeDocxTextboxEvidenceV1 as decode, type NativeDocxTextboxEvidenceV1} from './nativeTextboxInventoryV1.js'

const original = JSON.parse(readFileSync(new URL('../../../testdata/docx-native/document-v1.json', import.meta.url), 'utf8')) as NativeDocxDocumentV1
const options = {policy: 'source-textbox-inventory-v1' as const, read_only: true as const}
function fixture() {
  const document = structuredClone(original), paragraph = document.body.blocks[0]!.paragraph!
  const anchor = {...paragraph.anchor, path: paragraph.anchor.path + '/w:r[1]/w:pict[1]', start_byte: 200, end_byte: 400}
  document.unsupported = [{id: 'textbox:1', code: 'UNMODELED_DRAWING', capability: 'drawings', scope_id: paragraph.id, anchor, preservation: 'refuse-mutation', message: 'Drawing preserved'}]
  const evidence: NativeDocxTextboxEvidenceV1 = {items: [{package_sha256: document.source.package_sha256, part_sha256: 'sha256:' + 'a'.repeat(64), paragraph_id: paragraph.id, diagnostic_id: 'textbox:1', anchor: {...anchor}, kind: 'vml', status: 'supported', paragraphs: ['Source <script>text</script>', 'Second paragraph'], reason: ''}], omitted_count: 0}
  return {document, evidence, paragraph}
}
describe('read-only textbox source inventory', () => {
  it('requires explicit opt-in and preserves diagnostics and literal source text', () => {
    const {document, evidence} = fixture(), before = structuredClone({document, evidence})
    const result = create(document, options, evidence)
    expect(result.items[0]!.paragraphs).toEqual(['Source <script>text</script>', 'Second paragraph'])
    expect(result.source_diagnostics).toEqual(document.unsupported)
    result.items[0]!.paragraphs[0] = 'changed'
    result.items[0]!.anchor.path = 'changed'
    expect({document, evidence}).toEqual(before)
    expect(() => create(document, {...options, read_only: false} as unknown as typeof options, evidence)).toThrow()
    expect(() => create(document, {...options, extra: true} as typeof options, evidence)).toThrow()
    expect(decode(document, undefined)).toEqual({items: [], omitted_count: 0})
  })
  it('accepts direct DrawingML container anchors and explicit omissions', () => {
    for (const suffix of ['', '/wp:inline[1]', '/wp:anchor[1]', '/ns4a8a39ac:inline[1]/ns09d2785d:graphic[1]/ns09d2785d:graphicData[1]']) {
      const {document, evidence, paragraph} = fixture(), item = evidence.items[0]!
      item.kind = 'drawingml'; item.anchor.path = paragraph.anchor.path + '/w:r[1]/w:drawing[1]' + suffix
      document.unsupported[0]!.anchor = {...item.anchor}; document.unsupported[0]!.code = 'DRAWING_GRAPHIC_REQUIRED'
      expect(decode(document, evidence).items).toEqual(evidence.items)
      item.status = 'omitted'; item.paragraphs = []; item.reason = 'Text visibility is not qualified'
      expect(decode(document, evidence).items[0]!.status).toBe('omitted')
    }
  })
  it('rejects forged, stale, duplicate, reordered and invalid text evidence', () => {
    const mutations: Array<(e: NativeDocxTextboxEvidenceV1) => void> = [
      e => {e.items[0]!.package_sha256 = 'sha256:' + 'f'.repeat(64)},
      e => {e.items[0]!.anchor.xml_sha256 = 'sha256:' + 'f'.repeat(64)},
      e => {e.items[0]!.paragraph_id = 'missing'},
      e => {e.items[0]!.part_sha256 = 'invalid'},
      e => {Object.defineProperty(e.items[0]!, 'extra', {value: true})},
      e => {Object.assign(e.items[0]!, {kind: {toString() {throw Error('coercion must not run')}}})},
      e => {e.items.push(structuredClone(e.items[0]!))},
      e => {e.items[0]!.paragraphs = ['x'.repeat(4097)]},
      e => {e.items[0]!.paragraphs = []},
      e => {e.items[0]!.status = 'omitted'},
      e => {e.items[0]!.reason = 'not empty'},
      e => {e.omitted_count = -1},
      e => {e.items[0]!.paragraphs = Array(2)},
    ]
    for (const mutate of mutations) {const {document, evidence} = fixture(); mutate(evidence); expect(() => decode(document, evidence)).toThrow()}
    const {document, evidence} = fixture()
    let invoked = false
    Object.defineProperty(evidence.items[0]!, 'paragraphs', {get() {invoked = true; return []}})
    expect(() => decode(document, evidence)).toThrow('accessors'); expect(invoked).toBe(false)
  })
  it('rejects branch, nested and foreign-owner anchors even with matching diagnostics', () => {
    for (const suffix of ['/w:r[1]/mc:AlternateContent[1]/w:pict[1]', '/w:r[1]/w:pict[1]/v:shape[1]', '/w:pict[1]']) {
      const {document, evidence, paragraph} = fixture()
      evidence.items[0]!.anchor.path = paragraph.anchor.path + suffix
      document.unsupported[0]!.anchor = {...evidence.items[0]!.anchor}
      expect(() => decode(document, evidence)).toThrow('direct body drawing')
    }
    const {document, evidence} = fixture()
    document.unsupported[0]!.capability = 'runs'
    expect(() => decode(document, evidence)).toThrow('does not join')
  })
})

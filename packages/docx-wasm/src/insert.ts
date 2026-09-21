import type { NativeDocxDocumentV1 } from '@injoffice/docs/native-docx'

export interface DocxInsertPayload {
  mutations: Array<{
    target_kind: 'paragraph'; target_id: string; expected_xml_sha256: string
    operation: 'block.insert_after' | 'page_break.insert'
    text?: string
    image?: {data_base64: string; content_type: string; width_emu: number; height_emu: number; alt_text: string}
    split?: { run_id: string; offset_utf16: number }
  }>
}
export function validateInsert(document: NativeDocxDocumentV1, mutations: unknown[]): DocxInsertPayload {
  if (mutations.length !== 1) throw new TypeError('Insertions require one mutation.')
  const m = mutations[0] as DocxInsertPayload['mutations'][number]
  const keys = ['target_kind','target_id','expected_xml_sha256','operation',...(m.operation === 'page_break.insert' ? ['split'] : m.image ? ['image'] : ['text'])]
  if (Object.keys(m).length !== keys.length || Object.keys(m).some(k=>!keys.includes(k))) throw new TypeError('Invalid insertion fields.')
  const paragraph = document.body.blocks.find(b=>b.paragraph?.id===m.target_id)?.paragraph
  if (m.target_kind !== 'paragraph' || !paragraph || paragraph.anchor.xml_sha256 !== m.expected_xml_sha256 || !paragraph.edit_policy.allowed_operations.includes(m.operation)) throw new TypeError('Insertion target is not an advertised body paragraph.')
  if (m.operation === 'page_break.insert') {
    const text = paragraph.runs.find(r=>r.id===m.split?.run_id && r.kind==='text')?.text
    const offset = m.split?.offset_utf16!
    if (!m.split || Object.keys(m.split).length!==2 || Object.keys(m.split).some(k=>!['run_id','offset_utf16'].includes(k))) throw new TypeError('Invalid caret fields.')
    if (text === undefined || !Number.isSafeInteger(offset) || offset < 0 || offset > text.length || (offset > 0 && /[\uD800-\uDBFF]/.test(text[offset-1]) && /[\uDC00-\uDFFF]/.test(text[offset] ?? ''))) throw new TypeError('Caret is outside the run or splits a surrogate pair.')
  } else if (m.operation === 'block.insert_after') {
    if (m.image) {
      const im = m.image
      const wanted = ['data_base64','content_type','width_emu','height_emu','alt_text']
      if (typeof im !== 'object' || Object.keys(im).length !== wanted.length || Object.keys(im).some(k=>!wanted.includes(k)) || !['image/png','image/jpeg'].includes(im.content_type) || typeof im.data_base64!=='string' || im.data_base64.length>5592408 || !/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(im.data_base64) || !im.data_base64.length || typeof im.alt_text!=='string' || im.alt_text.length>255 || /[\u0000-\u001f]/.test(im.alt_text) || [im.width_emu,im.height_emu].some(v=>!Number.isSafeInteger(v)||v<1||v>20116800)) throw new TypeError('Invalid PNG/JPEG insertion.')
    } else if (m.text !== '') throw new TypeError('Only empty paragraph insertion is supported.')
  } else throw new TypeError('Unsupported insertion.')
  return {mutations:[m]}
}

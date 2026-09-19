export interface PdfRecoveryDraft {
  version: 1
  format: 'pdf'
  page: number
  tool: 'text' | 'note' | 'highlight' | 'underline' | 'strikeout' | 'rectangle' | 'ellipse' | 'line' | 'arrow' | 'form' | 'replace' | 'edit-note'
  fieldName?: string
  oldText?: string
  annotationRef?: string
  annotationSignature?: string
  text: string
  size: number
  color: string
  placement?: { at: [number, number]; end?: [number, number] }
}
const object = (value: unknown): value is Record<string, unknown> => !!value && typeof value === 'object' && !Array.isArray(value) && [Object.prototype, null].includes(Object.getPrototypeOf(value))
/** Recovery input is data, never commands; reject unknown fields and invalid geometry. */
export function parsePdfRecoveryDraft(value: unknown): PdfRecoveryDraft | null {
  if (!object(value)) return null
  const descriptors = Object.getOwnPropertyDescriptors(value)
  if (Object.getOwnPropertySymbols(value).length || Object.values(descriptors).some(item => !('value' in item))) return null
  if (Object.keys(value).some(key => !['version', 'format', 'page', 'tool', 'text', 'size', 'color', 'placement', 'fieldName', 'oldText', 'annotationRef', 'annotationSignature'].includes(key))) return null
  if (value.version !== 1 || value.format !== 'pdf' || !Number.isInteger(value.page) || Number(value.page) < 1 || Number(value.page) > 100000) return null
  if ((typeof value.tool !== 'string' || !['text', 'note', 'highlight', 'underline', 'strikeout', 'rectangle', 'ellipse', 'line', 'arrow', 'form', 'replace', 'edit-note'].includes(value.tool)) || typeof value.text !== 'string' || value.text.length > 10000 || typeof value.size !== 'number' || !Number.isFinite(value.size) || value.size < 6 || value.size > 144 || typeof value.color !== 'string' || !/^#[0-9a-f]{6}$/i.test(value.color)) return null
  const pair = (input: unknown): input is [number, number] => Array.isArray(input) && input.length === 2 && input.every(number => typeof number === 'number' && Number.isFinite(number) && Math.abs(number) <= 1000000)
  let placement: PdfRecoveryDraft['placement']
  if (value.placement !== undefined) {
    if (!object(value.placement) || Object.getOwnPropertySymbols(value.placement).length || Object.values(Object.getOwnPropertyDescriptors(value.placement)).some(item => !('value' in item)) || Object.keys(value.placement).some(key => !['at', 'end'].includes(key)) || !pair(value.placement.at) || (value.placement.end !== undefined && !pair(value.placement.end))) return null
    placement = { at: [...value.placement.at], ...(value.placement.end ? { end: [...value.placement.end] as [number, number] } : {}) }
  }
  if(value.tool==='edit-note') {
    if(typeof value.annotationRef!=='string'||!/^\d+ \d+ R$/.test(value.annotationRef)||typeof value.annotationSignature!=='string'||value.annotationSignature.length>32768||!value.annotationSignature||typeof value.oldText!=='string'||value.oldText.length>10000||!placement?.end)return null
  } else if(value.annotationRef!==undefined||value.annotationSignature!==undefined)return null
  if (value.tool === 'replace') { if (typeof value.oldText !== 'string' || !value.oldText.length || value.oldText.length > 10000 || !placement?.end || value.text.includes('\n') || value.text.includes('\r')) return null } else if (value.tool!=='edit-note'&&value.oldText !== undefined) return null
  if (value.tool === 'form') { if (typeof value.fieldName !== 'string' || !value.fieldName || value.fieldName.length > 1000 || placement) return null } else if (value.fieldName !== undefined || (!placement && !value.text)) return null
  if (['rectangle','ellipse','line','arrow','highlight','underline','strikeout'].includes(value.tool as string) && value.text !== '') return null
  return { version: 1, format: 'pdf', page: Number(value.page), tool: value.tool as PdfRecoveryDraft['tool'], text: value.text, size: value.size, color: value.color, ...(placement ? { placement } : {}), ...(['replace','edit-note'].includes(value.tool) ? { oldText: value.oldText as string } : {}), ...(value.tool==='edit-note'?{annotationRef:value.annotationRef as string,annotationSignature:value.annotationSignature as string}:{}), ...(value.tool === 'form' ? { fieldName: value.fieldName as string } : {}) }
}

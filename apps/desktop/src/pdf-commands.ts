import { PDFArray, PDFDict, PDFDocument, PDFHexString, PDFName, PDFTextField, PDFCheckBox, PDFDropdown, PDFOptionList, PDFRadioGroup, StandardFonts, degrees, rgb } from 'pdf-lib'
import fontkit from '@pdf-lib/fontkit'
import { applyDrawings } from '../../../packages/pdf/src/annotate/drawing'
import {annotationEntries,nameValue,numberArray,textValue} from '../../../packages/pdf/src/annotate/pdfObjects'
import { applyMarkups } from '../../../packages/pdf/src/annotate/markup'

export interface PdfFormField { name: string; kind: 'text' | 'checkbox' | 'choice' | 'radio'; value: string; options: string[]; readOnly: boolean; maxLength?: number }
export interface PdfAnnotationTarget {ref:string;subtype:string;rect:number[];contents:string;signature:string}
export type PdfCommand =
  | {kind:'annotation.delete';page:number;target:PdfAnnotationTarget}
  | {kind:'note.edit';page:number;target:Pick<PdfAnnotationTarget,'ref'|'signature'>;text:string}
  | { kind: 'line' | 'arrow'; page:number; from:[number,number]; to:[number,number]; color:string }
  | { kind: 'image'; page: number; bytes: Uint8Array }
  | { kind: 'form-value'; page: number; name: string; value: string; fontBytes?: Uint8Array }
  | { kind: 'rotate' | 'add-page' | 'delete-page'; page: number }
  | { kind: 'move-page'; page: number; to: number }
  | { kind: 'text'; page: number; at: [number, number]; text: string; size: number; color: string; fontBytes?: Uint8Array }
  | { kind: 'note'; page: number; at: [number, number]; text: string; color: string }
  | { kind: 'highlight' | 'underline' | 'strikeout' | 'rectangle' | 'ellipse'; page: number; rect: [number, number, number, number]; color: string }
export interface PdfSummary { fields: PdfFormField[]; pages: Array<{ width: number; height: number; rotation: number; annotations:PdfAnnotationTarget[] }>; editable: boolean; refusal?: string }
const options = { updateMetadata: false }
const colorValue = (value: string): [number, number, number] => {
  if (!/^#[0-9a-f]{6}$/i.test(value)) throw new Error('Choose a valid color.')
  return [1, 3, 5].map(index => parseInt(value.slice(index, index + 2), 16) / 255) as [number, number, number]
}
function hasSignature(doc: PDFDocument) {
  return doc.context.enumerateIndirectObjects().some(([, object]) => object instanceof PDFDict &&
    (object.has(PDFName.of('ByteRange')) || object.get(PDFName.of('Type'))?.toString() === '/Sig'))
}
export async function inspectPdf(bytes: Uint8Array): Promise<PdfSummary> {
  const doc = await PDFDocument.load(bytes.slice(), options)
  const graph=annotationGraph(doc)
  return { fields: readPdfFields(doc), pages: doc.getPages().map(page => ({ ...page.getSize(), rotation: page.getRotation().angle,annotations:editableAnnotations(page,graph) })), editable: !hasSignature(doc), ...(hasSignature(doc) ? { refusal: 'This PDF contains a digital signature. Editing is unavailable.' } : {}) }
}
function annotationGraph(doc:PDFDocument) {
  const counts=new Map<string,number>(),linked=new Set<string>()
  for(const page of doc.getPages())for(const entry of annotationEntries(page)) {
    if(entry.ref)counts.set(entry.ref.toString(),(counts.get(entry.ref.toString())??0)+1)
    for(const key of ['IRT','Popup','Parent']) {const value=entry.dict.get(PDFName.of(key));if(value)linked.add(value.toString())}
  }
  return {counts,linked}
}
function editableAnnotations(page:ReturnType<PDFDocument['getPage']>,graph=annotationGraph(page.doc)):PdfAnnotationTarget[] {
  const entries=annotationEntries(page)
  return entries.flatMap(entry=>{
    const subtype=nameValue(entry.dict,'Subtype'),rect=numberArray(entry.dict,'Rect'),signature=entry.dict.toString()
    const flags=Number(entry.dict.get(PDFName.of('F'))?.toString()??0)
    if(!entry.ref||!subtype||!['Text','Highlight','Underline','StrikeOut','Square','Circle','Line','Ink'].includes(subtype)||!rect||rect.length!==4||!rect.every(Number.isFinite)||signature.length>32768||!Number.isSafeInteger(flags)||(flags&(64|128|512)))return []
    // Linked comments/popups and shared annotation identities need a graph-aware edit.
    if(['IRT','Popup','Parent'].some(key=>entry.dict.has(PDFName.of(key))))return []
    if(graph.counts.get(entry.ref.toString())!==1||graph.linked.has(entry.ref.toString()))return []
    return [{ref:entry.ref.toString(),subtype,rect,contents:textValue(entry.dict,'Contents')??'',signature}]
  })
}
function readPdfFields(doc: PDFDocument): PdfFormField[] {
  return doc.getForm().getFields().flatMap((field): PdfFormField[] => {
    const base = { name: field.getName(), readOnly: field.isReadOnly(), options: [] as string[] }
    if (field instanceof PDFTextField) return [{ ...base, kind: 'text', value: field.getText() ?? '', maxLength: field.getMaxLength(), readOnly: base.readOnly || field.isPassword() }]
    if (field instanceof PDFCheckBox) return [{ ...base, kind: 'checkbox', value: String(field.isChecked()) }]
    if (field instanceof PDFRadioGroup) return [{ ...base, kind: 'radio', value: field.getSelected() ?? '', options: field.getOptions() }]
    if (field instanceof PDFDropdown || field instanceof PDFOptionList) return [{ ...base, kind: 'choice', value: field.getSelected()[0] ?? '', options: field.getOptions(), readOnly: base.readOnly || field.isMultiselect() }]
    return []
  })
}
function annotations(page: ReturnType<PDFDocument['getPage']>) {
  const array = page.node.lookupMaybe(PDFName.of('Annots'), PDFArray)
  return array ? array.asArray().map(ref => page.doc.context.lookup(ref, PDFDict)) : []
}
async function qualifiedPdfFont(doc: PDFDocument, text: string, bytes?: Uint8Array) {
  if (!bytes) return doc.embedFont(StandardFonts.Helvetica)
  if (!bytes.length || bytes.length > 8 * 1024 * 1024) throw new Error('The bundled PDF font is unavailable.')
  // pdf-lib writes advances but not arbitrary complex-script positioning.
  // Qualify the actual local font and layout instead of promising all Unicode.
  if (!/^[\p{Script=Latin}\p{Script=Greek}\p{Script=Cyrillic}\p{Number}\p{Punctuation}\p{Symbol}\p{Separator}\r\n]*$/u.test(text)) throw new Error('This text needs script shaping that is not supported yet. Use Latin, Greek, or Cyrillic text.')
  const face = fontkit.create(bytes)
  const features = { liga: false, kern: false }
  for (const line of text.split(/\r?\n/)) {
    if (Array.from(line).some(character => !face.hasGlyphForCodePoint(character.codePointAt(0)!))) throw new Error('The bundled font does not contain one or more characters.')
    const layout = face.layout(line, features)
    if (layout.direction === 'rtl' || layout.glyphs.length !== Array.from(line).length || layout.positions.some(position => position.xOffset !== 0 || position.yOffset !== 0 || position.yAdvance !== 0)) throw new Error('This text needs positioning that is not supported yet.')
  }
  doc.registerFontkit(fontkit)
  return doc.embedFont(bytes, { subset: true, features })
}
export async function applyPdfCommand(source: Uint8Array, command: PdfCommand): Promise<Uint8Array> {
  if (!source.length || source.length > 128 * 1024 * 1024) throw new Error('The PDF must be smaller than 128 MB.')
  const doc = await PDFDocument.load(source.slice(), options)
  if (hasSignature(doc)) throw new Error('Signed PDFs cannot be edited.')
  const pages = doc.getPages()
  if (!Number.isInteger(command.page) || command.page < 1 || command.page > pages.length) throw new Error('Choose an existing page.')
  const page = pages[command.page - 1]
  const refs = pages.map(value => value.ref.toString())
  const beforeAnnots = annotations(page).length
  const bounds = page.getCropBox()
  const point = (at: number[]) => {
    if (at.length !== 2 || at.some(value => !Number.isFinite(value)) || at[0] < bounds.x || at[0] > bounds.x + bounds.width || at[1] < bounds.y || at[1] > bounds.y + bounds.height) throw new Error('Place content inside the page.')
  }
  let saved: Uint8Array
  let expectedRotation = page.getRotation().angle
  switch (command.kind) {
    case 'note.edit': {
      if(typeof command.text!=='string'||!command.text.trim()||command.text.length>10000)throw new Error('Enter a note, up to 10,000 characters.')
      const target=editableAnnotations(page).find(value=>value.subtype==='Text'&&value.ref===command.target.ref&&value.signature===command.target.signature)
      if(!target)throw new Error('This note changed or cannot be edited safely. Select it again.')
      const entry=annotationEntries(page).find(value=>value.ref?.toString()===target.ref)!
      entry.dict.set(PDFName.of('Contents'),PDFHexString.fromText(command.text))
      saved=await doc.save()
      const check=await PDFDocument.load(saved,options),result=editableAnnotations(check.getPage(command.page-1)).find(value=>value.ref===target.ref)
      if(result?.contents!==command.text||annotations(check.getPage(command.page-1)).length!==beforeAnnots)throw new Error('The note edit could not be verified.')
      return saved
    }
    case 'annotation.delete': {
      const target=editableAnnotations(page).find(value=>value.ref===command.target.ref&&value.signature===command.target.signature)
      if(!target)throw new Error('This annotation changed or cannot be deleted safely. Select it again.')
      const entry=annotationEntries(page).find(value=>value.ref?.toString()===target.ref)!
      page.node.Annots()!.remove(entry.index)
      // Leave the detached object in the package; other non-annotation objects may reference it.
      saved=await doc.save()
      const check=await PDFDocument.load(saved,options),remaining=annotationEntries(check.getPage(command.page-1))
      if(remaining.length!==beforeAnnots-1||remaining.some(value=>value.ref?.toString()===target.ref))throw new Error('The annotation deletion could not be verified.')
      return saved
    }
    case 'image': {
      if (!command.bytes.length || command.bytes.length > 32 * 1024 * 1024) throw new Error('Images must be smaller than 32 MB.')
      const imageBytes = command.bytes.slice()
      const png = imageBytes.length >= 24 && [137,80,78,71,13,10,26,10].every((value,index) => imageBytes[index] === value)
      if (png) {
        const header = new DataView(imageBytes.buffer, imageBytes.byteOffset, imageBytes.byteLength)
        const width = header.getUint32(16), height = header.getUint32(20)
        if (!width || !height || width * height > 40000000) throw new Error('Images must contain at most 40 million pixels.')
      } else if (imageBytes[0] !== 255 || imageBytes[1] !== 216) throw new Error('Choose a PNG or JPEG image.')
      const image = png ? await doc.embedPng(imageBytes) : await doc.embedJpg(imageBytes)
      if (image.width * image.height > 40000000) throw new Error('Images must contain at most 40 million pixels.')
      const rotation = page.getRotation().angle, sideways = Math.abs(rotation % 180) === 90
      const scale = Math.min((sideways ? bounds.height : bounds.width) * .6 / image.width, (sideways ? bounds.width : bounds.height) * .6 / image.height)
      const width = image.width * scale, height = image.height * scale, angle = rotation * Math.PI / 180
      const x = bounds.x + bounds.width / 2 - width / 2 * Math.cos(angle) + height / 2 * Math.sin(angle)
      const y = bounds.y + bounds.height / 2 - width / 2 * Math.sin(angle) - height / 2 * Math.cos(angle)
      page.drawImage(image, { x, y, width, height, rotate: degrees(rotation) }); break
    }
    case 'form-value': {
      const info = readPdfFields(doc).find(field => field.name === command.name)
      if (!info || info.readOnly) throw new Error('This form field cannot be edited.')
      if (typeof command.value !== 'string' || command.value.length > 10000 || (info.maxLength !== undefined && command.value.length > info.maxLength)) throw new Error('The field value is too long.')
      const form = doc.getForm(), field = form.getField(command.name)
      if (field instanceof PDFTextField) field.setText(command.value)
      else if (field instanceof PDFCheckBox) { if (!['true', 'false'].includes(command.value)) throw new Error('Invalid checkbox value.'); if (command.value === 'true') field.check(); else field.uncheck() }
      else if (field instanceof PDFDropdown || field instanceof PDFOptionList || field instanceof PDFRadioGroup) { if (command.value && !info.options.includes(command.value)) throw new Error('Choose a listed field option.'); if (command.value) field.select(command.value); else field.clear() }
      try { form.updateFieldAppearances(await qualifiedPdfFont(doc, command.value, command.fontBytes)) } catch (error) { if (command.fontBytes && error instanceof Error) throw error; throw new Error('This form font cannot display the value. Use Latin text for now.') }
      saved = await doc.save({ updateFieldAppearances: false })
      const reopened = readPdfFields(await PDFDocument.load(saved, options)).find(field => field.name === command.name)
      if (reopened?.value !== command.value) throw new Error('The saved form value did not match the requested value.')
      return saved
    }
    case 'rotate': expectedRotation = (expectedRotation + 90) % 360; page.setRotation(degrees(expectedRotation)); break
    case 'add-page': {
      const added = doc.insertPage(command.page, [page.getWidth(), page.getHeight()])
      refs.splice(command.page, 0, added.ref.toString()); break
    }
    case 'delete-page':
      if (pages.length === 1) throw new Error('Keep at least one page.')
      if (annotations(page).some(value => value.get(PDFName.of('Subtype'))?.toString() === '/Widget')) throw new Error('Pages with form fields cannot be deleted yet.')
      doc.removePage(command.page - 1); refs.splice(command.page - 1, 1); break
    case 'move-page':
      if (!Number.isInteger(command.to) || command.to < 1 || command.to > pages.length || command.to === command.page) throw new Error('Choose a different page position.')
      doc.removePage(command.page - 1); doc.insertPage(command.to - 1, page)
      refs.splice(command.to - 1, 0, ...refs.splice(command.page - 1, 1)); break
    case 'text': {
      point(command.at)
      if (!command.text.trim() || command.text.length > 10000 || /[\r\n\t]/.test(command.text)) throw new Error('Enter one line of text, up to 10,000 characters.')
      if (!Number.isFinite(command.size) || command.size < 6 || command.size > 144) throw new Error('Text size must be 6–144 points.')
      const font = await qualifiedPdfFont(doc, command.text, command.fontBytes)
      let width: number
      try { width = font.widthOfTextAtSize(command.text, command.size) } catch { throw new Error('This font cannot display one or more characters. Use Latin text for now.') }
      const angle = page.getRotation().angle * Math.PI / 180
      for (const [x, y] of [[0, -command.size * .25], [width, -command.size * .25], [0, command.size], [width, command.size]]) point([command.at[0] + x * Math.cos(angle) - y * Math.sin(angle), command.at[1] + x * Math.sin(angle) + y * Math.cos(angle)])
      page.drawText(command.text, { x: command.at[0], y: command.at[1], size: command.size, font, color: rgb(...colorValue(command.color)), rotate: degrees(page.getRotation().angle) }); break
    }
    case 'note':
      point(command.at)
      if (!command.text.trim() || command.text.length > 10000) throw new Error('Enter a note, up to 10,000 characters.')
      saved = await applyDrawings(source.slice(), [{ kind: 'note', page: command.page, at: command.at, contents: command.text, color: colorValue(command.color) }]); return verifyAnnotation(saved, command.page, beforeAnnots, 'Text')
    case 'line': case 'arrow': {
      point(command.from); point(command.to)
      if(Math.hypot(command.to[0]-command.from[0],command.to[1]-command.from[1])<2)throw new Error('Draw a longer line on the page.')
      saved=await applyDrawings(source.slice(),[{kind:command.kind,page:command.page,from:command.from,to:command.to,width:1.5,color:colorValue(command.color)}])
      return verifyAnnotation(saved,command.page,beforeAnnots,'Line')
    }
    case 'highlight': case 'underline': case 'strikeout': case 'rectangle': case 'ellipse': {
      const [x1, y1, x2, y2] = command.rect
      point([x1, y1]); point([x2, y2]); if (x2 - x1 < 2 || y2 - y1 < 2) throw new Error('Drag a larger area on the page.')
      const rotation=((page.getRotation().angle%360)+360)%360
      const quads:Record<number,[number,number,number,number,number,number,number,number]>={0:[x1,y2,x2,y2,x1,y1,x2,y1],90:[x1,y1,x1,y2,x2,y1,x2,y2],180:[x2,y1,x1,y1,x2,y2,x1,y2],270:[x2,y2,x2,y1,x1,y2,x1,y1]}
      if(!quads[rotation])throw new Error('This page rotation is not supported for markup.')
      saved = command.kind === 'rectangle' || command.kind === 'ellipse'
        ? await applyDrawings(source.slice(), [{ kind: command.kind === 'ellipse' ? 'ellipse' : 'rect', page: command.page, rect: command.rect, width: 1.5, color: colorValue(command.color) }])
        : await applyMarkups(source.slice(), [{ page: command.page, type: command.kind, color: colorValue(command.color), quads: [quads[rotation]] }])
      return verifyAnnotation(saved, command.page, beforeAnnots, ({rectangle:'Square',ellipse:'Circle',highlight:'Highlight',underline:'Underline',strikeout:'StrikeOut'})[command.kind])
    }
  }
  saved = await doc.save()
  const check = await PDFDocument.load(saved, options)
  if (check.getPages().map(value => value.ref.toString()).join(',') !== refs.join(',')) throw new Error('Saved page order did not match the requested change.')
  if (command.kind === 'rotate' && check.getPage(command.page - 1).getRotation().angle !== expectedRotation) throw new Error('Saved page rotation did not match.')
  return saved
}
async function verifyAnnotation(saved: Uint8Array, page: number, previous: number, subtype: string) {
  const check = await PDFDocument.load(saved, options)
  const values = annotations(check.getPage(page - 1))
  if (values.length !== previous + 1 || values.at(-1)?.get(PDFName.of('Subtype'))?.toString() !== '/' + subtype) throw new Error('The annotation could not be verified after saving.')
  return saved
}

/** Byte snapshots make every supported command reversible, including page deletion. */
export class PdfHistory {
  private entries: Uint8Array[]
  private position = 0
  constructor(bytes: Uint8Array) { this.entries = [bytes.slice()] }
  get bytes() { return this.entries[this.position] }
  get canUndo() { return this.position > 0 }
  get canRedo() { return this.position + 1 < this.entries.length }
  push(bytes: Uint8Array) {
    this.entries = [...this.entries.slice(0, this.position + 1), bytes.slice()]
    while (this.entries.length > 2 && (this.entries.length > 21 || this.entries.reduce((sum, entry) => sum + entry.length, 0) > 96 * 1024 * 1024)) this.entries.shift()
    this.position = this.entries.length - 1
    return this.bytes
  }
  undo() { if (this.canUndo) this.position--; return this.bytes }
  redo() { if (this.canRedo) this.position++; return this.bytes }
}

export { parsePdfRecoveryDraft } from './pdf-recovery'

export function findPdfTextMatches(parts: readonly string[], query: string): Array<{ start: number; end: number; spans: number[] }> {
  const needle = query.trim().toLocaleLowerCase()
  if (!needle || needle.length > 500) return []
  const folded = parts.map(part => part.toLocaleLowerCase())
  const haystack = folded.join(' ')
  const ranges: Array<{ start: number; end: number }> = []
  let cursor = 0
  for (const part of folded) { ranges.push({ start: cursor, end: cursor + part.length }); cursor += part.length + 1 }
  const found = []
  for (let at = haystack.indexOf(needle); at >= 0 && found.length < 1000; at = haystack.indexOf(needle, at + Math.max(1, needle.length))) {
    found.push({ start: at, end: at + needle.length, spans: ranges.flatMap((range, index) => range.start < at + needle.length && range.end > at ? [index] : []) })
  }
  return found
}

export async function importPdfPages(source: Uint8Array, imported: Uint8Array, afterPage: number): Promise<Uint8Array> {
  if (!imported.length || imported.length > 32 * 1024 * 1024) throw new Error('Imported PDFs must be smaller than 32 MB.')
  const [doc, incoming] = await Promise.all([PDFDocument.load(source.slice(), options), PDFDocument.load(imported.slice(), options)])
  if (hasSignature(doc) || hasSignature(incoming)) throw new Error('Signed PDFs cannot be combined.')
  if (incoming.getForm().getFields().length) throw new Error('Importing pages with form fields is not supported yet.')
  if (!Number.isInteger(afterPage) || afterPage < 1 || afterPage > doc.getPageCount()) throw new Error('Choose an existing page.')
  const copied = await doc.copyPages(incoming, incoming.getPageIndices())
  copied.forEach((page, index) => doc.insertPage(afterPage + index, page))
  const expected = doc.getPages().map(page => page.ref.toString()).join(',')
  const saved = await doc.save()
  const check = await PDFDocument.load(saved, options)
  if (check.getPages().map(page => page.ref.toString()).join(',') !== expected) throw new Error('Imported page order could not be verified.')
  return saved
}
export function parsePdfPageRange(text: string, count: number): number[] {
  if (text.length > 1000 || !Number.isInteger(count) || count < 1) throw new Error('Enter page numbers, such as 1–3, 5.')
  const selected = new Set<number>()
  for (const segment of text.split(',')) {
    const match = segment.trim().match(/^(\d+)(?:\s*[-–]\s*(\d+))?$/)
    if (!match) throw new Error('Enter page numbers, such as 1-3, 5.')
    const from = Number(match[1]), to = Number(match[2] ?? match[1])
    if (from < 1 || to < from || to > count) throw new Error(`Choose pages between 1 and ${count}.`)
    for (let page = from; page <= to; page++) selected.add(page)
  }
  return [...selected].sort((a, b) => a - b)
}
export async function exportPdfPages(source: Uint8Array, selected: number[]): Promise<Uint8Array> {
  const doc = await PDFDocument.load(source.slice(), options)
  if (!selected.length || new Set(selected).size !== selected.length || selected.some(page => !Number.isInteger(page) || page < 1 || page > doc.getPageCount())) throw new Error('Choose valid pages to export.')
  if (doc.getForm().getFields().length) throw new Error('Exporting page ranges from PDFs with forms is not supported yet.')
  if (hasSignature(doc)) throw new Error('Exporting pages would remove the digital signature and is unavailable.')
  const output = await PDFDocument.create()
  for (const page of await output.copyPages(doc, selected.map(value => value - 1))) output.addPage(page)
  const saved = await output.save()
  if ((await PDFDocument.load(saved, options)).getPageCount() !== selected.length) throw new Error('Exported page count could not be verified.')
  return saved
}

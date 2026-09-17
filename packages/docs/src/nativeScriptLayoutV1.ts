import { scaleFontUnits, type FontResource } from '@injoffice/font-metrics/layout'

/** Font-authored OS/2 simulation metrics, not a Word-layout equivalence claim. */
export interface NativeDocxScriptTransformV1 {
  kind: 'subscript' | 'superscript'
  font_sha256: string
  units_per_em: number
  x_size: number
  y_size: number
  x_offset: number
  y_offset: number
}

export function validateNativeDocxScriptTransformV1(value: unknown): value is NativeDocxScriptTransformV1 {
  if (!value || typeof value !== 'object') return false
  const v = value as NativeDocxScriptTransformV1
  if (Object.keys(v).sort().join(',') !== 'font_sha256,kind,units_per_em,x_offset,x_size,y_offset,y_size' || (v.kind !== 'subscript' && v.kind !== 'superscript') || !/^sha256:[0-9a-f]{64}$/.test(v.font_sha256)) return false
  return [v.units_per_em,v.x_size,v.y_size,v.x_offset,v.y_offset].every((n) => Number.isSafeInteger(n) && !Object.is(n,-0)) && v.units_per_em >= 16 && v.units_per_em <= 16384 && v.x_size > 0 && v.x_size < v.units_per_em && v.y_size > 0 && v.y_size < v.units_per_em && Math.abs(v.x_offset) <= v.units_per_em && v.y_offset > 0 && v.y_offset <= v.units_per_em
}

/** Caller supplies an already digest-verified resource from the native resolver. */
export function readNativeDocxScriptTransformV1(resource: FontResource, kind: 'subscript' | 'superscript'): NativeDocxScriptTransformV1 {
  const bytes = resource.bytes, view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const bounds = (offset: number, size: number) => Number.isSafeInteger(offset) && offset >= 0 && size >= 0 && offset + size <= bytes.length
  if (!bounds(0,12)) throw new TypeError('Script font header is truncated')
  let offset = 0
  if (view.getUint32(0) === 0x74746366) {
    const count = view.getUint32(8), index = resource.face.collectionIndex ?? 0
    if (count < 1 || count > 256 || index < 0 || index >= count || !bounds(12,count*4)) throw new TypeError('Script font collection index is invalid')
    offset = view.getUint32(12 + index*4)
  } else if (resource.face.collectionIndex !== undefined && resource.face.collectionIndex !== 0) throw new TypeError('Script font collection index requires a collection')
  if (!bounds(offset,12)) throw new TypeError('Script sfnt header is truncated')
  const count = view.getUint16(offset+4)
  if (count < 1 || count > 256 || !bounds(offset+12,count*16)) throw new TypeError('Script sfnt table directory is unbounded')
  const tables = new Map<number,{offset:number;length:number}>()
  for (let i=0;i<count;i++) {
    const entry=offset+12+i*16, tag=view.getUint32(entry), start=view.getUint32(entry+8), length=view.getUint32(entry+12)
    if (tables.has(tag) || !bounds(start,length)) throw new TypeError('Script font tables are duplicate or out of bounds')
    tables.set(tag,{offset:start,length})
  }
  const head=tables.get(0x68656164), os2=tables.get(0x4f532f32)
  if (!head || head.length<20 || !os2 || os2.length<26) throw new TypeError('Font has no complete OS/2 script metrics')
  const units=view.getUint16(head.offset+18), base=os2.offset+(kind==='subscript'?10:18)
  if (units !== resource.metrics.unitsPerEm) throw new TypeError('Script metrics do not match the resolved font units')
  const result: NativeDocxScriptTransformV1 = { kind,font_sha256:resource.face.contentDigest,units_per_em:units,x_size:view.getInt16(base),y_size:view.getInt16(base+2),x_offset:view.getInt16(base+4),y_offset:view.getInt16(base+6) }
  if (!validateNativeDocxScriptTransformV1(result)) throw new TypeError('Font script size/offset is outside the bounded reduced-text profile')
  return result
}

export function nativeDocxScriptScaleV1(value: number, transform: NativeDocxScriptTransformV1, axis: 'x'|'y'): number {
  return scaleFontUnits(value,transform.units_per_em,axis==='x'?transform.x_size:transform.y_size)
}

/**
 * Word does not follow the font's `OS/2.ySuperscriptYOffset` when it raises a
 * superscript; measured against the Word 16.112 reference exports it raises the
 * baseline by exactly the amount the script size took away, so the reduced em
 * box sits flush with the top of the run's own em box:
 *
 *   raise = fontSize - ySuperscriptYSize/unitsPerEm * fontSize
 *
 * Confirmed on every superscript in the Office reference corpus whose own font
 * size is known (Word writes its PDF baselines on a 1/300in grid, and `raise`
 * is the exact difference of the two grid-snapped sizes in each case):
 *   24pt Arial     15.60pt script, 8.40pt raise   (tdf82173_footnoteStyle/_endnoteStyle)
 *   12pt Arial      7.92pt script, 4.08pt raise   (tdf123262_textFootnoteSeparators)
 *   11pt Calibri    6.96pt script, 4.08pt raise   (footnote, floatingtbl_with_formula)
 *   10pt Arial      6.48pt script, 3.60pt raise   (four footnote-area markers)
 * `ySuperscriptYOffset` would raise 11pt Calibri by 5.24pt instead of 3.85pt,
 * pushing the script ink above the run's ascent and growing the line box.
 *
 * The subscript drop keeps the font-authored `ySubscriptYOffset`: the corpus
 * holds exactly one subscript (floatingtbl_with_formula, which produces no page
 * here), so there is no second reference to move it against.
 */
export function nativeDocxScriptShiftV1(fontSize: number, transform: NativeDocxScriptTransformV1, axis: 'x'|'y'): number {
  if (axis==='x') return scaleFontUnits(transform.x_offset,transform.units_per_em,fontSize)
  if (transform.kind==='subscript') return -scaleFontUnits(transform.y_offset,transform.units_per_em,fontSize)
  return fontSize-nativeDocxScriptScaleV1(fontSize,transform,'y')
}

/**
 * Runs whose shaped fragments come from the shared shaped-text path that carries
 * the OS/2 script transform. Text runs, and footnote/endnote reference marks,
 * whose placed decimal number is shaped through that identical path. Controls,
 * drawings and comment marks emit glyphless or image atoms that carry no
 * transform, so a script transform on them stays refused.
 */
export function nativeDocxScriptTransformEligibleRunV1(run: { readonly kind: string; readonly reference?: { readonly kind: string } }): boolean {
  if (run.kind === 'text') return true
  return run.kind === 'reference' && (run.reference?.kind === 'footnote' || run.reference?.kind === 'endnote')
}

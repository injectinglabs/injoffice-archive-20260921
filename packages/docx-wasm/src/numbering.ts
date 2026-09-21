import type { NativeDocxParagraphPropertyPatchV1 } from '@injoffice/docs/native-docx'

const NUMBERING_KEYS = ['numbering_num_id', 'numbering_level', 'numbering_kind'] as const
const NUMBERING_KINDS = ['bullet', 'decimal'] as const

export function numberingPropertyNames(properties: Record<string, unknown>): boolean {
  return NUMBERING_KEYS.some((key) => properties[key] !== undefined)
}

/** Validates a paragraph numbering patch. Returns the bounded payload or throws. */
export function validateNumberingProperties(properties: Record<string, unknown>, targetKind: string, hasRange: boolean, index: number): NativeDocxParagraphPropertyPatchV1 {
  if (targetKind !== 'paragraph') throw new TypeError(`DOCX mutation ${index} numbering is a paragraph property and needs a paragraph target.`)
  if (hasRange) throw new TypeError(`DOCX mutation ${index} numbering applies to the whole paragraph and takes no range.`)
  for (const key of Object.getOwnPropertyNames(properties)) {
    if (!(NUMBERING_KEYS as readonly string[]).includes(key)) throw new TypeError(`DOCX mutation ${index} must patch numbering on its own.`)
  }
  const patch: NativeDocxParagraphPropertyPatchV1 = {}
  if (properties.numbering_kind !== undefined) {
    if (!(NUMBERING_KINDS as readonly string[]).includes(properties.numbering_kind as string)) throw new TypeError(`DOCX mutation ${index} numbering_kind must be bullet or decimal.`)
    patch.numbering_kind = properties.numbering_kind as NativeDocxParagraphPropertyPatchV1['numbering_kind']
  }
  if (properties.numbering_num_id !== undefined) {
    if (properties.numbering_num_id === null) patch.numbering_num_id = null
    else {
      const id = properties.numbering_num_id
      if (typeof id !== 'string' || (id !== '' && id !== '0' && !/^[A-Za-z0-9][A-Za-z0-9._:-]{0,255}$/.test(id))) {
        throw new TypeError(`DOCX mutation ${index} numbering_num_id must be a bounded native identifier.`)
      }
      patch.numbering_num_id = id
    }
  }
  if (properties.numbering_level !== undefined) {
    if (properties.numbering_level === null) patch.numbering_level = null
    else {
      const level = properties.numbering_level
      if (!Number.isSafeInteger(level) || (level as number) < 0 || (level as number) > 8) {
        throw new TypeError(`DOCX mutation ${index} numbering_level must be 0..8.`)
      }
      patch.numbering_level = level as number
    }
  }
  if (patch.numbering_kind && (patch.numbering_num_id === null || patch.numbering_num_id === '' || patch.numbering_num_id === '0')) {
    throw new TypeError(`DOCX mutation ${index} numbering_kind cannot remove a list.`)
  }
  if (patch.numbering_num_id === undefined && patch.numbering_kind === undefined) {
    throw new TypeError(`DOCX mutation ${index} numbering must set numbering_num_id or numbering_kind.`)
  }
  return patch
}

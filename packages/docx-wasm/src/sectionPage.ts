import type { NativeDocxDocumentV1, NativeDocxSectionPagePayloadV1, NativeDocxSectionPagePatchV1 } from '@injoffice/docs/native-docx'
import { NativeWasmError } from '@injoffice/native-runtime'

export function validateSectionPageMutation(document: NativeDocxDocumentV1, mutation: Record<string, unknown>): NativeDocxSectionPagePayloadV1 {
  const keys = ['target_kind', 'target_id', 'expected_xml_sha256', 'operation', 'page']
  if (Object.keys(mutation).length !== keys.length || Object.keys(mutation).some(key => !keys.includes(key)) || mutation.target_kind !== 'section') throw new TypeError('Page setup requires only a section target and page geometry.')
  const section = document.sections.find(section => section.id === mutation.target_id)
  if (!section || section.anchor.xml_sha256 !== mutation.expected_xml_sha256) throw new NativeWasmError('STALE_TARGET', 'The section anchor changed.')
  if (!section.edit_policy?.allowed_operations.includes('section.page.patch')) throw new NativeWasmError('UNSUPPORTED_SECTION_STRUCTURE', section.edit_policy?.refusal?.message ?? 'This section does not allow page setup.')
  const value = mutation.page
  if (!value || typeof value !== 'object' || Array.isArray(value) || (Object.getPrototypeOf(value) !== Object.prototype && Object.getPrototypeOf(value) !== null)) throw new TypeError('Page geometry must be a plain object.')
  const page = value as Record<string, unknown>
  const fields = ['width_twips', 'height_twips', 'orientation', 'margin_top_twips', 'margin_right_twips', 'margin_bottom_twips', 'margin_left_twips']
  if (Object.keys(page).length !== fields.length || Object.keys(page).some(key => !fields.includes(key))) throw new TypeError('Page geometry requires all seven fields.')
  for (const key of fields) {
    if (key === 'orientation') {
      if (page[key] !== 'portrait' && page[key] !== 'landscape') throw new TypeError('Invalid page orientation.')
    } else if (!Number.isSafeInteger(page[key]) || (page[key] as number) < (key.startsWith('margin_') ? 0 : 1) || (page[key] as number) > 31680) throw new TypeError(`Invalid page ${key}.`)
  }
  const patch = page as unknown as NativeDocxSectionPagePatchV1
  if (patch.margin_left_twips + patch.margin_right_twips + section.page.margins.gutter_twips >= patch.width_twips || patch.margin_top_twips + patch.margin_bottom_twips >= patch.height_twips) throw new TypeError('Page margins must leave positive body dimensions.')
  return { mutations: [{ target_kind: 'section', target_id: section.id, expected_xml_sha256: section.anchor.xml_sha256, operation: 'section.page.patch', page: { ...patch } }] }
}

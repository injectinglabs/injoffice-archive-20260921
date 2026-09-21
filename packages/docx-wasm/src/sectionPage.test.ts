import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { decodeNativeDocxDocument } from '@injoffice/docs/native-docx'
import { validateSectionPageMutation } from './sectionPage'

const fixture = JSON.parse(readFileSync(new URL('../../../go/officecompat/corpus/generated/expected/docx-inline-png-page-paint.json', import.meta.url), 'utf8'))
function input() {
  const decoded = decodeNativeDocxDocument(structuredClone(fixture.native))
  if (!decoded.ok) throw new Error('invalid fixture')
  const document = decoded.value, section = document.sections[0]!
  section.edit_policy = { mode: 'read-write', allowed_operations: ['section.page.patch'] }
  return { document, mutation: { target_kind: 'section', target_id: section.id, expected_xml_sha256: section.anchor.xml_sha256, operation: 'section.page.patch', page: { width_twips: 16838, height_twips: 11906, orientation: 'landscape', margin_top_twips: 720, margin_right_twips: 720, margin_bottom_twips: 720, margin_left_twips: 720 } } }
}

describe('section page WASM boundary', () => {
  it('carries every page field without changing units', () => {
    const { document, mutation } = input()
    expect(validateSectionPageMutation(document, mutation)).toEqual({ mutations: [mutation] })
  })
  it('refuses stale anchors, read-only policy and malformed geometry', () => {
    const { document, mutation } = input()
    expect(() => validateSectionPageMutation(document, { ...mutation, expected_xml_sha256: 'stale' })).toThrow(/anchor changed/)
    for (const page of [{ ...mutation.page, width_twips: 100 }, { ...mutation.page, margin_left_twips: -1 }, { ...mutation.page, width_twips: 1.5 }, { ...mutation.page, orientation: 'sideways' }, { ...mutation.page, extra: 1 }]) expect(() => validateSectionPageMutation(document, { ...mutation, page })).toThrow()
    const accessor = { ...mutation.page }
    Object.defineProperty(accessor, 'width_twips', { get() { throw new Error('getter was invoked') } })
    expect(() => validateSectionPageMutation(document, { ...mutation, page: accessor })).toThrow('only data fields')
    const hidden = { ...mutation.page }
    Object.defineProperty(hidden, 'extra', { value: 1 })
    expect(() => validateSectionPageMutation(document, { ...mutation, page: hidden })).toThrow('seven fields')
    document.sections[0]!.edit_policy = { mode: 'read-only', allowed_operations: [], refusal: { code: 'UNSUPPORTED_SECTION_STRUCTURE', message: 'Multiple sections', preservation: 'refuse-mutation' } }
    expect(() => validateSectionPageMutation(document, mutation)).toThrow('Multiple sections')
  })
})

import { describe, expect, it } from 'vitest'
import { DOCX_ABSENT_FONT_FAMILY_WARNING, DOCX_ABSENT_FONT_FAMILY_HOST_DEFAULT, validNativeDocxAbsentFontFamiliesV1, validNativeDocxApproximatedFontFamiliesV1, validNativeDocxHostDefaultFamilyPolicyV1 } from './nativeAbsentFontFamilyV1.js'

const hash = `sha256:${'a'.repeat(64)}`
const fact = { scope_kind: 'run' as const, scope_id: 'run:1', part_name: 'word/document.xml', path: '/w:document[1]/w:body[1]/w:p[1]/w:r[1]', package_sha256: hash }

describe('declared host-family source evidence', () => {
  it('bounds unique source facts and rejects malformed or foreign fields', () => {
    expect(validNativeDocxAbsentFontFamiliesV1([fact], hash)).toBe(true)
    expect(validNativeDocxAbsentFontFamiliesV1([{ ...fact, scope_kind: 'paragraph-mark', scope_id: 'paragraph:1' }], hash)).toBe(true)
    for (const candidate of [
      [fact, fact],
      [{ ...fact, package_sha256: `sha256:${'b'.repeat(64)}` }],
      [{ ...fact, scope_kind: 'marker' }],
      [{ ...fact, path: '/wrong' }],
      [{ ...fact, part_name: '../document.xml' }],
      // The source records the omission only; it never carries a face.
      [{ ...fact, font_family: 'Aptos' }],
      [{ ...fact, guessed: true }],
      Array.from({ length: 1001 }, (_, i) => ({ ...fact, scope_id: `run:${i}` })),
    ]) expect(validNativeDocxAbsentFontFamiliesV1(candidate, hash)).toBe(false)
  })
  it('requires one explicit fixed host choice and names where the value came from', () => {
    expect(validNativeDocxHostDefaultFamilyPolicyV1({ kind: 'host-default-family-v1', family: 'Aptos' })).toBe(true)
    for (const policy of [
      undefined,
      { kind: 'host-default-family-v1', family: 'Calibri' },
      { kind: 'host-default-family-v1' },
      { kind: 'word-default', family: 'Aptos' },
      { kind: 'host-default-family-v1', family: 'Aptos', exact: true },
    ]) expect(validNativeDocxHostDefaultFamilyPolicyV1(policy)).toBe(false)
    expect(DOCX_ABSENT_FONT_FAMILY_HOST_DEFAULT).toBe('Aptos')
    expect(DOCX_ABSENT_FONT_FAMILY_WARNING).toContain('Microsoft Word 16.112.4')
    expect(DOCX_ABSENT_FONT_FAMILY_WARNING).toContain('not a documented Microsoft Word default')
  })
  it('accepts applied families only as a non-empty subset of the retained evidence', () => {
    const other = { ...fact, scope_id: 'run:2', path: '/w:document[1]/w:body[1]/w:p[1]/w:r[2]' }
    const applied = { ...fact, chosen_family: 'Aptos' as const }
    expect(validNativeDocxApproximatedFontFamiliesV1([applied], [fact, other], hash)).toBe(true)
    expect(validNativeDocxApproximatedFontFamiliesV1([], [fact], hash)).toBe(false)
    expect(validNativeDocxApproximatedFontFamiliesV1([{ ...other, chosen_family: 'Aptos' }], [fact], hash)).toBe(false)
    expect(validNativeDocxApproximatedFontFamiliesV1([{ ...fact, chosen_family: 'Calibri' }], [fact], hash)).toBe(false)
    expect(validNativeDocxApproximatedFontFamiliesV1([fact], [fact], hash)).toBe(false)
    expect(validNativeDocxApproximatedFontFamiliesV1([applied, applied], [fact], hash)).toBe(false)
  })
})

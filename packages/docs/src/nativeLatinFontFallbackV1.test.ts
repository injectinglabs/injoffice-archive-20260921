import { describe, expect, it } from 'vitest'
import { DOCX_LATIN_FONT_FALLBACK_WARNING, validNativeDocxApproximatedFontFacesV1, validNativeDocxLatinFontFallbacksV1 } from './nativeLatinFontFallbackV1.js'

const hash = `sha256:${'a'.repeat(64)}`
const fact = { scope_kind: 'run' as const, scope_id: 'run:1', part_name: 'word/document.xml', path: '/w:document[1]/w:body[1]/w:p[1]/w:r[1]', font_family: 'Calibri', package_sha256: hash }

describe('Latin font fallback evidence', () => {
  it('bounds unique source facts and rejects malformed, foreign, or unattested fields', () => {
    expect(validNativeDocxLatinFontFallbacksV1([fact], hash)).toBe(true)
    expect(validNativeDocxLatinFontFallbacksV1([{ ...fact, scope_kind: 'paragraph-mark', scope_id: 'paragraph:1' }], hash)).toBe(true)
    for (const candidate of [
      [fact, fact],
      [{ ...fact, package_sha256: `sha256:${'b'.repeat(64)}` }],
      [{ ...fact, scope_kind: 'marker' }],
      [{ ...fact, font_family: '' }],
      [{ ...fact, font_family: ' Calibri' }],
      [{ ...fact, font_family: 'x'.repeat(257) }],
      [{ ...fact, path: '/wrong' }],
      [{ ...fact, part_name: '../document.xml' }],
      [{ ...fact, guessed: true }],
      Array.from({ length: 1001 }, (_, i) => ({ ...fact, scope_id: `run:${i}` })),
    ]) expect(validNativeDocxLatinFontFallbacksV1(candidate, hash)).toBe(false)
  })
  it('accepts applied faces only as a non-empty subset of the retained evidence', () => {
    const other = { ...fact, scope_id: 'run:2', path: '/w:document[1]/w:body[1]/w:p[1]/w:r[2]' }
    expect(validNativeDocxApproximatedFontFacesV1([fact], [fact, other], hash)).toBe(true)
    expect(validNativeDocxApproximatedFontFacesV1([], [fact], hash)).toBe(false)
    expect(validNativeDocxApproximatedFontFacesV1([{ ...fact, font_family: 'Arial' }], [fact], hash)).toBe(false)
    expect(validNativeDocxApproximatedFontFacesV1([other], [fact], hash)).toBe(false)
    expect(DOCX_LATIN_FONT_FALLBACK_WARNING).toContain('empty East-Asian or complex-script font slot')
  })
})

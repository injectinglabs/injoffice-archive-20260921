import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { decodeNativeDOCXFontInventoryV1, encodeNativeDOCXFontInventoryV1 } from './nativeFontInventoryV1.js'

const GOLDEN = readFileSync(new URL('../../../go/docxpatch/testdata/font-inventory-v1.json', import.meta.url), 'utf8').trim()

describe('canonical Go native DOCX font inventory v1', () => {
  it('decodes and byte-for-byte re-encodes the Go extractor golden', () => {
    const decoded = decodeNativeDOCXFontInventoryV1(GOLDEN)
    expect(decoded.main_sha256).toMatch(/^sha256:/)
    expect(decoded.families[0]!.faces[0]!.source).toMatchObject({ face_slot: 'embedRegular', relationship_target: 'Fonts/Face.ODTTF' })
    expect(encodeNativeDOCXFontInventoryV1(decoded)).toBe(GOLDEN)
  })

  it('rejects unknown, duplicate, incomplete, stale, and tampered attestations', () => {
    const attacks = [
      GOLDEN.replace('{"protocol":', '{"unknown":true,"protocol":'),
      GOLDEN.replace('{"protocol":', '{"protocol":"injoffice.docx.font-inventory","protocol":'),
      GOLDEN.replace(',"main_sha256":"sha256:', ',"main_sha256_missing":"sha256:'),
      GOLDEN.replace('rev:fcec6c78d5ae346010193ebd28219f13', 'rev:aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa'),
      GOLDEN.replace('Fonts/Face.ODTTF', 'Fonts/Other.ODTTF'),
      GOLDEN.replace('"embedding_rights":"installable"', '"embedding_rights":"editable"'),
    ]
    for (const attack of attacks) expect(() => decodeNativeDOCXFontInventoryV1(attack)).toThrow()
  })
})

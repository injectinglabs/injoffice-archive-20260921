import { createHash } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import {
  UNICODE_13_GENERATOR_SHA256,
  UNICODE_13_PROJECTION_SHA256,
  UNICODE_13_TABLES_ENCODING_SHA256,
  UNICODE_13_TABLES_RUNTIME_MATCH,
  UNICODE_13_UCD_SOURCE_SHA256,
  UNICODE_13_VERSION,
  isUnicode13Control,
  isUnicode13DefaultIgnorable,
  isUnicode13PrivateUse,
  isUnicode13TextWhiteSpace,
  unicode13Punctuation,
  unicode13Script,
} from './unicode13.js'

function digest(path: URL): `sha256:${string}` {
  return `sha256:${createHash('sha256').update(readFileSync(path)).digest('hex')}`
}

describe('generated Unicode 13 classification boundary', () => {
  it('binds official sources, committed projection, generator, and runtime encoding', () => {
    expect(UNICODE_13_VERSION).toBe('13.0.0')
    expect(UNICODE_13_TABLES_RUNTIME_MATCH).toBe(true)
    expect(UNICODE_13_UCD_SOURCE_SHA256).toEqual({
      'Scripts.txt': 'sha256:9a5ed1ec9b5f0d7147e9371ad792ab39203611af7637cff2aa4a5c663b172cde',
      'PropList.txt': 'sha256:485b5a3ed25dbf1f94dfa5a9b69d8b4550ffd0c33045ccc55ccfd7c80b2a40cf',
      'DerivedCoreProperties.txt': 'sha256:a5d45f59b39deaab3c72ce8c1a2e212a5e086dff11b1f9d5bb0e352642e82248',
      'UnicodeData.txt': 'sha256:bdbffbbfc8ad4d3a6d01b5891510458f3d36f7170422af4ea2bed3211a73e8bb',
      'DerivedAge.txt': 'sha256:e779a443d3aa2a3166a15becaa2b737c922480e32c0453d5956093633555078f',
    })
    expect(digest(new URL('../../../scripts/unicode13/ucd13-projection.txt', import.meta.url))).toBe(UNICODE_13_PROJECTION_SHA256)
    expect(digest(new URL('../../../scripts/generate-unicode13-tables.mjs', import.meta.url))).toBe(UNICODE_13_GENERATOR_SHA256)
    expect(UNICODE_13_TABLES_ENCODING_SHA256).toMatch(/^sha256:[a-f0-9]{64}$/)
  })

  it('uses Unicode 13 Script/Common/Inherited vectors independent of the host runtime', () => {
    expect(unicode13Script(0x0041)).toBe('Latn')
    expect(unicode13Script(0x05d0)).toBe('Hebr')
    expect(unicode13Script(0x0301)).toBe('Zinh')
    expect(unicode13Script(0x16fe2)).toBe('Zyyy')
    expect(unicode13Script(0x16fe3)).toBe('Zyyy')
  })

  it('pins whitespace, punctuation, default-ignorable, and control membership', () => {
    expect(isUnicode13TextWhiteSpace('\t \u00a0\u3000')).toBe(true)
    expect(isUnicode13TextWhiteSpace('\u200b')).toBe(false)
    expect(unicode13Punctuation('('.codePointAt(0)!)).toBe('open')
    expect(unicode13Punctuation(')'.codePointAt(0)!)).toBe('close')
    expect(unicode13Punctuation(0x16fe2)).toBe('close')
    expect(unicode13Punctuation(0x16fe3)).toBe('none')
    expect(isUnicode13DefaultIgnorable(0x200d)).toBe(true)
    expect(isUnicode13Control(0x2028)).toBe(true)
    expect(isUnicode13Control(0x0041)).toBe(false)
  })

  // Private Use is normative and version-independent, so it is a range test
  // rather than a table lookup: `Scripts.txt` lists no script for it at all,
  // which is why `unicode13Script` answers `'Other'` there. That answer must
  // not be read as "a script this build does not classify".
  it('separates Private Use from a script the pinned tables do not classify', () => {
    for (const codePoint of [0xe000, 0xf0a7, 0xf0b7, 0xf8ff, 0xf0000, 0xffffd, 0x100000, 0x10fffd]) {
      expect(isUnicode13PrivateUse(codePoint), codePoint.toString(16)).toBe(true)
      expect(unicode13Script(codePoint), codePoint.toString(16)).toBe('Other')
    }
    for (const codePoint of [0xdfff, 0xf900, 0xeffff, 0x100000 - 1, 0xffffe, 0x10fffe, 0x0e01, 0x0041, -1, 0x110000]) {
      expect(isUnicode13PrivateUse(codePoint), codePoint.toString(16)).toBe(false)
    }
    // Thai is the contrast case: a real script this build does not classify,
    // which answers 'Other' for a reason Private Use never does.
    expect(unicode13Script(0x0e01)).toBe('Other')
    expect(isUnicode13PrivateUse(0x0e01)).toBe(false)
  })

  it('keeps authoritative native text modules free of host Unicode and locale classification', () => {
    const sources = [
      new URL('./harfbuzz.ts', import.meta.url),
      new URL('../../docs/src/nativeShapingLines.ts', import.meta.url),
      new URL('../../docs/src/nativeResolvedLayout.ts', import.meta.url),
      new URL('../../docs/src/nativeFontInventoryV1.ts', import.meta.url),
      new URL('../../docs/src/nativeContract.ts', import.meta.url),
      new URL('../../docs/src/nativePaginationV1.ts', import.meta.url),
      new URL('../../docs/src/nativePaginatedLayoutContract.ts', import.meta.url),
      new URL('../../docs/src/nativePagePaintV1.ts', import.meta.url),
      new URL('../../docs/src/nativePagePaintCompilerV1.ts', import.meta.url),
      new URL('../../docs/src/nativeImagePagePaintV1.ts', import.meta.url),
      new URL('../../docs/src/nativeDeterminism.ts', import.meta.url),
    ].map((url) => readFileSync(url, 'utf8'))
    for (const source of sources) {
      expect(source).not.toMatch(/\\p\{|\.to(?:Locale)?(?:Lower|Upper)Case\(|\.localeCompare\(|\.normalize\(|\.trim\(/)
    }
  })
})

import { describe, expect, it } from 'vitest'
import {
  NATIVE_TEXT_LAYOUT_VERSION,
  classifyNativeOfficeLineBreak,
  classifyNativeOfficeLineBreakRanges,
  fontResolutionCacheKey,
  normalizeFontFamilyName,
  scaleFontUnits,
  scaleLineMetrics,
  shapingCacheKey,
  validateFontManifest,
  validateTextRunInput,
  type FontResolutionRequest,
  type NativeFontManifest,
  type NativeFontResolver,
  type NativeTextShaper,
  type ResolvedFontFace,
  type TextRunInput,
} from './layout.js'

const digest = `sha256:${'a'.repeat(64)}` as const

const manifest: NativeFontManifest = {
  version: NATIVE_TEXT_LAYOUT_VERSION,
  manifestId: 'injoffice.default',
  revision: '2026.08.27.1',
  faces: [
    {
      faceId: 'carlito.regular',
      family: 'Carlito',
      postscriptName: 'Carlito-Regular',
      weight: 400,
      style: 'normal',
      stretch: 100,
      source: { kind: 'bundled', resourceId: 'fonts/carlito-regular.ttf', contentDigest: digest },
      scripts: ['Latn'],
      languages: ['en-US'],
    },
  ],
  fallbackChains: [{ chainId: 'latin.default', faceIds: ['carlito.regular'], scripts: ['Latn'] }],
}

const run: TextRunInput = {
  version: NATIVE_TEXT_LAYOUT_VERSION,
  text: 'Office ffi',
  fontSizeMilliPoints: 12_000,
  font: {
    families: ['Calibri', 'Carlito'],
    weight: 400,
    style: 'normal',
    stretch: 100,
    fallbackChainIds: ['latin.default'],
  },
  script: 'Latn',
  language: 'en-US',
  direction: 'ltr',
  features: [{ tag: 'liga', value: 1 }],
}

describe('native Office line-break boundaries', () => {
  it('classifies glue, explicit opportunities, and East-Asian punctuation without provider hints', () => {
    expect(classifyNativeOfficeLineBreak('A', '\u00a0')).toBe('prohibited')
    expect(classifyNativeOfficeLineBreak('\u00a0', 'B')).toBe('prohibited')
    expect(classifyNativeOfficeLineBreak('A', '\u2060')).toBe('prohibited')
    expect(classifyNativeOfficeLineBreak('\u2060', 'B')).toBe('prohibited')
    expect(classifyNativeOfficeLineBreak('\u200b', 'B')).toBe('allowed')
    expect(classifyNativeOfficeLineBreak('-', 'B')).toBe('allowed')
    expect(classifyNativeOfficeLineBreak('/', 'B')).toBe('allowed')
    expect(classifyNativeOfficeLineBreak('\u6f22', '\u5b57')).toBe('allowed')
    expect(classifyNativeOfficeLineBreak('\u300c', '\u6f22')).toBe('prohibited')
    expect(classifyNativeOfficeLineBreak('\u6f22', '\u300d')).toBe('prohibited')
    expect(classifyNativeOfficeLineBreak('\u3000', 'B')).toBe('unsupported')
  })

  it('classifies shaped-cluster ranges without requiring substring allocation', () => {
    const text = 'left A-B right'
    expect(classifyNativeOfficeLineBreakRanges(text, 6, 7, text, 7, 8)).toBe('allowed')
    expect(classifyNativeOfficeLineBreakRanges(text, 5, 6, text, 6, 7)).toBe('prohibited')
    expect(classifyNativeOfficeLineBreakRanges(text, 7, 8, text, 8, 99)).toBe('unsupported')
  })

  it('reports unmodeled classes instead of guessing', () => {
    expect(classifyNativeOfficeLineBreak('A', '\ud83d\udca1')).toBe('unsupported')
    expect(classifyNativeOfficeLineBreak('\u00ad', 'B')).toBe('unsupported')
    expect(classifyNativeOfficeLineBreak('\u3041', '\u3042')).toBe('unsupported')
    expect(classifyNativeOfficeLineBreak('\u30fc', '\u30ab')).toBe('unsupported')
    expect(classifyNativeOfficeLineBreak('\u1100', '\u1161')).toBe('unsupported')
  })
})

const face: ResolvedFontFace = {
  faceId: 'carlito.regular',
  family: 'Carlito',
  postscriptName: 'Carlito-Regular',
  weight: 400,
  style: 'normal',
  stretch: 100,
  sourceKind: 'bundled',
  resourceId: 'fonts/carlito-regular.ttf',
  contentDigest: digest,
  resolution: 'substitute',
  matchedFamily: 'Calibri',
  fallbackChainId: 'latin.default',
}

const resolver: Pick<NativeFontResolver, 'providerId' | 'providerRevision'> = {
  providerId: 'fixture-resolver',
  providerRevision: '1',
}

const shaper: Pick<NativeTextShaper, 'providerId' | 'providerRevision'> = {
  providerId: 'fixture-harfbuzz',
  providerRevision: '8.3.0',
}

describe('validateFontManifest', () => {
  it('accepts a bounded content-addressed manifest and fallback chain', () => {
    expect(validateFontManifest(manifest)).toEqual({ ok: true, value: manifest })
  })

  it('rejects unknown fields, duplicate identities, bad digests, and dangling fallback references', () => {
    const invalid = {
      ...manifest,
      unexpected: true,
      faces: [
        manifest.faces[0],
        {
          ...manifest.faces[0],
          source: { ...manifest.faces[0]!.source, contentDigest: 'sha256:not-a-digest' },
        },
      ],
      fallbackChains: [{ chainId: 'latin.default', faceIds: ['missing.face'] }],
    }
    const result = validateFontManifest(invalid)
    expect(result.ok).toBe(false)
    if (result.ok) return
    expect(result.issues.map((issue) => [issue.path, issue.code])).toEqual(
      expect.arrayContaining([
        ['$.unexpected', 'unknown-field'],
        ['$.faces[1].faceId', 'duplicate'],
        ['$.faces[1].source.contentDigest', 'format'],
        ['$.fallbackChains[0].faceIds[0]', 'reference'],
      ]),
    )
  })
})

describe('validateTextRunInput', () => {
  it('accepts deterministic integer-unit shaping input', () => {
    expect(validateTextRunInput(run)).toEqual({ ok: true, value: run })
  })

  it('rejects malformed UTF-16, fractional sizes, duplicate features, invalid ranges, and unknown fields', () => {
    const invalid = {
      ...run,
      text: String.fromCharCode(0xd800),
      fontSizeMilliPoints: 12_000.5,
      extra: 'not-v1',
      features: [
        { tag: 'liga', value: 1, startUtf16: 3 },
        { tag: 'liga', value: 0, startUtf16: 3 },
      ],
    }
    const result = validateTextRunInput(invalid)
    expect(result.ok).toBe(false)
    if (result.ok) return
    expect(result.issues.map((issue) => [issue.path, issue.code])).toEqual(
      expect.arrayContaining([
        ['$.extra', 'unknown-field'],
        ['$.text', 'format'],
        ['$.fontSizeMilliPoints', 'range'],
        ['$.features[0]', 'required'],
        ['$.features[1]', 'required'],
        ['$.features[1]', 'duplicate'],
      ]),
    )
  })

  it('rejects NaN variation values and invalid script or direction values', () => {
    const result = validateTextRunInput({
      ...run,
      script: 'Latin',
      direction: 'auto',
      variations: [{ tag: 'wght', value: Number.NaN }],
    })
    expect(result.ok).toBe(false)
    if (result.ok) return
    expect(result.issues.map((issue) => issue.path)).toEqual(
      expect.arrayContaining(['$.script', '$.direction', '$.variations[0].value']),
    )
  })

  it('rejects OpenType feature ranges that split a surrogate pair', () => {
    const result = validateTextRunInput({
      ...run,
      text: 'A😀B',
      features: [{ tag: 'liga', value: 1, startUtf16: 1, endUtf16: 2 }],
    })
    expect(result.ok).toBe(false)
    if (result.ok) return
    expect(result.issues).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ path: '$.features[0].endUtf16', code: 'range' }),
      ]),
    )
  })
})

describe('font metrics', () => {
  it('scales signed design units to integer milli-points', () => {
    expect(scaleFontUnits(750, 1_000, 12_000)).toBe(9_000)
    expect(scaleFontUnits(-250, 1_000, 12_000)).toBe(-3_000)
    expect(() => scaleFontUnits(1, 0, 12_000)).toThrow(RangeError)
  })

  it('derives shared slide/DOCX line metrics without pixels', () => {
    expect(
      scaleLineMetrics(
        { unitsPerEm: 1_000, ascender: 800, descender: -200, lineGap: 100, capHeight: 700, underlinePosition: -100 },
        10_000,
      ),
    ).toEqual({
      fontSizeMilliPoints: 10_000,
      ascentMilliPoints: 8_000,
      descentMilliPoints: -2_000,
      lineGapMilliPoints: 1_000,
      lineHeightMilliPoints: 11_000,
      capHeightMilliPoints: 7_000,
      xHeightMilliPoints: undefined,
      underlinePositionMilliPoints: -1_000,
      underlineThicknessMilliPoints: undefined,
    })
  })

  it('refuses a line-height sum outside deterministic integer precision', () => {
    expect(() =>
      scaleLineMetrics(
        { unitsPerEm: 1, ascender: Number.MAX_SAFE_INTEGER, descender: -1, lineGap: 0 },
        1,
      ),
    ).toThrow(RangeError)
  })
})

describe('deterministic keys', () => {
  it('normalizes only pinned ASCII font-family case and whitespace', () => {
    expect(normalizeFontFamilyName('  Calibri   Light ')).toBe('calibri light')
    expect(normalizeFontFamilyName('Ｃａｌｉｂｒｉ')).toBe('Ｃａｌｉｂｒｉ')
  })

  it('keys font resolution by provider, manifest revision, request, and text coverage', () => {
    const request: FontResolutionRequest = { manifest, run }
    const first = fontResolutionCacheKey(request, resolver)
    const equivalent = fontResolutionCacheKey(
      { ...request, run: { ...run, font: { ...run.font, families: ['CALIBRI', 'carlito'] } } },
      resolver,
    )
    expect(equivalent).toBe(first)
    expect(fontResolutionCacheKey({ ...request, run: { ...run, text: 'different' } }, resolver)).not.toBe(first)
    expect(fontResolutionCacheKey({ ...request, run: { ...run, direction: 'rtl' } }, resolver)).not.toBe(first)
    expect(fontResolutionCacheKey({ ...request, run: { ...run, fontSizeMilliPoints: run.fontSizeMilliPoints + 1 } }, resolver)).not.toBe(first)
    expect(fontResolutionCacheKey({ ...request, run: { ...run, features: [{ tag: 'vert', value: 1 }] } }, resolver)).not.toBe(first)
    expect(fontResolutionCacheKey({ ...request, run: { ...run, variations: [{ tag: 'wght', value: 700 }] } }, resolver)).not.toBe(first)
    expect(fontResolutionCacheKey({ ...request, manifest: { ...manifest, revision: '2026.08.27.2' } }, resolver)).not.toBe(first)
  })

  it('canonicalizes feature and variation order but includes font bytes and provider revision', () => {
    const withOrderA: TextRunInput = {
      ...run,
      features: [
        { tag: 'kern', value: 1 },
        { tag: 'liga', value: 1 },
      ],
      variations: [
        { tag: 'wdth', value: 95 },
        { tag: 'wght', value: 450 },
      ],
    }
    const withOrderB: TextRunInput = {
      ...withOrderA,
      features: [...withOrderA.features!].reverse(),
      variations: [...withOrderA.variations!].reverse(),
    }
    expect(shapingCacheKey(withOrderA, face, shaper)).toBe(shapingCacheKey(withOrderB, face, shaper))
    expect(shapingCacheKey(withOrderA, { ...face, contentDigest: `sha256:${'b'.repeat(64)}` }, shaper)).not.toBe(
      shapingCacheKey(withOrderA, face, shaper),
    )
    expect(shapingCacheKey(withOrderA, face, { ...shaper, providerRevision: '8.4.0' })).not.toBe(
      shapingCacheKey(withOrderA, face, shaper),
    )
  })
})

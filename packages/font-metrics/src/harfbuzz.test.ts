import { createHash } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { describe, expect, it } from 'vitest'
import {
  HARFBUZZJS_ENTRY_SHA256,
  HARFBUZZJS_LOADER_SHA256,
  HARFBUZZJS_MANIFEST_SHA256,
  HARFBUZZ_RUNTIME_VERSION,
  HARFBUZZ_SHAPER_CONFIG_REVISION,
  HARFBUZZ_UNICODE_DATA_VERSION,
  HARFBUZZ_WASM_SHA256,
  createHarfBuzzTextShaperV1,
  inspectHarfBuzzFontMetricsV1,
} from './harfbuzz.js'
import {
  NATIVE_TEXT_LAYOUT_VERSION,
  shapingCacheKey,
  type FontResource,
  type NativeTextRefusal,
  type ResolvedFontFace,
  type ShapedSegment,
  type TextRunInput,
} from './layout.js'

const require = createRequire(import.meta.url)
const FONT_PATH = require.resolve('dejavu-fonts-ttf/ttf/DejaVuSans.ttf')
const FONT_BYTES = new Uint8Array(readFileSync(FONT_PATH))
const FONT_DIGEST = 'sha256:7da195a74c55bef988d0d48f9508bd5d849425c1770dba5d7bfc6ce9ed848954' as const

function digest(bytes: Uint8Array): `sha256:${string}` {
  return `sha256:${createHash('sha256').update(bytes).digest('hex')}`
}

function readU16(bytes: Uint8Array, offset: number): number {
  return (bytes[offset]! << 8) | bytes[offset + 1]!
}

function readU32(bytes: Uint8Array, offset: number): number {
  return ((bytes[offset]! * 0x1_000000) + (bytes[offset + 1]! << 16) + (bytes[offset + 2]! << 8) + bytes[offset + 3]!) >>> 0
}

function writeU32(bytes: Uint8Array, offset: number, value: number): void {
  bytes[offset] = value >>> 24
  bytes[offset + 1] = value >>> 16
  bytes[offset + 2] = value >>> 8
  bytes[offset + 3] = value
}

function tableRecord(bytes: Uint8Array, wanted: string, sfntOffset = 0): number {
  const count = readU16(bytes, sfntOffset + 4)
  for (let index = 0; index < count; index++) {
    const record = sfntOffset + 12 + index * 16
    const name = String.fromCharCode(...bytes.subarray(record, record + 4))
    if (name === wanted) return record
  }
  throw new Error(`fixture has no ${wanted} table`)
}

function mutatedResource(mutator: (bytes: Uint8Array) => void): FontResource {
  const bytes = FONT_BYTES.slice()
  mutator(bytes)
  return { ...font, bytes, face: { ...face, contentDigest: digest(bytes) } }
}

function oneFaceTtc(): FontResource {
  const bytes = new Uint8Array(FONT_BYTES.byteLength + 16)
  bytes.set([0x74, 0x74, 0x63, 0x66, 0, 1, 0, 0, 0, 0, 0, 1, 0, 0, 0, 16])
  bytes.set(FONT_BYTES, 16)
  const count = readU16(bytes, 20)
  for (let index = 0; index < count; index++) {
    const record = 28 + index * 16
    writeU32(bytes, record + 8, readU32(bytes, record + 8) + 16)
  }
  return { ...font, bytes, face: { ...face, collectionIndex: 0, contentDigest: digest(bytes) } }
}

const face: ResolvedFontFace = Object.freeze({
  faceId: 'fixture.dejavu-sans',
  family: 'DejaVu Sans',
  postscriptName: 'DejaVuSans',
  weight: 400,
  style: 'normal',
  stretch: 100,
  sourceKind: 'bundled',
  resourceId: 'fixture/dejavu-sans-2.37',
  contentDigest: FONT_DIGEST,
  resolution: 'exact',
  matchedFamily: 'DejaVu Sans',
})

const font: FontResource = Object.freeze({
  face,
  bytes: FONT_BYTES,
  metrics: Object.freeze({
    unitsPerEm: 2_048,
    ascender: 1_901,
    descender: -483,
    lineGap: 0,
    underlinePosition: -130,
    underlineThickness: 90,
  }),
})

function run(text: string, override: Partial<TextRunInput> = {}): TextRunInput {
  return {
    version: NATIVE_TEXT_LAYOUT_VERSION,
    text,
    fontSizeMilliPoints: 12_000,
    font: { families: ['DejaVu Sans'], weight: 400, style: 'normal', stretch: 100 },
    script: 'Latn',
    language: 'en-US',
    direction: 'ltr',
    ...override,
  }
}

function shape(text: string, override: Partial<TextRunInput> = {}, resource = font): ShapedSegment | NativeTextRefusal {
  const input = run(text, override)
  return createHarfBuzzTextShaperV1({ sourceRevision: 'git:test-suite' }).shape({ run: input, startUtf16: 0, endUtf16: input.text.length, font: resource })
}

function refusalCode(result: ShapedSegment | NativeTextRefusal): string | undefined {
  return 'status' in result ? result.decisions[0]?.code : undefined
}

describe('canonical HarfBuzz text shaper v1', () => {
  it('exports the exact pinned preflight metrics for authoritative resolver joins', () => {
    expect(inspectHarfBuzzFontMetricsV1({ bytes: FONT_BYTES, contentDigest: FONT_DIGEST })).toEqual(font.metrics)
    expect(() => inspectHarfBuzzFontMetricsV1({ bytes: FONT_BYTES, contentDigest: `sha256:${'0'.repeat(64)}` })).toThrow(/digest/)
    const mutated = FONT_BYTES.slice()
    mutated[0] ^= 0xff
    expect(() => inspectHarfBuzzFontMetricsV1({ bytes: mutated, contentDigest: digest(mutated) })).toThrow()
  })
  it('pins the exact redistributable font fixture instead of reading an OS font', () => {
    expect(digest(FONT_BYTES)).toBe(FONT_DIGEST)
    expect(FONT_PATH).toContain('node_modules/dejavu-fonts-ttf/ttf/DejaVuSans.ttf')
  })

  it('binds provider identity to runtime, WASM, configuration, and source revision', () => {
    const first = createHarfBuzzTextShaperV1({ sourceRevision: 'git:23f2ac8' })
    const replay = createHarfBuzzTextShaperV1({ sourceRevision: 'git:23f2ac8' })
    const changed = createHarfBuzzTextShaperV1({ sourceRevision: 'git:different' })
    expect(first.providerId).toBe('injoffice.harfbuzzjs')
    expect(first.providerRevision).toMatch(/^sha256:[a-f0-9]{64}$/)
    expect(replay.providerRevision).toBe(first.providerRevision)
    expect(changed.providerRevision).not.toBe(first.providerRevision)
    expect(first.provenance).toMatchObject({
      provider_revision: first.providerRevision,
      source_revision: 'git:23f2ac8',
      runtime_version: HARFBUZZ_RUNTIME_VERSION,
      runtime_wasm_sha256: HARFBUZZ_WASM_SHA256,
      runtime_entry_sha256: HARFBUZZJS_ENTRY_SHA256,
      runtime_loader_sha256: HARFBUZZJS_LOADER_SHA256,
      runtime_manifest_sha256: HARFBUZZJS_MANIFEST_SHA256,
      unicode_data_version: HARFBUZZ_UNICODE_DATA_VERSION,
      configuration_revision: HARFBUZZ_SHAPER_CONFIG_REVISION,
      direction_policy: 'explicit-complete-horizontal-ltr-rtl',
      default_feature_policy: 'harfbuzz-14.3.0-shape-defaults',
    })
    expect(Object.isFrozen(first)).toBe(true)
    expect(Object.isFrozen(first.provenance)).toBe(true)
  })

  it('shapes a real default ligature into exact deterministic glyphs and logical UTF-16 clusters', () => {
    const result = shape('office')
    expect('status' in result).toBe(false)
    if ('status' in result) return
    expect(result.face.contentDigest).toBe(FONT_DIGEST)
    expect(result.glyphs.map((glyph) => glyph.glyphId)).toEqual([82, 5_044, 70, 72])
    expect(result.glyphs.map((glyph) => glyph.advanceXMilliPoints)).toEqual([7_342, 11_602, 6_598, 7_383])
    expect(result.clusters.map((cluster) => [cluster.startUtf16, cluster.endUtf16, cluster.glyphStart, cluster.glyphEnd])).toEqual([
      [0, 1, 0, 1],
      [1, 4, 1, 2],
      [4, 5, 2, 3],
      [5, 6, 3, 4],
    ])
    expect(result.clusters[1]?.advanceInlineMilliPoints).toBe(result.glyphs[1]?.advanceXMilliPoints)
    expect(result.advanceInlineMilliPoints).toBe(result.glyphs.reduce((sum, glyph) => sum + glyph.advanceXMilliPoints, 0))
    expect(result.metrics).toEqual({
      fontSizeMilliPoints: 12_000,
      ascentMilliPoints: 11_139,
      descentMilliPoints: -2_830,
      lineGapMilliPoints: 0,
      lineHeightMilliPoints: 13_969,
      capHeightMilliPoints: undefined,
      xHeightMilliPoints: undefined,
      underlinePositionMilliPoints: -762,
      underlineThicknessMilliPoints: 527,
    })
  })

  it('honors a font-advertised feature range without leaking mutable font state', () => {
    const shaper = createHarfBuzzTextShaperV1({ sourceRevision: 'git:feature-state' })
    const disabledRun = run('office', { features: [{ tag: 'liga', value: 0, startUtf16: 0, endUtf16: 6 }] })
    const disabled = shaper.shape({ run: disabledRun, startUtf16: 0, endUtf16: disabledRun.text.length, font })
    expect('status' in disabled).toBe(false)
    if ('status' in disabled) return
    expect(disabled.glyphs.map((glyph) => glyph.glyphId)).toEqual([82, 73, 73, 76, 70, 72])
    const replayRun = run('office')
    const replay = shaper.shape({ run: replayRun, startUtf16: 0, endUtf16: replayRun.text.length, font })
    expect('status' in replay).toBe(false)
    if ('status' in replay) return
    expect(replay.glyphs.map((glyph) => glyph.glyphId)).toEqual([82, 5_044, 70, 72])

    const kerned = shape('AV')
    const unkerned = shape('AV', { features: [{ tag: 'kern', value: 0 }] })
    expect('status' in kerned).toBe(false)
    expect('status' in unkerned).toBe(false)
    if (!('status' in kerned) && !('status' in unkerned)) expect(kerned.advanceInlineMilliPoints).toBeLessThan(unkerned.advanceInlineMilliPoints)
  })

  it('allows disabling absent kern without enabling unadvertised features', () => {
    const noKern = mutatedResource(bytes => {
      for (const tag of ['GPOS','kern']) {
        const record = tableRecord(bytes,tag)
        bytes[record+3] = 'X'.charCodeAt(0)
      }
    })
    const shaper = createHarfBuzzTextShaperV1({sourceRevision:'git:no-kern'})
    const disabled = run('AV',{features:[{tag:'kern',value:0}]})
    const result = shaper.shape({run:disabled,startUtf16:0,endUtf16:2,font:noKern})
    expect('status' in result).toBe(false)
    const enabled = run('AV',{features:[{tag:'kern',value:1}]})
    expect(refusalCode(shaper.shape({run:enabled,startUtf16:0,endUtf16:2,font:noKern}))).toBe('unsupported-feature')
  })

  it('keeps combining marks and supplementary Unicode on complete UTF-16 clusters', () => {
    const combining = shape('A\u0301')
    expect('status' in combining).toBe(false)
    if ('status' in combining) return
    expect(combining.glyphs).toHaveLength(1)
    expect(combining.clusters).toEqual([{ startUtf16: 0, endUtf16: 2, glyphStart: 0, glyphEnd: 1, advanceInlineMilliPoints: 8_209 }])

    const supplementary = shape('😀')
    expect('status' in supplementary).toBe(false)
    if ('status' in supplementary) return
    expect(supplementary.glyphs.map((glyph) => glyph.glyphId)).toEqual([5_857])
    expect(supplementary.clusters[0]).toMatchObject({ startUtf16: 0, endUtf16: 2 })
  })

  it('is byte-for-byte deterministic and cache keys bind all exact shaping inputs', () => {
    const shaper = createHarfBuzzTextShaperV1({ sourceRevision: 'git:determinism' })
    const input = run('AV office A\u0301')
    const request = { run: input, startUtf16: 0, endUtf16: input.text.length, font }
    const first = shaper.shape(request)
    const second = shaper.shape(request)
    expect(JSON.stringify(second)).toBe(JSON.stringify(first))
    expect(shapingCacheKey(input, face, shaper)).not.toBe(shapingCacheKey({ ...input, language: 'fr-FR' }, face, shaper))
    expect(shapingCacheKey(input, face, shaper)).not.toBe(shapingCacheKey({ ...input, script: 'Grek' }, face, shaper))
    expect(shapingCacheKey(input, face, shaper)).not.toBe(shapingCacheKey(input, { ...face, contentDigest: `sha256:${'a'.repeat(64)}` }, shaper))
  })

  it('refuses malformed Unicode, partial runs, missing glyphs, and unknown request fields', () => {
    expect(refusalCode(shape('\ud800'))).toBe('invalid-contract')
    expect(refusalCode(shape('\u{a7c7}'))).toBe('missing-glyph')
    const input = run('complete')
    const shaper = createHarfBuzzTextShaperV1({ sourceRevision: 'git:contract' })
    expect(refusalCode(shaper.shape({ run: input, startUtf16: 1, endUtf16: input.text.length, font }))).toBe('invalid-contract')
    expect(refusalCode(shaper.shape({ run: input, startUtf16: 0, endUtf16: input.text.length, font, extra: true } as never))).toBe('invalid-contract')
  })

  it('shapes explicit Arabic/Hebrew RTL runs into logical cluster contracts', () => {
    const hebrew = shape('שלום', { direction: 'rtl', script: 'Hebr', language: 'he' })
    expect('status' in hebrew).toBe(false)
    if ('status' in hebrew) return
    expect(hebrew.clusters.map((cluster) => [cluster.startUtf16, cluster.endUtf16])).toEqual([[0, 1], [1, 2], [2, 3], [3, 4]])
    expect(hebrew.glyphs.map((glyph) => glyph.glyphId)).toEqual([1_344, 1_331, 1_324, 1_332])
    expect(hebrew.glyphs.every((glyph, index) => glyph.clusterIndex === index)).toBe(true)

    const arabic = shape('مرحبا', { direction: 'rtl', script: 'Arab', language: 'ar' })
    expect('status' in arabic).toBe(false)
    if ('status' in arabic) return
    expect(arabic.clusters.map((cluster) => cluster.startUtf16)).toEqual([0, 1, 2, 3, 4])
    expect(arabic.glyphs.map((glyph) => glyph.glyphId)).toEqual([5_341, 5_288, 5_277, 5_260, 5_256])
    expect(arabic.advanceInlineMilliPoints).toBe(arabic.glyphs.reduce((sum, glyph) => sum + glyph.advanceXMilliPoints, 0))

    const pointedHebrew = shape('שָלוֹם', { direction: 'rtl', script: 'Hebr', language: 'he' })
    expect('status' in pointedHebrew).toBe(false)
    if (!('status' in pointedHebrew)) expect(pointedHebrew.clusters).toEqual(expect.arrayContaining([expect.objectContaining({ startUtf16: 0, endUtf16: 2 })]))
    const markedArabic = shape('بَ', { direction: 'rtl', script: 'Arab', language: 'ar' })
    expect('status' in markedArabic).toBe(false)
    if (!('status' in markedArabic)) expect(markedArabic.clusters).toEqual([expect.objectContaining({ startUtf16: 0, endUtf16: 2 })])

    const rtlOpen = shape('(', { direction: 'rtl', script: 'Zyyy', language: 'und' })
    const ltrClose = shape(')', { direction: 'ltr', script: 'Zyyy', language: 'und' })
    expect('status' in rtlOpen || 'status' in ltrClose).toBe(false)
    if (!('status' in rtlOpen) && !('status' in ltrClose)) expect(rtlOpen.glyphs[0]!.glyphId).toBe(ltrClose.glyphs[0]!.glyphId)
  })

  it('fails closed for controls, unsupported scripts, variations, spacing, and features', () => {
    expect(refusalCode(shape('abc', { direction: 'ttb' }))).toBe('unsupported-direction')
    expect(refusalCode(shape('abc אב'))).toBe('unsupported-script')
    expect(refusalCode(shape('abc\u2067def'))).toBe('unsupported-direction')
    expect(refusalCode(shape('\u{10c80}', { script: 'Zyyy' }))).toBe('unsupported-script')
    expect(refusalCode(shape('\u{10900}', { script: 'Zyyy' }))).toBe('unsupported-script')
    expect(refusalCode(shape('\u{10a00}', { script: 'Zyyy' }))).toBe('unsupported-script')
    expect(refusalCode(shape('\u0301', { script: 'Zinh' }))).toBe('unsupported-script')
    expect(refusalCode(shape('a \u0301'))).toBe('unsupported-script')
    expect(refusalCode(shape('a\u200db'))).toBeUndefined()
    expect(refusalCode(shape('abc', { script: 'Grek' }))).toBe('unsupported-script')
    expect(refusalCode(shape('abc', { script: 'Arab', direction: 'rtl' }))).toBe('unsupported-script')
    for (const [text, script] of [['Ж', 'Cyrl'], ['Ω', 'Grek'], ['अ', 'Deva'], ['漢', 'Hani'], ['あ', 'Kana'], ['한', 'Hang']] as const) expect(refusalCode(shape(text, { script }))).toBe('unsupported-script')
    expect(refusalCode(shape('abc', { variations: [{ tag: 'wght', value: 500 }] }))).toBe('unsupported-feature')
    expect(refusalCode(shape('abc', { letterSpacingMilliPoints: 1 }))).toBe('unsupported-feature')
    expect(refusalCode(shape('abc', { wordSpacingMilliPoints: 1 }))).toBe('unsupported-feature')
    expect(refusalCode(shape('abc', { features: [{ tag: 'zzzz', value: 1 }] }))).toBe('unsupported-feature')
    expect(refusalCode(shape('abc', { features: [{ tag: 'calt', value: 0 }] }))).toBe('unsupported-feature')
    expect(refusalCode(shape('abc', {}, { ...font, face: { ...face, sourceKind: 'system' } }))).toBe('unsupported-font-format')
  })

  it('rejects digest mismatches, fabricated metrics, collection confusion, and malformed sfnt bytes before HarfBuzz', () => {
    const digestMismatch: FontResource = { ...font, face: { ...face, contentDigest: `sha256:${'0'.repeat(64)}` } }
    expect(refusalCode(shape('abc', {}, digestMismatch))).toBe('font-digest-mismatch')

    const fabricatedMetrics: FontResource = { ...font, metrics: { ...font.metrics, ascender: 1_900 } }
    expect(refusalCode(shape('abc', {}, fabricatedMetrics))).toBe('font-metrics-unavailable')
    const omittedMetrics: FontResource = { ...font, metrics: { unitsPerEm: 2_048, ascender: 1_901, descender: -483, lineGap: 0 } }
    expect(refusalCode(shape('abc', {}, omittedMetrics))).toBe('font-metrics-unavailable')

    const collectionConfusion: FontResource = { ...font, face: { ...face, collectionIndex: 0 } }
    expect(refusalCode(shape('abc', {}, collectionConfusion))).toBe('unsupported-font-format')

    const malformedBytes = FONT_BYTES.slice()
    malformedBytes[0] = 0x77
    malformedBytes[1] = 0x4f
    malformedBytes[2] = 0x46
    malformedBytes[3] = 0x46
    const malformed: FontResource = { ...font, bytes: malformedBytes, face: { ...face, contentDigest: digest(malformedBytes) } }
    expect(refusalCode(shape('abc', {}, malformed))).toBe('unsupported-font-format')

    const badMagic = mutatedResource((bytes) => {
      const head = tableRecord(bytes, 'head')
      writeU32(bytes, readU32(bytes, head + 8) + 12, 0)
    })
    const overlapping = mutatedResource((bytes) => {
      const head = tableRecord(bytes, 'head')
      const hmtx = tableRecord(bytes, 'hmtx')
      writeU32(bytes, hmtx + 8, readU32(bytes, head + 8))
    })
    const descendingLoca = mutatedResource((bytes) => {
      const headOffset = readU32(bytes, tableRecord(bytes, 'head') + 8)
      const maxpOffset = readU32(bytes, tableRecord(bytes, 'maxp') + 8)
      const locaOffset = readU32(bytes, tableRecord(bytes, 'loca') + 8)
      const glyphCount = readU16(bytes, maxpOffset + 4)
      const longLoca = readU16(bytes, headOffset + 50) === 1
      if (longLoca) writeU32(bytes, locaOffset + glyphCount * 4, 0)
      else {
        bytes[locaOffset + glyphCount * 2] = 0
        bytes[locaOffset + glyphCount * 2 + 1] = 0
      }
    })
    const badChecksum = mutatedResource((bytes) => {
      const hmtxOffset = readU32(bytes, tableRecord(bytes, 'hmtx') + 8)
      bytes[hmtxOffset] ^= 1
    })
    for (const resource of [badMagic, overlapping, descendingLoca, badChecksum]) expect(refusalCode(shape('abc', {}, resource))).toBe('unsupported-font-format')
  })

  it('selects an explicit face from deterministic TTC bytes', () => {
    const result = shape('office', {}, oneFaceTtc())
    expect('status' in result).toBe(false)
    if ('status' in result) return
    expect(result.face.collectionIndex).toBe(0)
    expect(result.glyphs.map((glyph) => glyph.glyphId)).toEqual([82, 5_044, 70, 72])
  })

  it('enforces run and font resource bounds without partial output', () => {
    expect(refusalCode(shape('a'.repeat(262_145)))).toBe('invalid-contract')
    const oversized = new Uint8Array(64 * 1024 * 1024 + 1)
    const resource: FontResource = { ...font, bytes: oversized, face: { ...face, contentDigest: digest(oversized) } }
    const result = shape('abc', {}, resource)
    expect(refusalCode(result)).toBe('font-bytes-unavailable')
    expect('status' in result && result.status === 'refused' && !('glyphs' in result)).toBe(true)
  })

  it('does not mutate caller-owned text, face, metrics, or font bytes', () => {
    const beforeBytes = digest(FONT_BYTES)
    const beforeFace = JSON.stringify(face)
    const beforeMetrics = JSON.stringify(font.metrics)
    const input = run('office')
    const beforeRun = JSON.stringify(input)
    const result = createHarfBuzzTextShaperV1({ sourceRevision: 'git:ownership' }).shape({ run: input, startUtf16: 0, endUtf16: input.text.length, font })
    expect(refusalCode(result)).toBeUndefined()
    expect(digest(FONT_BYTES)).toBe(beforeBytes)
    expect(JSON.stringify(face)).toBe(beforeFace)
    expect(JSON.stringify(font.metrics)).toBe(beforeMetrics)
    expect(JSON.stringify(input)).toBe(beforeRun)
    expect(Object.isFrozen(result)).toBe(true)
  })

  it('rejects shared byte storage instead of racing a content digest snapshot', () => {
    const shared = new Uint8Array(new SharedArrayBuffer(FONT_BYTES.byteLength))
    shared.set(FONT_BYTES)
    const resource: FontResource = { ...font, bytes: shared }
    expect(refusalCode(shape('abc', {}, resource))).toBe('invalid-contract')
  })
})

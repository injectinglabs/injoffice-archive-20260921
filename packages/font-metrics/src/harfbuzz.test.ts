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
/** A minimal well-formed STAT v1.1 header: no design axes, no axis values, elidedFallbackNameID 2. */
const STAT_AXIS_VALUE_ONLY = new Uint8Array([0, 1, 0, 1, 0, 8, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 2])

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

/** Rebuilds the fixture with extra top-level tables, keeping every sfnt invariant the preflight checks. */
function resourceWithTables(extra: ReadonlyArray<readonly [string, Uint8Array]>): FontResource {
  const count = readU16(FONT_BYTES, 4)
  const existing = [...Array(count).keys()].map((index) => {
    const record = 12 + index * 16
    return {
      tag: String.fromCharCode(...FONT_BYTES.subarray(record, record + 4)),
      offset: readU32(FONT_BYTES, record + 8),
      length: readU32(FONT_BYTES, record + 12),
    }
  })
  const laidOut = [...existing.map((table) => ({ ...table, source: FONT_BYTES.subarray(table.offset, table.offset + table.length) })), ...extra.map(([tag, payload]) => ({ tag, offset: 0, length: payload.byteLength, source: payload }))]
    .sort((left, right) => (left.tag < right.tag ? -1 : left.tag > right.tag ? 1 : 0))
  const directoryEnd = 12 + laidOut.length * 16
  let cursor = directoryEnd + ((4 - (directoryEnd % 4)) % 4)
  for (const table of laidOut) {
    table.offset = cursor
    cursor += table.length + ((4 - (table.length % 4)) % 4)
  }
  const bytes = new Uint8Array(cursor)
  bytes.set(FONT_BYTES.subarray(0, 12))
  writeU32(bytes, 4, (laidOut.length << 16) | (readU32(FONT_BYTES, 4) & 0xffff))
  laidOut.forEach((table, index) => {
    const record = 12 + index * 16
    for (let byte = 0; byte < 4; byte++) bytes[record + byte] = table.tag.charCodeAt(byte)
    bytes.set(table.source, table.offset)
    writeU32(bytes, record + 8, table.offset)
    writeU32(bytes, record + 12, table.length)
    let checksum = 0
    for (let relative = 0; relative < table.length; relative += 4) {
      let word = 0
      for (let byte = 0; byte < 4; byte++) {
        const tableIndex = relative + byte
        word = (word * 256 + (tableIndex < table.length && !(table.tag === 'head' && tableIndex >= 8 && tableIndex < 12) ? bytes[table.offset + tableIndex]! : 0)) >>> 0
      }
      checksum = (checksum + word) >>> 0
    }
    writeU32(bytes, record + 4, checksum)
  })
  return { ...font, bytes, face: { ...face, contentDigest: digest(bytes) } }
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
    strikeoutPosition: 530,
    strikeoutThickness: 102,
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
      strikeoutPositionMilliPoints: 3_105,
      strikeoutThicknessMilliPoints: 598,
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

  it('allows disabling absent kern and makes enabling one a no-op', () => {
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
    // A face with no kern lookup has nothing for an explicit kern=1 to run, so
    // asking for it is the same shaping as not asking - not a refusal. Refusing
    // it made an inherited w:kern unpaintable on every face without GSUB/GPOS
    // kern, which is most CJK faces.
    const enabled = shaper.shape({run:run('AV',{features:[{tag:'kern',value:1}]}),startUtf16:0,endUtf16:2,font:noKern})
    const plain = shaper.shape({run:run('AV',{features:[]}),startUtf16:0,endUtf16:2,font:noKern})
    expect('status' in enabled).toBe(false)
    if ('status' in enabled || 'status' in plain) return
    expect(enabled.advanceInlineMilliPoints).toBe(plain.advanceInlineMilliPoints)
    expect(enabled.glyphs).toEqual(plain.glyphs)
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

  /**
   * LRM, RLM and ALM are Default_Ignorable_Code_Point: the Unicode core
   * specification requires a renderer to give them no visible glyph and no
   * advance, and HarfBuzz implements that by substituting the face's space
   * glyph and zeroing the advance. Word's own export agrees: in
   * tdf118361_RTLfootnoteSeparator.docx the trailing paragraph-mark glyph sits
   * the same 6.43 pt after the last authored content in the paragraph that
   * carries LRM+RLM as it does in the note-separator paragraph that carries
   * neither. Measure the advance rather than assume it.
   */
  it('shapes the implicit directional marks to a zero advance and keeps the scoped controls refused', () => {
    const spaceRun = shape(' ')
    expect('status' in spaceRun).toBe(false)
    for (const [label, text, override] of [
      ['LRM', '\u200e', {}],
      ['RLM', '\u200f', { direction: 'rtl' as const }],
      ['ALM', '\u061c', { direction: 'rtl' as const, script: 'Arab' }],
      ['LRM+RLM', '\u200e\u200f', {}],
    ] as const) {
      const result = shape(text, override)
      expect(refusalCode(result), label).toBeUndefined()
      if ('status' in result) continue
      expect(result.advanceInlineMilliPoints, label).toBe(0)
      expect(result.clusters.map((cluster) => cluster.advanceInlineMilliPoints), label).toEqual(result.clusters.map(() => 0))
      expect(result.glyphs.map((glyph) => glyph.advanceXMilliPoints), label).toEqual(result.glyphs.map(() => 0))
      // It is the face's own space glyph, never the .notdef sentinel, and the
      // cluster reports that it inks nothing so the painter accepts the empty
      // outline that glyph has.
      expect(result.glyphs.every((glyph) => glyph.glyphId !== 0), label).toBe(true)
      expect(result.clusters.every((cluster) => cluster.whitespace === true), label).toBe(true)
      if (!('status' in spaceRun)) expect(result.glyphs.every((glyph) => glyph.glyphId === spaceRun.glyphs[0]!.glyphId), label).toBe(true)
      if (!('status' in spaceRun)) expect(spaceRun.advanceInlineMilliPoints, label).toBeGreaterThan(0)
    }
    // A mark beside real text still shapes, and the text keeps its own advance.
    const marked = shape('ab\u200ecd')
    const plain = shape('abcd')
    expect('status' in marked || 'status' in plain).toBe(false)
    if (!('status' in marked) && !('status' in plain)) expect(marked.advanceInlineMilliPoints).toBe(plain.advanceInlineMilliPoints)
    // Every embedding, override and isolate control still refuses.
    for (const control of ['\u202a', '\u202b', '\u202c', '\u202d', '\u202e', '\u2066', '\u2067', '\u2068', '\u2069']) {
      expect(refusalCode(shape(`abc${control}def`)), control).toBe('unsupported-direction')
    }
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
    for (const [text, script] of [['Ж', 'Cyrl'], ['Ω', 'Grek'], ['अ', 'Deva']] as const) expect(refusalCode(shape(text, { script }))).toBe('unsupported-script')
    // Hani/Kana/Hang are qualified scripts now, so this Latin-only face refuses
    // them for the honest reason: it has no glyph, not an unknown script.
    for (const [text, script] of [['漢', 'Hani'], ['あ', 'Kana'], ['한', 'Hang']] as const) expect(refusalCode(shape(text, { script }))).toBe('missing-glyph')
    // A scalar that does not belong to the declared East-Asian script still refuses as one.
    expect(refusalCode(shape('あ', { script: 'Hani' }))).toBe('unsupported-script')
    expect(refusalCode(shape('abc', { variations: [{ tag: 'wght', value: 500 }] }))).toBe('unsupported-feature')
    expect(refusalCode(shape('abc', { wordSpacingMilliPoints: 1 }))).toBe('unsupported-feature')
    expect(refusalCode(shape('abc', { features: [{ tag: 'zzzz', value: 1 }] }))).toBe('unsupported-feature')
    expect(refusalCode(shape('abc', { features: [{ tag: 'calt', value: 0 }] }))).toBe('unsupported-feature')
    expect(refusalCode(shape('abc', {}, { ...font, face: { ...face, sourceKind: 'system' } }))).toBe('unsupported-font-format')
  })

  // A Private Use scalar has no Unicode script at all: `Scripts.txt` lists no
  // block for it, so it reaches the classifier as `'Other'` for a different
  // reason than a script this build does not qualify. Word's symbol-font
  // encoding depends on exactly that: a legacy symbol face's byte 0xB7 is
  // authored as U+F0B7, and the resolved font slot, not Unicode, decides the
  // glyph. Word paints `tdf118812_tableStyles-comprehensive.docx`'s bullet from
  // SymbolMT in its own PDF export.
  it('shapes a Private Use scalar from the face that maps it instead of calling it an unknown script', () => {
    // U+F001 and U+F002 are in this fixture's own cmap (format 12 group F000..F003).
    const mapped = shape('\uf001\uf002', { script: 'Zyyy' })
    expect(refusalCode(mapped)).toBeUndefined()
    expect('status' in mapped).toBe(false)
    if (!('status' in mapped)) {
      expect(mapped.glyphs).toHaveLength(2)
      expect(mapped.glyphs.every((glyph) => glyph.glyphId !== 0)).toBe(true)
      expect(mapped.advanceInlineMilliPoints).toBeGreaterThan(0)
    }
    // Mixed with other Common scalars in one run, which is how a Word list
    // marker and its separator arrive from the itemizer.
    expect(refusalCode(shape('\uf001 \uf002', { script: 'Zyyy' }))).toBeUndefined()
    // A Private Use scalar this face does not map refuses for the honest reason
    // - no glyph - rather than for a script that does not exist.
    expect(refusalCode(shape('\ue500', { script: 'Zyyy' }))).toBe('missing-glyph')
    // Both supplementary Private Use planes are the same fact.
    expect(refusalCode(shape('\u{f0000}', { script: 'Zyyy' }))).toBe('missing-glyph')
    expect(refusalCode(shape('\u{100000}', { script: 'Zyyy' }))).toBe('missing-glyph')
  })

  // Inertness guard: every case here already refused before Private Use was
  // admitted and must keep refusing, so it cannot fail before the change.
  it('keeps refusing a real but unqualified script carried in a Common run', () => {
    for (const text of ['\u0e01', '\u0f40', '\u0905', '\u{10c80}']) {
      expect(refusalCode(shape(text, { script: 'Zyyy' })), text).toBe('unsupported-script')
    }
    // Private Use is admitted in the Common run only: a run that declares a
    // real script still rejects a scalar that does not belong to it.
    expect(refusalCode(shape('a\uf001', { script: 'Latn' }))).toBe('unsupported-script')
  })

  // Character tracking: DOCX w:spacing on w:rPr (ECMA-376 17.3.2.35).
  //
  // Word 16's own export of StyleRef-DE.docx settles where the delta goes. The
  // Subtitle style there carries w:spacing w:val="15" (15 twips = 0.75 pt) and
  // Word writes that one run with `0.0651 Tc` at Tm scale 46 under a 0.24 CTM,
  // against a small negative justification Tc on every other run in the file.
  // The run begins at x=300 and the next text object begins at x=1540.55, so
  // Word's own advance for its 57 glyphs is 1240.55 text-space units against
  // 1070.01 unspaced. A delta after EVERY glyph predicts 1240.70, out by 0.15,
  // which is the file's own /Widths quantisation (the untracked control run in
  // the same file is out by 0.03 over 10 glyphs). A delta only BETWEEN glyphs
  // predicts 1237.70, out by 2.85 - a whole tracking step. Word uses the first.
  it('places character tracking at every cluster trailing edge, including the last', () => {
    const text = 'Handgloves'
    const untracked = shape(text)
    if ('status' in untracked) throw new Error('baseline run must shape')
    const delta = 750 // 15 twips, the measured StyleRef-DE Subtitle tracking
    const tracked = shape(text, { letterSpacingMilliPoints: delta })
    if ('status' in tracked) throw new Error('tracked run must shape')

    // n, not n-1: the run grows by one delta per cluster, the last one included.
    expect(tracked.clusters).toHaveLength(untracked.clusters.length)
    expect(tracked.advanceInlineMilliPoints).toBe(untracked.advanceInlineMilliPoints + delta * untracked.clusters.length)
    expect(tracked.advanceInlineMilliPoints).not.toBe(untracked.advanceInlineMilliPoints + delta * (untracked.clusters.length - 1))

    // The delta sits between clusters, never inside one: every cluster keeps its
    // own shaped advance plus exactly one delta, and no glyph is redistributed.
    for (const [index, cluster] of tracked.clusters.entries()) {
      expect(cluster.advanceInlineMilliPoints).toBe(untracked.clusters[index]!.advanceInlineMilliPoints + delta)
      expect(cluster.startUtf16).toBe(untracked.clusters[index]!.startUtf16)
      expect(cluster.endUtf16).toBe(untracked.clusters[index]!.endUtf16)
    }
    expect(tracked.glyphs.map((glyph) => glyph.glyphId)).toEqual(untracked.glyphs.map((glyph) => glyph.glyphId))
    expect(tracked.glyphs.map((glyph) => [glyph.offsetXMilliPoints, glyph.offsetYMilliPoints]))
      .toEqual(untracked.glyphs.map((glyph) => [glyph.offsetXMilliPoints, glyph.offsetYMilliPoints]))
    expect(tracked.metrics).toEqual(untracked.metrics)

    // Word's measure is signed: a negative w:spacing condenses by the same rule.
    const condensed = shape(text, { letterSpacingMilliPoints: -delta })
    if ('status' in condensed) throw new Error('condensed run must shape')
    expect(condensed.advanceInlineMilliPoints).toBe(untracked.advanceInlineMilliPoints - delta * untracked.clusters.length)
  })

  it('leaves an absent or zero character tracking exactly as inert as it was', () => {
    const text = 'Handgloves'
    const untracked = shape(text)
    if ('status' in untracked) throw new Error('baseline run must shape')
    for (const zero of [{ letterSpacingMilliPoints: 0 }, {}] as const) {
      const result = shape(text, zero)
      if ('status' in result) throw new Error('inert run must shape')
      expect(result).toEqual(untracked)
    }
  })

  it('refuses the tracking it cannot place at a cluster edge exactly', () => {
    // A zero-advance cluster takes no delta, so an implicit directional mark
    // keeps the zero advance the Unicode core specification requires of it.
    const marked = shape('a\u200eb', { letterSpacingMilliPoints: 750 })
    if ('status' in marked) throw new Error('implicit-mark run must shape')
    const mark = marked.clusters.find((cluster) => cluster.startUtf16 === 1)
    expect(mark?.advanceInlineMilliPoints).toBe(0)

    // Condensing a cluster past a zero advance has no cluster edge to place.
    expect(refusalCode(shape('Handgloves', { letterSpacingMilliPoints: -1_000_000 }))).toBe('unsupported-feature')
    // Word spacing is still not modeled and still refuses, tracking or not.
    expect(refusalCode(shape('Handgloves', { letterSpacingMilliPoints: 750, wordSpacingMilliPoints: 1 }))).toBe('unsupported-feature')
  })

  it('copies and hashes a face once per cached face, not once per shaped run', () => {
    // `Uint8Array.from` copies through the source's iterator, so a counting
    // subclass records exactly how many times the provider takes an owned copy
    // of the caller's font bytes — the same copies it hashes.
    let copies = 0
    class CountedBytes extends Uint8Array {}
    Object.defineProperty(CountedBytes.prototype, Symbol.iterator, {
      value(this: Uint8Array) {
        copies += 1
        return Uint8Array.prototype[Symbol.iterator].call(this)
      },
      writable: true,
      configurable: true,
    })
    const bytes = new CountedBytes(FONT_BYTES.byteLength)
    bytes.set(FONT_BYTES)
    const resource: FontResource = { ...font, bytes }
    const shaper = createHarfBuzzTextShaperV1({ sourceRevision: 'git:test-suite' })
    const shapeOnce = (text: string) => shaper.shape({ run: run(text), startUtf16: 0, endUtf16: text.length, font: resource })

    for (const text of ['alpha', 'beta', 'gamma', 'delta']) expect(refusalCode(shapeOnce(text))).toBeUndefined()
    expect(copies).toBe(1)

    // The time-of-check/time-of-use guard still holds against a cached face.
    bytes[bytes.byteLength - 1] = bytes[bytes.byteLength - 1]! ^ 0xff
    expect(refusalCode(shapeOnce('epsilon'))).toBe('font-digest-mismatch')
    expect(copies).toBe(1)
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
    // A table record whose length no longer covers its own header: the cmap
    // becomes unwalkable, which is the integrity the preflight actually depends on.
    const truncatedTable = mutatedResource((bytes) => {
      writeU32(bytes, tableRecord(bytes, 'cmap') + 12, 8)
    })
    // A four-byte-aligned table offset that leaves the font bytes entirely.
    const offsetOutsideBytes = mutatedResource((bytes) => {
      writeU32(bytes, tableRecord(bytes, 'maxp') + 8, bytes.byteLength - (bytes.byteLength % 4))
    })
    // `hhea.numberOfHMetrics` claiming more metrics than `hmtx` can hold.
    const inconsistentHmtx = mutatedResource((bytes) => {
      writeU32(bytes, tableRecord(bytes, 'hmtx') + 12, 8)
    })
    for (const resource of [badMagic, overlapping, descendingLoca, truncatedTable, offsetOutsideBytes, inconsistentHmtx]) expect(refusalCode(shape('abc', {}, resource))).toBe('unsupported-font-format')
  })

  it('admits a face whose sfnt directory states a stale table checksum', () => {
    // sfnt table checksums are advisory build metadata, not a parse
    // precondition. Word 16 ships `symbol.ttf` with a `cmap` record stating
    // 0x42f6990a over bytes that sum to 0x2796fb28, and that face is the only
    // one on a stock macOS install that maps the Symbol bullet U+F0B7, so
    // refusing it substituted nothing - it removed the document. The bytes are
    // still pinned: `contentDigest` is a sha-256 over the whole file, verified
    // before and after the preflight, which is strictly stronger than the sum.
    const staleChecksum = mutatedResource((bytes) => {
      const record = tableRecord(bytes, 'cmap')
      writeU32(bytes, record + 4, (readU32(bytes, record + 4) ^ 0xdead_beef) >>> 0)
    })
    expect(inspectHarfBuzzFontMetricsV1({ bytes: staleChecksum.bytes, contentDigest: staleChecksum.face.contentDigest })).toEqual(font.metrics)
    const stale = shape('office bullets', {}, staleChecksum)
    const pristine = shape('office bullets')
    expect('status' in stale).toBe(false)
    expect('status' in pristine).toBe(false)
    // The table bytes are untouched, so the face shapes exactly as it does when
    // the advisory number agrees.
    expect((stale as ShapedSegment).glyphs).toEqual((pristine as ShapedSegment).glyphs)
    expect((stale as ShapedSegment).clusters).toEqual((pristine as ShapedSegment).clusters)
    expect((stale as ShapedSegment).advanceInlineMilliPoints).toBe((pristine as ShapedSegment).advanceInlineMilliPoints)
  })

  it('accepts an empty optional prep table, including when it shares loca\'s offset', () => {
    const emptyPrep = mutatedResource((bytes) => {
      const prep = tableRecord(bytes, 'prep')
      const loca = tableRecord(bytes, 'loca')
      writeU32(bytes, prep + 4, 0)
      writeU32(bytes, prep + 8, readU32(bytes, loca + 8))
      writeU32(bytes, prep + 12, 0)
    })
    expect(inspectHarfBuzzFontMetricsV1({ bytes: emptyPrep.bytes, contentDigest: emptyPrep.face.contentDigest })).toEqual(font.metrics)
    const result = shape('office', {}, emptyPrep)
    expect('status' in result).toBe(false)
    if ('status' in result) return
    expect(result.glyphs.map((glyph) => glyph.glyphId)).toEqual([82, 5_044, 70, 72])
  })

  it('still refuses an empty required head table', () => {
    const emptyHead = mutatedResource((bytes) => {
      const head = tableRecord(bytes, 'head')
      writeU32(bytes, head + 4, 0)
      writeU32(bytes, head + 12, 0)
    })
    expect(refusalCode(shape('office', {}, emptyHead))).toBe('unsupported-font-format')
    expect(() => inspectHarfBuzzFontMetricsV1({ bytes: emptyHead.bytes, contentDigest: emptyHead.face.contentDigest })).toThrow(/empty or outside/)
  })

  it('admits a static face that declares STAT without fvar, and shapes it identically', () => {
    // Every Aptos face Word 16 ships (Aptos.ttf, Aptos-Narrow.ttf, Aptos-Light.ttf, ...) is a
    // static glyf/loca instance that carries STAT and no fvar. STAT names where a face sits in
    // its family's design space; it binds no outline, advance or line metric to an axis, so the
    // face is as fully determined as one without it.
    const withStat = resourceWithTables([['STAT', STAT_AXIS_VALUE_ONLY]])
    expect(inspectHarfBuzzFontMetricsV1({ bytes: withStat.bytes, contentDigest: withStat.face.contentDigest })).toEqual(font.metrics)
    const result = shape('office', {}, withStat)
    expect('status' in result).toBe(false)
    if ('status' in result) return
    expect(result.glyphs.map((glyph) => glyph.glyphId)).toEqual([82, 5_044, 70, 72])
    expect(result.advanceInlineMilliPoints).toBe((shape('office') as ShapedSegment).advanceInlineMilliPoints)
  })

  it('still refuses a genuinely variable face, whose metrics need an axis position the contract never carries', () => {
    for (const variationTable of ['avar', 'cvar', 'fvar', 'gvar', 'HVAR', 'MVAR', 'VVAR']) {
      const variable = resourceWithTables([[variationTable, new Uint8Array(32)]])
      expect(refusalCode(shape('office', {}, variable))).toBe('unsupported-font-format')
      expect(() => inspectHarfBuzzFontMetricsV1({ bytes: variable.bytes, contentDigest: variable.face.contentDigest })).toThrow(/variable-font table .* is outside the fixed-font v1 contract/)
    }
  })

  it('refuses a face that declares STAT alongside fvar, so STAT is never a bypass', () => {
    const variable = resourceWithTables([['STAT', STAT_AXIS_VALUE_ONLY], ['fvar', new Uint8Array(32)]])
    expect(refusalCode(shape('office', {}, variable))).toBe('unsupported-font-format')
    expect(() => inspectHarfBuzzFontMetricsV1({ bytes: variable.bytes, contentDigest: variable.face.contentDigest })).toThrow(/fvar/)
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

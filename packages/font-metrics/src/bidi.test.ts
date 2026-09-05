import { describe, expect, it } from 'vitest'
import { createHash } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { gunzipSync } from 'node:zlib'
import {
  BIDI_JS_FACTORY_SHA256,
  BIDI_JS_ENTRY_SHA256,
  BIDI_JS_MANIFEST_SHA256,
  BIDI_JS_PACKAGE_VERSION,
  BIDI_UNICODE_VERSION,
  NATIVE_BIDI_LIMITS,
  NATIVE_BIDI_PROVIDER_REVISION,
  reorderNativeBidiLineV1,
  resolveNativeBidiParagraphV1,
  type NativeBidiExplicitRangeV1,
} from './bidi.js'

function resolved(text: string, baseDirection: 'ltr' | 'rtl' = 'ltr') {
  const result = resolveNativeBidiParagraphV1({ text, baseDirection })
  expect(result.ok, result.ok ? '' : result.message).toBe(true)
  if (!result.ok) throw new Error(result.message)
  return result.value
}

describe('pinned native Unicode bidi boundary', () => {
  it('pins implementation, Unicode data, and deterministic provider identity', () => {
    expect(BIDI_JS_PACKAGE_VERSION).toBe('1.0.3')
    expect(BIDI_UNICODE_VERSION).toBe('13.0.0')
    expect(BIDI_JS_FACTORY_SHA256).toBe('sha256:ce1928a26521e7eca2dde603a76f2332f7a1ab85977a4a10651af3f31f952a26')
    expect(BIDI_JS_ENTRY_SHA256).toBe('sha256:5b58433ed951be70376bee55cb9a98cdce788e00cd2fa5f881cbcb1dcad03102')
    expect(BIDI_JS_MANIFEST_SHA256).toBe('sha256:c2579f7705ab96dcc6192ab5a317d87760d2be7c67d72976d85edb4181bec5de')
    expect(NATIVE_BIDI_PROVIDER_REVISION).toMatch(/^sha256:[a-f0-9]{64}$/)
    expect(JSON.stringify(resolved('abc אבג 123'))).toBe(JSON.stringify(resolved('abc אבג 123')))
  })

  it('resolves mixed Hebrew, European numerals, spaces, and neutrals then emits inverse maps', () => {
    const paragraph = resolved('abc אבג 123')
    expect(paragraph.levels).toEqual([0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2])
    const order = reorderNativeBidiLineV1(paragraph.levels, paragraph.baseLevel, paragraph.levels.map((_, index) => index === 3 || index === 7))
    expect(order).toEqual({ ok: true, value: {
      visualToLogical: [0, 1, 2, 3, 8, 9, 10, 7, 6, 5, 4],
      logicalToVisual: [0, 1, 2, 3, 10, 9, 8, 7, 4, 5, 6],
    } })
  })

  it('keeps Arabic/Hebrew numerals LTR inside an RTL paragraph and resets trailing spaces per line', () => {
    const paragraph = resolved('אבג (123) xyz ', 'rtl')
    const order = reorderNativeBidiLineV1(paragraph.levels, 1, [...paragraph.levels].map((_, index) => index === paragraph.levels.length - 1))
    expect(order.ok).toBe(true)
    if (!order.ok) return
    expect(order.value.visualToLogical.slice(0, 4)).toEqual([13, 10, 11, 12])
    expect(order.value.visualToLogical.slice(6, 9)).toEqual([5, 6, 7])
    expect(order.value.visualToLogical.at(-1)).toBe(0)
  })

  it('preserves combining-mark levels and represents explicit run direction as an isolate', () => {
    expect(resolved('a\u0301 אב').levels).toEqual([0, 0, 0, 1, 1])
    const explicit = resolveNativeBidiParagraphV1({ text: 'abc', baseDirection: 'rtl', explicitRanges: [{ startUtf16: 0, endUtf16: 3, direction: 'ltr' }] })
    expect(explicit).toEqual(expect.objectContaining({ ok: true, value: expect.objectContaining({ baseLevel: 1, levels: [2, 2, 2] }) }))
  })

  it('fails closed on malformed UTF-16, authored controls, overlapping ranges, and hostile bounds', () => {
    expect(resolveNativeBidiParagraphV1({ text: '\ud800', baseDirection: 'ltr' })).toMatchObject({ ok: false, code: 'invalid-input' })
    expect(resolveNativeBidiParagraphV1({ text: 'a\u2067b', baseDirection: 'ltr' })).toMatchObject({ ok: false, code: 'unsupported-control' })
    expect(resolveNativeBidiParagraphV1({ text: 'abcd', baseDirection: 'ltr', explicitRanges: [{ startUtf16: 0, endUtf16: 3, direction: 'ltr' }, { startUtf16: 2, endUtf16: 4, direction: 'rtl' }] })).toMatchObject({ ok: false, code: 'invalid-input' })
    expect(resolveNativeBidiParagraphV1({ text: 'a'.repeat(NATIVE_BIDI_LIMITS.maxUtf16 + 1), baseDirection: 'ltr' })).toMatchObject({ ok: false, code: 'resource-limit' })
    expect(reorderNativeBidiLineV1(new Array(NATIVE_BIDI_LIMITS.maxClustersPerLine + 1).fill(0), 0, new Array(NATIVE_BIDI_LIMITS.maxClustersPerLine + 1).fill(false))).toMatchObject({ ok: false, code: 'invalid-input' })

    const exactRanges = Array.from({ length: NATIVE_BIDI_LIMITS.maxExplicitRanges }, (_unused, index) => ({ startUtf16: index * 2, endUtf16: index * 2 + 1, direction: index % 2 === 0 ? 'ltr' as const : 'rtl' as const }))
    expect(resolveNativeBidiParagraphV1({ text: 'a '.repeat(NATIVE_BIDI_LIMITS.maxExplicitRanges), baseDirection: 'ltr', explicitRanges: exactRanges })).toMatchObject({ ok: true })
    expect(resolveNativeBidiParagraphV1({ text: 'a '.repeat(NATIVE_BIDI_LIMITS.maxExplicitRanges + 1), baseDirection: 'ltr', explicitRanges: [...exactRanges, { startUtf16: NATIVE_BIDI_LIMITS.maxExplicitRanges * 2, endUtf16: NATIVE_BIDI_LIMITS.maxExplicitRanges * 2 + 1, direction: 'ltr' }] })).toMatchObject({ ok: false, code: 'resource-limit' })
  })

  it('gates the assigned Unicode 13 repertoire and snapshots hostile accessors once', () => {
    expect(resolveNativeBidiParagraphV1({ text: '\u0860', baseDirection: 'rtl' })).toMatchObject({ ok: true })
    expect(resolveNativeBidiParagraphV1({ text: '\u0870', baseDirection: 'rtl' })).toMatchObject({ ok: false, code: 'unsupported-scalar' })
    let textReads = 0
    const changing = {
      get text() { textReads += 1; return textReads === 1 ? 'abc' : 'x'.repeat(NATIVE_BIDI_LIMITS.maxUtf16 + 1) },
      baseDirection: 'ltr' as const,
    }
    expect(resolveNativeBidiParagraphV1(changing)).toMatchObject({ ok: true, value: { textLengthUtf16: 3 } })
    expect(textReads).toBe(1)
    const hostileRanges = new Proxy([] as NativeBidiExplicitRangeV1[], { get(target, property, receiver) { if (property === '0') throw new Error('hostile range'); return Reflect.get(target, property, receiver) } })
    hostileRanges.length = 1
    expect(() => resolveNativeBidiParagraphV1({ text: 'a', baseDirection: 'ltr', explicitRanges: hostileRanges })).not.toThrow()
    expect(resolveNativeBidiParagraphV1({ text: 'a', baseDirection: 'ltr', explicitRanges: hostileRanges })).toMatchObject({ ok: false, code: 'invalid-input' })
    const hostileLevels = new Proxy([0], { get(target, property, receiver) { if (property === '0') throw new Error('hostile level'); return Reflect.get(target, property, receiver) } })
    expect(() => reorderNativeBidiLineV1(hostileLevels, 0, [false])).not.toThrow()
    expect(reorderNativeBidiLineV1(hostileLevels, 0, [false])).toMatchObject({ ok: false, code: 'invalid-input' })
  })

  it('atomically refuses astral Adlam and emoji before the UTF-16 bidi-js engine', () => {
    expect(resolveNativeBidiParagraphV1({ text: 'א\u{1E900}ב', baseDirection: 'rtl' })).toMatchObject({ ok: false, code: 'unsupported-scalar', message: expect.stringContaining('BMP') })
    expect(resolveNativeBidiParagraphV1({ text: 'א😀ב', baseDirection: 'rtl' })).toMatchObject({ ok: false, code: 'unsupported-scalar', message: expect.stringContaining('atomically') })
  })

  it('rejects malformed cluster-level line inputs without partial maps', () => {
    expect(reorderNativeBidiLineV1([0, 126], 0, [false, false])).toEqual(expect.objectContaining({ ok: false, code: 'invalid-input' }))
    expect(reorderNativeBidiLineV1([0, 1], 0, [false])).toEqual(expect.objectContaining({ ok: false, code: 'invalid-input' }))
  })

  it('conforms to every admitted BMP-only explicit-base Unicode 13 BidiCharacterTest vector', { timeout: 20_000 }, () => {
    const compressed = readFileSync(new URL('../../../scripts/unicode13/bidi/BidiCharacterTest-13.0.0.txt.gz', import.meta.url))
    expect(createHash('sha256').update(compressed).digest('hex')).toBe('cbb9d664ea46fcf22dc42bee73cda0f2ffef733e527b7ea375c3d93a50affc5a')
    const raw = gunzipSync(compressed)
    expect(createHash('sha256').update(raw).digest('hex')).toBe('b0939c352a162034a58e89c792cdcfa5bb4db5d9fe506e36201d5a78744cf714')
    let admitted = 0
    for (const sourceLine of raw.toString('utf8').split(/\r?\n/)) {
      const line = sourceLine.replace(/#.*/, '').replace(/^[\u0009\u0020]+|[\u0009\u0020]+$/g, '')
      if (line.length === 0) continue
      const fields = line.split(';').map((field) => field.replace(/^[\u0009\u0020]+|[\u0009\u0020]+$/g, ''))
      if (fields.length !== 5 || (fields[1] !== '0' && fields[1] !== '1')) continue
      const codePoints = fields[0]!.split(' ').map((value) => Number.parseInt(value, 16))
      if (codePoints.some((value) => value > 0xffff || (value >= 0xd800 && value <= 0xdfff))) continue
      const text = String.fromCodePoint(...codePoints)
      if (/[\u061c\u200e\u200f\u202a-\u202e\u2066-\u2069]/u.test(text)) continue
      const expectedLevels = fields[3]!.split(' ').map(Number)
      const expectedOrder = fields[4] === '' ? [] : fields[4]!.split(' ').map(Number)
      if (expectedLevels.some((value) => !Number.isSafeInteger(value)) || expectedOrder.length !== text.length) continue
      const result = resolveNativeBidiParagraphV1({ text, baseDirection: fields[1] === '0' ? 'ltr' : 'rtl' })
      expect(result.ok, `Unicode 13 BidiCharacterTest vector failed: ${fields[0]}`).toBe(true)
      if (!result.ok) continue
      expect(result.value.levels, `Unicode 13 levels drifted: ${fields[0]}`).toEqual(expectedLevels)
      const trailing = codePoints.map((codePoint) => (codePoint >= 0x09 && codePoint <= 0x0d) || (codePoint >= 0x1c && codePoint <= 0x1e) || codePoint === 0x20 || codePoint === 0x85 || codePoint === 0x2028 || codePoint === 0x2029)
      const reordered = reorderNativeBidiLineV1(result.value.levels, result.value.baseLevel, trailing)
      expect(reordered.ok, `Unicode 13 reorder failed: ${fields[0]}`).toBe(true)
      if (reordered.ok) expect(reordered.value.visualToLogical, `Unicode 13 order drifted: ${fields[0]}`).toEqual(expectedOrder)
      admitted += 1
    }
    expect(admitted).toBeGreaterThan(10_000)
  })
})

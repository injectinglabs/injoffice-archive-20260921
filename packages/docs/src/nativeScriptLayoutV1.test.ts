import { readFileSync } from 'node:fs'
import { createHash } from 'node:crypto'
import { createRequire } from 'node:module'
import { describe, expect, it } from 'vitest'
import type { FontResource } from '@injoffice/font-metrics/layout'
import { readNativeDocxScriptTransformV1, validateNativeDocxScriptTransformV1, nativeDocxScriptTransformEligibleRunV1, nativeDocxScriptScaleV1, nativeDocxScriptShiftV1 } from './nativeScriptLayoutV1.js'
const require = createRequire(import.meta.url)
const bytes = new Uint8Array(readFileSync(require.resolve('dejavu-fonts-ttf/ttf/DejaVuSans.ttf')))
const resource = { bytes, metrics:{unitsPerEm:2048}, face:{contentDigest:`sha256:${createHash('sha256').update(bytes).digest('hex')}`} } as FontResource
describe('native OS/2 script metrics', () => {
  it('reads distinct horizontal and vertical font-authored sizes without a guessed ratio', () => {
    for (const kind of ['subscript','superscript'] as const) {
      const transform = readNativeDocxScriptTransformV1(resource,kind)
      expect(validateNativeDocxScriptTransformV1(transform)).toBe(true)
      expect(transform.x_size).not.toBe(transform.y_size)
      for (const patch of [{x_size:0},{y_size:2048},{y_offset:-1},{x_offset:2049},{units_per_em:0},{font_sha256:'unknown'},{unknown:1}]) expect(validateNativeDocxScriptTransformV1({...transform,...patch})).toBe(false)
    }
  })
  it('admits a script transform only on runs the shaped-text path produces', () => {
    for (const run of [{kind:'text'},{kind:'reference',reference:{kind:'footnote'}},{kind:'reference',reference:{kind:'endnote'}}]) expect(nativeDocxScriptTransformEligibleRunV1(run)).toBe(true)
    for (const run of [{kind:'control'},{kind:'drawing'},{kind:'reference'},{kind:'reference',reference:{kind:'comment'}},{kind:'reference',reference:{kind:'comment-range-start'}},{kind:'reference',reference:{kind:'comment-range-end'}}]) expect(nativeDocxScriptTransformEligibleRunV1(run)).toBe(false)
  })
  it('raises a superscript by the size the script reduction took away, not by the font-authored offset', () => {
    // Word 16.112 reference exports: the superscript baseline sits exactly one
    // "shrink" above the run baseline, so the reduced em box is flush with the
    // top of the run's own em box. DejaVuSans reduces to 1433/2048, so an 11pt
    // run raises 11pt - 11pt*1433/2048, and never by ySuperscriptYOffset (983).
    const transform = readNativeDocxScriptTransformV1(resource,'superscript')
    expect([transform.y_size,transform.y_offset]).toEqual([1433,983])
    for (const fontSize of [11_000,24_000,7_500]) {
      const reduced = nativeDocxScriptScaleV1(fontSize,transform,'y')
      expect(nativeDocxScriptShiftV1(fontSize,transform,'y')).toBe(fontSize-reduced)
      expect(nativeDocxScriptShiftV1(fontSize,transform,'y')).not.toBe(Math.round(fontSize*transform.y_offset/transform.units_per_em))
    }
    expect(nativeDocxScriptShiftV1(11_000,transform,'y')).toBe(3_303)
  })
  it('leaves the font-authored script size, the horizontal offset and the subscript drop alone', () => {
    const superscript = readNativeDocxScriptTransformV1(resource,'superscript'), subscript = readNativeDocxScriptTransformV1(resource,'subscript')
    // Size stays the OS/2-derived reduction on both axes and for both kinds.
    for (const transform of [superscript,subscript]) for (const fontSize of [11_000,24_000]) {
      expect(nativeDocxScriptScaleV1(fontSize,transform,'y')).toBe(Math.round(fontSize*transform.y_size/transform.units_per_em))
      expect(nativeDocxScriptScaleV1(fontSize,transform,'x')).toBe(Math.round(fontSize*transform.x_size/transform.units_per_em))
    }
    // The subscript still drops by the font-authored ySubscriptYOffset, negated.
    expect(nativeDocxScriptShiftV1(11_000,subscript,'y')).toBe(-Math.round(11_000*subscript.y_offset/subscript.units_per_em))
    expect(nativeDocxScriptShiftV1(11_000,subscript,'y')).toBe(-1_536)
    // The horizontal shift stays font-authored and unsigned for both kinds.
    for (const transform of [superscript,subscript]) expect(nativeDocxScriptShiftV1(11_000,transform,'x')).toBe(Math.round(11_000*transform.x_offset/transform.units_per_em))
  })
  it('refuses truncated font directories, missing metrics and mismatched resolved units', () => {
    for (const truncated of [new Uint8Array(),bytes.slice(0,20),bytes.slice(0,1000)]) expect(()=>readNativeDocxScriptTransformV1({...resource,bytes:truncated},'superscript')).toThrow()
    expect(()=>readNativeDocxScriptTransformV1({...resource,metrics:{...resource.metrics,unitsPerEm:1000}},'subscript')).toThrow(/resolved font units/)
  })
})

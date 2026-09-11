import { readFileSync } from 'node:fs'
import { createHash } from 'node:crypto'
import { createRequire } from 'node:module'
import { describe, expect, it } from 'vitest'
import type { FontResource } from '@injoffice/font-metrics/layout'
import { readNativeDocxScriptTransformV1, validateNativeDocxScriptTransformV1 } from './nativeScriptLayoutV1.js'
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
  it('refuses truncated font directories, missing metrics and mismatched resolved units', () => {
    for (const truncated of [new Uint8Array(),bytes.slice(0,20),bytes.slice(0,1000)]) expect(()=>readNativeDocxScriptTransformV1({...resource,bytes:truncated},'superscript')).toThrow()
    expect(()=>readNativeDocxScriptTransformV1({...resource,metrics:{...resource.metrics,unitsPerEm:1000}},'subscript')).toThrow(/resolved font units/)
  })
})

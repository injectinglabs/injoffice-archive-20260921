import { fileURLToPath } from 'node:url'
import { build } from 'vite'
import { describe, expect, it } from 'vitest'
import { decodeNativeDocxPagePaintV1 } from './nativePagePaintOutputV1.js'
import { decodeNativeDocxPagePaintV1 as compilerDecode } from './nativePagePaintV1.js'

describe('browser-safe page paint output entry', () => {
  it('shares the identical strict decoder with the server compiler', () => {
    expect(decodeNativeDocxPagePaintV1).toBe(compilerDecode)
    expect(decodeNativeDocxPagePaintV1({ protocol: 'injoffice.docx.page-paint', version: 1, pages: [] }).ok).toBe(false)
  })
  it('bundles for browsers without Node or native provider dependencies', async () => {
    const forbidden: string[] = []
    await build({ configFile: false, logLevel: 'silent', plugins: [{ name: 'reject-native-provider', enforce: 'pre', resolveId(id) {
      if (id.startsWith('node:') || /font-metrics\/(bidi|harfbuzz)(?:$|\.)/.test(id)) { forbidden.push(id); throw new Error(`Browser output decoder imported ${id}`) }
      return null
    } }], build: { write: false, lib: { entry: { source: fileURLToPath(new URL('./nativePagePaintOutputV1.ts', import.meta.url)), published: fileURLToPath(new URL('../dist/nativePagePaintOutputV1.js', import.meta.url)) }, formats: ['es'] } } })
    expect(forbidden).toEqual([])
  })
})

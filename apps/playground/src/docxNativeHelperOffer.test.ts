import { createHash } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { digestNativeDocxPackage, nativeDocxHelperOffer } from './docxNativeHelperOffer'

const digest = (bytes: Uint8Array) => `sha256:${createHash('sha256').update(bytes).digest('hex')}`

describe('native DOCX helper offer', () => {
  it('hashes opened package bytes with the native sha256: prefix', async () => {
    const bytes = Uint8Array.from([1, 2, 3])
    await expect(digestNativeDocxPackage(bytes)).resolves.toBe(digest(bytes))
  })

  it('offers helper controls from opened bytes after extract refusal', () => {
    const bytes = Uint8Array.from([0x50, 0x4b])
    const packageDigest = digest(bytes)
    expect(nativeDocxHelperOffer({ apiBase: '', bytes, packageDigest })).toBeNull()
    expect(nativeDocxHelperOffer({ apiBase: 'https://helper.test/', bytes: null, packageDigest })).toBeNull()
    expect(nativeDocxHelperOffer({ apiBase: 'https://helper.test/', bytes, packageDigest: 'sha256:nope' })).toBeNull()
    expect(nativeDocxHelperOffer({ apiBase: 'https://helper.test/', bytes, packageDigest })).toEqual({
      bytes, packageDigest, apiBase: 'https://helper.test',
    })
  })

  it('does not gate Docs playground helper controls on a successful extract', () => {
    const source = readFileSync(new URL('./pages/DocsPage.tsx', import.meta.url), 'utf8')
    expect(source).toContain('nativeDocxHelperOffer')
    expect(source).toContain('digestNativeDocxPackage')
    expect(source).not.toMatch(/authoritativeBytes && document && <NativeDocxPages/)
  })
})

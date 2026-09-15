import { createHash } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { digestNativeDocxPackage, nativeDocxHelperOffer, nativeDocxHelperPackageDigest } from './docxNativeHelperOffer'

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
    expect(nativeDocxHelperOffer({ apiBase: 'https://helper.test/', bytes, packageDigest: nativeDocxHelperPackageDigest(undefined, packageDigest) })).toEqual({
      bytes, packageDigest, apiBase: 'https://helper.test',
    })
  })

  it('pairs replaced bytes with the current document digest, not the originally opened digest', () => {
    const original = Uint8Array.from([1])
    const mutated = Uint8Array.from([2])
    const originalDigest = digest(original)
    const mutatedDigest = digest(mutated)
    expect(nativeDocxHelperPackageDigest(undefined, originalDigest)).toBe(originalDigest)
    expect(nativeDocxHelperPackageDigest(mutatedDigest, originalDigest)).toBe(mutatedDigest)
    expect(nativeDocxHelperOffer({ apiBase: 'https://helper.test', bytes: original, packageDigest: nativeDocxHelperPackageDigest(undefined, originalDigest) })).toEqual({
      bytes: original, packageDigest: originalDigest, apiBase: 'https://helper.test',
    })
    expect(nativeDocxHelperOffer({ apiBase: 'https://helper.test', bytes: mutated, packageDigest: nativeDocxHelperPackageDigest(mutatedDigest, originalDigest) })).toEqual({
      bytes: mutated, packageDigest: mutatedDigest, apiBase: 'https://helper.test',
    })
  })

  it('does not gate Docs playground helper controls on a successful extract', () => {
    const source = readFileSync(new URL('./pages/DocsPage.tsx', import.meta.url), 'utf8')
    expect(source).toContain('nativeDocxHelperOffer')
    expect(source).toContain('nativeDocxHelperPackageDigest')
    expect(source).toContain('digestNativeDocxPackage')
    expect(source).toContain('setOpenedDigest(next.source.package_sha256)')
    expect(source).not.toMatch(/authoritativeBytes && document && <NativeDocxPages/)
  })
})

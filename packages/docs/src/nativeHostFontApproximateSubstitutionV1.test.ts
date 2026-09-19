import {describe, expect, it} from 'vitest'
import {
  discloseNativeDocxHostFontApproximateSubstitutionReasonsV1,
  DOCX_HOST_FONT_APPROXIMATE_SUBSTITUTION_WARNING,
  nativeDocxHostFontApproximateSubstitutionReasonV1,
  nativeDocxHostFontApproximateSubstitutionV1,
  selectLoadedHostManifestSubstituteV1,
} from './nativeHostFontApproximateSubstitutionV1.js'
import type {NativeFontManifest} from '@injoffice/font-metrics/layout'

const digest = `sha256:${'a'.repeat(64)}` as const
function face(family: string, faceId = `host:${family}`, kind: 'host' | 'document' = 'host'): NativeFontManifest['faces'][number] {
  return {faceId, family, weight: 400, style: 'normal', stretch: 100, source: {kind, resourceId: faceId, contentDigest: digest}}
}
function weighted(family: string, weight: number, style: 'normal' | 'italic' = 'normal'): NativeFontManifest['faces'][number] {
  const faceId = `host:${family}:${weight}:${style}`
  return {faceId, family, weight, style, stretch: 100, source: {kind: 'host', resourceId: faceId, contentDigest: digest}}
}

describe('approximate loaded-host font substitution', () => {
  it('selects a host face already in the manifest and loaded, never Calibri', () => {
    const dejavu = face('DejaVu Sans')
    const loaded = new Set([dejavu.faceId])
    const selected = selectLoadedHostManifestSubstituteV1([dejavu], loaded, {family: 'Candara', weight: 400, style: 'normal'})
    expect(selected).toEqual(dejavu)
    const record = nativeDocxHostFontApproximateSubstitutionV1({family: 'Candara', weight: 400, style: 'normal'}, selected!)
    expect(record).toMatchObject({source_family: 'Candara', selected_family: 'DejaVu Sans', reason: 'loaded-host-manifest-face-v1'})
    expect(record.selected_family).not.toBe('Calibri')
    const reasons = discloseNativeDocxHostFontApproximateSubstitutionReasonsV1([record])
    expect(reasons[0]).toBe(DOCX_HOST_FONT_APPROXIMATE_SUBSTITUTION_WARNING)
    expect(reasons[1]).toContain('Candara / 400 / normal -> DejaVu Sans')
  })

  it('does not invent a family that is not loaded', () => {
    const calibri = face('Calibri')
    const documentFace = face('DejaVu Sans', 'doc:dejavu', 'document')
    expect(selectLoadedHostManifestSubstituteV1([calibri], new Set(), {family: 'Candara', weight: 400, style: 'normal'})).toBeUndefined()
    expect(selectLoadedHostManifestSubstituteV1([documentFace], new Set([documentFace.faceId]), {family: 'Candara', weight: 400, style: 'normal'})).toBeUndefined()
    expect(selectLoadedHostManifestSubstituteV1([calibri], new Set([calibri.faceId]), {family: 'Candara', weight: 700, style: 'normal'})).toBeUndefined()
    expect(selectLoadedHostManifestSubstituteV1([calibri], new Set([calibri.faceId]), {family: 'Candara', weight: 400, style: 'italic'})).toBeUndefined()
    expect(() => nativeDocxHostFontApproximateSubstitutionV1({family: 'Candara', weight: 400, style: 'normal'}, documentFace)).toThrow(/loaded host face/)
    expect(discloseNativeDocxHostFontApproximateSubstitutionReasonsV1([])).toEqual([])
  })

  it('resolves one source family to one target family across weights', () => {
    // Manifest order is exactly the corpus shape that produced the split:
    // the first loaded 400 face is Calibri, the first loaded 700 face is Aptos.
    const calibri400 = weighted('Calibri', 400)
    const aptos700 = weighted('Aptos', 700)
    const calibri700 = weighted('Calibri', 700)
    const faces = [calibri400, aptos700, calibri700]
    const loaded = new Set(faces.map(f => f.faceId))
    const regular = selectLoadedHostManifestSubstituteV1(faces, loaded, {family: 'Segoe UI', weight: 400, style: 'normal'})
    expect(regular?.family).toBe('Calibri')
    const bold = selectLoadedHostManifestSubstituteV1(faces, loaded, {family: 'Segoe UI', weight: 700, style: 'normal'}, regular!.family)
    expect(bold?.family).toBe('Calibri')
    expect(bold?.faceId).toBe(calibri700.faceId)
    // Without the preferred family the naive first-match picks a different typeface.
    expect(selectLoadedHostManifestSubstituteV1(faces, loaded, {family: 'Segoe UI', weight: 700, style: 'normal'})?.family).toBe('Aptos')
    const records = [
      nativeDocxHostFontApproximateSubstitutionV1({family: 'Segoe UI', weight: 400, style: 'normal'}, regular!),
      nativeDocxHostFontApproximateSubstitutionV1({family: 'Segoe UI', weight: 700, style: 'normal'}, bold!),
    ]
    expect(new Set(records.map(r => r.selected_family))).toEqual(new Set(['Calibri']))
    for (const reason of discloseNativeDocxHostFontApproximateSubstitutionReasonsV1(records).slice(1)) {
      expect(reason).not.toContain('family-inconsistent')
    }
  })

  it('never silently splits a family when the preferred family lacks the weight', () => {
    const calibri400 = weighted('Calibri', 400)
    const aptos700 = weighted('Aptos', 700)
    const faces = [calibri400, aptos700]
    const loaded = new Set(faces.map(f => f.faceId))
    const bold = selectLoadedHostManifestSubstituteV1(faces, loaded, {family: 'Segoe UI', weight: 700, style: 'normal'}, 'Calibri')
    expect(bold?.family).toBe('Aptos')
    const record = nativeDocxHostFontApproximateSubstitutionV1({family: 'Segoe UI', weight: 700, style: 'normal'}, bold!, {familyFallbackFrom: 'Calibri'})
    expect(record.family_fallback_from).toBe('Calibri')
    const reason = nativeDocxHostFontApproximateSubstitutionReasonV1(record)
    expect(reason).toContain('Segoe UI / 700 / normal -> Aptos [family-inconsistent: Calibri has no 700/normal face]')
    // The disclosure stays visible to a reader that stops at the first parenthesis.
    expect(reason.slice(0, reason.indexOf('('))).toContain('family-inconsistent')
    // The preserved invariants still hold for the fallback path.
    expect(() => nativeDocxHostFontApproximateSubstitutionV1({family: 'Segoe UI', weight: 700, style: 'normal'}, bold!, {familyFallbackFrom: 'Aptos'})).toThrow(/family fallback/)
  })

  it('folds casing variants of one authored family to the same target', () => {
    const calibri400 = weighted('Calibri', 400)
    const aptos700 = weighted('Aptos', 700)
    const calibri700 = weighted('Calibri', 700)
    const faces = [calibri400, aptos700, calibri700]
    const loaded = new Set(faces.map(f => f.faceId))
    expect(selectLoadedHostManifestSubstituteV1(faces, loaded, {family: 'SEGOE UI', weight: 700, style: 'normal'}, 'calibri')?.family).toBe('Calibri')
  })

  it('keeps operator manifest order instead of Word default families', () => {
    const dejavu = face('DejaVu Sans')
    const calibri = face('Calibri')
    expect(selectLoadedHostManifestSubstituteV1([dejavu, calibri], new Set([dejavu.faceId, calibri.faceId]), {family: 'Trebuchet MS', weight: 400, style: 'normal'})?.family).toBe('DejaVu Sans')
  })
})

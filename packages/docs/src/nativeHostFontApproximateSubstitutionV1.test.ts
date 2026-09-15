import {describe, expect, it} from 'vitest'
import {
  discloseNativeDocxHostFontApproximateSubstitutionReasonsV1,
  DOCX_HOST_FONT_APPROXIMATE_SUBSTITUTION_WARNING,
  nativeDocxHostFontApproximateSubstitutionV1,
  selectLoadedHostManifestSubstituteV1,
} from './nativeHostFontApproximateSubstitutionV1.js'
import type {NativeFontManifest} from '@injoffice/font-metrics/layout'

const digest = `sha256:${'a'.repeat(64)}` as const
function face(family: string, faceId = `host:${family}`, kind: 'host' | 'document' = 'host'): NativeFontManifest['faces'][number] {
  return {faceId, family, weight: 400, style: 'normal', stretch: 100, source: {kind, resourceId: faceId, contentDigest: digest}}
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

  it('keeps operator manifest order instead of Word default families', () => {
    const dejavu = face('DejaVu Sans')
    const calibri = face('Calibri')
    expect(selectLoadedHostManifestSubstituteV1([dejavu, calibri], new Set([dejavu.faceId, calibri.faceId]), {family: 'Trebuchet MS', weight: 400, style: 'normal'})?.family).toBe('DejaVu Sans')
  })
})

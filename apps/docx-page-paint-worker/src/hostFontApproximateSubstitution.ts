/** Worker copy of packages/docs/src/nativeHostFontApproximateSubstitutionV1.ts.
 * The compiled worker cannot import docs source; keep reason/selector in lockstep. */
import {createHash} from 'node:crypto'
import type {NativeFontManifest} from '@injoffice/font-metrics/layout'

export const DOCX_HOST_FONT_APPROXIMATE_SUBSTITUTION_REASON = 'loaded-host-manifest-face-v1' as const
export const DOCX_HOST_FONT_APPROXIMATE_SUBSTITUTION_WARNING = 'Approximate page preview substituted missing authored fonts with a host face already loaded from the operator font manifest. This is not Microsoft Word font matching; metrics and layout may differ. Source font names and original file bytes are unchanged.'

export interface NativeDocxHostFontApproximateSubstitutionV1 {
  source_family: string
  selected_family: string
  face_id: string
  font_digest: string
  weight: number
  style: 'normal' | 'italic'
  reason: typeof DOCX_HOST_FONT_APPROXIMATE_SUBSTITUTION_REASON
}

const fold = (value: string) => value.replace(/[A-Z]/g, character => character.toLowerCase())
const evidence = (value: unknown) => `sha256:${createHash('sha256').update(JSON.stringify(value)).digest('hex')}`

/** First already-loaded host face of the same weight/style in manifest order. Never invents a family. */
export function selectLoadedHostManifestSubstituteV1(
  faces: NativeFontManifest['faces'],
  loadedFaceIds: ReadonlySet<string>,
  reference: {family: string; weight: number; style: string},
): NativeFontManifest['faces'][number] | undefined {
  return faces.find(face =>
    face.source.kind === 'host'
    && loadedFaceIds.has(face.faceId)
    && !!face.source.contentDigest
    && face.weight === reference.weight
    && face.style === reference.style
    && face.stretch === 100
    && fold(face.family) !== fold(reference.family)
  )
}

export function nativeDocxHostFontApproximateSubstitutionV1(
  reference: {family: string; weight: number; style: 'normal' | 'italic'},
  face: NativeFontManifest['faces'][number],
): NativeDocxHostFontApproximateSubstitutionV1 {
  if (face.source.kind !== 'host' || !face.source.contentDigest) throw new TypeError('Approximate host substitution requires a loaded host face')
  if (face.weight !== reference.weight || face.style !== reference.style || face.stretch !== 100) throw new TypeError('Approximate host substitution requires matching weight and style')
  if (fold(face.family) === fold(reference.family)) throw new TypeError('Approximate host substitution cannot target the authored family')
  return {
    source_family: reference.family,
    selected_family: face.family,
    face_id: face.faceId,
    font_digest: face.source.contentDigest,
    weight: reference.weight,
    style: reference.style,
    reason: DOCX_HOST_FONT_APPROXIMATE_SUBSTITUTION_REASON,
  }
}

export function nativeDocxHostFontApproximateSubstitutionReasonV1(record: NativeDocxHostFontApproximateSubstitutionV1): string {
  return `Approximate host substitution: ${record.source_family} / ${record.weight} / ${record.style} -> ${record.selected_family} (${record.reason}, face ${record.face_id}); evidence ${evidence(record)}. Font metrics and layout may differ.`
}

export function discloseNativeDocxHostFontApproximateSubstitutionReasonsV1(records: readonly NativeDocxHostFontApproximateSubstitutionV1[]): string[] {
  if (!records.length) return []
  return [DOCX_HOST_FONT_APPROXIMATE_SUBSTITUTION_WARNING, ...records.map(nativeDocxHostFontApproximateSubstitutionReasonV1)]
}

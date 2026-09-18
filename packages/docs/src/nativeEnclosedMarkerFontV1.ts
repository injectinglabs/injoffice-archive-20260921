import { decodeNativeDocxDocument, type NativeDocxDocumentV1, type NativeDocxParagraphV1 } from './nativeContract.js'
import { decodeNativeDocxResolvedLayout, nativeDocxResolvedNumberingModelSha256V1, type NativeDocxResolvedLayoutInputV1 } from './nativeResolvedLayout.js'

/** The declared host family this preview paints an enclosed-number list marker
 * in when the marker's own font slot resolves a face authored for Latin text.
 *
 * The choice is a preview-host choice and is NOT validated against Microsoft
 * Word. It is bounded by three facts that were measured, not assumed:
 *   * repertoire: Cambria covers the whole `decimalEnclosedCircle` repertoire
 *     U+2460..U+2473 (so do Calibri, Calibri Light, Carlito and Candara; Aptos
 *     and Aptos Narrow carry U+2460 but stop before U+2473, and Arial, Times
 *     New Roman, Liberation Serif/Sans/Mono, Courier New, Verdana, Trebuchet
 *     MS, Symbol and Wingdings have none of it).
 *   * fit: Cambria's U+2460 advances 1.2007 em against Calibri's and Carlito's
 *     1.3281 em, so a `(1).` marker fits a 360-twip hanging label region at
 *     12 pt (16.87 pt) where the Calibri-metric faces overflow it (18.97 pt)
 *     and the numbered paragraph is refused atomically at its numbering tab.
 *   * design: Cambria is a serif face, which is what the corpus instance's own
 *     High ANSI slot names (Liberation Serif).
 */
export const DOCX_ENCLOSED_MARKER_FONT_HOST_DEFAULT = 'Cambria' as const
export const DOCX_ENCLOSED_MARKER_FONT_WARNING =
  `Approximate read-only preview: a list marker whose generated text uses Enclosed Alphanumerics (U+2460..U+24FF) is painted in the declared ${DOCX_ENCLOSED_MARKER_FONT_HOST_DEFAULT} host family instead of the face its High ANSI font slot resolves, because this tier reads no cmap and cannot attest that face's repertoire. The substituted family is a preview-host choice, is not the authored face and is not validated against Microsoft Word; its advances and glyph design are not Word's.` as const

/** Extractor evidence: this numbered paragraph's marker text uses the Enclosed
 * Alphanumerics block while its font slot resolves the paragraph's Latin face.
 * The source records the scope and the family it resolved, never a substitute. */
export interface NativeDocxEnclosedMarkerFontV1 {
  scope_kind: 'numbering-marker'
  scope_id: string
  part_name: string
  path: string
  source_family: string
  package_sha256: string
}

export interface NativeDocxApproximatedEnclosedMarkerFontV1 extends NativeDocxEnclosedMarkerFontV1 {
  chosen_family: typeof DOCX_ENCLOSED_MARKER_FONT_HOST_DEFAULT
}

const FACT_KEYS = 'package_sha256,part_name,path,scope_id,scope_kind,source_family'

export function validNativeDocxEnclosedMarkerFontsV1(value: unknown, packageSHA256: string): value is NativeDocxEnclosedMarkerFontV1[] {
  if (!Array.isArray(value) || value.length > 1000) return false
  const ids = new Set<string>()
  for (const fact of value) {
    if (!fact || typeof fact !== 'object' || Object.keys(fact).sort().join(',') !== FACT_KEYS
      || fact.scope_kind !== 'numbering-marker'
      || typeof fact.scope_id !== 'string' || !/^[A-Za-z0-9][A-Za-z0-9._:-]{0,255}$/.test(fact.scope_id)
      || fact.package_sha256 !== packageSHA256 || !/^sha256:[0-9a-f]{64}$/.test(fact.package_sha256)
      || typeof fact.part_name !== 'string' || fact.part_name.length < 1 || fact.part_name.length > 1024 || fact.part_name.startsWith('/') || /[\\\u0000-\u001f\u007f]/.test(fact.part_name) || fact.part_name.split('/').some((s: string) => !s || s === '.' || s === '..')
      || typeof fact.path !== 'string' || fact.path.length > 4096 || !fact.path.startsWith('/w:') || /[\u0000-\u001f\u007f]/.test(fact.path)
      || typeof fact.source_family !== 'string' || fact.source_family.length < 1 || fact.source_family.length > 256 || /[\u0000-\u001f\u007f]/.test(fact.source_family) || fact.source_family.trim() !== fact.source_family) return false
    if (ids.has(fact.scope_id)) return false
    ids.add(fact.scope_id)
  }
  return true
}

/** Applied substitutions must each be one of the retained source facts carrying
 * the declared host family; the preview may apply a subset when the loaded
 * manifest does not attest every weight/style the markers need. */
export function validNativeDocxApproximatedEnclosedMarkerFontsV1(value: unknown, source: readonly NativeDocxEnclosedMarkerFontV1[], packageSHA256: string): value is NativeDocxApproximatedEnclosedMarkerFontV1[] {
  if (!Array.isArray(value) || value.length === 0 || value.length > source.length
    || value.some((f) => !f || typeof f !== 'object' || f.chosen_family !== DOCX_ENCLOSED_MARKER_FONT_HOST_DEFAULT || Object.keys(f).sort().join(',') !== `chosen_family,${FACT_KEYS}`)) return false
  const facts = (value as NativeDocxApproximatedEnclosedMarkerFontV1[]).map(({ chosen_family: _chosen, ...fact }) => fact)
  if (!validNativeDocxEnclosedMarkerFontsV1(facts, packageSHA256)) return false
  const retained = new Set(source.map((fact) => JSON.stringify(Object.entries(fact).sort())))
  return facts.every((fact) => retained.has(JSON.stringify(Object.entries(fact).sort())))
}

function storyParagraphs(document: NativeDocxDocumentV1): NativeDocxParagraphV1[] {
  const stories = [document.body, ...document.headers, ...document.footers, ...document.notes, ...document.comment_stories]
  return stories.flatMap((story) => story.blocks.flatMap((block) => block.paragraph ? [block.paragraph] : (block.table?.rows ?? []).flatMap((row) => row.cells.flatMap((cell) => cell.paragraphs))))
}

/** Whether the host font manifest attests a face for this family/weight/style. */
export type AttestedFaceFn = (family: string, weight: 400 | 700, style: 'normal' | 'italic') => boolean

function target(resolved: NativeDocxResolvedLayoutInputV1, fact: NativeDocxEnclosedMarkerFontV1) {
  return resolved.paragraphs.filter((p) => p.paragraph_id === fact.scope_id).map((p) => p.numbering?.marker_properties)
}

/** @internal Repoints an enclosed-number marker at the declared host family, and
 * only where the marker still carries exactly the family the source evidence
 * recorded and the loaded host manifest attests the declared family at that
 * marker's own weight and style. A scope whose substitute is not attested keeps
 * the authored face and stays exactly as it was; no run, paragraph mark or
 * marker outside the evidence is touched. */
export function projectNativeDocxEnclosedMarkerFontsV1(documentValue: unknown, resolvedValue: unknown, facts: readonly NativeDocxEnclosedMarkerFontV1[], attested: AttestedFaceFn = () => false): { resolved: NativeDocxResolvedLayoutInputV1; applied: NativeDocxApproximatedEnclosedMarkerFontV1[] } {
  const document = decodeNativeDocxDocument(documentValue), layout = decodeNativeDocxResolvedLayout(resolvedValue)
  if (!document.ok || !layout.ok || document.value.document_id !== layout.value.document_id || document.value.revision !== layout.value.revision || !validNativeDocxEnclosedMarkerFontsV1(facts, document.value.source.package_sha256)) throw new TypeError('Enclosed-marker font evidence does not exact-join')
  const paragraphs = storyParagraphs(document.value)
  const resolved = structuredClone(layout.value)
  if (facts.length === 0) return { resolved, applied: [] }
  const applied: NativeDocxApproximatedEnclosedMarkerFontV1[] = []
  for (const fact of facts) {
    const candidates = paragraphs.filter((p) => p.id === fact.scope_id)
    if (candidates.length !== 1 || candidates[0]!.anchor.part_name !== fact.part_name || candidates[0]!.anchor.path !== fact.path) throw new TypeError('Enclosed-marker scope anchor does not exact-join')
    const targets = target(resolved, fact)
    if (targets.length !== 1 || !targets[0] || targets[0].font_family !== fact.source_family) throw new TypeError('Enclosed-marker evidence does not join the resolved marker family')
    const weight = targets[0].bold === true ? 700 : 400, style = targets[0].italic === true ? 'italic' : 'normal'
    if (!attested(DOCX_ENCLOSED_MARKER_FONT_HOST_DEFAULT, weight, style)) continue
    targets[0].font_family = DOCX_ENCLOSED_MARKER_FONT_HOST_DEFAULT
    applied.push({ ...fact, chosen_family: DOCX_ENCLOSED_MARKER_FONT_HOST_DEFAULT })
  }
  // A repointed marker family is part of the resolved marker model, so the source
  // attestation is recomputed over the projected markers; the strict view below
  // restores the authored family and the original digest.
  if (resolved.numbering_source && applied.length) resolved.numbering_source.model_sha256 = nativeDocxResolvedNumberingModelSha256V1(resolved.paragraphs, resolved.numbering_source)
  return { resolved, applied }
}

/** @internal The strict view of a projected layout: the authored marker families
 * restored so the strict font inventory references still exact-join their scopes. */
export function stripNativeDocxEnclosedMarkerFontsV1(resolved: NativeDocxResolvedLayoutInputV1, facts: readonly NativeDocxEnclosedMarkerFontV1[]): NativeDocxResolvedLayoutInputV1 {
  if (facts.length === 0) return resolved
  const stripped = structuredClone(resolved)
  let restored = false
  for (const fact of facts) for (const properties of target(stripped, fact)) if (properties?.font_family === DOCX_ENCLOSED_MARKER_FONT_HOST_DEFAULT) { properties.font_family = fact.source_family; restored = true }
  if (stripped.numbering_source && restored) stripped.numbering_source.model_sha256 = nativeDocxResolvedNumberingModelSha256V1(stripped.paragraphs, stripped.numbering_source)
  return stripped
}

import { decodeNativeDocxDocument, type NativeDocxDocumentV1, type NativeDocxParagraphV1 } from './nativeContract.js'
import { decodeNativeDocxResolvedLayout, type NativeDocxResolvedLayoutInputV1 } from './nativeResolvedLayout.js'

/** The explicitly selected host family for scopes a package leaves font-less.
 * The value is measured, not assumed: ink-cluster extents of Microsoft Word
 * 16.112.4's own rendering of packages that state no font anywhere match Aptos
 * at every measured string, and no other candidate family does. It remains a
 * preview-host choice, not an authored face and not a documented Word default. */
export const DOCX_ABSENT_FONT_FAMILY_HOST_DEFAULT = 'Aptos' as const
export const DOCX_ABSENT_FONT_FAMILY_VALUE_SOURCE = 'measured against Microsoft Word 16.112.4 references' as const
export const DOCX_ABSENT_FONT_FAMILY_WARNING =
  `Approximate read-only preview: this package selects no font anywhere, so source-absent font families use the explicitly selected ${DOCX_ABSENT_FONT_FAMILY_HOST_DEFAULT} host default (${DOCX_ABSENT_FONT_FAMILY_VALUE_SOURCE}); this is not an authored face and not a documented Microsoft Word default.` as const

/** Extractor evidence: the package carries no w:rFonts at all, so this scope has
 * no family to read. The source records the omission and never a face. */
export interface NativeDocxAbsentFontFamilyV1 {
  scope_kind: 'run' | 'paragraph-mark'
  scope_id: string
  part_name: string
  path: string
  package_sha256: string
}

export interface NativeDocxHostDefaultFamilyPolicyV1 {
  kind: 'host-default-family-v1'
  family: typeof DOCX_ABSENT_FONT_FAMILY_HOST_DEFAULT
}

export interface NativeDocxApproximatedFontFamilyV1 extends NativeDocxAbsentFontFamilyV1 {
  chosen_family: typeof DOCX_ABSENT_FONT_FAMILY_HOST_DEFAULT
}

const FACT_KEYS = 'package_sha256,part_name,path,scope_id,scope_kind'

export function validNativeDocxHostDefaultFamilyPolicyV1(value: unknown): value is NativeDocxHostDefaultFamilyPolicyV1 {
  if (!value || typeof value !== 'object') return false
  const p = value as Record<string, unknown>
  return Object.keys(p).sort().join(',') === 'family,kind' && p.kind === 'host-default-family-v1' && p.family === DOCX_ABSENT_FONT_FAMILY_HOST_DEFAULT
}

export function validNativeDocxAbsentFontFamiliesV1(value: unknown, packageSHA256: string): value is NativeDocxAbsentFontFamilyV1[] {
  if (!Array.isArray(value) || value.length > 1000) return false
  const ids = new Set<string>()
  for (const fact of value) {
    if (!fact || typeof fact !== 'object' || Object.keys(fact).sort().join(',') !== FACT_KEYS
      || !['run', 'paragraph-mark'].includes(fact.scope_kind)
      || typeof fact.scope_id !== 'string' || !/^[A-Za-z0-9][A-Za-z0-9._:-]{0,255}$/.test(fact.scope_id)
      || fact.package_sha256 !== packageSHA256 || !/^sha256:[0-9a-f]{64}$/.test(fact.package_sha256)
      || typeof fact.part_name !== 'string' || fact.part_name.length < 1 || fact.part_name.length > 1024 || fact.part_name.startsWith('/') || /[\\\u0000-\u001f\u007f]/.test(fact.part_name) || fact.part_name.split('/').some((s: string) => !s || s === '.' || s === '..')
      || typeof fact.path !== 'string' || fact.path.length > 4096 || !fact.path.startsWith('/w:') || /[\u0000-\u001f\u007f]/.test(fact.path)) return false
    const id = `${fact.scope_kind}:${fact.scope_id}`
    if (ids.has(id)) return false
    ids.add(id)
  }
  return true
}

/** Applied families must each be one of the retained source facts carrying the
 * declared host family; the preview may apply a subset when the manifest does
 * not attest every weight/style the document needs. */
export function validNativeDocxApproximatedFontFamiliesV1(value: unknown, source: readonly NativeDocxAbsentFontFamilyV1[], packageSHA256: string): value is NativeDocxApproximatedFontFamilyV1[] {
  if (!Array.isArray(value) || value.length === 0 || value.length > source.length
    || value.some((f) => !f || typeof f !== 'object' || f.chosen_family !== DOCX_ABSENT_FONT_FAMILY_HOST_DEFAULT || Object.keys(f).sort().join(',') !== `chosen_family,${FACT_KEYS}`)) return false
  const facts = (value as NativeDocxApproximatedFontFamilyV1[]).map(({ chosen_family: _chosen, ...fact }) => fact)
  if (!validNativeDocxAbsentFontFamiliesV1(facts, packageSHA256)) return false
  const retained = new Set(source.map((fact) => JSON.stringify(Object.entries(fact).sort())))
  return facts.every((fact) => retained.has(JSON.stringify(Object.entries(fact).sort())))
}

function storyParagraphs(document: NativeDocxDocumentV1): NativeDocxParagraphV1[] {
  const stories = [document.body, ...document.headers, ...document.footers, ...document.notes, ...document.comment_stories]
  return stories.flatMap((story) => story.blocks.flatMap((block) => block.paragraph ? [block.paragraph] : (block.table?.rows ?? []).flatMap((row) => row.cells.flatMap((cell) => cell.paragraphs))))
}

type ReferenceFn = (resolved: NativeDocxResolvedLayoutInputV1) => ReadonlyArray<{ family: string; weight: number; style: string }>
/** Whether the host font manifest attests a face for this family/weight/style. */
export type AttestedFaceFn = (family: string, weight: 400 | 700, style: 'normal' | 'italic') => boolean

function target(resolved: NativeDocxResolvedLayoutInputV1, fact: NativeDocxAbsentFontFamilyV1) {
  return fact.scope_kind === 'paragraph-mark'
    ? resolved.paragraphs.filter((p) => p.paragraph_id === fact.scope_id).map((p) => p.paragraph_mark_properties)
    : resolved.runs.filter((r) => r.run_id === fact.scope_id).map((r) => r.properties)
}

/** @internal Applies the declared host family only to proven source omissions,
 * and only where the loaded host manifest attests that family at the scope's own
 * weight and style. A scope whose face is not attested keeps no family and stays
 * unshaped exactly as before; no resolved family is ever overridden. */
export function projectNativeDocxAbsentFontFamiliesV1(documentValue: unknown, resolvedValue: unknown, facts: readonly NativeDocxAbsentFontFamilyV1[], policy: unknown, references: ReferenceFn, attested: AttestedFaceFn = () => false): { resolved: NativeDocxResolvedLayoutInputV1; applied: NativeDocxApproximatedFontFamilyV1[] } {
  if (!validNativeDocxHostDefaultFamilyPolicyV1(policy)) throw new TypeError(`Missing-family approximation requires the explicit ${DOCX_ABSENT_FONT_FAMILY_HOST_DEFAULT} host policy`)
  const document = decodeNativeDocxDocument(documentValue), layout = decodeNativeDocxResolvedLayout(resolvedValue)
  if (!document.ok || !layout.ok || document.value.document_id !== layout.value.document_id || document.value.revision !== layout.value.revision || !validNativeDocxAbsentFontFamiliesV1(facts, document.value.source.package_sha256)) throw new TypeError('Missing-family source evidence does not exact-join')
  const paragraphs = storyParagraphs(document.value)
  const resolved = structuredClone(layout.value)
  if (facts.length === 0) return { resolved, applied: [] }
  // The evidence is whole-package: it exists only for a source that selects no
  // font anywhere, so the strict layout must carry no font reference at all.
  if (references(resolved).length !== 0) throw new TypeError('Missing-family evidence requires a package that resolves no font reference at all')
  const applied: NativeDocxApproximatedFontFamilyV1[] = []
  for (const fact of facts) {
    const candidates = fact.scope_kind === 'paragraph-mark' ? paragraphs.filter((p) => p.id === fact.scope_id) : paragraphs.flatMap((p) => p.runs.filter((r) => r.id === fact.scope_id))
    if (candidates.length !== 1 || candidates[0]!.anchor.part_name !== fact.part_name || candidates[0]!.anchor.path !== fact.path) throw new TypeError('Missing-family scope anchor does not exact-join')
    const targets = target(resolved, fact)
    if (targets.length !== 1 || !targets[0] || targets[0].font_family !== undefined) throw new TypeError('Host family policy cannot override an authored/resolved font family')
    const weight = targets[0].bold === true ? 700 : 400, style = targets[0].italic === true ? 'italic' : 'normal'
    if (!attested(policy.family, weight, style)) continue
    targets[0].font_family = policy.family
    applied.push({ ...fact, chosen_family: policy.family })
  }
  return { resolved, applied }
}

/** @internal The strict view of a projected layout: projected families removed so
 * the strict font inventory references still exact-join their original scopes. */
export function stripNativeDocxAbsentFontFamiliesV1(resolved: NativeDocxResolvedLayoutInputV1, facts: readonly NativeDocxAbsentFontFamilyV1[]): NativeDocxResolvedLayoutInputV1 {
  if (facts.length === 0) return resolved
  const stripped = structuredClone(resolved)
  for (const fact of facts) for (const properties of target(stripped, fact)) if (properties.font_family === DOCX_ABSENT_FONT_FAMILY_HOST_DEFAULT) delete properties.font_family
  return stripped
}

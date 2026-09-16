import { decodeNativeDocxDocument, type NativeDocxDocumentV1, type NativeDocxParagraphV1 } from './nativeContract.js'
import { decodeNativeDocxResolvedLayout, type NativeDocxResolvedLayoutInputV1 } from './nativeResolvedLayout.js'

/** Declared approximate policy for runs whose strict font selection stays unresolved
 * only because of an empty East-Asian or complex-script font slot. */
export const DOCX_LATIN_FONT_FALLBACK_WARNING = 'Approximate read-only preview: runs whose font selection strict resolution leaves unresolved because of an empty East-Asian or complex-script font slot use the authored ascii/hAnsi face; Word may select the complex-script font for right-to-left or complex-script text.' as const

/** Extractor evidence: strict resolution has no face for this scope; the authored Latin face is recorded, never applied, by the source. */
export interface NativeDocxLatinFontFallbackV1 {
  scope_kind: 'run' | 'paragraph-mark'
  scope_id: string
  part_name: string
  path: string
  font_family: string
  package_sha256: string
}

const FACT_KEYS = 'font_family,package_sha256,part_name,path,scope_id,scope_kind'

export function validNativeDocxLatinFontFallbacksV1(value: unknown, packageSHA256: string): value is NativeDocxLatinFontFallbackV1[] {
  if (!Array.isArray(value) || value.length > 1000) return false
  const ids = new Set<string>()
  for (const fact of value) {
    if (!fact || typeof fact !== 'object' || Object.keys(fact).sort().join(',') !== FACT_KEYS
      || !['run', 'paragraph-mark'].includes(fact.scope_kind)
      || typeof fact.scope_id !== 'string' || !/^[A-Za-z0-9][A-Za-z0-9._:-]{0,255}$/.test(fact.scope_id)
      || fact.package_sha256 !== packageSHA256 || !/^sha256:[0-9a-f]{64}$/.test(fact.package_sha256)
      || typeof fact.part_name !== 'string' || fact.part_name.length < 1 || fact.part_name.length > 1024 || fact.part_name.startsWith('/') || /[\\\u0000-\u001f\u007f]/.test(fact.part_name) || fact.part_name.split('/').some((s: string) => !s || s === '.' || s === '..')
      || typeof fact.path !== 'string' || fact.path.length > 4096 || !fact.path.startsWith('/w:') || /[\u0000-\u001f\u007f]/.test(fact.path)
      || typeof fact.font_family !== 'string' || fact.font_family.length < 1 || fact.font_family.length > 256 || /[\u0000-\u001f\u007f]/.test(fact.font_family) || fact.font_family.trim() !== fact.font_family) return false
    const id = `${fact.scope_kind}:${fact.scope_id}`
    if (ids.has(id)) return false
    ids.add(id)
  }
  return true
}

/** Applied faces must each be one of the retained source facts; the preview may apply a subset. */
export function validNativeDocxApproximatedFontFacesV1(value: unknown, source: readonly NativeDocxLatinFontFallbackV1[], packageSHA256: string): value is NativeDocxLatinFontFallbackV1[] {
  if (!validNativeDocxLatinFontFallbacksV1(value, packageSHA256) || value.length === 0) return false
  const retained = new Set(source.map((fact) => JSON.stringify(Object.entries(fact).sort())))
  return value.every((fact) => retained.has(JSON.stringify(Object.entries(fact).sort())))
}

function storyParagraphs(document: NativeDocxDocumentV1): NativeDocxParagraphV1[] {
  const stories = [document.body, ...document.headers, ...document.footers, ...document.notes, ...document.comment_stories]
  return stories.flatMap((story) => story.blocks.flatMap((block) => block.paragraph ? [block.paragraph] : (block.table?.rows ?? []).flatMap((row) => row.cells.flatMap((cell) => cell.paragraphs))))
}

type ReferenceFn = (resolved: NativeDocxResolvedLayoutInputV1) => ReadonlyArray<{ family: string; weight: number; style: string }>
/** Whether the host font manifest attests a face for this family/weight/style. */
export type AttestedFaceFn = (family: string, weight: 400 | 700, style: 'normal' | 'italic') => boolean

function referenceKeys(references: ReturnType<ReferenceFn>): string {
  return JSON.stringify([...new Set(references.map((entry) => `${entry.family.toLowerCase()}\u0000${entry.weight}\u0000${entry.style}`))].sort())
}

function target(resolved: NativeDocxResolvedLayoutInputV1, fact: NativeDocxLatinFontFallbackV1) {
  return fact.scope_kind === 'paragraph-mark'
    ? resolved.paragraphs.filter((p) => p.paragraph_id === fact.scope_id).map((p) => p.paragraph_mark_properties)
    : resolved.runs.filter((r) => r.run_id === fact.scope_id).map((r) => r.properties)
}

/** @internal Applies authored Latin faces only where strict resolution has none
 * and only when the face/weight/style is either already a reference of the strict
 * inventory or attested by the explicitly supplied host manifest; other facts are
 * skipped and their scopes stay unshaped exactly as before. */
export function projectNativeDocxLatinFontFallbacksV1(documentValue: unknown, resolvedValue: unknown, facts: readonly NativeDocxLatinFontFallbackV1[], references: ReferenceFn, attested: AttestedFaceFn = () => false): { resolved: NativeDocxResolvedLayoutInputV1; applied: NativeDocxLatinFontFallbackV1[] } {
  const document = decodeNativeDocxDocument(documentValue), layout = decodeNativeDocxResolvedLayout(resolvedValue)
  if (!document.ok || !layout.ok || document.value.document_id !== layout.value.document_id || document.value.revision !== layout.value.revision || !validNativeDocxLatinFontFallbacksV1(facts, document.value.source.package_sha256)) throw new TypeError('Latin font fallback evidence does not exact-join')
  const paragraphs = storyParagraphs(document.value)
  const resolved = structuredClone(layout.value)
  const baseline = referenceKeys(references(resolved))
  const applied: NativeDocxLatinFontFallbackV1[] = []
  for (const fact of facts) {
    const candidates = fact.scope_kind === 'paragraph-mark' ? paragraphs.filter((p) => p.id === fact.scope_id) : paragraphs.flatMap((p) => p.runs.filter((r) => r.id === fact.scope_id))
    if (candidates.length !== 1 || candidates[0]!.anchor.part_name !== fact.part_name || candidates[0]!.anchor.path !== fact.path) throw new TypeError('Latin font fallback scope anchor does not exact-join')
    const targets = target(resolved, fact)
    if (targets.length !== 1 || !targets[0] || targets[0].font_family !== undefined) throw new TypeError('Latin font fallback cannot override a resolved font family')
    targets[0].font_family = fact.font_family
    const weight = targets[0].bold === true ? 700 : 400, style = targets[0].italic === true ? 'italic' : 'normal'
    if (referenceKeys(references(resolved)) === baseline || attested(fact.font_family, weight, style)) applied.push({ ...fact })
    else delete targets[0].font_family
  }
  return { resolved, applied }
}

/** @internal The strict view of a projected layout: projected faces removed so the
 * strict font inventory references still exact-join their original scopes. */
export function stripNativeDocxLatinFontFallbacksV1(resolved: NativeDocxResolvedLayoutInputV1, facts: readonly NativeDocxLatinFontFallbackV1[]): NativeDocxResolvedLayoutInputV1 {
  if (facts.length === 0) return resolved
  const stripped = structuredClone(resolved)
  for (const fact of facts) for (const properties of target(stripped, fact)) if (properties.font_family === fact.font_family) delete properties.font_family
  return stripped
}

import type { NativeDocxDocumentV1, NativeDocxStoryV1 } from './nativeContract.js'
import { decodeNativeDocxPaginationRequestV1, type NativeDocxPaginationRequestV1, type NativeDocxPaginatedLayoutV1 } from './nativePaginationV1.js'
import type { NativeDocxShapedLinesV1 } from './nativeShapingLines.js'
import { nativeDocxPageNumberV1 } from './nativePageNumbersV1.js'
import {qualifyNativeDocxFontSubstitutionsV1,isQualifiedNativeDocxFontDiagnosticV1,type NativeDocxFontSubstitutionV1} from './nativeFontSubstitutionEvidenceV1.js'
import type {NativeFontManifest} from '@injoffice/font-metrics/layout'

export interface NativeDocxFontVariantPolicyV1 {manifest:NativeFontManifest;policy:unknown;descriptors?:{eligibility:unknown;inventoryJSON:string}}
/** Internal read-only route; every variant is independently source-qualified. */
export function validateNativeDocxFontPageFieldVariantsV1(request:NativeDocxPaginationRequestV1,layout:NativeDocxPaginatedLayoutV1,input:unknown,font:NativeDocxFontVariantPolicyV1){
 return validatePageFieldVariants(request,layout,input,font)
}

export interface NativeDocxPageFieldVariantV1 { page_id: string; shaped_lines: NativeDocxShapedLinesV1 }
export const DOCX_PAGE_FIELD_LIMITS = { maxPages: 64, maxFragments: 100_000 } as const

function runs(story: NativeDocxStoryV1) {
  return story.blocks.flatMap((block) => block.paragraph?.runs ?? block.table?.rows.flatMap((row) => row.cells.flatMap((cell) => cell.paragraphs.flatMap((p) => p.runs))) ?? [])
}

/** Header/footer expansion; body fields retain markers and are separately source-replayed. */
export function hasNativeDocxPageFieldsV1(document: NativeDocxDocumentV1): boolean {
  for (const story of [...document.notes, ...document.comment_stories]) if (runs(story).some((run) => run.page_field)) throw new TypeError('PAGE/NUMPAGES fields in notes and comments remain refused')
  for (const run of runs(document.body)) if (run.page_field && !/^[1-9][0-9]{0,5}$/.test(run.text ?? '')) throw new TypeError('Body fields require layout-derived decimal text before header expansion')
  let found = false
  for (const story of [...document.headers, ...document.footers]) for (const run of runs(story)) if (run.page_field) {
    found = true
    if (run.kind !== 'text' || run.text !== '') throw new TypeError('Source page-field text must be empty; cached results are not authoritative')
    if (!run.anchor.path.includes('/w:fldSimple[') && !/^.*\/w:p\[\d+\]\/w:r\[\d+\]\/w:t\[\d+\]$/.test(run.anchor.path)) throw new TypeError('Page-field source must be anchored in a simple field or a paragraph-local flat field result text')
    // Modeled header/footer parts are not passthrough parts. Their root anchor
    // hashes the raw source XML containing both instruction and cached result;
    // document.source.package_sha256 binds the complete OPC bytes as well.
    if (!story.anchor.xml_sha256 || story.anchor.part_name !== story.part_name) throw new TypeError('Page-field source story requires its digest-bound raw XML root')
    const paragraph = story.blocks.find((block) => block.paragraph?.runs.some((candidate) => candidate.id === run.id))?.paragraph
    if (!paragraph || paragraph.edit_policy.mode !== 'read-only') throw new TypeError('Page-field paragraphs must remain read-only and cannot be table content')
  }
  return found
}

/** Only substitution permitted by this profile; original document is never mutated. */
export function nativeDocxPageFieldDocumentV1(document: NativeDocxDocumentV1, ordinal: number, count: number, displayNumber?: number): NativeDocxDocumentV1 {
  hasNativeDocxPageFieldsV1(document)
  if (displayNumber === undefined && document.sections.some(section => section.page_number_start !== undefined)) throw new TypeError('Section restarts require a final pagination-derived display number')
  if (displayNumber !== undefined && (!Number.isInteger(displayNumber) || displayNumber < 0 || displayNumber > 999999)) throw new RangeError('Display number exceeds bounded decimal profile')
  if (!Number.isInteger(count) || count < 1 || count > DOCX_PAGE_FIELD_LIMITS.maxPages || !Number.isInteger(ordinal) || ordinal < 0 || ordinal >= count) throw new RangeError('Page-field expansion exceeds its bounded page profile')
  const derived = structuredClone(document)
  for (const story of [...derived.headers, ...derived.footers]) for (const run of runs(story)) if (run.page_field) {
    run.text = String(run.page_field === 'PAGE' ? displayNumber ?? ordinal + 1 : count)
    delete run.page_field
  }
  return derived
}

/** Reuse full existing shaping/source validation after deterministic field substitution. */
export function validateNativeDocxPageFieldVariantsV1(request: NativeDocxPaginationRequestV1, layout: NativeDocxPaginatedLayoutV1, input: unknown): NativeDocxPageFieldVariantV1[] | undefined {
  return validatePageFieldVariants(request,layout,input)
}
function validatePageFieldVariants(request: NativeDocxPaginationRequestV1, layout: NativeDocxPaginatedLayoutV1, input: unknown,font?:NativeDocxFontVariantPolicyV1): NativeDocxPageFieldVariantV1[] | undefined {
  const qualify=(shaped:NativeDocxShapedLinesV1,document:NativeDocxDocumentV1)=>font?qualifyNativeDocxFontSubstitutionsV1(shaped,request.resolved_layout,font.manifest,font.policy,document,font.descriptors):[]
  const baseRecords=qualify(request.shaped_lines,request.document)
  const choices=new Map(baseRecords.map(r=>[JSON.stringify([r.source_id,r.source_role]),JSON.stringify(r)]))
  const headerScopes=new Set([...request.document.headers,...request.document.footers].flatMap(s=>s.blocks.flatMap(b=>b.paragraph?[b.paragraph.id,...b.paragraph.runs.map(r=>r.id)]:[])))
  const metadata=(shaped:NativeDocxShapedLinesV1,records:NativeDocxFontSubstitutionV1[])=>font?{...shaped,paragraphs:[],font_substitutions:records.filter(r=>!headerScopes.has(r.source_id)),diagnostics:shaped.diagnostics.filter(d=>!headerScopes.has(d.scope_id)||!isQualifiedNativeDocxFontDiagnosticV1(d,records))}:{...shaped,paragraphs:[]}
  const hasFields = hasNativeDocxPageFieldsV1(request.document)
  if (!hasFields) { if (input !== undefined) throw new TypeError('Page-field variants require authored page fields'); return undefined }
  if (layout.status !== 'paginated') { if (input !== undefined) throw new TypeError('Refused pagination cannot carry field variants'); return undefined }
  if (!Array.isArray(input) || input.length !== layout.pages.length || input.length > DOCX_PAGE_FIELD_LIMITS.maxPages) throw new TypeError('Page-field variants must exactly cover bounded final pagination')
  let fragmentCount = 0
  return input.map((entry, index) => {
    const page = layout.pages[index]!
    if (!entry || typeof entry !== 'object' || Object.keys(entry).sort().join(',') !== 'page_id,shaped_lines' || entry.page_id !== page.id) throw new TypeError('Page-field variants must use exact final page order and identity')
    const document = nativeDocxPageFieldDocumentV1(request.document, page.ordinal, layout.pages.length, nativeDocxPageNumberV1(request.document, layout, page.ordinal))
    const decoded = decodeNativeDocxPaginationRequestV1({ ...request, document, shaped_lines: entry.shaped_lines })
    if (!decoded.ok) throw new TypeError(`Page-field shaping source join failed: ${decoded.issues[0]?.message}`)
    const shaped = decoded.value.shaped_lines
    const records=qualify(shaped,document)
    for(const r of records){const key=JSON.stringify([r.source_id,r.source_role]),choice=JSON.stringify(r);if(choices.has(key)&&choices.get(key)!==choice)throw new TypeError('Page variants conflict on a source font choice');choices.set(key,choice)}
    const originalNonStories = request.shaped_lines.paragraphs.filter((p) => p.story_kind !== 'header' && p.story_kind !== 'footer')
    const derivedNonStories = shaped.paragraphs.filter((p) => p.story_kind !== 'header' && p.story_kind !== 'footer')
    if (JSON.stringify(originalNonStories) !== JSON.stringify(derivedNonStories) || JSON.stringify(metadata(request.shaped_lines,baseRecords)) !== JSON.stringify(metadata(shaped,records))) throw new TypeError('Page fields must not change body/note shaping, dimensions or provider identity')
    fragmentCount += shaped.paragraphs.reduce((total, p) => total + p.lines.reduce((n, line) => n + line.fragments.length, 0), 0)
    if (fragmentCount > DOCX_PAGE_FIELD_LIMITS.maxFragments) throw new RangeError('Cumulative page-field shaping exceeds fragment budget')
    return { page_id: page.id, shaped_lines: shaped }
  })
}

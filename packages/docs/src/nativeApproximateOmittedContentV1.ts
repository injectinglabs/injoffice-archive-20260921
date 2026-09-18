import type { NativeDocxDocumentV1, NativeDocxParagraphV1 } from './nativeContract.js'
import type { NativeDocxResolvedLayoutInputV1 } from './nativeResolvedLayout.js'
import type { NativeDocxShapedLinesV1 } from './nativeShapingLines.js'
import type { NativeDocxPagePaintV1 } from './nativePagePaintV1.js'
import { isRenderNeutralLayoutDiagnostic } from './nativeRenderDiagnostics.js'

/** Bounded explicit disclosure of source content the approximate preview did
 * not paint. Formatting-only approximations (colors, spacing, font metadata)
 * remain in the original diagnostics; this list is limited to content whose
 * glyphs, images, or blocks are absent from the painted pages. */
export const DOCX_APPROXIMATE_OMITTED_CONTENT_LIMIT = 64 as const
export const DOCX_APPROXIMATE_OMITTED_CONTENT_WARNING = 'Approximate preview omitted source content it cannot paint; see omitted_content for each dropped item.' as const

/** Source/resolution codes whose approximate omission drops visible content
 * rather than only approximating a property of painted content. */
export const DOCX_APPROXIMATE_OMITTED_CONTENT_CODES = new Set([
  'UNMODELED_RUN_CONTENT',
  'UNMODELED_PARAGRAPH_CONTENT',
  'UNMODELED_BODY_BLOCK',
  'NESTED_TABLE_OR_CELL_MARKUP',
  'PICTURE_GRAPHIC_REQUIRED',
  'PICTURE_TRANSFORM_PRESERVED',
  'UNMODELED_DRAWING',
  'UNRESOLVED_COMMENT_RANGE',
  'UNRESOLVED_COMMENT_REFERENCE',
  'FIELD_SEMANTICS',
  'WRAPPED_RUN_MARKUP',
  // The vertical rule w:cols w:sep="1" asks for is ink Word draws and this
  // tier does not, so it is a dropped visible mark rather than an approximated
  // property of painted content.
  'COLUMN_SEPARATOR_UNSUPPORTED',
  // A Word 2010 run-typography extension the approximate tier paints without:
  // the requested stylistic set, ligature mode, digit form or digit spacing
  // selects glyphs and advances the painted run does not carry, and w14:props3d
  // is extrusion, bevel and contour ink Word draws (it rasterises such a run
  // into an image in its own PDF export) and this tier does not. Each is a
  // visible mark absent from the painted page, not an approximated property of
  // it, so it is named here and the preview reports content_status 'partial'.
  'STYLISTIC_SET_UNAPPLIED',
  'LIGATURE_MODE_UNAPPLIED',
  'NUMBER_FORM_UNAPPLIED',
  'NUMBER_SPACING_UNAPPLIED',
  'TEXT_EFFECT_3D_UNAPPLIED',
  // The rest of the Word 2010 run decorations - w14:glow, w14:shadow,
  // w14:reflection, w14:textOutline, w14:scene3d - and a DISABLED w14:cntxtAlts.
  // The decorations are ink Word draws around or through glyphs this tier paints
  // bare; the disabled contextual alternates leave the run painted WITH calt,
  // which selects glyph forms Word does not paint. Each is a visible mark absent
  // from the painted page, so it is named here and the preview reports
  // content_status 'partial'.
  'TEXT_DECORATION_UNAPPLIED',
  'CONTEXTUAL_ALTERNATES_UNAPPLIED',
  // The level's picture bullet is a marker the source asks for and the painted
  // page does not carry: the paragraph is painted with no marker at all. Word
  // paints no bullet ink for an empty one either, but the request is still a
  // mark this preview does not produce, so it is named here.
  'EMPTY_PICTURE_BULLET_UNPAINTED',
  // A rotated cell's text is painted, but not where or how Word paints it: the
  // lines run along the other axis, wrap against the row height rather than the
  // cell width, and Word clips whatever overflows the cell. That is a visual
  // result this preview does not produce, so it is disclosed here rather than
  // left to a formatting-only note.
  'CELL_TEXT_DIRECTION_UNSUPPORTED',
  // A vertically centred or bottom-seated cell's paragraphs are painted, but
  // not where Word paints them: they sit at the top of the cell box instead of
  // being distributed against the row's height. That is a visual result this
  // preview does not produce, so it is disclosed here rather than left to a
  // formatting-only note.
  'CELL_VERTICAL_ALIGNMENT_UNSUPPORTED',
  // Word compresses a paragraph whose w:line is negative until its lines
  // overlap; this tier paints them at the inherited spacing instead. The
  // compressed line box Word draws is absent from the page, so it is disclosed
  // here rather than left as an approximated property of painted content.
  'NEGATIVE_LINE_SPACING_UNAPPLIED',
  // A list marker whose generated text uses Enclosed Alphanumerics is painted,
  // but not in the face the source resolves for it: the marker's font slot names
  // the paragraph's Latin face and this tier reads no cmap, so it paints a
  // declared host family instead. The glyph design and advances on the page are
  // another face's, which is a visual result this preview does not source from
  // the document, so it is named here and the preview reports 'partial'.
  'ENCLOSED_NUMBER_MARKER_FONT_PRESERVED',
])

/** Shaping diagnostics approximate preview ignores; strict paint refuses them. */
const OMITTED_SHAPING_SEVERITY = 'unsupported'
/** Body paragraph present in the source but absent from shaped lines. */
export const DOCX_APPROXIMATE_UNSHAPED_PARAGRAPH_CODE = 'PARAGRAPH_NOT_SHAPED' as const

export const DOCX_APPROXIMATE_OMITTED_CATEGORIES = ['drawing', 'equation', 'field', 'table', 'comment', 'content-control', 'revision', 'reference', 'text', 'block', 'other'] as const
export type NativeDocxOmittedContentCategoryV1 = typeof DOCX_APPROXIMATE_OMITTED_CATEGORIES[number]
export type NativeDocxOmittedContentOriginV1 = 'source' | 'resolution' | 'shaping' | 'pagination'

export interface NativeDocxOmittedContentV1 {
  /** Original diagnostic code (source/resolution/shaping) or PARAGRAPH_NOT_SHAPED. */
  code: string
  /** Which stage skipped the diagnostic in approximate mode. */
  origin: NativeDocxOmittedContentOriginV1
  category: NativeDocxOmittedContentCategoryV1
  /** Paragraph, run, story, table or document id the diagnostic is scoped to. */
  scope_id: string
  part_name?: string
  path?: string
  message: string
  /** Identical (code, scope, path, message) records merged into this entry. */
  count: number
}

export type NativeDocxApproximateContentStatusV1 = 'complete' | 'partial'

export interface NativeDocxApproximateOmissionsV1 {
  content_status: NativeDocxApproximateContentStatusV1
  omitted_content: NativeDocxOmittedContentV1[]
  /** Total omitted records before bounding and merging. */
  omitted_content_total: number
  /** Every painted page id without any paint command. A blank page can never decode as complete. */
  unpainted_pages: string[]
}

export type NativeDocxApproximateOmissionSourceV1 = { document: NativeDocxDocumentV1; resolved_layout: NativeDocxResolvedLayoutInputV1; shaped_lines: NativeDocxShapedLinesV1 }

const MAX_STRING = 8192
const ID = /^[^\u0000\r\n]{1,512}$/

/** Source paths carry parser-owned namespace prefixes (unknown namespaces are
 * hashed, e.g. `ns4d2aa588:AlternateContent`), so local names decide. */
const local = (names: string) => new RegExp(`/(?:[A-Za-z0-9_.-]+:)?(?:${names})\\[`)
const EQUATION = local('oMath|oMathPara')
const DRAWING = local('drawing|pict|object|anchor|inline|shape|group|txbxContent|AlternateContent|graphic|graphicData')
const FIELD = local('fldSimple|fldChar|instrText')
const CONTENT_CONTROL = local('sdt|sdtContent')
const REVISION = local('ins|del|moveFrom|moveTo')
const COMMENT = local('commentRangeStart|commentRangeEnd|commentReference')
const TABLE = local('tbl')
/** Non-visual markers: Word paints nothing for them either, so their
 * preservation is not omitted content. They stay in the original diagnostics. */
const MARKER = local('bookmarkStart|bookmarkEnd|proofErr|permStart|permEnd')
export function nativeDocxNonVisualMarkerPathV1(path: string | undefined): boolean { return MARKER.test(path ?? '') }

export function nativeDocxOmittedContentCategoryV1(code: string, path: string | undefined): NativeDocxOmittedContentCategoryV1 {
  const at = path ?? ''
  if (EQUATION.test(at)) return 'equation'
  if (DRAWING.test(at) || code === 'UNMODELED_DRAWING' || code === 'PICTURE_GRAPHIC_REQUIRED' || code === 'PICTURE_TRANSFORM_PRESERVED' || code === 'drawing-layout-unsupported') return 'drawing'
  if (FIELD.test(at) || code === 'FIELD_SEMANTICS') return 'field'
  if (CONTENT_CONTROL.test(at)) return 'content-control'
  if (REVISION.test(at) || code === 'WRAPPED_RUN_MARKUP') return 'revision'
  if (COMMENT.test(at) || code === 'UNRESOLVED_COMMENT_RANGE' || code === 'UNRESOLVED_COMMENT_REFERENCE') return 'comment'
  if (TABLE.test(at) || code === 'NESTED_TABLE_OR_CELL_MARKUP' || code === 'table-layout-unsupported') return 'table'
  if (code === 'reference-layout-unsupported') return 'reference'
  if (code === 'UNMODELED_BODY_BLOCK') return 'block'
  if (code === 'UNMODELED_RUN_CONTENT' || code === 'UNMODELED_PARAGRAPH_CONTENT' || code === DOCX_APPROXIMATE_UNSHAPED_PARAGRAPH_CODE || code.startsWith('unsupported-numbering-') || code.endsWith('-unresolved') || code.endsWith('-unsupported')) return 'text'
  // A run-typography feature the painted run does not carry is a deviation in
  // painted text, so it is disclosed under the text category next to it.
  if (code === 'STYLISTIC_SET_UNAPPLIED' || code === 'LIGATURE_MODE_UNAPPLIED' || code === 'NUMBER_FORM_UNAPPLIED' || code === 'NUMBER_SPACING_UNAPPLIED' || code === 'TEXT_EFFECT_3D_UNAPPLIED' || code === 'TEXT_DECORATION_UNAPPLIED' || code === 'CONTEXTUAL_ALTERNATES_UNAPPLIED' || code === 'NEGATIVE_LINE_SPACING_UNAPPLIED' || code === 'EMPTY_PICTURE_BULLET_UNPAINTED' || code === 'ENCLOSED_NUMBER_MARKER_FONT_PRESERVED') return 'text'
  return 'other'
}

function bodyParagraphs(document: NativeDocxDocumentV1): NativeDocxParagraphV1[] {
  const out: NativeDocxParagraphV1[] = []
  for (const block of document.body.blocks) {
    if (block.paragraph) out.push(block.paragraph)
    for (const row of block.table?.rows ?? []) for (const cell of row.cells) out.push(...cell.paragraphs)
  }
  return out
}

/** Every diagnostic the approximate run skipped instead of refusing, limited
 * to content omissions, derived from the same inputs the run consulted. */
export function collectNativeDocxApproximateOmissionsV1(source: NativeDocxApproximateOmissionSourceV1, paint: Pick<NativeDocxPagePaintV1, 'status' | 'pages'>): NativeDocxApproximateOmissionsV1 {
  if (paint.status !== 'painted') return nativeDocxApproximateRefusalOmissionsV1()
  const merged = new Map<string, NativeDocxOmittedContentV1>()
  let total = 0
  const record = (entry: Omit<NativeDocxOmittedContentV1, 'count'>) => {
    total += 1
    const key = [entry.origin, entry.code, entry.scope_id, entry.part_name ?? '', entry.path ?? '', entry.message].join('\u0000')
    const existing = merged.get(key)
    if (existing) existing.count += 1
    else merged.set(key, { ...entry, count: 1 })
  }
  for (const entry of source.document.unsupported) {
    if (!DOCX_APPROXIMATE_OMITTED_CONTENT_CODES.has(entry.code) || nativeDocxNonVisualMarkerPathV1(entry.anchor?.path)) continue
    record({ code: entry.code, origin: 'source', category: nativeDocxOmittedContentCategoryV1(entry.code, entry.anchor?.path), scope_id: entry.scope_id, ...(entry.anchor ? { part_name: entry.anchor.part_name, path: entry.anchor.path } : {}), message: entry.message })
  }
  for (const entry of source.resolved_layout.diagnostics) {
    if (!DOCX_APPROXIMATE_OMITTED_CONTENT_CODES.has(entry.code) || isRenderNeutralLayoutDiagnostic(entry, source.resolved_layout) || nativeDocxNonVisualMarkerPathV1(entry.path)) continue
    record({ code: entry.code, origin: 'resolution', category: nativeDocxOmittedContentCategoryV1(entry.code, entry.path), scope_id: entry.scope_id, ...(entry.part_name !== undefined ? { part_name: entry.part_name } : {}), ...(entry.path !== undefined ? { path: entry.path } : {}), message: entry.message })
  }
  const anchors = new Map<string, { part_name: string; path: string }>()
  for (const paragraph of bodyParagraphs(source.document)) {
    anchors.set(paragraph.id, { part_name: paragraph.anchor.part_name, path: paragraph.anchor.path })
    for (const run of paragraph.runs) anchors.set(run.id, { part_name: run.anchor.part_name, path: run.anchor.path })
  }
  for (const entry of source.shaped_lines.diagnostics) {
    if (entry.severity !== OMITTED_SHAPING_SEVERITY) continue
    const anchor = anchors.get(entry.source_id ?? '') ?? anchors.get(entry.scope_id)
    record({ code: entry.code, origin: 'shaping', category: nativeDocxOmittedContentCategoryV1(entry.code, anchor?.path), scope_id: entry.source_id ?? entry.scope_id, ...(anchor ?? {}), message: entry.message })
  }
  // Backstop only: a dropped paragraph whose paragraph or run scope already has
  // an entry above is that entry, not a second item.
  const disclosed = new Set([...merged.values()].map((entry) => entry.scope_id))
  const shaped = new Set(source.shaped_lines.paragraphs.map((paragraph) => paragraph.paragraph_id))
  for (const paragraph of bodyParagraphs(source.document)) {
    if (shaped.has(paragraph.id) || disclosed.has(paragraph.id) || paragraph.runs.some((run) => disclosed.has(run.id))) continue
    const content = paragraph.runs.filter((run) => (run.kind === 'text' && (run.text ?? '') !== '') || run.kind === 'drawing' || run.kind === 'reference')
    if (content.length === 0) continue
    const category: NativeDocxOmittedContentCategoryV1 = content.some((run) => run.kind === 'drawing') ? 'drawing' : content.some((run) => run.kind === 'text') ? 'text' : 'reference'
    record({ code: DOCX_APPROXIMATE_UNSHAPED_PARAGRAPH_CODE, origin: 'pagination', category, scope_id: paragraph.id, part_name: paragraph.anchor.part_name, path: paragraph.anchor.path, message: `Body paragraph with ${category} content has no shaped lines and was not painted, and no other diagnostic explains it` })
  }
  const omitted = [...merged.values()].sort((left, right) => compare(left, right)).slice(0, DOCX_APPROXIMATE_OMITTED_CONTENT_LIMIT)
  const unpainted = paint.pages.filter((page) => page.commands.length === 0).map((page) => page.id)
  return { content_status: omitted.length > 0 || unpainted.length > 0 ? 'partial' : 'complete', omitted_content: omitted, omitted_content_total: total, unpainted_pages: unpainted }
}

function compare(left: NativeDocxOmittedContentV1, right: NativeDocxOmittedContentV1): number {
  for (const key of ['category', 'code', 'scope_id', 'path', 'message'] as const) {
    const a = left[key] ?? '', b = right[key] ?? ''
    if (a !== b) return a < b ? -1 : 1
  }
  return 0
}

/** Refusals paint nothing, so their disclosure is a fixed empty partial record. */
export function nativeDocxApproximateRefusalOmissionsV1(): NativeDocxApproximateOmissionsV1 {
  return { content_status: 'partial', omitted_content: [], omitted_content_total: 0, unpainted_pages: [] }
}

/** Structural validation for the wire; content_status must agree with the
 * disclosed omissions, and unpainted_pages must list exactly the painted pages
 * without commands, so a painted-but-blank result can never decode as complete. */
export function validNativeDocxApproximateOmissionsV1(value: Partial<NativeDocxApproximateOmissionsV1>, paint: Pick<NativeDocxPagePaintV1, 'status' | 'pages'>): boolean {
  const { content_status, omitted_content, omitted_content_total, unpainted_pages } = value
  if (content_status !== 'complete' && content_status !== 'partial') return false
  if (!Array.isArray(omitted_content) || omitted_content.length > DOCX_APPROXIMATE_OMITTED_CONTENT_LIMIT) return false
  if (!Number.isSafeInteger(omitted_content_total) || omitted_content_total! < 0) return false
  const keys = new Set<string>()
  let counted = 0
  for (const entry of omitted_content) {
    if (!entry || typeof entry !== 'object' || Object.keys(entry).some((key) => !['code', 'origin', 'category', 'scope_id', 'part_name', 'path', 'message', 'count'].includes(key))) return false
    if (typeof entry.code !== 'string' || !ID.test(entry.code) || typeof entry.scope_id !== 'string' || !ID.test(entry.scope_id) || typeof entry.message !== 'string' || entry.message.length < 1 || entry.message.length > MAX_STRING) return false
    if (!['source', 'resolution', 'shaping', 'pagination'].includes(entry.origin) || !(DOCX_APPROXIMATE_OMITTED_CATEGORIES as readonly string[]).includes(entry.category)) return false
    if ((entry.part_name !== undefined && (typeof entry.part_name !== 'string' || !ID.test(entry.part_name))) || (entry.path !== undefined && (typeof entry.path !== 'string' || entry.path.length < 1 || entry.path.length > MAX_STRING))) return false
    if (!Number.isSafeInteger(entry.count) || entry.count < 1) return false
    const key = [entry.origin, entry.code, entry.scope_id, entry.part_name ?? '', entry.path ?? '', entry.message].join('\u0000')
    if (keys.has(key)) return false
    keys.add(key)
    counted += entry.count
  }
  if (counted > omitted_content_total!) return false
  if (!Array.isArray(unpainted_pages) || unpainted_pages.some((id) => typeof id !== 'string') || new Set(unpainted_pages).size !== unpainted_pages.length) return false
  if (paint.status !== 'painted') return content_status === 'partial' && omitted_content.length === 0 && unpainted_pages.length === 0
  const blank = paint.pages.filter((page) => page.commands.length === 0).map((page) => page.id)
  if (unpainted_pages.length !== blank.length || unpainted_pages.some((id, index) => id !== blank[index])) return false
  return content_status === (omitted_content.length > 0 || unpainted_pages.length > 0 ? 'partial' : 'complete')
}

/** Human summary such as "3 items not rendered: drawings (2), equations (1)". */
export function nativeDocxOmittedContentSummaryV1(value: Pick<NativeDocxApproximateOmissionsV1, 'omitted_content' | 'omitted_content_total' | 'unpainted_pages'>): string | null {
  const counts = new Map<NativeDocxOmittedContentCategoryV1, number>()
  let listed = 0
  for (const entry of value.omitted_content) { counts.set(entry.category, (counts.get(entry.category) ?? 0) + entry.count); listed += entry.count }
  const parts: string[] = []
  if (listed > 0) {
    const groups = [...counts.entries()].sort((a, b) => b[1] - a[1] || (a[0] < b[0] ? -1 : 1)).map(([category, count]) => `${plural(category)} (${count})`)
    const hidden = value.omitted_content_total - listed
    parts.push(`${value.omitted_content_total} item${value.omitted_content_total === 1 ? '' : 's'} not rendered: ${groups.join(', ')}${hidden > 0 ? `, ${hidden} more not listed` : ''}`)
  }
  if (value.unpainted_pages.length > 0) parts.push(`${value.unpainted_pages.length} page${value.unpainted_pages.length === 1 ? '' : 's'} painted nothing`)
  return parts.length ? parts.join('; ') : null
}

function plural(category: NativeDocxOmittedContentCategoryV1): string {
  switch (category) {
    case 'drawing': return 'drawings'
    case 'equation': return 'equations'
    case 'field': return 'fields'
    case 'table': return 'tables'
    case 'comment': return 'comments'
    case 'content-control': return 'content controls'
    case 'revision': return 'tracked changes'
    case 'reference': return 'references'
    case 'text': return 'text runs'
    case 'block': return 'blocks'
    case 'other': return 'other items'
  }
}

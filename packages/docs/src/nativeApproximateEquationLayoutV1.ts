/** Approximate OMML equation layout and paint for the read-only approximate
 * page preview.
 *
 * The strict extractor keeps refusing m:oMath / m:oMathPara content. The Go
 * sidecar (`InspectNativeApproximateEquationsV1`) joins those refusals to their
 * source nodes and describes a bounded construct tree with resolved runs. This
 * module validates that sidecar against the current document, lays every
 * supported equation out as TeX-style boxes (script scaling 0.7 / 0.5, fraction
 * rule on the math axis, n-ary limits above and below in display style,
 * delimiters stretched by uniform glyph scaling), reserves the equation extent
 * as a glyphless inline atom in the internal body copy, and after pagination
 * paints glyph runs shaped with the declared math face plus rules. Italic
 * correction and OpenType MATH tables are not consulted; that is disclosed.
 * Nothing here touches strict paint or source bytes. */
import { NATIVE_TEXT_LAYOUT_VERSION, type FontResource, type NativeFontFaceManifest, type NativeFontManifest, type NativeFontResolver, type NativeTextShaper, type ResolvedFontFace, type ShapedSegment, type TextRunInput } from '@injoffice/font-metrics/layout'
import type { NativeDocxDocumentV1, NativeDocxParagraphV1, NativeDocxRunV1, NativeDocxSourceAnchorV1 } from './nativeContract.js'
import type { NativeDocxResolvedLayoutInputV1 } from './nativeResolvedLayout.js'
import type { NativeDocxShapedLinesV1 } from './nativeShapingLines.js'
import { ID, RGB, preflightWire, paintCommandID } from './nativePagePaintWireV1.js'
import { nativeDocxPlaceGlyphPathV1, nativeDocxCaptureGlyphOutlineV1, type NativeDocxContentAddressedFaceV1, type NativeDocxFillGlyphPathCommandV1, type NativeDocxFillTableCellCommandV1, type NativeDocxFillTextHighlightCommandV1, type NativeDocxGlyphDesignPathCommandV1, type NativeDocxGlyphOutlineProviderV1, type NativeDocxGlyphOutlineResultV1, type NativeDocxPagePaintCommandV1, type NativeDocxPagePaintSuccessV1, type NativeDocxPaintLineV1, type NativeDocxPaintPageV1 } from './nativePagePaintV1.js'
import { qualifyNativeDocxInlineTextboxV1 } from './nativeTextboxInlineV1.js'
import { asciiLowerNative } from './nativeDeterminism.js'
import { collectNativeDocxApproximateOmissionsV1, DOCX_APPROXIMATE_OMITTED_CONTENT_WARNING, type NativeDocxApproximateOmissionsV1 } from './nativeApproximateOmittedContentV1.js'

export const DOCX_APPROXIMATE_EQUATIONS_PROTOCOL = 'injoffice.docx.approximate-equations' as const
export const DOCX_APPROXIMATE_EQUATION_POLICY = 'docx.approximate-equation-preview-v1' as const
export const DOCX_APPROXIMATE_EQUATION_CODE = 'docx.approximate-equation-preview' as const
export const DOCX_APPROXIMATE_EQUATION_OMITTED_CODE = 'docx.approximate-equation-omitted' as const
export const DOCX_APPROXIMATE_EQUATION_FONT_CODE = 'docx.approximate-equation-font' as const
export const DOCX_APPROXIMATE_EQUATION_WARNING = `${DOCX_APPROXIMATE_EQUATION_CODE}: OMML equations are painted approximately with TeX-style box layout (script scaling 0.7/0.5, fraction rule on the math axis, display limits above and below n-ary operators, delimiters stretched by uniform glyph scaling); italic correction, OpenType MATH metrics and Word's equation line breaking are not modeled. The equation extent is reserved above the line baseline. Original equation refusals and source bytes are unchanged.` as const
/** Rule primitives carry this table id so consumers can tell equation paint from table paint. */
export const DOCX_APPROXIMATE_EQUATION_TABLE_ID = DOCX_APPROXIMATE_EQUATION_POLICY

const MAX_EQUATIONS = 64
const MAX_LINES = 64
const MAX_NODES = 512
const MAX_DEPTH = 32
const MAX_TEXT_UNITS = 4096
const MAX_FONT_REQUESTS = 32
const MAX_GLYPHS = 20_000
const MAX_REASONS = 24
/** Interactive viewers read at most 16 MiB; keep body paint plus equation paint under this. */
const MAX_ENVELOPE_BYTES = 15 * 1024 * 1024
const MIN_FONT_SIZE = 5_000
const MAX_FONT_SIZE = 1_638_000
const EMU_PER_MILLIPOINT = 12.7
const SCRIPT_SCALE = [1, 0.7, 0.5] as const
/** Declared math faces the preview accepts in place of an absent authored math face. */
export const DOCX_APPROXIMATE_MATH_SUBSTITUTE_FAMILIES = ['Cambria Math', 'STIX Two Math', 'STIX Math', 'STIXGeneral', 'XITS Math', 'Latin Modern Math', 'Asana Math', 'Libertinus Math', 'TeX Gyre Termes Math', 'TeX Gyre Pagella Math', 'DejaVu Math TeX Gyre', 'Fira Math', 'Noto Sans Math'] as const

export type NativeDocxApproximateMathKindV1 = 'row' | 'text' | 'fraction' | 'superscript' | 'subscript' | 'subsuperscript' | 'nary' | 'delimiter' | 'radical' | 'function' | 'bar' | 'accent' | 'limit-lower' | 'limit-upper'
export interface NativeDocxApproximateEquationFontV1 { family: string; weight: number; style: string }
export interface NativeDocxApproximateMathRunV1 { font_family: string; font_size_half_points: number; bold: boolean; italic: boolean; color?: string; style?: 'p' | 'b' | 'i' | 'bi'; normal?: boolean }
export interface NativeDocxApproximateMathNodeV1 {
  kind: NativeDocxApproximateMathKindV1
  text?: string
  run?: NativeDocxApproximateMathRunV1
  children?: NativeDocxApproximateMathNodeV1[]
  chr?: string
  beg_chr?: string
  end_chr?: string
  sep_chr?: string
  bar?: boolean
  limit_location?: 'undOvr' | 'subSup'
  sub_hide?: boolean
  sup_hide?: boolean
  degree_hide?: boolean
  position?: 'top' | 'bot'
  grow?: boolean
}
export interface NativeDocxApproximateEquationV1 {
  id: string
  paragraph_id: string
  diagnostic_ids: string[]
  anchor: NativeDocxSourceAnchorV1
  status: 'supported' | 'omitted'
  reason?: string
  display: boolean
  justification?: 'left' | 'right' | 'center' | 'centerGroup'
  lines?: NativeDocxApproximateMathNodeV1[]
  notes?: string[]
}
export interface NativeDocxApproximateEquationsV1 {
  protocol: typeof DOCX_APPROXIMATE_EQUATIONS_PROTOCOL
  version: 1
  policy: typeof DOCX_APPROXIMATE_EQUATION_POLICY
  package_sha256: string
  part_sha256: string
  items: NativeDocxApproximateEquationV1[]
  omitted_count: number
  font_requests: NativeDocxApproximateEquationFontV1[]
}

function record(value: unknown): value is Record<string, unknown> { return value !== null && typeof value === 'object' && !Array.isArray(value) }
function exactKeys(value: Record<string, unknown>, required: readonly string[], optional: readonly string[] = []): boolean {
  const keys = Object.keys(value)
  return required.every(key => key in value) && keys.every(key => required.includes(key) || optional.includes(key))
}
function anchorValid(value: unknown, part: string): value is NativeDocxSourceAnchorV1 {
  return record(value) && exactKeys(value, ['part_name', 'path', 'start_byte', 'end_byte', 'xml_sha256']) && value.part_name === part && typeof value.path === 'string' && Number.isSafeInteger(value.start_byte) && Number.isSafeInteger(value.end_byte) && (value.start_byte as number) >= 0 && (value.end_byte as number) > (value.start_byte as number) && typeof value.xml_sha256 === 'string' && /^sha256:[0-9a-f]{64}$/.test(value.xml_sha256)
}
function within(inner: NativeDocxSourceAnchorV1, outer: NativeDocxSourceAnchorV1): boolean { return inner.part_name === outer.part_name && inner.start_byte >= outer.start_byte && inner.end_byte <= outer.end_byte }
function sameAnchor(left: NativeDocxSourceAnchorV1, right: NativeDocxSourceAnchorV1): boolean { return left.part_name === right.part_name && left.path === right.path && left.start_byte === right.start_byte && left.end_byte === right.end_byte && left.xml_sha256 === right.xml_sha256 }
function oneCodePoint(value: unknown): value is string { return typeof value === 'string' && [...value].length === 1 }
function fontValid(value: unknown): value is NativeDocxApproximateEquationFontV1 { return record(value) && exactKeys(value, ['family', 'weight', 'style']) && typeof value.family === 'string' && value.family.length > 0 && value.family.length <= 128 && (value.weight === 400 || value.weight === 700) && (value.style === 'normal' || value.style === 'italic') }

function validateRun(value: unknown): value is NativeDocxApproximateMathRunV1 {
  if (!record(value) || !exactKeys(value, ['font_family', 'font_size_half_points', 'bold', 'italic'], ['color', 'style', 'normal'])) return false
  if (typeof value.font_family !== 'string' || !value.font_family.length || value.font_family.length > 128 || !Number.isSafeInteger(value.font_size_half_points) || (value.font_size_half_points as number) < 1 || (value.font_size_half_points as number) > 3276) return false
  if (typeof value.bold !== 'boolean' || typeof value.italic !== 'boolean') return false
  if (value.color !== undefined && (typeof value.color !== 'string' || !RGB.test(value.color))) return false
  if (value.style !== undefined && !['p', 'b', 'i', 'bi'].includes(String(value.style))) return false
  return value.normal === undefined || typeof value.normal === 'boolean'
}

const ARITY: Record<NativeDocxApproximateMathKindV1, number | undefined> = { row: undefined, text: 0, fraction: 2, superscript: 2, subscript: 2, subsuperscript: 3, nary: 3, delimiter: undefined, radical: 2, function: 2, bar: 1, accent: 1, 'limit-lower': 2, 'limit-upper': 2 }
const NODE_OPTIONAL: Record<NativeDocxApproximateMathKindV1, readonly string[]> = {
  row: ['children'], text: ['text', 'run'], fraction: ['children', 'run', 'bar'], superscript: ['children', 'run'], subscript: ['children', 'run'], subsuperscript: ['children', 'run'],
  nary: ['children', 'run', 'chr', 'limit_location', 'sub_hide', 'sup_hide', 'grow'], delimiter: ['children', 'run', 'beg_chr', 'end_chr', 'sep_chr', 'grow'], radical: ['children', 'run', 'degree_hide'],
  function: ['children', 'run'], bar: ['children', 'run', 'position'], accent: ['children', 'run', 'chr'], 'limit-lower': ['children', 'run'], 'limit-upper': ['children', 'run'],
}

function validateNode(value: unknown, depth: number, budget: { nodes: number; units: number }): value is NativeDocxApproximateMathNodeV1 {
  if (depth > MAX_DEPTH || ++budget.nodes > MAX_NODES || !record(value) || typeof value.kind !== 'string' || !(value.kind in ARITY)) return false
  const kind = value.kind as NativeDocxApproximateMathKindV1
  if (!exactKeys(value, ['kind'], NODE_OPTIONAL[kind])) return false
  if (value.run !== undefined && !validateRun(value.run)) return false
  if (kind === 'text') {
    if (typeof value.text !== 'string' || !value.text.length || !validateRun(value.run)) return false
    budget.units += value.text.length
    return budget.units <= MAX_TEXT_UNITS
  }
  if (value.text !== undefined) return false
  for (const key of ['chr', 'beg_chr', 'end_chr', 'sep_chr'] as const) if (value[key] !== undefined && (typeof value[key] !== 'string' || (key === 'chr' ? !oneCodePoint(value[key]) : (value[key] as string).length !== 0 && !oneCodePoint(value[key])))) return false
  for (const key of ['bar', 'sub_hide', 'sup_hide', 'degree_hide', 'grow'] as const) if (value[key] !== undefined && typeof value[key] !== 'boolean') return false
  if (value.limit_location !== undefined && value.limit_location !== 'undOvr' && value.limit_location !== 'subSup') return false
  if (value.position !== undefined && value.position !== 'top' && value.position !== 'bot') return false
  const children = value.children ?? []
  if (!Array.isArray(children) || children.length > MAX_NODES) return false
  const arity = ARITY[kind]
  if (arity !== undefined ? children.length !== arity : kind === 'delimiter' && children.length === 0) return false
  return children.every(child => validateNode(child, depth + 1, budget))
}

/** Bounded structural validation plus source joins: every equation must name a
 * body paragraph, carry its own retained UNMODELED_PARAGRAPH_CONTENT refusal at
 * exactly the same anchor, sit directly under the paragraph, and not overlap
 * modeled runs. A malformed sidecar refuses as a whole. */
export function decodeNativeDocxApproximateEquationsV1(value: unknown, document: NativeDocxDocumentV1): NativeDocxApproximateEquationsV1 {
  if (preflightWire(value, 'approximate equations', 200_000, 10_000).length) throw new TypeError('Approximate equations exceed their bounded wire')
  const input = structuredClone(value) as NativeDocxApproximateEquationsV1
  if (!record(input) || !exactKeys(input as unknown as Record<string, unknown>, ['protocol', 'version', 'policy', 'package_sha256', 'part_sha256', 'items', 'omitted_count', 'font_requests']) || input.protocol !== DOCX_APPROXIMATE_EQUATIONS_PROTOCOL || input.version !== 1 || input.policy !== DOCX_APPROXIMATE_EQUATION_POLICY || input.package_sha256 !== document.source.package_sha256 || typeof input.part_sha256 !== 'string' || !/^sha256:[0-9a-f]{64}$/.test(input.part_sha256) || !Array.isArray(input.items) || input.items.length > MAX_EQUATIONS || !Number.isSafeInteger(input.omitted_count) || input.omitted_count < 0 || input.omitted_count > 1_000_000) throw new TypeError('Approximate equations do not exact-join the source document')
  if (!Array.isArray(input.font_requests) || input.font_requests.length > MAX_FONT_REQUESTS || !input.font_requests.every(fontValid)) throw new TypeError('Approximate equation font requests are invalid')
  const paragraphs = new Map(document.body.blocks.flatMap(block => block.paragraph ? [[block.id, block.paragraph] as const] : []))
  const diagnostics = new Map(document.unsupported.map(entry => [entry.id, entry]))
  const ids = new Set<string>()
  const main = document.source.main_part
  for (const item of input.items) {
    if (!record(item) || !exactKeys(item as unknown as Record<string, unknown>, ['id', 'paragraph_id', 'diagnostic_ids', 'anchor', 'status', 'display'], ['reason', 'justification', 'lines', 'notes'])) throw new TypeError('Approximate equation has unknown or missing fields')
    if (typeof item.id !== 'string' || !ID.test(item.id) || ids.has(item.id)) throw new TypeError('Approximate equation id is invalid or duplicated')
    ids.add(item.id)
    const paragraph = paragraphs.get(item.paragraph_id)
    if (!paragraph || !anchorValid(item.anchor, main) || !within(item.anchor, paragraph.anchor) || !item.anchor.path.startsWith(paragraph.anchor.path + '/') || item.anchor.path.slice(paragraph.anchor.path.length + 1).includes('/') || !/:oMath(?:Para)?\[[1-9][0-9]*\]$/.test(item.anchor.path)) throw new TypeError('Approximate equation does not exact-join its body paragraph')
    if (!Array.isArray(item.diagnostic_ids) || item.diagnostic_ids.length === 0 || item.diagnostic_ids.length > 64 || item.diagnostic_ids.some(id => { const d = typeof id === 'string' ? diagnostics.get(id) : undefined; return !d || d.scope_id !== item.paragraph_id || !d.anchor || !within(d.anchor, item.anchor) })) throw new TypeError('Approximate equation must join retained source equation diagnostics')
    if (!item.diagnostic_ids.some(id => { const d = diagnostics.get(id); return d?.code === 'UNMODELED_PARAGRAPH_CONTENT' && d.anchor && sameAnchor(d.anchor, item.anchor) })) throw new TypeError('Approximate equation must join the retained refusal of its own root')
    if (paragraph.runs.some(run => within(run.anchor, item.anchor) || within(item.anchor, run.anchor))) throw new TypeError('Approximate equation overlaps modeled text')
    if (item.status !== 'supported' && item.status !== 'omitted' || typeof item.display !== 'boolean') throw new TypeError('Approximate equation status is invalid')
    if (item.reason !== undefined && (typeof item.reason !== 'string' || item.reason.length > 256)) throw new TypeError('Approximate equation reason is unbounded')
    if (item.notes !== undefined && (!Array.isArray(item.notes) || item.notes.length > 32 || item.notes.some(note => typeof note !== 'string' || note.length > 512))) throw new TypeError('Approximate equation notes are unbounded')
    if (item.justification !== undefined && (!item.display || !['left', 'right', 'center', 'centerGroup'].includes(item.justification))) throw new TypeError('Approximate equation justification is invalid')
    if (item.status === 'omitted') { if (item.lines !== undefined || typeof item.reason !== 'string' || !item.reason.length) throw new TypeError('Omitted approximate equation requires a reason and no layout'); continue }
    if (!Array.isArray(item.lines) || item.lines.length === 0 || item.lines.length > MAX_LINES || (item.lines.length > 1 && !item.display)) throw new TypeError('Supported approximate equation requires bounded layout lines')
    const budget = { nodes: 0, units: 0 }
    if (!item.lines.every(line => validateNode(line, 0, budget) && line.kind === 'row')) throw new TypeError('Approximate equation layout tree is invalid or exceeds its budget')
  }
  return input
}

export type NativeDocxApproximateMathFacePolicyV1 = 'exact' | 'declared-math-substitute' | 'text-face-fallback'
export interface NativeDocxApproximateMathFaceChoiceV1 { requested: NativeDocxApproximateEquationFontV1; face: NativeFontFaceManifest; policy: NativeDocxApproximateMathFacePolicyV1; weight_style_match: boolean }

function faceFamilies(face: NativeFontFaceManifest): string[] { return [face.family, ...(face.aliases ?? [])].map(asciiLowerNative) }
function loadable(face: NativeFontFaceManifest): boolean { return !!face.source.contentDigest && face.source.kind !== 'system' && face.stretch === 100 }
function pickFace(candidates: NativeFontFaceManifest[], weight: number, style: string): NativeFontFaceManifest | undefined {
  return candidates.find(face => face.weight === weight && face.style === style) ?? candidates.find(face => face.style === style) ?? candidates.find(face => face.weight === weight) ?? candidates[0]
}

/** Math face policy: the authored family when the manifest provides it, else a
 * declared math substitute family, else the paragraph text face. Never a
 * system lookup, never an invented family. */
export function selectNativeDocxApproximateMathFaceV1(manifest: NativeFontManifest, requested: NativeDocxApproximateEquationFontV1, textFamily: string | undefined): NativeDocxApproximateMathFaceChoiceV1 | undefined {
  const faces = manifest.faces.filter(loadable)
  const family = asciiLowerNative(requested.family)
  const exact = pickFace(faces.filter(face => faceFamilies(face).includes(family)), requested.weight, requested.style)
  if (exact) return { requested, face: exact, policy: 'exact', weight_style_match: exact.weight === requested.weight && exact.style === requested.style }
  const substitutes = DOCX_APPROXIMATE_MATH_SUBSTITUTE_FAMILIES.map(asciiLowerNative)
  const math = pickFace(faces.filter(face => faceFamilies(face).some(name => substitutes.includes(name))), requested.weight, requested.style)
  if (math) return { requested, face: math, policy: 'declared-math-substitute', weight_style_match: math.weight === requested.weight && math.style === requested.style }
  const text = textFamily ? asciiLowerNative(textFamily) : undefined
  const textFaces = text ? faces.filter(face => faceFamilies(face).includes(text)) : []
  const fallback = pickFace(textFaces, requested.weight, requested.style) ?? pickFace(faces.filter(face => face.source.kind === 'host'), requested.weight, requested.style) ?? pickFace(faces, requested.weight, requested.style)
  return fallback ? { requested, face: fallback, policy: 'text-face-fallback', weight_style_match: fallback.weight === requested.weight && fallback.style === requested.style } : undefined
}

export interface NativeDocxApproximateEquationRuntimeV1 {
  manifest: NativeFontManifest
  resolver: NativeFontResolver
  shaper: NativeTextShaper
  outlineProvider: NativeDocxGlyphOutlineProviderV1
}

/** Glyph placed relative to the equation origin: x from the left edge, y from
 * the equation baseline, both in millipoints with y growing downward. */
interface MathGlyph { x: number; y: number; size: number; face: NativeDocxContentAddressedFaceV1; glyph_id: number; color: string; units_per_em: number; path: NativeDocxGlyphDesignPathCommandV1[] }
interface MathRule { x: number; y: number; width: number; height: number; color: string }
type AtomClass = 'ord' | 'op' | 'bin' | 'rel' | 'open' | 'close' | 'punct' | 'inner'
interface MathBox { width: number; ascent: number; depth: number; cls: AtomClass; glyphs: MathGlyph[]; rules: MathRule[] }

export interface NativeDocxApproximateEquationLayoutV1 {
  id: string
  width: number
  ascent: number
  depth: number
  glyphs: MathGlyph[]
  rules: MathRule[]
  notes: string[]
}
export interface NativeDocxApproximateEquationLayoutsV1 {
  layouts: Map<string, NativeDocxApproximateEquationLayoutV1>
  omitted: Array<{ id: string; reason: string }>
  faces: NativeDocxApproximateMathFaceChoiceV1[]
  notes: string[]
}

const BIN = new Set(['+', '−', '-', '±', '∓', '×', '÷', '⋅', '∗', '∘', '∙', '*', '/', '∧', '∨', '∩', '∪', '⊕', '⊗'])
const REL = new Set(['=', '<', '>', '≤', '≥', '≠', '≈', '≡', '∼', '≃', '≅', '∝', '→', '←', '↔', '⇒', '⇐', '⇔', '∈', '∉', '⊂', '⊃', '⊆', '⊇', '∣', '∥', ':', '≪', '≫', '⊥', '≐'])
const OPEN = new Set(['(', '[', '{', '⟨', '⌈', '⌊', '|'])
const CLOSE = new Set([')', ']', '}', '⟩', '⌉', '⌋'])
const PUNCT = new Set([',', ';'])
const LETTER = /^\p{L}$/u
const DIGIT = /^[\p{Nd}.]$/u
const SPACING: Record<AtomClass, Record<AtomClass, number>> = {
  ord: { ord: 0, op: 3, bin: 4, rel: 5, open: 0, close: 0, punct: 0, inner: 3 },
  op: { ord: 3, op: 3, bin: 0, rel: 5, open: 0, close: 0, punct: 0, inner: 3 },
  bin: { ord: 4, op: 4, bin: 0, rel: 0, open: 4, close: 0, punct: 0, inner: 4 },
  rel: { ord: 5, op: 5, bin: 0, rel: 0, open: 5, close: 0, punct: 0, inner: 5 },
  open: { ord: 0, op: 0, bin: 0, rel: 0, open: 0, close: 0, punct: 0, inner: 0 },
  close: { ord: 0, op: 3, bin: 4, rel: 5, open: 0, close: 0, punct: 0, inner: 3 },
  punct: { ord: 3, op: 3, bin: 0, rel: 3, open: 3, close: 0, punct: 3, inner: 3 },
  inner: { ord: 3, op: 3, bin: 4, rel: 5, open: 3, close: 0, punct: 3, inner: 3 },
}

function classify(ch: string): AtomClass {
  if (BIN.has(ch)) return 'bin'
  if (REL.has(ch)) return 'rel'
  if (OPEN.has(ch)) return 'open'
  if (CLOSE.has(ch)) return 'close'
  if (PUNCT.has(ch)) return 'punct'
  return 'ord'
}
function scriptTag(text: string): string { return /\p{Script=Greek}/u.test(text) ? 'Grek' : /\p{Script=Cyrillic}/u.test(text) ? 'Cyrl' : 'Latn' }
function emptyBox(cls: AtomClass = 'ord'): MathBox { return { width: 0, ascent: 0, depth: 0, cls, glyphs: [], rules: [] } }
function shifted(box: MathBox, dx: number, dy: number): MathBox {
  return { ...box, glyphs: box.glyphs.map(glyph => ({ ...glyph, x: glyph.x + dx, y: glyph.y + dy })), rules: box.rules.map(rule => ({ ...rule, x: rule.x + dx, y: rule.y + dy })) }
}
/** Horizontal concatenation with per-gap spacing; boxes share one baseline. */
function hbox(parts: Array<{ box: MathBox; gap: number }>, cls: AtomClass): MathBox {
  const result = emptyBox(cls)
  let x = 0
  for (const { box, gap } of parts) {
    x += gap
    const placed = shifted(box, x, 0)
    result.glyphs.push(...placed.glyphs)
    result.rules.push(...placed.rules)
    result.ascent = Math.max(result.ascent, box.ascent)
    result.depth = Math.max(result.depth, box.depth)
    x += box.width
  }
  result.width = x
  return result
}
/** Raise a box by `up` millipoints (its glyphs move toward the top). */
function raised(box: MathBox, up: number): MathBox { return { ...shifted(box, 0, -up), ascent: box.ascent + up, depth: box.depth - up } }
function overlay(boxes: MathBox[], cls: AtomClass): MathBox {
  const result = emptyBox(cls)
  for (const box of boxes) {
    result.glyphs.push(...box.glyphs)
    result.rules.push(...box.rules)
    result.width = Math.max(result.width, box.width)
    result.ascent = Math.max(result.ascent, box.ascent)
    result.depth = Math.max(result.depth, box.depth)
  }
  return result
}
function centered(box: MathBox, width: number): MathBox { return { ...shifted(box, Math.max(0, Math.round((width - box.width) / 2)), 0), width: Math.max(width, box.width) } }

interface Style { display: boolean; level: 0 | 1 | 2 }
const down = (style: Style): Style => ({ display: false, level: style.level === 0 ? 1 : 2 })
const noDisplay = (style: Style): Style => ({ display: false, level: style.level })

interface LayoutContext {
  runtime: NativeDocxApproximateEquationRuntimeV1
  base: number
  rule: number
  axis: number
  xHeight: number
  textFamily: string | undefined
  faces: Map<string, NativeDocxApproximateMathFaceChoiceV1 | null>
  resources: Map<string, FontResource>
  shaped: Map<string, ShapedSegment | null>
  outlines: Map<string, NativeDocxGlyphOutlineResultV1>
  notes: Set<string>
  glyphs: number
  faceChoices: NativeDocxApproximateMathFaceChoiceV1[]
}

const em = (context: LayoutContext, size: number, units: number) => Math.round(size * units / 18)

function sizeFor(context: LayoutContext, run: NativeDocxApproximateMathRunV1, style: Style): number {
  return Math.max(MIN_FONT_SIZE, Math.min(MAX_FONT_SIZE, Math.round(run.font_size_half_points * 500 * SCRIPT_SCALE[style.level])))
}

async function faceFor(context: LayoutContext, family: string, weight: number, style: string): Promise<NativeDocxApproximateMathFaceChoiceV1 | undefined> {
  const key = `${asciiLowerNative(family)}\0${weight}\0${style}`
  let choice = context.faces.get(key)
  if (choice === undefined) {
    choice = selectNativeDocxApproximateMathFaceV1(context.runtime.manifest, { family, weight, style }, context.textFamily) ?? null
    context.faces.set(key, choice)
    if (choice) context.faceChoices.push(choice)
  }
  return choice ?? undefined
}

async function resourceFor(context: LayoutContext, face: NativeFontFaceManifest): Promise<FontResource | undefined> {
  const cached = context.resources.get(face.faceId)
  if (cached) return cached
  const run: TextRunInput = { version: NATIVE_TEXT_LAYOUT_VERSION, text: 'A', fontSizeMilliPoints: 10_000, font: { families: [face.family], weight: face.weight, style: face.style as 'normal' | 'italic', stretch: 100 }, script: 'Latn', language: 'und', direction: 'ltr' }
  const resolution = await context.runtime.resolver.resolve({ manifest: context.runtime.manifest, run })
  if (resolution.status !== 'resolved' || resolution.face.faceId !== face.faceId) return undefined
  const resource = await context.runtime.resolver.load(resolution.face)
  if (!('bytes' in resource)) return undefined
  context.resources.set(face.faceId, resource)
  return resource
}

async function shapeText(context: LayoutContext, resource: FontResource, text: string, size: number): Promise<ShapedSegment | undefined> {
  const key = `${resource.face.faceId}\0${size}\0${text}`
  if (context.shaped.has(key)) return context.shaped.get(key) ?? undefined
  const run: TextRunInput = Object.freeze({ version: NATIVE_TEXT_LAYOUT_VERSION, text, fontSizeMilliPoints: size, features: [{ tag: 'kern', value: 0 }], font: { families: [resource.face.family], weight: resource.face.weight, style: resource.face.style as 'normal' | 'italic', stretch: 100 }, script: scriptTag(text), language: 'und', direction: 'ltr' as const })
  let segment: ShapedSegment | undefined
  try {
    const shaped = await context.runtime.shaper.shape(Object.freeze({ run, startUtf16: 0, endUtf16: text.length, font: resource }))
    if (!('status' in shaped) && shaped.glyphs.length && shaped.glyphs.every(glyph => glyph.glyphId !== 0)) segment = shaped
  } catch { segment = undefined }
  context.shaped.set(key, segment ?? null)
  return segment
}

function contentFace(face: ResolvedFontFace): NativeDocxContentAddressedFaceV1 { return { face_id: face.faceId, content_digest: face.contentDigest, ...(face.collectionIndex !== undefined ? { collection_index: face.collectionIndex } : {}) } }

async function outlineFor(context: LayoutContext, face: NativeDocxContentAddressedFaceV1, glyphID: number): Promise<NativeDocxGlyphOutlineResultV1 | undefined> {
  const key = `${face.content_digest}\0${face.collection_index ?? ''}\0${glyphID}`
  const cached = context.outlines.get(key)
  if (cached) return cached
  try {
    const live = await context.runtime.outlineProvider.getGlyphOutline(Object.freeze({ face: Object.freeze({ ...face }), glyph_id: glyphID }))
    const captured = nativeDocxCaptureGlyphOutlineV1(structuredClone(live), face, glyphID)
    if (captured) context.outlines.set(key, captured)
    return captured
  } catch { return undefined }
}

function inkBounds(path: NativeDocxGlyphDesignPathCommandV1[]): { xMin: number; xMax: number; yMin: number; yMax: number } | undefined {
  let xMin = Infinity, xMax = -Infinity, yMin = Infinity, yMax = -Infinity
  const point = (x: number, y: number) => { xMin = Math.min(xMin, x); xMax = Math.max(xMax, x); yMin = Math.min(yMin, y); yMax = Math.max(yMax, y) }
  for (const command of path) {
    if (command.kind === 'close_path') continue
    if (command.kind === 'quadratic_to') point(command.control_x, command.control_y)
    if (command.kind === 'cubic_to') { point(command.control_1_x, command.control_1_y); point(command.control_2_x, command.control_2_y) }
    point(command.x, command.y)
  }
  return xMax > xMin && yMax > yMin ? { xMin, xMax, yMin, yMax } : undefined
}

/** Shape one token with one face and measure its ink bounds from outlines. */
async function glyphBox(context: LayoutContext, choice: NativeDocxApproximateMathFaceChoiceV1, text: string, size: number, color: string, cls: AtomClass): Promise<MathBox | undefined> {
  const resource = await resourceFor(context, choice.face)
  if (!resource) return undefined
  const segment = await shapeText(context, resource, text, size)
  if (!segment) return undefined
  const face = contentFace(resource.face)
  const box = emptyBox(cls)
  let x = 0
  let ink = false
  for (const glyph of segment.glyphs) {
    if (context.glyphs >= MAX_GLYPHS) { context.notes.add('equation glyphs beyond the preview glyph budget were dropped'); break }
    const outline = await outlineFor(context, face, glyph.glyphId)
    if (!outline) return undefined
    if (outline.status === 'outlined') {
      const bounds = inkBounds(outline.path)
      if (bounds) {
        const scale = size / outline.units_per_em
        box.ascent = Math.max(box.ascent, Math.round(bounds.yMax * scale) - glyph.offsetYMilliPoints)
        box.depth = Math.max(box.depth, Math.round(-bounds.yMin * scale) + glyph.offsetYMilliPoints)
        ink = true
      }
      box.glyphs.push({ x: x + glyph.offsetXMilliPoints, y: -glyph.offsetYMilliPoints, size, face, glyph_id: glyph.glyphId, color, units_per_em: outline.units_per_em, path: outline.path })
      context.glyphs += 1
    }
    x += glyph.advanceXMilliPoints
  }
  box.width = Math.max(0, x)
  if (!ink) { box.ascent = Math.round(resource.metrics.ascender * size / resource.metrics.unitsPerEm * 0.5); box.depth = 0 }
  return box
}

/** Shape a symbol with the requested face, then with any other loadable face
 * that covers it. Returns undefined when no manifest face has the glyph. */
async function symbolBox(context: LayoutContext, run: NativeDocxApproximateMathRunV1 | undefined, text: string, size: number, cls: AtomClass, upright = true): Promise<MathBox | undefined> {
  const color = run?.color ?? '000000'
  const family = run?.font_family ?? context.textFamily ?? ''
  const weight = run?.bold ? 700 : 400
  const style = upright ? 'normal' : run?.italic ? 'italic' : 'normal'
  const choice = await faceFor(context, family, weight, style)
  if (choice) { const box = await glyphBox(context, choice, text, size, color, cls); if (box) return box }
  for (const face of context.runtime.manifest.faces.filter(loadable)) {
    if (face.faceId === choice?.face.faceId) continue
    const alternative: NativeDocxApproximateMathFaceChoiceV1 = { requested: { family, weight, style }, face, policy: 'text-face-fallback', weight_style_match: false }
    const box = await glyphBox(context, alternative, text, size, color, cls)
    if (box) { context.notes.add(`symbol ${text} (U+${(text.codePointAt(0) ?? 0).toString(16).toUpperCase().padStart(4, '0')}) shaped with loaded face ${face.family} because the math face lacks it`); return box }
  }
  return undefined
}

/** Tokenize a math text run: letters lean (unless m:sty p/b or m:nor), digits
 * and operators stay upright; operator classes drive TeX-style spacing. */
async function textRow(context: LayoutContext, node: NativeDocxApproximateMathNodeV1, style: Style): Promise<MathBox[]> {
  const run = node.run!
  const text = node.text ?? ''
  const size = sizeFor(context, run, style)
  const color = run.color ?? '000000'
  const lean = run.normal ? false : run.style ? run.style === 'i' || run.style === 'bi' : true
  const tokens: Array<{ text: string; cls: AtomClass; italic: boolean }> = []
  for (const ch of text) {
    if (/\s/u.test(ch)) { tokens.push({ text: ch, cls: 'ord', italic: false }); continue }
    if (run.normal) { const last = tokens[tokens.length - 1]; if (last && last.cls === 'ord' && !last.italic && !/\s/u.test(last.text)) last.text += ch; else tokens.push({ text: ch, cls: 'ord', italic: false }); continue }
    if (LETTER.test(ch)) { const last = tokens[tokens.length - 1]; if (last && last.italic === lean && last.cls === 'ord' && LETTER.test([...last.text].at(-1) ?? '')) last.text += ch; else tokens.push({ text: ch, cls: 'ord', italic: lean }); continue }
    if (DIGIT.test(ch)) { const last = tokens[tokens.length - 1]; if (last && !last.italic && last.cls === 'ord' && DIGIT.test([...last.text].at(-1) ?? '')) last.text += ch; else tokens.push({ text: ch, cls: 'ord', italic: false }); continue }
    tokens.push({ text: ch, cls: classify(ch), italic: false })
  }
  const boxes: MathBox[] = []
  for (const token of tokens) {
    const choice = await faceFor(context, run.font_family, run.bold ? 700 : 400, token.italic ? 'italic' : 'normal')
    let box = choice ? await glyphBox(context, choice, token.text, size, color, token.cls) : undefined
    if (!box && token.cls !== 'ord') box = await symbolBox(context, run, token.text, size, token.cls)
    if (!box) {
      if (/\s/u.test(token.text)) { boxes.push({ ...emptyBox(), width: em(context, size, 6) }); continue }
      throw new Error(`no loaded face covers "${token.text}"`)
    }
    boxes.push(box)
  }
  return boxes
}

function spaced(context: LayoutContext, boxes: MathBox[], style: Style, size: number): MathBox {
  const parts: Array<{ box: MathBox; gap: number }> = []
  let previous: AtomClass | undefined
  for (const [index, box] of boxes.entries()) {
    let cls = box.cls
    // A binary operator without a left operand, or following another operator, is ordinary (TeX rule 5).
    if (cls === 'bin' && (previous === undefined || previous === 'bin' || previous === 'op' || previous === 'rel' || previous === 'open' || previous === 'punct')) cls = 'ord'
    const next = boxes[index + 1]
    if (cls === 'bin' && (!next || next.cls === 'rel' || next.cls === 'close' || next.cls === 'punct')) cls = 'ord'
    const units = previous === undefined || style.level > 0 ? 0 : SPACING[previous][cls]
    parts.push({ box: { ...box, cls }, gap: em(context, size, units) })
    previous = cls
  }
  const row = hbox(parts, boxes.length === 1 ? boxes[0]!.cls : 'ord')
  return row
}

function runSize(context: LayoutContext, node: NativeDocxApproximateMathNodeV1, style: Style): number {
  const run = node.run ?? firstRun(node)
  return run ? sizeFor(context, run, style) : Math.max(MIN_FONT_SIZE, Math.round(context.base * SCRIPT_SCALE[style.level]))
}
function firstRun(node: NativeDocxApproximateMathNodeV1): NativeDocxApproximateMathRunV1 | undefined {
  if (node.run) return node.run
  for (const child of node.children ?? []) { const run = firstRun(child); if (run) return run }
  return undefined
}
function controlRun(node: NativeDocxApproximateMathNodeV1): NativeDocxApproximateMathRunV1 | undefined { return node.run ?? firstRun(node) }

async function layoutRow(context: LayoutContext, node: NativeDocxApproximateMathNodeV1, style: Style): Promise<MathBox> {
  const boxes: MathBox[] = []
  for (const child of node.children ?? []) {
    if (child.kind === 'text') boxes.push(...await textRow(context, child, style))
    else boxes.push(await layoutNode(context, child, style))
  }
  if (!boxes.length) return emptyBox()
  return spaced(context, boxes, style, runSize(context, node, style))
}

async function layoutNode(context: LayoutContext, node: NativeDocxApproximateMathNodeV1, style: Style): Promise<MathBox> {
  const children = node.children ?? []
  const size = runSize(context, node, style)
  const control = controlRun(node)
  const color = control?.color ?? '000000'
  const rule = Math.max(1, Math.round(context.rule * size / context.base))
  const axis = Math.round(context.axis * size / context.base)
  const xHeight = Math.round(context.xHeight * size / context.base)
  switch (node.kind) {
    case 'row': return layoutRow(context, node, style)
    case 'text': return spaced(context, await textRow(context, node, style), style, size)
    case 'fraction': {
      const inner = down(style)
      const numerator = await layoutRow(context, children[0]!, style.display ? noDisplay(style) : inner)
      const denominator = await layoutRow(context, children[1]!, style.display ? noDisplay(style) : inner)
      const width = Math.max(numerator.width, denominator.width) + em(context, size, 2)
      const bar = node.bar !== false
      const gap = (style.display ? 3 : 1) * rule
      const thickness = bar ? rule : 0
      const numeratorUp = axis + thickness / 2 + gap + numerator.depth
      const denominatorDown = -(axis - thickness / 2 - gap - denominator.ascent)
      const box = overlay([raised(centered(numerator, width), Math.round(numeratorUp)), raised(centered(denominator, width), -Math.round(denominatorDown))], 'inner')
      if (bar) box.rules.push({ x: em(context, size, 1), y: -Math.round(axis + thickness / 2), width: width - em(context, size, 2), height: thickness, color })
      box.width = width
      return box
    }
    case 'superscript': case 'subscript': case 'subsuperscript': {
      const base = await layoutRow(context, children[0]!, style)
      const scriptStyle = down(style)
      const sub = node.kind === 'superscript' ? undefined : await layoutRow(context, children[1]!, scriptStyle)
      const sup = node.kind === 'subscript' ? undefined : await layoutRow(context, children[node.kind === 'superscript' ? 1 : 2]!, scriptStyle)
      const scriptSize = Math.round(size * SCRIPT_SCALE[scriptStyle.level] / SCRIPT_SCALE[style.level])
      const parts: MathBox[] = [base]
      let up = 0, downShift = 0
      if (sup) up = Math.max(base.ascent - Math.round(scriptSize * 0.35), Math.round(size * (style.display ? 0.42 : 0.36)), sup.depth + Math.round(xHeight * 0.25))
      if (sub) downShift = Math.max(base.depth + Math.round(scriptSize * 0.2), Math.round(size * 0.2), sub.ascent - Math.round(xHeight * 0.8))
      if (sup && sub) { const clearance = (up - sup.depth) - (sub.ascent - downShift); if (clearance < 4 * rule) up += 4 * rule - clearance }
      const kern = em(context, size, 0.5)
      if (sup) parts.push(shifted(raised(sup, up), base.width + kern, 0))
      if (sub) parts.push(shifted(raised(sub, -downShift), base.width + kern, 0))
      const box = overlay(parts, base.cls)
      box.width = base.width + (sup || sub ? kern + Math.max(sup?.width ?? 0, sub?.width ?? 0) : 0)
      return box
    }
    case 'nary': {
      const location = node.limit_location ?? (style.display ? 'undOvr' : 'subSup')
      const operatorSize = style.display ? Math.round(size * 1.4) : size
      const operator = (await symbolBox(context, control, node.chr ?? '∑', operatorSize, 'op')) ?? await symbolBox(context, control, '∑', operatorSize, 'op')
      if (!operator) throw new Error(`no loaded face covers the n-ary operator ${node.chr ?? '∑'}`)
      // Center the operator on the math axis.
      const shift = axis - Math.round((operator.ascent - operator.depth) / 2)
      let core = raised(operator, shift)
      const scriptStyle = down(style)
      const sub = node.sub_hide ? undefined : await layoutRow(context, children[0]!, scriptStyle)
      const sup = node.sup_hide ? undefined : await layoutRow(context, children[1]!, scriptStyle)
      const body = await layoutRow(context, children[2]!, style)
      if (location === 'undOvr' && (sub || sup)) {
        const gap = Math.max(2 * rule, em(context, size, 1.5))
        const width = Math.max(core.width, sub?.width ?? 0, sup?.width ?? 0)
        const parts = [centered(core, width)]
        if (sup) parts.push(raised(centered(sup, width), core.ascent + gap + sup.depth))
        if (sub) parts.push(raised(centered(sub, width), -(core.depth + gap + sub.ascent)))
        core = overlay(parts, 'op')
        core.width = width
      } else if (sub || sup) {
        const scriptSize = Math.round(size * SCRIPT_SCALE[scriptStyle.level] / SCRIPT_SCALE[style.level])
        const parts = [core]
        const up = sup ? Math.max(core.ascent - Math.round(scriptSize * 0.35), sup.depth + Math.round(xHeight * 0.25)) : 0
        const downShift = sub ? Math.max(core.depth + Math.round(scriptSize * 0.2), sub.ascent - Math.round(xHeight * 0.8)) : 0
        if (sup) parts.push(shifted(raised(sup, up), core.width, 0))
        if (sub) parts.push(shifted(raised(sub, -downShift), core.width, 0))
        const width = core.width + Math.max(sup?.width ?? 0, sub?.width ?? 0)
        core = overlay(parts, 'op')
        core.width = width
      }
      const box = hbox([{ box: core, gap: 0 }, { box: body, gap: em(context, size, 3) }], 'inner')
      return box
    }
    case 'delimiter': {
      const inner: MathBox[] = []
      for (const child of children) inner.push(await layoutRow(context, child, style))
      const separator = node.sep_chr ?? '|'
      const open = node.beg_chr ?? '('
      const close = node.end_chr ?? ')'
      const content = overlay(inner, 'inner')
      const target = 2 * Math.max(content.ascent - axis, content.depth + axis) + 2 * rule
      const fence = async (ch: string): Promise<MathBox | undefined> => {
        if (!ch) return undefined
        const probe = await symbolBox(context, control, ch, size, 'ord')
        if (!probe) { context.notes.add(`delimiter ${ch} drawn as a rule because no loaded face covers it`); return { width: rule * 3, ascent: Math.round(target / 2 + axis), depth: Math.round(target / 2 - axis), cls: 'ord', glyphs: [], rules: [{ x: rule, y: -Math.round(target / 2 + axis), width: rule, height: target, color }] } }
        const height = probe.ascent + probe.depth
        const scale = height > 0 ? Math.min(2.5, Math.max(1, target / height)) : 1
        const grown = scale > 1.02 ? (await symbolBox(context, control, ch, Math.min(MAX_FONT_SIZE, Math.round(size * scale)), 'ord')) ?? probe : probe
        // Center the stretched fence on the math axis.
        return raised(grown, axis - Math.round((grown.ascent - grown.depth) / 2))
      }
      const parts: Array<{ box: MathBox; gap: number }> = []
      const opening = await fence(open)
      if (opening) parts.push({ box: { ...opening, cls: 'open' }, gap: 0 })
      for (const [index, box] of inner.entries()) {
        if (index > 0) { const sep = await fence(separator); if (sep) parts.push({ box: sep, gap: em(context, size, 1) }) }
        parts.push({ box, gap: index > 0 && separator ? em(context, size, 1) : 0 })
      }
      const closing = await fence(close)
      if (closing) parts.push({ box: { ...closing, cls: 'close' }, gap: 0 })
      return hbox(parts, 'inner')
    }
    case 'radical': {
      const body = await layoutRow(context, children[1]!, style)
      const clearance = (style.display ? 3 : 1) * rule + rule
      const target = body.ascent + body.depth + clearance + rule
      const probe = await symbolBox(context, control, '√', size, 'ord')
      if (!probe) throw new Error('no loaded face covers the radical sign')
      const height = probe.ascent + probe.depth
      const scale = height > 0 ? Math.min(2.5, Math.max(1, target / height)) : 1
      const sign = scale > 1.02 ? (await symbolBox(context, control, '√', Math.min(MAX_FONT_SIZE, Math.round(size * scale)), 'ord')) ?? probe : probe
      // Align the sign's bottom with the body's bottom, then rule across the body from the sign's top.
      const signShift = -(body.depth - sign.depth)
      const placedSign = raised(sign, signShift)
      const top = Math.max(placedSign.ascent, body.ascent + clearance)
      const gap = em(context, size, 1)
      const parts: MathBox[] = [placedSign, shifted(body, placedSign.width, 0)]
      const box = overlay(parts, 'ord')
      box.rules.push({ x: Math.max(0, placedSign.width - rule), y: -top, width: body.width + rule + gap, height: rule, color })
      box.width = placedSign.width + body.width + gap
      box.ascent = Math.max(box.ascent, top + rule)
      if (!node.degree_hide && (children[0]!.children ?? []).length) {
        const degree = await layoutRow(context, children[0]!, { display: false, level: 2 })
        const raise = Math.round(top * 0.6)
        const placedDegree = raised(degree, raise)
        const shiftX = Math.max(0, placedDegree.width - Math.round(placedSign.width * 0.4))
        return { ...overlay([placedDegree, shifted(box, shiftX, 0)], 'ord'), width: box.width + shiftX }
      }
      return box
    }
    case 'function': {
      const name = await layoutRow(context, children[0]!, style)
      const argument = await layoutRow(context, children[1]!, style)
      return hbox([{ box: { ...name, cls: 'op' }, gap: 0 }, { box: argument, gap: em(context, size, style.level > 0 ? 0 : 3) }], 'inner')
    }
    case 'bar': {
      const body = await layoutRow(context, children[0]!, style)
      const box = { ...body, cls: 'ord' as const, rules: [...body.rules], glyphs: [...body.glyphs] }
      if (node.position === 'top') { box.rules.push({ x: 0, y: -(body.ascent + 3 * rule), width: body.width, height: rule, color }); box.ascent = body.ascent + 4 * rule } else { box.rules.push({ x: 0, y: body.depth + 2 * rule, width: body.width, height: rule, color }); box.depth = body.depth + 3 * rule }
      return box
    }
    case 'accent': {
      const body = await layoutRow(context, children[0]!, style)
      const accent = await symbolBox(context, control, node.chr ?? '̂', size, 'ord')
      if (!accent) { context.notes.add(`accent ${node.chr ?? 'U+0302'} omitted because no loaded face covers it`); return body }
      const raise = body.ascent - Math.min(accent.depth, 0) + rule - Math.max(0, accent.depth) + Math.max(0, accent.depth)
      const placedAccent = raised(centered({ ...accent, width: Math.max(accent.width, 1) }, body.width), Math.max(raise, xHeight))
      return { ...overlay([body, placedAccent], 'ord'), width: body.width }
    }
    case 'limit-lower': case 'limit-upper': {
      const base = await layoutRow(context, children[0]!, style)
      const limit = await layoutRow(context, children[1]!, down(style))
      const width = Math.max(base.width, limit.width)
      const gap = Math.max(2 * rule, em(context, size, 1.5))
      const placed = node.kind === 'limit-upper' ? raised(centered(limit, width), base.ascent + gap + limit.depth) : raised(centered(limit, width), -(base.depth + gap + limit.ascent))
      return { ...overlay([centered(base, width), placed], 'op'), width }
    }
  }
}

function resolvedTextFamily(resolved: NativeDocxResolvedLayoutInputV1, paragraphID: string): string | undefined {
  return resolved.paragraphs.find(paragraph => paragraph.paragraph_id === paragraphID)?.paragraph_mark_properties.font_family
}

/** Lay out every supported equation. Failures omit that equation only. */
export async function layoutNativeDocxApproximateEquationsV1(equations: NativeDocxApproximateEquationsV1, resolved: NativeDocxResolvedLayoutInputV1, runtime: NativeDocxApproximateEquationRuntimeV1): Promise<NativeDocxApproximateEquationLayoutsV1> {
  const result: NativeDocxApproximateEquationLayoutsV1 = { layouts: new Map(), omitted: [], faces: [], notes: [] }
  const faces = new Map<string, NativeDocxApproximateMathFaceChoiceV1 | null>()
  const resources = new Map<string, FontResource>()
  const shaped = new Map<string, ShapedSegment | null>()
  const outlines = new Map<string, NativeDocxGlyphOutlineResultV1>()
  const faceChoices: NativeDocxApproximateMathFaceChoiceV1[] = []
  let glyphs = 0
  for (const equation of equations.items) {
    if (equation.status !== 'supported' || !equation.lines) { result.omitted.push({ id: equation.id, reason: equation.reason ?? 'unsupported' }); continue }
    const run = equation.lines.map(firstRun).find(Boolean)
    if (!run) { result.omitted.push({ id: equation.id, reason: 'equation has no text runs' }); continue }
    const base = run.font_size_half_points * 500
    const notes = new Set<string>()
    const context: LayoutContext = { runtime, base, rule: Math.max(300, Math.round(base * 0.045)), axis: Math.round(base * 0.25), xHeight: Math.round(base * 0.5), textFamily: resolvedTextFamily(resolved, equation.paragraph_id), faces, resources, shaped, outlines, notes, glyphs, faceChoices }
    try {
      // Axis and x-height come from the first resolvable face's design metrics.
      const first = await faceFor(context, run.font_family, run.bold ? 700 : 400, 'normal')
      const resource = first ? await resourceFor(context, first.face) : undefined
      if (resource?.metrics.xHeight) { context.xHeight = Math.round(resource.metrics.xHeight * base / resource.metrics.unitsPerEm); context.axis = Math.round(context.xHeight / 2) }
      const style: Style = { display: equation.display, level: 0 }
      const lines: MathBox[] = []
      for (const line of equation.lines) lines.push(await layoutRow(context, line, style))
      const width = Math.max(1, ...lines.map(line => line.width))
      const justify = (box: MathBox): MathBox => equation.justification === 'left' ? box : equation.justification === 'right' ? shifted(box, width - box.width, 0) : centered(box, width)
      let stacked = justify(lines[0]!)
      let cursor = stacked.depth
      const gap = em(context, base, 6)
      for (const line of lines.slice(1)) { const placed = raised(justify(line), -(cursor + gap + line.ascent)); cursor = placed.depth; stacked = { ...overlay([stacked, placed], 'ord'), width } }
      if (!stacked.glyphs.length && !stacked.rules.length) { result.omitted.push({ id: equation.id, reason: 'equation produced no paint' }); continue }
      const layout: NativeDocxApproximateEquationLayoutV1 = { id: equation.id, width: Math.max(1, Math.round(stacked.width)), ascent: Math.max(1, Math.round(stacked.ascent)), depth: Math.max(0, Math.round(stacked.depth)), glyphs: stacked.glyphs, rules: stacked.rules, notes: [...notes, ...(equation.notes ?? [])] }
      result.layouts.set(equation.id, layout)
      glyphs = context.glyphs
    } catch (error) {
      result.omitted.push({ id: equation.id, reason: `layout failed: ${(error instanceof Error ? error.message : 'unknown').slice(0, 160)}` })
    }
    for (const note of notes) if (!result.notes.includes(note)) result.notes.push(note)
  }
  result.faces = faceChoices
  return result
}

export interface NativeDocxApproximateEquationProjectionV1 {
  document: NativeDocxDocumentV1
  resolved: NativeDocxResolvedLayoutInputV1
  /** Equation id to the synthetic glyphless drawing run id reserved in the body copy. */
  inlineRuns: Map<string, string>
  notes: string[]
}

function emuFor(document: NativeDocxDocumentV1, millipoints: number): number {
  // Reserve exactly the laid-out extent under the inline textbox atom scale.
  let emu = Math.max(1, Math.round(millipoints * EMU_PER_MILLIPOINT))
  for (let attempt = 0; attempt < 4; attempt++) {
    const probe = qualifyNativeDocxInlineTextboxV1(document, 'probe', { id: 'probe', anchor: { part_name: document.source.main_part, path: '/probe', start_byte: 0, end_byte: 1, xml_sha256: `sha256:${'0'.repeat(64)}` }, placement: 'inline', width_emu: emu, height_emu: emu, textbox_text: 'probe', edit_policy: { mode: 'read-only', allowed_operations: [], refusal: { code: 'PROBE', message: 'probe', preservation: 'refuse-mutation' } } } as never)
    if (!probe.ok) break
    if (probe.value.width_millipoints === millipoints) break
    emu += probe.value.width_millipoints < millipoints ? 1 : -1
  }
  return Math.max(1, emu)
}

/** Equal-width column extent of the section owning a body block, in millipoints. */
function columnWidth(document: NativeDocxDocumentV1, paragraphID: string): number | undefined {
  const blockIndices = new Map(document.body.blocks.map((block, index) => [block.id, index]))
  const blockIndex = blockIndices.get(paragraphID) ?? 0
  let section = document.sections[0]
  for (const candidate of document.sections) {
    const start = blockIndices.get(candidate.starts_at_block_id)
    if (start !== undefined && start <= blockIndex) section = candidate
  }
  if (!section) return undefined
  const page = section.page
  const columns = Math.max(1, page.columns)
  const text = page.width_twips - page.margins.left_twips - page.margins.right_twips - page.margins.gutter_twips
  const width = Math.floor((text - page.column_spacing_twips * (columns - 1)) / columns)
  return width > 0 ? width * 50 : undefined
}

/** Reserve laid-out equations as glyphless textbox atoms so surrounding text
 * reflows around them; drop the joined refusals from this internal copy only.
 * Display equations in otherwise empty paragraphs follow their oMathPara
 * justification (Word centers by default). */
export function projectNativeDocxApproximateEquationsV1(document: NativeDocxDocumentV1, resolved: NativeDocxResolvedLayoutInputV1, equations: NativeDocxApproximateEquationsV1, layouts: NativeDocxApproximateEquationLayoutsV1): NativeDocxApproximateEquationProjectionV1 {
  const projected = structuredClone(document)
  const projectedResolved = structuredClone(resolved)
  const inlineRuns = new Map<string, string>()
  const removed = new Set<string>()
  const notes: string[] = []
  const paragraphs = new Map(projected.body.blocks.flatMap(block => block.paragraph ? [[block.id, block.paragraph] as const] : []))
  for (const equation of equations.items) {
    const layout = layouts.layouts.get(equation.id)
    if (equation.status !== 'supported' || !layout) continue
    const paragraph = paragraphs.get(equation.paragraph_id)
    if (!paragraph) continue
    for (const id of equation.diagnostic_ids) removed.add(id)
    let width = layout.width
    const column = columnWidth(projected, equation.paragraph_id)
    if (column !== undefined && width > column) { width = column; notes.push(`${equation.id}: equation width clamped to its column width`) }
    const height = layout.ascent + layout.depth
    const runID = `${equation.id}:run`
    const run: NativeDocxRunV1 = {
      kind: 'drawing', id: runID, anchor: equation.anchor,
      drawing: {
        id: `${equation.id}:drawing`, anchor: equation.anchor, placement: 'inline', width_emu: emuFor(projected, width), height_emu: emuFor(projected, height),
        // The atom carries no painted text; the id keeps the required non-empty marker honest.
        textbox_text: equation.id, textbox_fill_rgb: 'FFFFFF', textbox_line_rgb: '000000',
        edit_policy: { mode: 'read-only', allowed_operations: [], refusal: { code: 'APPROXIMATE_EQUATION_PREVIEW', message: 'Equation reserved for read-only approximate preview', preservation: 'refuse-mutation' } },
      },
    }
    let index = paragraph.runs.findIndex(existing => existing.anchor.start_byte > equation.anchor.start_byte)
    if (index < 0) index = paragraph.runs.length
    paragraph.runs.splice(index, 0, run)
    projectedResolved.runs.push({ run_id: runID, paragraph_id: paragraph.id, applied_paragraph_styles: [], applied_character_styles: [], properties: {} })
    inlineRuns.set(equation.id, runID)
    if (equation.display && paragraph.runs.every(existing => existing.kind === 'drawing' && inlineRuns.has(existing.id.replace(/:run$/, '')))) {
      const resolvedParagraph = projectedResolved.paragraphs.find(entry => entry.paragraph_id === paragraph.id)
      const alignment = equation.justification === 'left' ? 'left' : equation.justification === 'right' ? 'right' : 'center'
      if (resolvedParagraph && resolvedParagraph.properties.alignment !== alignment) { resolvedParagraph.properties.alignment = alignment; notes.push(`${equation.id}: display equation paragraph aligned ${alignment} per oMathPara justification`) }
    }
  }
  projected.unsupported = projected.unsupported.filter(entry => !removed.has(entry.id))
  return { document: projected, resolved: projectedResolved, inlineRuns, notes }
}

export interface NativeDocxApproximateEquationPaintResultV1 {
  painted: string[]
  omitted: Array<{ id: string; reason: string }>
  reasons: string[]
}

function findHighlight(pages: NativeDocxPaintPageV1[], runID: string): { page: NativeDocxPaintPageV1; line: NativeDocxPaintLineV1; command: NativeDocxFillTextHighlightCommandV1 } | undefined {
  for (const page of pages) {
    const command = page.commands.find((entry): entry is NativeDocxFillTextHighlightCommandV1 => entry.kind === 'fill_text_highlight' && entry.source_id === runID)
    if (!command) continue
    const line = page.lines.find(entry => entry.command_ids.includes(command.id))
    if (line) return { page, line, command }
  }
  return undefined
}

/** Rebuild page paint order: behind floats, table fills, the line-owned
 * commands in line order, table borders, front floats. Every original command
 * keeps its category; only equation paint is added or the reservation removed. */
function rebuildCommands(page: NativeDocxPaintPageV1, extra: Map<string, NativeDocxPagePaintCommandV1>, rules: NativeDocxPagePaintCommandV1[], removed: Set<string>): void {
  const byID = new Map<string, NativeDocxPagePaintCommandV1>()
  const behindFloats: NativeDocxPagePaintCommandV1[] = [], frontFloats: NativeDocxPagePaintCommandV1[] = [], fills: NativeDocxPagePaintCommandV1[] = [], borders: NativeDocxPagePaintCommandV1[] = []
  for (const command of page.commands) {
    if (removed.has(command.id)) continue
    if (command.kind === 'paint_floating_image') (command.layer === 'behind' ? behindFloats : frontFloats).push(command)
    else if (command.kind === 'fill_table_cell') fills.push(command)
    else if (command.kind === 'stroke_table_border') borders.push(command)
    else byID.set(command.id, command)
  }
  for (const [id, command] of extra) byID.set(id, command)
  const ordinary: NativeDocxPagePaintCommandV1[] = []
  for (const line of page.lines) for (const id of line.command_ids) { const command = byID.get(id); if (command) ordinary.push(command) }
  page.commands = [...behindFloats, ...fills, ...ordinary, ...borders, ...rules, ...frontFloats]
}

/** Append approximate equation paint to already painted approximate pages at
 * the reserved atom positions. Pages are mutated in place; the caller
 * re-validates the whole envelope. */
export function paintNativeDocxApproximateEquationsV1(paint: Pick<NativeDocxPagePaintSuccessV1, 'pages'>, equations: NativeDocxApproximateEquationsV1, layouts: NativeDocxApproximateEquationLayoutsV1, projection: NativeDocxApproximateEquationProjectionV1): NativeDocxApproximateEquationPaintResultV1 {
  const result: NativeDocxApproximateEquationPaintResultV1 = { painted: [], omitted: [...layouts.omitted], reasons: [] }
  const extraByPage = new Map<string, Map<string, NativeDocxPagePaintCommandV1>>()
  const rulesByPage = new Map<string, NativeDocxPagePaintCommandV1[]>()
  const removed = new Set<string>()
  let byteBudget = MAX_ENVELOPE_BYTES - JSON.stringify(paint.pages).length
  let droppedGlyphs = 0
  for (const equation of equations.items) {
    const layout = layouts.layouts.get(equation.id)
    if (equation.status !== 'supported' || !layout) continue
    const runID = projection.inlineRuns.get(equation.id)
    const found = runID ? findHighlight(paint.pages, runID) : undefined
    if (!found) { result.omitted.push({ id: equation.id, reason: 'equation-not-placed' }); continue }
    const { page, line, command } = found
    // The reservation highlight only located the atom; the equation paints itself.
    removed.add(command.id)
    line.command_ids = line.command_ids.filter(id => id !== command.id)
    const originX = command.x_millipoints
    const baselineY = command.y_millipoints + command.height_millipoints - layout.depth
    const extra = extraByPage.get(page.id) ?? new Map<string, NativeDocxPagePaintCommandV1>()
    extraByPage.set(page.id, extra)
    const fragmentID = `${equation.id}:eq`
    const maxX = page.width_millipoints, maxY = page.height_millipoints
    let glyphIndex = 0
    for (const glyph of layout.glyphs) {
      const path = nativeDocxPlaceGlyphPathV1(glyph.path, Math.round(originX + glyph.x), Math.round(baselineY + glyph.y), glyph.size, glyph.units_per_em)
      if (!path) continue
      const paintCommand: NativeDocxFillGlyphPathCommandV1 = { kind: 'fill_glyph_path', id: paintCommandID(line.placed_line_id, fragmentID, glyphIndex), line_id: line.line_id, fragment_id: fragmentID, source_id: runID!, glyph_index: glyphIndex, face: glyph.face, glyph_id: glyph.glyph_id, font_size_millipoints: glyph.size, fill_rgb: glyph.color, fill_rule: 'nonzero', outline_kind: 'path', path }
      const bytes = JSON.stringify(paintCommand).length + 1
      if (bytes > byteBudget) { droppedGlyphs += 1; continue }
      byteBudget -= bytes
      glyphIndex += 1
      extra.set(paintCommand.id, paintCommand)
      line.command_ids.push(paintCommand.id)
    }
    const rules = rulesByPage.get(page.id) ?? []
    rulesByPage.set(page.id, rules)
    for (const [ruleIndex, rule] of layout.rules.entries()) {
      const x = Math.min(Math.max(Math.round(originX + rule.x), 0), maxX), y = Math.min(Math.max(Math.round(baselineY + rule.y), 0), maxY)
      const width = Math.min(Math.max(1, Math.round(rule.width)), Math.max(1, maxX - x)), height = Math.min(Math.max(1, Math.round(rule.height)), Math.max(1, maxY - y))
      const fill: NativeDocxFillTableCellCommandV1 = { kind: 'fill_table_cell', id: `${equation.id}:rule:${ruleIndex}`, table_id: DOCX_APPROXIMATE_EQUATION_TABLE_ID, row_id: equation.id, cell_id: `rule:${ruleIndex}`, x_millipoints: x, y_millipoints: y, width_millipoints: width, height_millipoints: height, fill_rgb: rule.color }
      rules.push(fill)
    }
    result.painted.push(equation.id)
  }
  for (const page of paint.pages) if (extraByPage.has(page.id) || rulesByPage.has(page.id) || removed.size) rebuildCommands(page, extraByPage.get(page.id) ?? new Map(), rulesByPage.get(page.id) ?? [], removed)
  result.reasons = buildReasons(equations, layouts, projection, result, droppedGlyphs)
  return result
}

function buildReasons(equations: NativeDocxApproximateEquationsV1, layouts: NativeDocxApproximateEquationLayoutsV1, projection: NativeDocxApproximateEquationProjectionV1, result: NativeDocxApproximateEquationPaintResultV1, droppedGlyphs: number): string[] {
  const reasons: string[] = []
  if (result.painted.length) {
    const notes = new Set<string>([...layouts.notes, ...projection.notes])
    for (const id of result.painted) for (const note of layouts.layouts.get(id)?.notes ?? []) notes.add(note)
    reasons.push(DOCX_APPROXIMATE_EQUATION_WARNING)
    reasons.push(`${DOCX_APPROXIMATE_EQUATION_CODE}: painted ${result.painted.length} of ${equations.items.length} refused OMML equations (${equations.items.filter(equation => result.painted.includes(equation.id)).map(equation => `${equation.id} ${equation.display ? 'display' : 'inline'} for diagnostics ${equation.diagnostic_ids.join(',')}`).join('; ')})${notes.size ? `; approximations: ${[...notes].join('; ')}` : ''}`.slice(0, 8000))
    const seen = new Set<string>()
    for (const choice of layouts.faces) {
      const key = `${asciiLowerNative(choice.requested.family)}\0${choice.requested.weight}\0${choice.requested.style}`
      if (seen.has(key)) continue
      seen.add(key)
      if (reasons.length >= MAX_REASONS - 2) break
      reasons.push(`${DOCX_APPROXIMATE_EQUATION_FONT_CODE}: ${choice.requested.family} / ${choice.requested.weight} / ${choice.requested.style} -> ${choice.face.family} / ${choice.face.weight} / ${choice.face.style} (${choice.policy}${choice.weight_style_match ? '' : ', weight/style differ'}; loaded face ${choice.face.faceId})${choice.policy === 'exact' ? '' : '; no authored math face is loaded, glyph shapes and metrics differ from Word'}`)
    }
  }
  if (result.omitted.length || equations.omitted_count || droppedGlyphs) reasons.push(`${DOCX_APPROXIMATE_EQUATION_OMITTED_CODE}: ${result.omitted.map(entry => `${entry.id} (${entry.reason})`).join('; ')}${equations.omitted_count ? `; ${equations.omitted_count} equations beyond the ${MAX_EQUATIONS} equation budget` : ''}${droppedGlyphs ? `; ${droppedGlyphs} glyphs beyond the preview size budget` : ''}`.slice(0, 8000))
  return reasons.slice(0, MAX_REASONS)
}

/** Re-derive the omitted-content disclosure after equation paint changed page
 * commands; pages that only gained equation paint are no longer blank. */
export function discloseNativeDocxApproximateEquationOmissionsV1(result: NativeDocxApproximateOmissionsV1 & { reasons: string[]; status: 'painted' | 'refused'; pages: NativeDocxPaintPageV1[] }, source: { document: NativeDocxDocumentV1; resolved_layout: NativeDocxResolvedLayoutInputV1; shaped_lines: NativeDocxShapedLinesV1 }): void {
  const omissions = collectNativeDocxApproximateOmissionsV1(source, result)
  result.content_status = omissions.content_status
  result.omitted_content = omissions.omitted_content
  result.omitted_content_total = omissions.omitted_content_total
  result.unpainted_pages = omissions.unpainted_pages
  const needsWarning = omissions.omitted_content.length > 0 || omissions.unpainted_pages.length > 0
  const has = result.reasons.includes(DOCX_APPROXIMATE_OMITTED_CONTENT_WARNING)
  if (needsWarning && !has) result.reasons.push(DOCX_APPROXIMATE_OMITTED_CONTENT_WARNING)
  if (!needsWarning && has) result.reasons = result.reasons.filter(reason => reason !== DOCX_APPROXIMATE_OMITTED_CONTENT_WARNING)
}

export interface NativeDocxApproximateEquationStageV1 {
  equations: NativeDocxApproximateEquationsV1
  layouts: NativeDocxApproximateEquationLayoutsV1
  projection: NativeDocxApproximateEquationProjectionV1
}

/** Compiler entry before body pagination: validate the same-bytes sidecar
 * against the source document, lay every supported equation out with the
 * declared math face policy, and reserve the extents in an internal body copy. */
export async function prepareNativeDocxApproximateEquationStageV1(sidecar: unknown, document: NativeDocxDocumentV1, resolved: NativeDocxResolvedLayoutInputV1, runtime: NativeDocxApproximateEquationRuntimeV1): Promise<NativeDocxApproximateEquationStageV1> {
  const equations = decodeNativeDocxApproximateEquationsV1(sidecar, document)
  const layouts = await layoutNativeDocxApproximateEquationsV1(equations, resolved, runtime)
  const projection = projectNativeDocxApproximateEquationsV1(document, resolved, equations, layouts)
  return { equations, layouts, projection }
}

/** Compiler entry after body pagination: paint the laid-out equations at their
 * reserved atoms, append the declared reasons and re-derive omitted content. */
export function completeNativeDocxApproximateEquationStageV1(result: NativeDocxApproximateOmissionsV1 & { reasons: string[]; status: 'painted' | 'refused'; pages: NativeDocxPaintPageV1[] }, stage: NativeDocxApproximateEquationStageV1, source: { document: NativeDocxDocumentV1; resolved_layout: NativeDocxResolvedLayoutInputV1; shaped_lines: NativeDocxShapedLinesV1 }): NativeDocxApproximateEquationPaintResultV1 {
  const painted = paintNativeDocxApproximateEquationsV1(result, stage.equations, stage.layouts, stage.projection)
  for (const reason of painted.reasons) if (!result.reasons.includes(reason) && result.reasons.length < 260) result.reasons.push(reason)
  discloseNativeDocxApproximateEquationOmissionsV1(result, source)
  return painted
}

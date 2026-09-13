import { snapshotNativePlainData } from './nativePlainData.js'
import type { NativeWorkbookV2 } from './nativeContractV2.generated.js'
import type { NativeWorkbookObjectsV1 } from './nativeObjectsPreviewV1.js'

export interface NativeRichTextRunV1 {
  text: string
  properties: 'direct' | 'cell-inherited'
  font_scheme?: 'major' | 'minor'
  declared_font_name?: string
  theme_part?: string
  theme_sha256?: string
  font_name?: string
  font_size_points?: number
  font_color?: string
  bold?: boolean
  italic?: boolean
  underline?: 'single' | 'none'
  underline_origin?: 'explicit-val' | 'default-val'
  omitted: string[]
}
export interface NativeRichTextCellV1 {
  sheet_id: string; sheet_part: string; row: number; column: number; ref: string; style_id: number
  storage: 'inline' | 'shared'; source_part: string; shared_index: string; text: string
  status: 'available' | 'omitted'; warnings: string[]; runs?: NativeRichTextRunV1[]
}
export interface NativeRichTextPreviewV1 { cells: NativeRichTextCellV1[]; warnings: string[] }
const PROPERTIES = ['font_name', 'font_size_points', 'font_color', 'bold', 'italic'] as const
const SCHEME = ['font_scheme', 'declared_font_name', 'theme_part', 'theme_sha256'] as const
const OMITTED = ['font-family-hint', 'baseline', 'underline-none', 'strike-false']
const safeText = (s: string) => !/[\u0000-\u001f\u007f\u2028\u2029]/.test(s)
const path = (s: unknown): s is string => typeof s === 'string' && s.length > 0 && s.length <= 1024 && safeText(s) && !s.startsWith('/') && !s.includes('\\') && !s.split('/').some(p => !p || p === '.' || p === '..')
const int = (v: unknown, max: number): v is number => Number.isSafeInteger(v) && !Object.is(v, -0) && Number(v) >= 0 && Number(v) <= max
const ref = (row: number, column: number) => { let s = ''; for (let c = column + 1; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + (c - 1) % 26) + s; return s + String(row + 1) }
/** Closed, bounded additive data. Does not change native style qualification. */
export function decodeNativeRichTextPreviewV1(input: unknown): NativeRichTextPreviewV1 {
  const fail = (): never => { throw new TypeError('Invalid rich-text preview') }
  const exact = (v: unknown, required: string[], optional: string[] = []) => {
    if (!v || typeof v !== 'object' || Array.isArray(v)) return fail()
    const o = v as Record<string, unknown>
    if (required.some(k => !Object.hasOwn(o, k)) || Object.keys(o).some(k => !required.includes(k) && !optional.includes(k))) return fail()
    return o
  }
  const warnings = (v: unknown, empty = false): string[] => {
    if (!Array.isArray(v) || v.length < (empty ? 0 : 1) || v.length > 8 || v.some(s => typeof s !== 'string' || !s || s.length > 1024 || !safeText(s))) return fail()
    return v as string[]
  }
  const o = exact(snapshotNativePlainData(input, { maxDepth: 8, maxNodes: 30000 }), ['cells', 'warnings'])
  if (!Array.isArray(o.cells) || o.cells.length > 256) return fail()
  let runCount = 0, textCount = 0
  const ids = new Set<string>(), sheets = new Map<string, string>(), parts = new Map<string, string>()
  const cells = o.cells.map(raw => {
    const c = exact(raw, ['sheet_id', 'sheet_part', 'row', 'column', 'ref', 'style_id', 'storage', 'source_part', 'shared_index', 'text', 'status', 'warnings'], ['runs'])
    if (typeof c.sheet_id !== 'string' || !/^[1-9][0-9]{0,9}$/.test(c.sheet_id) || Number(c.sheet_id) > 0xffffffff || !path(c.sheet_part) || !path(c.source_part) || !int(c.row, 1048575) || !int(c.column, 16383) || c.ref !== ref(c.row, c.column) || !int(c.style_id, 65535)) return fail()
    if (sheets.has(c.sheet_id) && sheets.get(c.sheet_id) !== c.sheet_part || parts.has(c.sheet_part) && parts.get(c.sheet_part) !== c.sheet_id) return fail()
    sheets.set(c.sheet_id, c.sheet_part); parts.set(c.sheet_part, c.sheet_id)
    const id = `${c.sheet_id}:${c.ref}`; if (ids.has(id)) return fail(); ids.add(id)
    if (typeof c.text !== 'string' || c.text.length > 2048 || (textCount += c.text.length) > 32768 || typeof c.shared_index !== 'string') return fail()
    if (c.storage === 'inline' ? c.source_part !== c.sheet_part || c.shared_index !== '' : c.storage !== 'shared' || !/^[0-9]{1,128}$/.test(c.shared_index) || c.status === 'available' && !/^(0|[1-9][0-9]{0,8})$/.test(c.shared_index)) return fail()
    const base = { ...c, warnings: warnings(c.warnings) } as unknown as NativeRichTextCellV1
    if (c.status === 'omitted') { if (Object.hasOwn(c, 'runs')) return fail(); return base }
    if (c.status !== 'available' || !safeText(c.text) || !Array.isArray(c.runs) || !c.runs.length || c.runs.length > 64 || (runCount += c.runs.length) > 1024) return fail()
    const runs = c.runs.map(raw => {
      const r = exact(raw, ['text', 'properties', 'omitted'], [...PROPERTIES, ...SCHEME, 'underline', 'underline_origin'])
      if (typeof r.text !== 'string' || !r.text || !safeText(r.text) || r.text.length > 2048 || !['direct', 'cell-inherited'].includes(String(r.properties)) || !Array.isArray(r.omitted) || r.omitted.length > 4 || new Set(r.omitted).size !== r.omitted.length || r.omitted.some(v => !OMITTED.includes(String(v)))) return fail()
      const hasScheme = SCHEME.some(k => Object.hasOwn(r, k))
      if (hasScheme && (SCHEME.some(k => !Object.hasOwn(r, k)) || r.properties !== 'direct' || !['major', 'minor'].includes(String(r.font_scheme)) || typeof r.font_scheme !== 'string' || typeof r.font_name !== 'string' || typeof r.declared_font_name !== 'string' || !/^[A-Za-z0-9][A-Za-z0-9 ._-]{0,127}$/.test(r.declared_font_name) || !path(r.theme_part) || typeof r.theme_sha256 !== 'string' || !/^sha256:[a-f0-9]{64}$/.test(r.theme_sha256) || !/^[\x20-\x7e]+$/.test(r.text))) return fail()
      const hasUnderline = Object.hasOwn(r, 'underline')
      if (hasUnderline !== Object.hasOwn(r, 'underline_origin') || hasUnderline && (typeof r.underline !== 'string' || typeof r.underline_origin !== 'string' || !['single', 'none'].includes(String(r.underline)) || !['explicit-val', 'default-val'].includes(String(r.underline_origin)) || r.underline_origin === 'default-val' && r.underline !== 'single' || r.omitted.includes('underline-none'))) return fail()
      if (r.properties === 'cell-inherited' && (hasUnderline || PROPERTIES.some(k => Object.hasOwn(r, k)) || r.omitted.length)) return fail()
      if (Object.hasOwn(r, 'font_name') && (typeof r.font_name !== 'string' || !/^[A-Za-z0-9][A-Za-z0-9 ._-]{0,127}$/.test(r.font_name))) return fail()
      if (Object.hasOwn(r, 'font_color') && (typeof r.font_color !== 'string' || !/^#[0-9A-F]{6}$/.test(r.font_color))) return fail()
      if (Object.hasOwn(r, 'font_size_points') && (typeof r.font_size_points !== 'number' || !Number.isFinite(r.font_size_points) || r.font_size_points < 1 || r.font_size_points > 409)) return fail()
      if (['bold', 'italic'].some(k => Object.hasOwn(r, k) && typeof r[k] !== 'boolean')) return fail()
      return r as unknown as NativeRichTextRunV1
    })
    if (runs.map(r => r.text).join('') !== c.text) return fail()
    return { ...base, runs }
  })
  return { cells, warnings: warnings(o.warnings, true) }
}

/** Require current package, worksheet, cell, string index, style and run joins.
 * Underline is a raw inspector attestation; the extracted model does not
 * independently represent or authenticate that property.
 * Unknown/unsupported raw formatting stays an explicit plaintext omission. */
export function selectNativeRichTextPreviewV1(workbook: NativeWorkbookV2, sheetId: string, objects: NativeWorkbookObjectsV1): NativeRichTextPreviewV1 {
  const fail = (): never => { throw new TypeError('Rich-text preview does not join the opened source') }
  if (objects.protocol !== 'injoffice.xlsx.preview-objects' || objects.version !== 1 || !/^sha256:[a-f0-9]{64}$/.test(workbook.source.package_sha256) || objects.package_sha256 !== workbook.source.package_sha256) return fail()
  const sheets = workbook.sheets.filter(s => s.id === sheetId); if (sheets.length !== 1) return fail()
  const sheet = sheets[0]!, preview = decodeNativeRichTextPreviewV1(objects.rich_text ?? { cells: [], warnings: [] })
  const entries = preview.cells.filter(c => c.sheet_id === sheetId || c.sheet_part === sheet.part_name)
  const cells = new Map<string, typeof sheet.cells[number]>()
  for (const c of sheet.cells) { if (c.ref !== ref(c.row, c.column) || cells.has(c.ref)) return fail(); cells.set(c.ref, c) }
  for (const e of entries) {
    const c = cells.get(e.ref), v = c?.value
    if (e.sheet_id !== sheetId || e.sheet_part !== sheet.part_name || !c || e.row !== c.row || e.column !== c.column || e.style_id !== c.style_id || v?.kind !== 'string' || !v.rich || v.text !== e.text || v.storage !== e.storage || (v.storage === 'shared' && v.lexical !== e.shared_index)) return fail()
    if (e.status === 'omitted') continue
    const style = workbook.styles.find(s => s.id === c.style_id)?.effective
    if (workbook.unsupported?.some(u => u.part_name === sheet.part_name && u.cell_ref === c.ref && u.code !== 'RICH_CELL_STRING') || c.formula || sheet.cells.some(c => c.formula && c.formula.type !== 'normal') || objects.tables.some(t => t.sheet_part === sheet.part_name) || sheet.merged_ranges.some(m => c.row >= m.row && c.row <= m.end_row && c.column >= m.column && c.column <= m.end_column) || !style || style.projection !== 'full' || style.unsupported.length || style.wrap_text || style.shrink_to_fit || style.text_rotation) return fail()
    if (e.runs?.length !== v.runs?.length) return fail()
    e.runs!.forEach((r, i) => { const original = v.runs![i]!; if (r.text !== original.text || PROPERTIES.some(k => Object.hasOwn(r, k) !== Object.hasOwn(original, k) || r[k] !== original[k])) return fail() })
  }
  return { cells: entries, warnings: preview.warnings }
}

/** Every span starts from this cell's base, never the preceding run. */
export function nativeRichTextRunDisclosureV1(run: NativeRichTextRunV1): string {
  const missing = PROPERTIES.filter(k => !Object.hasOwn(run, k))
  const scheme = run.font_scheme ? `; Source ${run.font_scheme} scheme: declared ${run.declared_font_name}, resolved theme Latin face ${run.font_name} from ${run.theme_part} (${run.theme_sha256}); printable ASCII only` : ''
  const underline = run.underline ? `Underline ${run.underline} (${run.underline_origin === 'default-val' ? 'present u, schema-default single' : 'explicit source val'}); browser decoration metrics are approximate` : run.omitted.includes('underline-none') ? 'Explicit source underline none (legacy evidence)' : 'No direct underline declaration; undecorated host fallback'
  return `${run.properties === 'cell-inherited' ? 'No rPr: cell font inherited' : `Direct properties; missing ${missing.join(', ') || 'none'} use approximate cell-font fallback`}${run.omitted.length ? `; declared ${run.omitted.join(', ')}` : ''}${scheme}; ${underline}. Host font matching and shaping are approximate.`
}

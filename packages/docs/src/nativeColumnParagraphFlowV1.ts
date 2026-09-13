/** Bounded, source-bound whole-paragraph flow across two unequal columns. */
import type { NativeDocxDocumentV1, NativeDocxParagraphV1 } from './nativeContract.js'
import type { NativeDocxResolvedLayoutInputV1 } from './nativeResolvedLayout.js'
import type { NativeDocxPaginationSettingsV1 } from './nativePaginationSettings.js'
import type { NativeDocxShapedLinesV1 } from './nativeShapingLines.js'
import { qualifyNativeDocxSectionColumnsV1, type NativeDocxQualifiedSectionGeometryV1 } from './nativeSectionColumnsV1.js'

export interface NativeDocxColumnParagraphProfileV1 {
  geometry: NativeDocxQualifiedSectionGeometryV1
  paragraphs: NativeDocxParagraphV1[]
}
export interface NativeDocxColumnParagraphFlowV1 extends NativeDocxColumnParagraphProfileV1 {
  shaped_lines: NativeDocxShapedLinesV1
  placements: Array<{ paragraph_id: string; page_ordinal: number; column_ordinal: number }>
}

export function qualifyNativeDocxColumnParagraphProfileV1(document: NativeDocxDocumentV1, resolved: NativeDocxResolvedLayoutInputV1, settings: NativeDocxPaginationSettingsV1): NativeDocxColumnParagraphProfileV1 | undefined {
  if (settings.profile !== 'word-modern-default' || settings.no_column_balance !== true || settings.even_and_odd_headers || settings.mirror_margins || settings.gutter_at_top || settings.diagnostics.length) return undefined
  if (document.sections.length !== 1 || document.headers.length || document.footers.length || document.notes.length || document.comments.length || document.comment_stories.length || resolved.tables.length || resolved.numbering_source) return undefined
  const section = document.sections[0]!
  if (section.page.columns !== 2 || section.page.column_layout !== 'explicit' || section.header_refs.length || section.footer_refs.length || section.title_page || section.page_number_start !== undefined || section.break_type !== 'next-page') return undefined
  const geometry = qualifyNativeDocxSectionColumnsV1(section, { allowUnequalWidths: true })
  if (!geometry.ok || geometry.value.columns[0]!.width_millipoints === geometry.value.columns[1]!.width_millipoints) return undefined
  const paragraphs: NativeDocxParagraphV1[] = []
  if (!document.body.blocks.length || document.body.blocks.length > 256 || section.starts_at_block_id !== document.body.blocks[0]!.id) return undefined
  let sourceTextLength = 0
  for (const block of document.body.blocks) {
    if (block.kind !== 'paragraph' || !block.paragraph) return undefined
    const paragraph = block.paragraph
    sourceTextLength += paragraph.runs.reduce((sum, run) => sum + (run.text ?? '').length, 0)
    if (sourceTextLength > 100000) return undefined
    const properties = resolved.paragraphs.find((entry) => entry.paragraph_id === paragraph.id)
    if (!properties || properties.numbering || paragraph.properties.numbering || properties.properties.bidi || properties.properties.keep_next || properties.properties.page_break_before) return undefined
    if ((properties.properties.line_rule ?? 'auto') !== 'auto' || (properties.properties.line ?? 240) !== 240) return undefined
    if ((properties.properties.spacing_before_twips ?? 0) !== 0 || (properties.properties.spacing_after_twips ?? 0) !== 0) return undefined
    if (!paragraph.runs.length || paragraph.runs.some((run) => run.kind !== 'text' || run.page_field !== undefined || run.layout_page_field !== undefined || !/^[\x20-\x7e]*$/.test(run.text ?? ''))) return undefined
    if (!paragraph.runs.some((run) => (run.text ?? '').length > 0)) return undefined
    paragraphs.push(paragraph)
  }
  if (resolved.paragraphs.length !== paragraphs.length || resolved.runs.some((run) => run.properties.rtl || run.properties.hidden)) return undefined
  // Keep the extractor's diagnostic in the immutable source. Only this exact,
  // section-local record is eligible for discharge after both candidates pass.
  if (document.unsupported.length !== 1) return undefined
  const diagnostic = document.unsupported[0]!
  if (diagnostic.code !== 'UNEQUAL_SECTION_COLUMNS' || diagnostic.capability !== 'sections' || diagnostic.scope_id !== section.id || diagnostic.preservation !== 'refuse-mutation' || !diagnostic.anchor || JSON.stringify(diagnostic.anchor) !== JSON.stringify(section.anchor) || diagnostic.message !== 'Unequal column widths require per-column shaping outside the exact v1 slice') return undefined
  return { geometry: geometry.value, paragraphs }
}

export function planNativeDocxColumnParagraphFlowV1(document: NativeDocxDocumentV1, resolved: NativeDocxResolvedLayoutInputV1, settings: NativeDocxPaginationSettingsV1, candidates: readonly NativeDocxShapedLinesV1[]): NativeDocxColumnParagraphFlowV1 | undefined {
  const profile = qualifyNativeDocxColumnParagraphProfileV1(document, resolved, settings)
  if (!profile || candidates.length !== 2) return undefined
  const base = candidates[0]!
  const canonical = (value: unknown): string => JSON.stringify(value)
  const runFontIdentities = new Map<string, string>()
  const sourceScripts = new Map<string, string>()
  for (const [ordinal, candidate] of candidates.entries()) {
    if (candidate.document_id !== document.document_id || candidate.revision !== document.revision || candidate.available_width_millipoints !== profile.geometry.columns[ordinal]!.width_millipoints || candidate.tab_interval_millipoints !== settings.default_tab_stop_twips * 50 || candidate.numbering_source || candidate.font_substitutions?.length) return undefined
    if (canonical(candidate.font_manifest) !== canonical(base.font_manifest) || canonical(candidate.providers) !== canonical(base.providers) || canonical(candidate.diagnostics) !== canonical(base.diagnostics) || candidate.paragraphs.length !== profile.paragraphs.length) return undefined
    for (const [index, paragraph] of candidate.paragraphs.entries()) {
      const source = profile.paragraphs[index]!
      const properties = resolved.paragraphs.find((entry) => entry.paragraph_id === source.id)!.properties
      if (paragraph.paragraph_id !== source.id || paragraph.story_id !== document.body.id || paragraph.story_kind !== 'body' || paragraph.direction !== 'ltr' || paragraph.list_marker || paragraph.spacing_before_millipoints !== 0 || paragraph.spacing_after_millipoints !== 0 || paragraph.indent_start_millipoints !== 0 || paragraph.indent_end_millipoints !== 0 || paragraph.first_line_delta_millipoints !== 0 || paragraph.lines.length === 0 || paragraph.lines.length > 256) return undefined
      if (paragraph.lines.length > 1 && properties.keep_lines !== true) return undefined
      if (paragraph.lines.some((line) => line.available_width_millipoints !== candidate.available_width_millipoints || line.hard_break_after || line.exclusion_start_millipoints !== undefined || line.advance_inline_millipoints > line.available_width_millipoints)) return undefined
      const starts = new Map<string, number>()
      let sourceLength = 0
      for (const run of source.runs) { starts.set(run.id, sourceLength); sourceLength += (run.text ?? '').length }
      let covered = 0
      for (const line of paragraph.lines) {
        if (!line.fragments.length || line.ascent_millipoints !== Math.max(...line.fragments.map((fragment) => fragment.ascent_millipoints)) || line.descent_millipoints !== Math.min(...line.fragments.map((fragment) => fragment.descent_millipoints)) || line.line_gap_millipoints !== Math.max(...line.fragments.map((fragment) => fragment.line_gap_millipoints)) || line.line_height_millipoints !== line.ascent_millipoints - line.descent_millipoints + line.line_gap_millipoints) return undefined
        for (const fragment of line.fragments) {
        const run = source.runs.find((entry) => entry.id === fragment.source_id)
        if (!run || fragment.source_kind !== 'run' || fragment.direction !== 'ltr' || fragment.bidi_level !== 0 || fragment.start_utf16 >= fragment.end_utf16 || starts.get(run.id)! + fragment.start_utf16 !== covered || run.text!.slice(fragment.start_utf16, fragment.end_utf16) !== fragment.text || !fragment.face_id || fragment.script_transform) return undefined
        covered += fragment.end_utf16 - fragment.start_utf16
        const identity = JSON.stringify([fragment.face_id, fragment.ascent_millipoints, fragment.descent_millipoints, fragment.line_gap_millipoints, fragment.language, fragment.underline_position_millipoints, fragment.underline_thickness_millipoints])
        const prior = runFontIdentities.get(run.id)
        if (prior !== undefined && prior !== identity) return undefined
        runFontIdentities.set(run.id, identity)
        for (let offset = fragment.start_utf16; offset < fragment.end_utf16; offset += 1) {
          const key = `${run.id}:${offset}`
          const priorScript = sourceScripts.get(key)
          if (priorScript !== undefined && priorScript !== fragment.script) return undefined
          sourceScripts.set(key, fragment.script)
        }
      }
      }
      if (covered !== sourceLength) return undefined
      const height = paragraph.lines.reduce((sum, line) => sum + line.line_height_millipoints, 0)
      if (!Number.isSafeInteger(height) || height <= 0 || paragraph.block_advance_millipoints !== height || height > profile.geometry.body_height_millipoints) return undefined
    }
  }
  let page = 0, column = 0, used = 0
  const selected: NativeDocxShapedLinesV1['paragraphs'] = []
  const placements: NativeDocxColumnParagraphFlowV1['placements'] = []
  for (let index = 0; index < profile.paragraphs.length; index += 1) {
    let paragraph = candidates[column]!.paragraphs[index]!
    if (used + paragraph.block_advance_millipoints > profile.geometry.body_height_millipoints) {
      column += 1
      if (column === 2) { column = 0; page += 1 }
      used = 0
      paragraph = candidates[column]!.paragraphs[index]!
    }
    selected.push(paragraph)
    placements.push({ paragraph_id: paragraph.paragraph_id, page_ordinal: page, column_ordinal: column })
    used += paragraph.block_advance_millipoints
  }
  return { ...profile, shaped_lines: { ...base, paragraphs: selected }, placements }
}

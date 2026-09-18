import type { NativeDocxResolvedLayoutInputV1 } from './nativeResolvedLayout.js'

/** A `w:tblStyle` whose `w:val` matches no `w:styleId` in the styles part.
 *
 * ECMA-376 17.7.2 binds a style reference to the `w:style` whose `w:styleId` it
 * names, so a reference that names no such style selects no style: there is no
 * definition to apply and none to defer. The resolver already reads it exactly
 * that way -- it records this absence and emits a resolved table carrying only
 * the dangling id, with no borders, no cell shading and no geometry -- so the
 * diagnostic reports a style that contributed nothing rather than formatting
 * this tier failed to reproduce. That is what makes it render-neutral, and the
 * conditions below are what prove it: any resolved style effect on the same
 * table, or any missing part provenance, disqualifies the exemption.
 *
 * This is not a general "unresolved style" exemption. A style that exists but
 * whose definition this tier cannot reproduce keeps its own diagnostics
 * (`TABLE_STYLE_EFFECTS_PRESERVED`, `CONDITIONAL_TABLE_STYLE_PRESERVED`) and
 * keeps blocking.
 */
export function nativeDocxUnresolvableTableStyleV1(
  diagnostic: NativeDocxResolvedLayoutInputV1['diagnostics'][number],
  resolved: NativeDocxResolvedLayoutInputV1,
): boolean {
  if (diagnostic.code !== 'MISSING_TABLE_STYLE' || diagnostic.severity !== 'unsupported' || diagnostic.preservation !== 'preserve-verbatim') return false
  if (diagnostic.part_name === undefined || diagnostic.part_name !== resolved.source_parts.styles_part || diagnostic.path !== undefined) return false
  const table = resolved.tables.find((entry) => entry.table_id === diagnostic.scope_id)
  return table !== undefined && table.style_id !== undefined
    && table.borders === undefined && table.cell_shading_rgb === undefined
    && table.geometry === undefined && table.automatic_border_preview === undefined
}

/** Only the extractor's exact-schema metadata below is render-neutral.
 * Keep the original diagnostic and preservation/edit policies intact. Unknown
 * latent markup and unresolved active style definitions are never exempted.
 * Font matching descriptors additionally require exact supplied faces in both
 * the compiler preflight and the shaper; this does not enable substitution.
 */
export function isRenderNeutralLayoutDiagnostic(
  diagnostic: NativeDocxResolvedLayoutInputV1['diagnostics'][number],
  resolved: NativeDocxResolvedLayoutInputV1,
): boolean {
  if (diagnostic.code === 'EMPTY_NUMBERING_STYLE_PRESERVED') return diagnostic.scope_id === resolved.document_id
    && diagnostic.severity === 'unsupported' && diagnostic.preservation === 'preserve-verbatim'
    && diagnostic.part_name !== undefined && diagnostic.part_name === resolved.source_parts.styles_part
    && /^\/w:styles\[1\]\/w:style\[[1-9][0-9]*\]$/.test(diagnostic.path ?? '')
  // A repeated w:font for a family the table already describes selects
  // nothing: the resolver keeps the first description, and the repeat carries
  // only the matching metadata this tier never consults. It is disclosed at the
  // exact w:font element so the source fact stays visible, and it is neutral to
  // every advance, so it does not block shaping.
  if (diagnostic.code === 'DUPLICATE_FONT_TABLE_ENTRY') return diagnostic.scope_id === resolved.document_id
    && diagnostic.severity === 'unsupported'
    && diagnostic.preservation === 'preserve-verbatim'
    && diagnostic.part_name !== undefined
    && diagnostic.part_name === resolved.source_parts.font_table_part
    && /^\/w:fonts\[1\]\/w:font\[[1-9][0-9]*\]$/.test(diagnostic.path ?? '')
  if (diagnostic.code === 'MISSING_TABLE_STYLE') return nativeDocxUnresolvableTableStyleV1(diagnostic, resolved)
  // The same dangling reference is reported twice. `resolveTableGeometry` walks
  // the style chain for the geometry it might have carried, and the first hop of
  // that walk is the referenced style itself, so its absence is also recorded as
  // a missing basedOn ancestor. There is no ancestor: the head is what is
  // missing, the chain is empty, and no layer was dropped. Exempt it only while
  // the table's own MISSING_TABLE_STYLE is itself exempt, which is what proves
  // the head -- and therefore the whole chain -- resolved to nothing.
  // ECMA-376 17.7.2 says a w:pStyle or w:rStyle naming a style the package does
  // not define is ignored, so the consumer cascades from the document defaults
  // and no layer was dropped. The resolver states exactly that case under its
  // own code, at the styles part of the scope that referenced it.
  if (diagnostic.code === 'UNDEFINED_STYLE_REFERENCE') return diagnostic.severity === 'unsupported'
    && diagnostic.preservation === 'preserve-verbatim'
    && diagnostic.part_name !== undefined && diagnostic.part_name === resolved.source_parts.styles_part
    && diagnostic.path === undefined
  if (diagnostic.code === 'MISSING_STYLE_REFERENCE') return diagnostic.severity === 'unsupported'
    && diagnostic.preservation === 'preserve-verbatim'
    && diagnostic.part_name !== undefined && diagnostic.part_name === resolved.source_parts.styles_part
    && diagnostic.path === undefined
    && resolved.diagnostics.some((entry) => entry.code === 'MISSING_TABLE_STYLE' && entry.scope_id === diagnostic.scope_id
      && entry.part_name === diagnostic.part_name && nativeDocxUnresolvableTableStyleV1(entry, resolved))
  if (diagnostic.code === 'FONT_MATCHING_METADATA_PRESERVED') return diagnostic.scope_id === resolved.document_id
    && diagnostic.severity === 'unsupported'
    && diagnostic.preservation === 'preserve-verbatim'
    && diagnostic.part_name !== undefined
    && diagnostic.part_name === resolved.source_parts.font_table_part
    && /^\/w:fonts\[1\]\/w:font\[[1-9][0-9]*\]\/w:(?:panose1|charset|family|pitch|sig|notTrueType)\[1\]$/.test(diagnostic.path ?? '')
  return diagnostic.code === 'LATENT_STYLE_BEHAVIOR_PRESERVED'
    && diagnostic.scope_id === resolved.document_id
    && diagnostic.severity === 'unsupported'
    && diagnostic.preservation === 'preserve-verbatim'
    && diagnostic.part_name !== undefined
    && diagnostic.part_name === resolved.source_parts.styles_part
    && diagnostic.path === '/w:styles[1]/w:latentStyles[1]'
}

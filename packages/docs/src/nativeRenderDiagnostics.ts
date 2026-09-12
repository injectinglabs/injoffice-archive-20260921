import type { NativeDocxResolvedLayoutInputV1 } from './nativeResolvedLayout.js'

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
  if (diagnostic.code === 'FONT_MATCHING_METADATA_PRESERVED') return diagnostic.scope_id === resolved.document_id
    && diagnostic.severity === 'unsupported'
    && diagnostic.preservation === 'preserve-verbatim'
    && diagnostic.part_name !== undefined
    && diagnostic.part_name === resolved.source_parts.font_table_part
    && /^\/w:fonts\[1\]\/w:font\[[1-9][0-9]*\]\/w:(?:panose1|charset|family|pitch|sig)\[1\]$/.test(diagnostic.path ?? '')
  return diagnostic.code === 'LATENT_STYLE_BEHAVIOR_PRESERVED'
    && diagnostic.scope_id === resolved.document_id
    && diagnostic.severity === 'unsupported'
    && diagnostic.preservation === 'preserve-verbatim'
    && diagnostic.part_name !== undefined
    && diagnostic.part_name === resolved.source_parts.styles_part
    && diagnostic.path === '/w:styles[1]/w:latentStyles[1]'
}

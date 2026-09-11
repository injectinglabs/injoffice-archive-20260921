import type { NativeDocxResolvedLayoutInputV1 } from './nativeResolvedLayout.js'

/** Only the extractor's exact-schema latent UI/locking metadata is render-neutral.
 * Keep the original diagnostic and preservation/edit policies intact. Unknown
 * latent markup and unresolved active style definitions are never exempted.
 */
export function isRenderNeutralLayoutDiagnostic(
  diagnostic: NativeDocxResolvedLayoutInputV1['diagnostics'][number],
  resolved: NativeDocxResolvedLayoutInputV1,
): boolean {
  return diagnostic.code === 'LATENT_STYLE_BEHAVIOR_PRESERVED'
    && diagnostic.scope_id === resolved.document_id
    && diagnostic.severity === 'unsupported'
    && diagnostic.preservation === 'preserve-verbatim'
    && diagnostic.part_name !== undefined
    && diagnostic.part_name === resolved.source_parts.styles_part
    && diagnostic.path === '/w:styles[1]/w:latentStyles[1]'
}

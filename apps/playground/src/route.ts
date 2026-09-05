export const SURFACES = [
  'overview',
  'sheets',
  'charts',
  'pivots',
  'shapes',
  'connectors',
  'formulas',
  'docs',
  'slides',
  'pdf',
  'collab',
  'history',
  'font-metrics',
  'pptx-authored',
  'pptx-native',
  'pptx-render',
] as const
export type Surface = (typeof SURFACES)[number]

const LEGACY_ALIASES: Record<string, Surface> = { native: 'sheets' }

export function parseSurface(hash = typeof location === 'undefined' ? '' : location.hash): Surface {
  const path = hash.replace(/^#\/?/, '').split(/[/?#]/)[0]?.toLowerCase() ?? ''
  if (path in LEGACY_ALIASES) return LEGACY_ALIASES[path]
  return (SURFACES as readonly string[]).includes(path) ? (path as Surface) : 'overview'
}

export function surfaceHref(surface: Surface): string {
  return `#/${surface}`
}

export type SheetsView = 'editor' | 'native' | 'tools'

export function parseSheetsView(hash = typeof location === 'undefined' ? '' : location.hash): SheetsView {
  const query = hash.split('?')[1]?.split('#')[0] ?? ''
  const view = new URLSearchParams(query).get('view')
  return view === 'native' || view === 'tools' ? view : 'editor'
}

export function isDocsHash(hash = typeof location === 'undefined' ? '' : location.hash): boolean {
  const path = hash.replace(/^#\/?/, '').split(/[/?]/)[0]?.toLowerCase() ?? ''
  return path === 'guides'
}

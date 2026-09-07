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
  'agent',
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

export const AGENT_TOOLS = [
  {
    tool: 'sheets',
    format: 'xlsx',
    fileType: 'XLSX',
    label: 'Sheets',
    title: 'AI change sets · Sheets',
    description: 'Inspect a workbook, propose one cell write, preview the isolated change, and verify the replacement revision. No model SDK or network request.',
  },
  {
    tool: 'docs',
    format: 'docx',
    fileType: 'DOCX',
    label: 'Docs',
    title: 'AI change sets · Docs',
    description: 'Inspect a document, replace one guarded paragraph, preview the isolated change, and verify the replacement revision. No model SDK or network request.',
  },
  {
    tool: 'slides',
    format: 'pptx',
    fileType: 'PPTX',
    label: 'Slides',
    title: 'AI change sets · Slides',
    description: 'Inspect a deck, update one stable metric shape, preview the isolated change, and verify the replacement revision. No model SDK or network request.',
  },
  {
    tool: 'pdf',
    format: 'pdf',
    fileType: 'PDF',
    label: 'PDF',
    title: 'AI change sets · PDF',
    description: 'Inspect a PDF, rotate one page, preview the isolated change, and verify the replacement revision. No model SDK or network request.',
  },
] as const
export type AgentTool = (typeof AGENT_TOOLS)[number]['tool']
export type AgentToolFormat = (typeof AGENT_TOOLS)[number]['format']

export function parseAgentTool(hash = typeof location === 'undefined' ? '' : location.hash): AgentTool {
  const query = hash.split('?')[1]?.split('#')[0] ?? ''
  const format = new URLSearchParams(query).get('format')?.toLowerCase() ?? ''
  if (format === 'docs' || format === 'docx') return 'docs'
  if (format === 'slides' || format === 'pptx') return 'slides'
  if (format === 'pdf') return 'pdf'
  return 'sheets'
}

export function agentHref(tool: AgentTool): string {
  return `#/agent?format=${tool}`
}

export function agentFormatFromTool(tool: AgentTool): AgentToolFormat {
  return AGENT_TOOLS.find((item) => item.tool === tool)?.format ?? 'xlsx'
}

export function isDocsHash(hash = typeof location === 'undefined' ? '' : location.hash): boolean {
  const path = hash.replace(/^#\/?/, '').split(/[/?]/)[0]?.toLowerCase() ?? ''
  return path === 'guides'
}

export function isDesignSystemHash(hash = typeof location === 'undefined' ? '' : location.hash): boolean {
  const path = hash.replace(/^#\/?/, '').split(/[/?]/)[0]?.toLowerCase() ?? ''
  return path === 'design-system'
}

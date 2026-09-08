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
    description: 'Try a simulated AI request with the built-in mock agent—no model, credentials, or setup. Review and approve a real XLSX change through public tools, then inspect native verification. Explore stale approvals, safe retries, and readback failures.',
  },
  {
    tool: 'docs',
    format: 'docx',
    fileType: 'DOCX',
    label: 'Docs',
    title: 'AI change sets · Docs',
    description: 'Explore approval and revision guards with a simulated document edit. This lifecycle example uses document-shaped data, not DOCX bytes. No model service.',
  },
  {
    tool: 'slides',
    format: 'pptx',
    fileType: 'PPTX',
    label: 'Slides',
    title: 'AI change sets · Slides',
    description: 'Explore approval and revision guards with a simulated slide update. This lifecycle example uses document-shaped data, not PPTX bytes. No model service.',
  },
  {
    tool: 'pdf',
    format: 'pdf',
    fileType: 'PDF',
    label: 'PDF',
    title: 'AI change sets · PDF',
    description: 'Explore approval and revision guards with a simulated page rotation. This lifecycle example uses document-shaped data, not PDF bytes. No model service.',
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

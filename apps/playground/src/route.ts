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
    title: 'Update a workstream status',
    description: 'Choose a workstream and its new status in the launch tracker. Review the difference, approve the edit, and download a verified XLSX. Simulated agent, real file operations; no LLM or setup.',
  },
  {
    tool: 'docs',
    format: 'docx',
    fileType: 'DOCX',
    label: 'Docs',
    title: 'Revise document text',
    description: 'Choose inspected text in the launch brief and write its replacement. Review and approve the exact edit, then download a verified DOCX. Simulated agent, real file operations; no LLM or setup.',
  },
  {
    tool: 'slides',
    format: 'pptx',
    fileType: 'PPTX',
    label: 'Slides',
    title: 'Edit presentation text',
    description: 'Update an exact text element in the launch review. Review the difference, approve it, and download a verified PPTX. Simulated agent, real file operations; no LLM or setup.',
  },
  {
    tool: 'pdf',
    format: 'pdf',
    fileType: 'PDF',
    label: 'PDF',
    title: 'Rotate a PDF page',
    description: 'Choose a page and rotation in the operating review. Compare page metadata, approve the change, and download a verified PDF. Simulated agent, real file operations; no LLM or setup.',
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

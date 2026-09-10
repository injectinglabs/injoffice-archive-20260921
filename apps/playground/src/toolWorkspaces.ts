import { parseAgentTool, type AgentTool, type Surface } from './route'

export type ToolWorkspaceFeature = { id: string; label: string; group: string; source: Exclude<Surface, 'overview'>; description: string }
export type ToolWorkspaceDefinition = { tool: AgentTool; title: string; description: string; features: readonly ToolWorkspaceFeature[] }

const agent: ToolWorkspaceFeature = { id: 'agent', label: 'Guided agent task', group: 'AI', source: 'agent', description: 'A simulated agent proposes a real file edit. Review, approve, verify, and download without a model service.' }
const collab: ToolWorkspaceFeature = { id: 'collab', label: 'Two-editor collaboration', group: 'Collaborate', source: 'collab', description: 'Try synchronized edits and presence with an independent collaboration sample.' }
const history: ToolWorkspaceFeature = { id: 'history', label: 'History and diffs', group: 'Collaborate', source: 'history', description: 'Compare structured changes and inspect a sample version timeline.' }
const shapes: ToolWorkspaceFeature = { id: 'shapes', label: 'Shape gallery', group: 'More tools', source: 'shapes', description: 'Inspect shared shape identifiers and renderer approximations; the native wire sample is spreadsheet-oriented.' }
const fonts: ToolWorkspaceFeature = { id: 'font-metrics', label: 'Typography and layout', group: 'More tools', source: 'font-metrics', description: 'Inspect layout contracts and the explicit Node boundary for font resolution and shaping.' }

export const TOOL_WORKSPACES: readonly ToolWorkspaceDefinition[] = [
  { tool: 'sheets', title: 'Sheets', description: 'Update a launch tracker. Review the change and download a verified workbook.', features: [
    { id: 'editor', label: 'Workbook editor', group: 'Edit', source: 'sheets', description: 'Edit the interactive sample workbook in your browser.' },
    { id: 'native', label: 'XLSX file round trip', group: 'Edit', source: 'sheets', description: 'Open real XLSX bytes, apply a bounded change, verify, and download.' },
    { id: 'tools', label: 'Workbook package samples', group: 'Edit', source: 'sheets', description: 'Explore sparklines and outlines. Print uses a sample callback, and exchange uses a demonstration codec—not real printing or XLSX import.' },
    agent,
    { id: 'charts', label: 'Charts', group: 'Analyze', source: 'charts', description: 'Render sample data with the real ECharts adapter. SVG export is a title-only sample, not the rendered chart image.' },
    { id: 'pivots', label: 'Pivot tables', group: 'Analyze', source: 'pivots', description: 'Group, filter, and aggregate an independent values grid.' },
    { id: 'formulas', label: 'Formula coverage', group: 'Analyze', source: 'formulas', description: 'Search real audited coverage. The separate calculation example is a limited + / SUM demo, not the complete formula engine.' },
    { id: 'connectors', label: 'Data connectors', group: 'Analyze', source: 'connectors', description: 'Normalize sample JSON and CSV into bound ranges.' },
    collab, history, shapes,
  ] },
  { tool: 'docs', title: 'Docs', description: 'Revise a launch brief. Approve the exact text change before saving.', features: [
    { id: 'editor', label: 'Document editor', group: 'Edit', source: 'docs', description: 'Open and update DOCX text, then download the real file. The preview is not Word pagination.' },
    agent, collab, history, fonts,
  ] },
  { tool: 'slides', title: 'Slides', description: 'Update a presentation. Review the changed text and download the verified deck.', features: [
    { id: 'editor', label: 'Presentation editor', group: 'Edit', source: 'slides', description: 'Build a DeckSpec sample with themes, layout checks, transitions, and a live canvas; this view does not emit PPTX file bytes.' },
    { id: 'pptx-native', label: 'Native PPTX editing', group: 'Edit', source: 'pptx-native', description: 'Edit guarded text or an exact supported shape in real PPTX bytes, then verify and download.' },
    { id: 'pptx-authored', label: 'PPTX authoring', group: 'Edit', source: 'pptx-authored', description: 'Compile a sample DeckSpec into strict native JSON presentation objects, not emitted PPTX file bytes.' },
    agent, collab, shapes,
    { id: 'pptx-render', label: 'PPTX rendering', group: 'More tools', source: 'pptx-render', description: 'Inspect deterministic render trees and renderer-independent paint commands.' },
    fonts,
  ] },
  { tool: 'pdf', title: 'PDF', description: 'Rotate a review page. Approve the change and download the verified PDF.', features: [
    { id: 'editor', label: 'PDF editor', group: 'Edit', source: 'pdf', description: 'Open, annotate, organize, and download a PDF. Advanced server tools remain explicitly separate.' },
    agent, collab, fonts,
  ] },
]

export function workspaceHref(tool: AgentTool, feature = 'agent'): string {
  const definition = TOOL_WORKSPACES.find((item) => item.tool === tool)!
  const selected = definition.features.some((item) => item.id === feature) ? feature : 'agent'
  return selected === 'agent' ? `#/${tool}` : `#/${tool}?feature=${encodeURIComponent(selected)}`
}

/** Keep shell and mounted views in sync even before an old URL is replaced. */
export function workspaceNavigationHash(hash: string): string {
  return resolveToolWorkspace(hash) ? hash : workspaceHref('sheets')
}

/** Canonical tool routes and old individual-demo bookmarks share one resolver. */
export function resolveToolWorkspace(hash: string): { tool: AgentTool; feature: string } | null {
  const path = hash.replace(/^#\/?/, '').split(/[/?#]/)[0]?.toLowerCase() ?? ''
  const query = new URLSearchParams(hash.split('?')[1]?.split('#')[0] ?? '')
  const direct = TOOL_WORKSPACES.find((item) => item.tool === path)
  if (direct) {
    const requested = query.get('feature') ?? (direct.tool === 'sheets' ? query.get('view') : null) ?? 'agent'
    return { tool: direct.tool, feature: direct.features.some((item) => item.id === requested) ? requested : 'agent' }
  }
  if (path === 'native') return { tool: 'sheets', feature: 'native' }
  if (path === 'agent' || path === 'collab') return { tool: parseAgentTool(hash), feature: path }
  if (path === 'history') return { tool: parseAgentTool(hash) === 'docs' ? 'docs' : 'sheets', feature: 'history' }
  if (path === 'font-metrics') {
    const requestedTool = parseAgentTool(hash)
    return { tool: requestedTool === 'slides' || requestedTool === 'pdf' ? requestedTool : 'docs', feature: path }
  }
  if (path === 'shapes') return { tool: parseAgentTool(hash) === 'slides' ? 'slides' : 'sheets', feature: path }
  if (['charts', 'pivots', 'connectors', 'formulas'].includes(path)) return { tool: 'sheets', feature: path }
  if (['pptx-authored', 'pptx-native', 'pptx-render'].includes(path)) return { tool: 'slides', feature: path }
  return null
}

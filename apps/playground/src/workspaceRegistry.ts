import { DEMO_BY_SURFACE, type DemoDefinition } from './demoRegistry'
import type { DemoRecipe } from './demoRecipes'
import { preloadableLazyNamed } from './preloadableLazy'
import { type AgentTool, type Surface } from './route'
import { TOOL_WORKSPACES, resolveToolWorkspace } from './toolWorkspaces'

const components = {
  sheets: preloadableLazyNamed(() => import('./components/ToolWorkspace'), 'SheetsWorkspace'),
  docs: preloadableLazyNamed(() => import('./components/ToolWorkspace'), 'DocsWorkspace'),
  slides: preloadableLazyNamed(() => import('./components/ToolWorkspace'), 'SlidesWorkspace'),
  pdf: preloadableLazyNamed(() => import('./components/ToolWorkspace'), 'PdfWorkspace'),
}

const sheetsViewRecipes: Record<string, DemoRecipe> = {
  editor: {
    id: 'sheets-live-editor',
    title: 'Explore the live workbook',
    outcome: 'Edit sample cells and create charts, pivots, and shapes in the live spreadsheet. Native file editing uses a separate sample in XLSX file round trip.',
    minutes: 3,
    steps: [
      { id: 'edit', title: 'Edit a cell', instruction: 'Select a sample cell and enter a new value.', evidence: 'The live grid shows your edit.' },
      { id: 'chart', title: 'Visualize a range', instruction: 'Select a data range, choose Chart type, and select Add chart.', evidence: 'A chart appears for the selected range.' },
      { id: 'pivot', title: 'Summarize the data', instruction: 'Select a data block including its header row and choose Add pivot table.', evidence: 'The pivot panel opens for the selected data.' },
    ],
    sources: [{ label: 'Live spreadsheet editor', path: 'apps/playground/src/UniverEditor.tsx', href: 'https://github.com/injectinglabs/injoffice/blob/main/apps/playground/src/UniverEditor.tsx' }],
    related: [],
  },
  tools: {
    id: 'sheets-package-tools',
    title: 'Explore spreadsheet package contracts',
    outcome: 'Try independent sparkline, outline, print, and exchange samples. Print and exchange demonstrate host callbacks, not full workbook printing or file import.',
    minutes: 2,
    steps: [
      { id: 'sparkline', title: 'Change a sparkline', instruction: 'Choose line, column, or win-loss from Sparkline type.', evidence: 'The SVG changes using the same sample values.' },
      { id: 'outline', title: 'Expand a row group', instruction: 'Select Expand Q1 rows, then collapse them again.', evidence: 'The outline manager changes which sample rows are visible.' },
      { id: 'exchange', title: 'Inspect a sample exchange job', instruction: 'Select Import XLSX job.', evidence: 'The sample codec reports a completed job. Use Edit → XLSX file round trip for real workbook bytes.' },
    ],
    sources: [{ label: 'Spreadsheet package samples', path: 'apps/playground/src/pages/SheetsToolsPage.tsx', href: 'https://github.com/injectinglabs/injoffice/blob/main/apps/playground/src/pages/SheetsToolsPage.tsx' }],
    related: [],
  },
}

/** Four public examples; the capability registry is implementation detail. */
export const WORKSPACE_DEMOS: DemoDefinition[] = TOOL_WORKSPACES.map(workspace => ({
  ...DEMO_BY_SURFACE.get(workspace.tool)!,
  title: workspace.title,
  navTitle: workspace.title,
  description: workspace.description,
  ...components[workspace.tool],
}))

export function preloadWorkspace(surface: Surface): Promise<void> {
  return components[surface as AgentTool]?.preload() ?? Promise.resolve()
}

export function preloadWorkspaceOnIntent(surface: Surface): void {
  void preloadWorkspace(surface).catch(() => undefined)
}

/** Guide/source follows the selected capability, not a stale workspace default. */
export function workspaceProofDemo(hash: string): DemoDefinition {
  const route = resolveToolWorkspace(hash)
  const workspace = TOOL_WORKSPACES.find(item => item.tool === route?.tool) ?? TOOL_WORKSPACES[0]!
  const feature = workspace.features.find(item => item.id === (route?.feature ?? 'agent'))!
  const demo = DEMO_BY_SURFACE.get(feature.source) ?? DEMO_BY_SURFACE.get(workspace.tool)!
  const recipe = workspace.tool === 'sheets' ? sheetsViewRecipes[feature.id] ?? demo.recipe : demo.recipe
  return { ...demo, recipe, description: feature.description, title: `${workspace.title}: ${feature.label}` }
}

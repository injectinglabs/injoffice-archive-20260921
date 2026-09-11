import { DEMO_BY_SURFACE, type DemoDefinition } from './demoRegistry'
import { preloadableLazyNamed } from './preloadableLazy'
import { type AgentTool, type Surface } from './route'
import { TOOL_WORKSPACES } from './toolWorkspaces'

const components = {
  sheets: preloadableLazyNamed(() => import('./components/ToolWorkspace'), 'SheetsWorkspace'),
  docs: preloadableLazyNamed(() => import('./components/ToolWorkspace'), 'DocsWorkspace'),
  slides: preloadableLazyNamed(() => import('./components/ToolWorkspace'), 'SlidesWorkspace'),
  pdf: preloadableLazyNamed(() => import('./components/ToolWorkspace'), 'PdfWorkspace'),
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

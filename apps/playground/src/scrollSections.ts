import type { DemoDefinition } from './demoRegistry'
import { surfaceHref, type Surface } from './route'
import { WORKSPACE_DEMOS } from './workspaceRegistry'
import { resolveToolWorkspace } from './toolWorkspaces'
export { workspaceNavigationHash } from './toolWorkspaces'

export type ScrollSection = { key: string; surface: Surface; href: string; demo?: DemoDefinition }
export const SCROLL_SECTIONS: ScrollSection[] = [
  ...WORKSPACE_DEMOS.map(demo => ({ key: demo.surface, surface: demo.surface, href: surfaceHref(demo.surface), demo })),
]

export function sectionForHash(hash: string): ScrollSection {
  const target = resolveToolWorkspace(hash)
  return SCROLL_SECTIONS.find(section => section.surface === target?.tool) ?? SCROLL_SECTIONS[0]!
}

/** Coordinates are relative to the viewport. Reading never moves focus or scroll. */
export function activeSectionKey(sections: readonly { key: string; top: number }[], probe: number, atEnd = false): string | undefined {
  if (atEnd) return sections.filter(section => Number.isFinite(section.top)).at(-1)?.key
  return sections.filter(section => section.top <= probe).at(-1)?.key ?? sections[0]?.key
}

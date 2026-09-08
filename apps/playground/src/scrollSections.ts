import { DEMOS, DEMO_GROUPS, type DemoDefinition } from './demoRegistry'
import { AGENT_TOOLS, agentHref, parseAgentTool, parseSurface, surfaceHref, type AgentTool, type Surface } from './route'

export type ScrollSection = { key: string; surface: Surface; href: string; demo?: DemoDefinition; tool?: AgentTool }
export const SCROLL_SECTIONS: ScrollSection[] = [
  { key: 'overview', surface: 'overview', href: surfaceHref('overview') },
  ...DEMO_GROUPS.flatMap(group => DEMOS.filter(demo => demo.group === group).flatMap<ScrollSection>(demo =>
    demo.surface === 'agent'
      ? AGENT_TOOLS.map(({ tool }) => ({ key: `agent-${tool}`, surface: demo.surface, href: agentHref(tool), demo, tool }))
      : [{ key: demo.surface, surface: demo.surface, href: surfaceHref(demo.surface), demo }],
  )),
]

export function sectionForHash(hash: string): ScrollSection {
  const surface = parseSurface(hash)
  return SCROLL_SECTIONS.find(section => section.surface === surface && (surface !== 'agent' || section.tool === parseAgentTool(hash))) ?? SCROLL_SECTIONS[0]!
}

/** Coordinates are relative to the viewport. Reading never moves focus or scroll. */
export function activeSectionKey(sections: readonly { key: string; top: number }[], probe: number): string | undefined {
  return sections.filter(section => section.top <= probe).at(-1)?.key ?? sections[0]?.key
}

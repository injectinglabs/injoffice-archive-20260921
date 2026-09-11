import { describe, expect, it } from 'vitest'
import { DEMOS } from './demoRegistry'
import { TOOL_WORKSPACES } from './toolWorkspaces'
import { WORKSPACE_DEMOS } from './workspaceRegistry'

describe('four public tool examples', () => {
  it('reduces navigation to four tools without removing any implemented capability surface', () => {
    expect(WORKSPACE_DEMOS.map(item => item.surface)).toEqual(['sheets', 'docs', 'slides', 'pdf'])
    const covered = new Set(TOOL_WORKSPACES.flatMap(item => item.features.map(feature => feature.source)))
    expect(covered).toEqual(new Set(DEMOS.map(item => item.surface)))
    expect(TOOL_WORKSPACES[0]?.features.map(item => item.id)).toEqual(expect.arrayContaining(['editor', 'native', 'tools', 'charts', 'pivots', 'formulas', 'connectors', 'agent', 'collab', 'history', 'shapes']))
    for (const tool of TOOL_WORKSPACES) expect(tool.features.map(item => item.id)).toEqual(expect.arrayContaining(['editor', 'agent', 'collab']))
  })

})

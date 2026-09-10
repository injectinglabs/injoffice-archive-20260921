import { describe, expect, it } from 'vitest'
import { DEMOS } from './demoRegistry'
import { TOOL_WORKSPACES, workspaceHref } from './toolWorkspaces'
import { WORKSPACE_DEMOS, workspaceProofDemo } from './workspaceRegistry'

describe('four public tool examples', () => {
  it('reduces navigation to four tools without removing any implemented capability surface', () => {
    expect(WORKSPACE_DEMOS.map(item => item.surface)).toEqual(['sheets', 'docs', 'slides', 'pdf'])
    const covered = new Set(TOOL_WORKSPACES.flatMap(item => item.features.map(feature => feature.source)))
    expect(covered).toEqual(new Set(DEMOS.map(item => item.surface)))
    expect(TOOL_WORKSPACES[0]?.features.map(item => item.id)).toEqual(expect.arrayContaining(['editor', 'native', 'tools', 'charts', 'pivots', 'formulas', 'connectors', 'agent', 'collab', 'history', 'shapes']))
    for (const tool of TOOL_WORKSPACES) expect(tool.features.map(item => item.id)).toEqual(expect.arrayContaining(['editor', 'agent', 'collab']))
  })

  it('selects the actual active feature guide/source through canonical and legacy URLs', () => {
    expect(workspaceProofDemo('#/sheets?feature=editor').recipe.id).toBe('sheets-live-editor')
    expect(workspaceProofDemo('#/sheets?feature=tools').recipe.id).toBe('sheets-package-tools')
    expect(workspaceProofDemo('#/sheets?feature=native').recipe.id).toBe('sheets-native-round-trip')
    expect(workspaceProofDemo('#/charts').surface).toBe('charts')
    expect(workspaceProofDemo('#/sheets?feature=charts').surface).toBe('charts')
    expect(workspaceProofDemo('#/slides?feature=pptx-native').surface).toBe('pptx-native')
    expect(workspaceProofDemo('#/agent?format=docs').surface).toBe('agent')
    for (const workspace of TOOL_WORKSPACES) {
      for (const feature of workspace.features) {
        expect(workspaceProofDemo(workspaceHref(workspace.tool, feature.id)).title).toBe(`${workspace.title}: ${feature.label}`)
      }
    }
  })
})

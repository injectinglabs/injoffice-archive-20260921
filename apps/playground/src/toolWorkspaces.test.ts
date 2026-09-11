import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { SURFACES } from './route'
import { resolveToolWorkspace, TOOL_WORKSPACES, workspaceExamples, workspaceHref } from './toolWorkspaces'

describe('four tool workspace catalogue', () => {
  it('has four tools, small primary groups, and complete existing surface coverage', () => {
    expect(TOOL_WORKSPACES.map((item) => item.tool)).toEqual(['sheets', 'docs', 'slides', 'pdf'])
    for (const workspace of TOOL_WORKSPACES) {
      expect(workspace.features[0]?.id).toBe('editor')
      expect(new Set(workspace.features.map((item) => item.id)).size).toBe(workspace.features.length)
      expect(new Set(workspace.features.map((item) => item.group)).size).toBeGreaterThanOrEqual(4)
      expect(new Set(workspace.features.map((item) => item.group)).size).toBeLessThanOrEqual(5)
      expect(workspace.features.some((item) => item.id === 'agent')).toBe(true)
      expect(workspace.features.some((item) => item.id === 'collab')).toBe(true)
    }
    const sources = new Set(TOOL_WORKSPACES.flatMap((item) => item.features.map((feature) => feature.source)))
    expect([...sources].sort()).toEqual(SURFACES.filter((surface) => surface !== 'overview').sort())
    expect(TOOL_WORKSPACES[0]!.features.map((feature) => feature.id)).toEqual(expect.arrayContaining(['editor', 'native', 'tools', 'charts', 'pivots', 'connectors', 'formulas', 'shapes', 'history']))
  })

  it('round-trips every canonical feature route', () => {
    for (const workspace of TOOL_WORKSPACES) for (const feature of workspace.features) {
      expect(resolveToolWorkspace(workspaceHref(workspace.tool, feature.id))).toEqual({ tool: workspace.tool, feature: feature.id })
    }
    expect(workspaceHref('docs')).toBe('#/docs')
    expect(workspaceHref('docs', 'charts')).toBe('#/docs')
    expect(resolveToolWorkspace('#/docs?feature=charts')).toEqual({ tool: 'docs', feature: 'agent' })
    expect(resolveToolWorkspace('#/sheets?feature=agent&view=native')).toEqual({ tool: 'sheets', feature: 'agent' })
  })

  it.each([
    ['#/agent', 'sheets', 'agent'], ['#/agent?format=docx', 'docs', 'agent'], ['#/agent?format=pptx', 'slides', 'agent'], ['#/agent?format=pdf', 'pdf', 'agent'],
    ['#/sheets?view=tools', 'sheets', 'tools'], ['#/sheets?view=native', 'sheets', 'native'], ['#/native', 'sheets', 'native'],
    ['#/charts', 'sheets', 'charts'], ['#/pivots', 'sheets', 'pivots'], ['#/formulas', 'sheets', 'formulas'], ['#/connectors', 'sheets', 'connectors'],
    ['#/collab?format=docs', 'docs', 'collab'], ['#/history', 'sheets', 'history'], ['#/history?format=docs', 'docs', 'history'],
    ['#/shapes', 'sheets', 'shapes'], ['#/shapes?format=slides', 'slides', 'shapes'],
    ['#/font-metrics', 'docs', 'font-metrics'], ['#/font-metrics?format=pdf', 'pdf', 'font-metrics'],
    ['#/pptx-authored', 'slides', 'pptx-authored'], ['#/pptx-native', 'slides', 'pptx-native'], ['#/pptx-render', 'slides', 'pptx-render'],
  ])('preserves legacy bookmark %s', (hash, tool, feature) => {
    expect(resolveToolWorkspace(hash)).toEqual({ tool, feature })
  })

  it.each(['', '#/', '#/overview', '#/unknown'])('does not turn unknown route %s into an unrelated capability', (hash) => {
    expect(resolveToolWorkspace(hash)).toBeNull()
  })

  it('lists every example, lazily retains its content, and keeps failures local', () => {
    const source = readFileSync(new URL('./components/ToolWorkspace.tsx', import.meta.url), 'utf8')
    expect(source).toContain('visited.includes(feature.id)')
    expect(source).toContain('examples.map((feature) => <section')
    expect(source).toContain('new IntersectionObserver')
    expect(source).not.toContain('hidden={feature.id !== current.id}')
    expect(source).toContain('previous.includes(feature) ? previous : [...previous, feature]')
    expect(source).toContain('data-workspace-error')
    expect(source).toContain('data-workspace-retry')
    expect(source).toContain('data-workspace-loading')
    expect(source).toContain('useMemo(() => lazy(() => loadFeature(tool, feature, initialHash))')
    expect(source).toContain('key={`${feature.id}-${retryKeys[feature.id] ?? 0}`}')
    expect(source).toContain('requestAnimationFrame(() => window.dispatchEvent(new Event(\'resize\')))')
    expect(source).toContain('Each view has its own sample and state; edits do not transfer between views.')
    expect(source).toContain('fixedTool={tool}')
    expect(source).toContain('fixedFormat={tool}')
    expect(source).not.toContain("import('../pages/SheetsPage')")
    expect(source).toContain('if (!(feature in featureHashes.current))')
    expect(source).toContain('route?.tool === tool && route.feature === feature ? hash : undefined')
    expect(source).toContain('activate(route.feature, window.location.hash)')
    expect(source).toContain('initialHash={featureHashes.current[feature.id]}')
    expect(source).not.toContain('inert={feature.id !== current.id}')
    for (const workspace of TOOL_WORKSPACES) {
      expect(workspaceExamples(workspace)[0]?.id).toBe('agent')
      expect(workspaceExamples(workspace)).toHaveLength(workspace.features.length)
      expect(new Set(workspaceExamples(workspace).map(feature => feature.id)).size).toBe(workspace.features.length)
    }
  })

  it('uses a sidebar index and named, sequential sections without an examples dropdown', () => {
    const source = readFileSync(new URL('./components/ToolWorkspace.tsx', import.meta.url), 'utf8')
    expect(source).toContain('const instanceId = useId()')
    expect(source).toContain('aria-labelledby=')
    expect(source).toContain('tabIndex={-1}')
    expect(source).not.toContain('More examples')
    const shell = readFileSync(new URL('./App.tsx', import.meta.url), 'utf8')
    expect(shell).toContain('workspaceExamples(workspace).map(feature => <li')
    expect(shell).toContain('href={workspaceHref(workspace.tool, feature.id)}')
    expect(source).not.toContain('role="tablist"')
    expect(source).not.toContain('<select')
  })

  it('opens the same-document assistant by default while preserving explicit editor links', () => {
    for (const { tool } of TOOL_WORKSPACES) {
      expect(resolveToolWorkspace(`#/${tool}`)).toEqual({ tool, feature: 'agent' })
      expect(resolveToolWorkspace(`#/${tool}?feature=editor`)).toEqual({ tool, feature: 'editor' })
    }
  })

  it('listens to passive scrolling without writing navigation history from panels', () => {
    const source = readFileSync(new URL('./components/ToolWorkspace.tsx', import.meta.url), 'utf8')
    expect(source).toContain("window.addEventListener('injoffice:workspace-scroll', sync)")
    expect(source).toContain("window.removeEventListener('injoffice:workspace-scroll', sync)")
    expect(source).not.toContain('window.location.hash =')
  })
})

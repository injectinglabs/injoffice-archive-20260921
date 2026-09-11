import { existsSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'

const source = readFileSync(new URL('./App.tsx', import.meta.url), 'utf8')

describe('direct four-tool demo', () => {
  it('removes the introduction and obsolete navigation instead of hiding them', () => {
    expect(existsSync(new URL('./pages/OverviewPage.tsx', import.meta.url))).toBe(false)
    for (const removed of ['OverviewPage', 'demo-overview', 'Explore the four document tools', 'All demos', "surfaceHref('overview')"]) expect(source).not.toContain(removed)
    expect(source).toContain('TOOL_WORKSPACES.map')
    expect(source).toContain("const Heading = demo.surface === 'sheets' ? 'h1' : 'h2'")
  })

  it('replaces retired route history and makes the brand return to the first tool', () => {
    expect(source).toContain('const destination = workspaceNavigationHash(location.hash)')
    expect(source).toContain("history.replaceState(history.state, '', destination)")
    expect(source).toContain('aria-label="InjOffice demo start"')
    expect(source).toContain("location.hash === surfaceHref('sheets')")
    const workspace = readFileSync(new URL('./components/ToolWorkspace.tsx', import.meta.url), 'utf8')
    expect(workspace).toContain('resolveToolWorkspace(workspaceNavigationHash(window.location.hash))')
  })
})

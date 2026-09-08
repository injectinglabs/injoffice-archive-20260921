import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'

const source = readFileSync(new URL('./pages/OverviewPage.tsx', import.meta.url), 'utf8')

describe('four-tool overview', () => {
  it('offers only the four comprehensive workspaces, not a separate feature catalogue', () => {
    expect(source).toContain('Explore the four document tools')
    expect(source).toContain('id="showcase-title" tabIndex={-1}')
    expect(source).toContain('TOOL_WORKSPACES.map')
    expect(source).toContain('data-workspace-entry={workspace.tool}')
    expect(source).toContain('workspaceHref(workspace.tool)')
    expect(source).not.toContain('showcaseItems')
    expect(source).not.toContain('showcase-catalog')
    expect(source).not.toContain('showcase-query')
  })

  it('describes implemented features and keeps sample and model boundaries explicit', () => {
    expect(source).toContain('Simulated agent · real document operations')
    expect(source).toContain('No LLM, API key, or external AI service')
    expect(source).toContain('independent samples')
    expect(source).toContain('package-contract simulations')
    expect(source).toContain('host-only capabilities')
    expect(source).not.toContain('sha256:')
    expect(source).not.toContain('full Microsoft Office parity')
  })

  it('preserves continuous scrolling and explains state retention', () => {
    expect(source).toContain('Scroll between the four tools')
    expect(source).toContain('navigation follows your position')
    expect(source).toContain('until you reset or close that workspace')
    expect(source).toContain('preloadWorkspaceOnIntent')
  })
})

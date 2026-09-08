import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'

const source = readFileSync(new URL('./pages/OverviewPage.tsx', import.meta.url), 'utf8')

describe('overview guided workflow', () => {
  it('introduces document tasks before the searchable catalogue in one continuous demo', () => {
    expect(source).toContain('Scroll through the demos below, or jump to a section from the navigation.')
    expect(source).toContain('<h1 id="showcase-title" tabIndex={-1}>Give an agent a document task</h1>')
    expect(source).toContain('<h2 id="showcase-catalog-title">')
    expect(source.indexOf('className="task-launch"')).toBeLessThan(source.indexOf('className="showcase-catalog"'))
    expect(source).toContain('Browse all demos')
  })
  it('offers four bounded starter links without triggering writes or approval', () => {
    for (const tool of ['sheets', 'docs', 'slides', 'pdf']) expect(source).toContain(`tool: '${tool}'`)
    for (const task of ['Update a workstream status', 'Revise a document title', 'Update a presentation title', 'Rotate a PDF page']) expect(source).toContain(task)
    expect(source).toContain('href={agentHref(starter.tool)}')
    expect(source).toContain('Choosing a task never applies or approves a change.')
    expect(source).not.toMatch(/\.commit\(|\.approve\(/)
  })

  it('clearly separates simulated proposals from real document operations without invented evidence', () => {
    expect(source).toContain('Simulated agent · real document operations')
    expect(source).toContain('No LLM, API key, or external AI service.')
    expect(source).toContain('These guided tasks cover specific edits in bundled files, not unrestricted AI.')
    for (const step of ['Choose a task', 'Review the difference', 'Approve the edit', 'Download verified output']) expect(source).toContain(step)
    expect(source).not.toContain('sha256:')
    expect(source).not.toContain('value confirmed')
    expect(source).not.toContain('Open-source')
  })

  it('states the package-to-surface coverage without implying full Office parity', () => {
    expect(source).toContain('Filter sixteen focused surfaces across twenty-six packages')
    expect(source).toContain('not a claim of unrestricted Microsoft Office parity')
    expect(source).not.toMatch(/alternative to Google and Microsoft Office/i)
  })

  it('offers task, file type, and text filters with announced results and recovery', () => {
    expect(source).toContain('role="search"')
    expect(source).toContain('DEMO_TASKS.map')
    expect(source).toContain('DEMO_FORMATS.map')
    expect(source).toContain('showcaseItems(DEMOS)')
    expect(source).toContain('filterShowcaseItems')
    expect(source).toContain('aria-pressed={selected}')
    expect(source).toContain('aria-live="polite"')
    expect(source).toContain('No working proof matches those filters')
    expect(source).toContain('Show all proofs')
  })
})

import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'

const read = (path: string) => readFileSync(new URL(path, import.meta.url), 'utf8')

describe('document-first presentation', () => {
  it('places task controls in the assistant beside the same inspected artifact', () => {
    const page = read('./pages/AgentPage.tsx')
    const assistant = page.indexOf('aria-label="Document assistant"')
    const technical = page.indexOf('<details data-agent-technical')
    expect(page.indexOf('<ArtifactView')).toBeLessThan(assistant)
    expect(page.indexOf('<GuidedAgentTaskForm')).toBeGreaterThan(assistant)
    expect(page.indexOf('<GuidedAgentTaskForm')).toBeLessThan(technical)
    expect(page).toContain('content={shownContent}')
    expect(page).toContain('highlighted={Boolean(preview && validation?.ok)}')
    expect(page).toContain('Simulated agent. Real file edits. No LLM.')
    expect(page).toContain('hidden={!validation?.ok}')
    expect(page).toContain('hidden={!receipt && state !== \'committing\'}')
  })

  it('removes redundant shell labels and makes runtime details opt-in', () => {
    const shell = read('./App.tsx')
    for (const retired of ['demo-breadcrumb', 'demo-chips', 'demo-preview__bar', 'RuntimePill', 'Browser engines ready']) expect(shell).not.toContain(retired)
    expect(shell).toContain('<summary>About this demo</summary>')
    expect(shell).toContain('Checking optional server connection')
    expect(shell).not.toContain('sidecar-status--')
    expect(shell).toContain('<summary>Options</summary>')
  })

  it('keeps compact screen support and readable disclosures explicit', () => {
    const styles = read('./document-first.css')
    expect(styles).toContain('@media (max-width: 1100px)')
    expect(styles).toContain('@media (max-width: 760px)')
    expect(styles).toContain('grid-template-columns: minmax(0, 1fr) minmax(320px, 370px)')
    expect(styles).toContain('font-size: 14px; line-height: 1.6;')
    expect(styles).toContain('> section[hidden] { display: none; }')
  })
})

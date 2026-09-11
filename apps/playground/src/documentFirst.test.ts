import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'

const read = (path: string) => readFileSync(new URL(path, import.meta.url), 'utf8')

describe('document-first presentation', () => {
  it('leads XLSX with a file task and keeps configuration details collapsed', () => {
    const page = read('./pages/NativeRoundTripPage.tsx')
    expect(page).toContain('<h2>Open a spreadsheet</h2>')
    expect(page).toContain('Try sample spreadsheet')
    expect(page).toContain('<details className="ds-panel" data-xlsx-technical>')
    expect(page).toContain('<summary>Technical details</summary>')
    expect(page).toContain('{SERVER_FALLBACK_CONFIGURED\n')
    expect(page).toContain('A server API is configured.')
    expect(page).toContain('Server fallback is unavailable in this build.')
    expect(page).not.toContain('Server fallback stays disabled until')
    expect(page).not.toContain('Run the bundled proof')
    expect(page).toContain('Original bytes stay in this browser')
    expect(page).toContain('Uploads bytes to the configured API')
  })

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

  it('keeps the header free of demo disclosures and retains workspace options', () => {
    const shell = read('./App.tsx')
    for (const retired of ['demo-breadcrumb', 'demo-chips', 'demo-preview__bar', 'RuntimePill', 'Browser engines ready']) expect(shell).not.toContain(retired)
    expect(shell).not.toContain('About this demo')
    expect(shell).not.toContain('runtime-information')
    expect(shell).not.toContain('Checking optional server connection')
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

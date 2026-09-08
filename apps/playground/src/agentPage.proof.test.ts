import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { DEMOS } from './demoRegistry'
import { parseSurface } from './route'

describe('agent workflow page proof', () => {
  const page = readFileSync(new URL('./pages/AgentPage.tsx', import.meta.url), 'utf8')
  const runtime = readFileSync(new URL('./agentDemoRuntime.ts', import.meta.url), 'utf8')
  const guide = readFileSync(new URL('../../../docs/AGENT-CHANGESETS.md', import.meta.url), 'utf8')

  it('is routable, discoverable, and provider independent', () => {
    expect(parseSurface('#/agent')).toBe('agent')
    expect(DEMOS.find((demo) => demo.surface === 'agent')).toMatchObject({ packageName: '@injoffice/agent-tools', runtime: 'Browser' })
    expect(page).toContain('No language model, provider credentials, model service, or external AI service')
    expect(page).toContain('Simulated agent · real document operations')
    expect(page).not.toContain('Lifecycle simulation')
    expect(page).toContain('not a full-fidelity page or slide rendering')
    expect(page).toContain('not a rendered PDF page')
    for (const factory of ['createNativeAgentSessionInput', 'createNativeDocxAgentSessionInput', 'createNativePptxAgentSessionInput', 'createNativePdfAgentSessionInput']) expect(page).toContain(`await ${factory}(mode`)
    expect(page).toContain('terminateLoading()')
    expect(page).toContain('data-agent-download')
    expect(page).toContain('Run agent')
    expect(page).toContain('office.inspect')
    expect(page).toContain('office.plan')
    expect(page).toContain('AGENT_TOOLS')
    expect(page).toContain('label="Office tool"')
    expect(page).toContain('data-agent-tool={tool}')
    expect(runtime).toContain("from '@injoffice/agent-tools'")
    expect(runtime).toContain('createAgentSession({')
  })

  it('shows the complete guarded lifecycle and an explicit approval boundary', () => {
    for (const label of ['Inspect', 'Plan', 'Preview + diff', 'Validate', 'Approve', 'Commit', 'Verify']) expect(page).toContain(`'${label}'`)
    expect(page).toContain('I reviewed this exact diff')
    expect(page).toContain('Refused before write')
    expect(page).toContain('Source fingerprint')
    expect(page).toContain('Output fingerprint')
    expect(page).toContain("const updateRequest = useCallback((request: string) => { reset(); setPrompt(request) }, [reset])")
    expect(page).toContain("if (!changeSet || !inspection || !approved || !validation?.ok) return")
    expect(page).toContain("state === 'verified' && verification?.ok && downloadURL")
  })

  it('leads with guided tasks and keeps the public demo strictly mock-only', () => {
    expect(page).toContain('<GuidedAgentTaskForm')
    expect(page).toContain('context={proposalContext}')
    expect(page).toContain('onRequestChange={updateRequest}')
    expect(page).toContain('!proposalContext || !prompt.trim()')
    expect(page).toContain('data-agent-advanced-request')
    expect(page).toContain('readOnly={!advancedRequest || mode !== \'safe\'}')
    expect(page).toContain('Advanced request active')
    expect(page).toContain('setAdvancedRequest(false)')
    expect(page).toContain('reset(); setAdvancedRequest(event.target.checked)')
    expect(page).toContain('await requestMockAgentProposal(request, controller.signal)')
    for (const obsolete of ['proposalSource', 'liveContext', 'data-agent-live-consent', '/api/agent/propose', '/api/agent/proposal-status', 'fetch(']) expect(page).not.toContain(obsolete)
    for (const label of ['Choose a task', 'Inspect and preview', 'Review and approve', 'Verify and download']) expect(page).toContain(label)
    expect(page).toContain("Inspect: inspection ? 'done'")
    expect(page).toContain("Commit: receipt ? 'done'")
    expect(page).toContain("Verify: verification ? verification.ok ? 'done' : 'refused'")
  })

  it('keeps technical material collapsed without hiding failure or approval', () => {
    expect(page).toContain('<details data-agent-technical className="agent-technical">')
    expect(page).toContain('<summary>Technical details</summary>')
    expect(page).toContain('<details data-agent-safety-details')
    const technicalStart = page.indexOf('<details data-agent-technical')
    for (const marker of ['data-agent-request', 'aria-label="Actual office.* tool calls"', 'aria-label="Agent change set stages"', '<dt>Source fingerprint', 'data-agent-proposal-trace', 'data-agent-trace']) expect(page.indexOf(marker)).toBeGreaterThan(technicalStart)
    for (const marker of ['<section className="agent-refusal', '<section className="agent-approval', '<p className="tool-error" role="alert"', 'data-agent-download']) expect(page.indexOf(marker)).toBeLessThan(technicalStart)
  })

  it('keeps retained format sections independent of global hash navigation', () => {
    expect(page).toContain('fixedTool?: AgentTool')
    expect(page).toContain('if (fixedTool !== undefined) return')
    expect(page).toContain("if (parseSurface() === 'agent') setRoutedTool(parseAgentTool())")
    expect(page).toContain('const tool = fixedTool ?? routedTool')
    expect(page).toContain('{!fixedTool && <DsSegment')
    expect(page).toContain('const instanceId = useId()')
    expect(page).toContain('htmlFor={promptId}')
    expect(page).toContain('id={promptId}')
    expect(page).toContain('aria-labelledby={artifactTitleId}')
    expect(page).toContain('id={artifactTitleId}')
    expect(page).not.toContain('id="agent-prompt"')
    expect(page).not.toContain('id="agent-artifact-title"')
  })

  it('does not reset a retained spreadsheet view when an unrelated section changes the hash', () => {
    const sheets = readFileSync(new URL('./pages/SheetsPage.tsx', import.meta.url), 'utf8')
    expect(sheets).toContain("if (parseSurface() === 'sheets') setView(parseSheetsView())")
    expect(sheets).toContain('parseSheetsView(initialHash)')
    expect(sheets).toContain("window.removeEventListener('hashchange', syncView)")
  })

  it('documents the architecture, supported formats, safety, and non-goals', () => {
    for (const heading of ['## Architecture', '## Formats', '## Safety contract', '## Non-goals']) expect(guide).toContain(heading)
    for (const format of ['XLSX', 'DOCX', 'PPTX', 'PDF']) expect(guide).toContain(`**${format}:**`)
    expect(guide).toContain('do not contain a model SDK')
    for (const href of ['#/agent?format=sheets', '#/agent?format=docs', '#/agent?format=slides', '#/agent?format=pdf']) {
      expect(guide).toContain(href)
    }
  })
})

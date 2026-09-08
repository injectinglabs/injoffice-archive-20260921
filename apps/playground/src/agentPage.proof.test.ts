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
    expect(page).toContain('no model service')
    expect(page).toContain('Simulated AI proposal')
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

import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { DEMOS } from './demoRegistry'
import { parseSurface } from './route'

describe('agent workflow page proof', () => {
  const page = readFileSync(new URL('./pages/AgentPage.tsx', import.meta.url), 'utf8')
  const runtime = readFileSync(new URL('./agentDemoRuntime.ts', import.meta.url), 'utf8')
  const guide = readFileSync(new URL('../../docs/src/pages/AgentWorkflowsPage.tsx', import.meta.url), 'utf8')

  it('is routable, discoverable, and provider independent', () => {
    expect(parseSurface('#/agent')).toBe('agent')
    expect(DEMOS.find((demo) => demo.surface === 'agent')).toMatchObject({ packageName: '@injoffice/agent-tools', runtime: 'Browser' })
    expect(page).toContain('no model SDK or network request')
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
    for (const heading of ['architecture', 'capabilities', 'safety', 'non-goals']) expect(guide).toContain(`id="${heading}"`)
    for (const format of ['XLSX', 'DOCX', 'PPTX', 'PDF']) expect(guide).toContain(`<strong>${format}:</strong>`)
    expect(guide).toContain('do not contain a model SDK')
  })
})

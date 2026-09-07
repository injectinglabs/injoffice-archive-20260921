import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'

const source = readFileSync(new URL('./pages/OverviewPage.tsx', import.meta.url), 'utf8')

describe('overview native workflow', () => {
  it('leads with the source-authoritative workflow and opens the native XLSX proof directly', () => {
    expect(source).toContain('Original Office bytes stay authoritative')
    expect(source).toContain("${surfaceHref('sheets')}?view=native")
    expect(source).toContain('Reopen the exact returned bytes')
    expect(source).toContain('Unsafe or stale changes produce no replacement file')
  })

  it('states the package-to-surface coverage without implying full Office parity', () => {
    expect(source).toContain('<dt>24</dt><dd>composable packages</dd>')
    expect(source).toContain('<dt>15</dt><dd>proof surfaces</dd>')
    expect(source).toContain('Shared infrastructure appears inside the workflows it powers')
    expect(source).toContain('not a claim of unrestricted Microsoft Office parity')
    expect(source).not.toMatch(/alternative to Google and Microsoft Office/i)
  })

  it('offers task, file type, and text filters with announced results and recovery', () => {
    expect(source).toContain('role="search"')
    expect(source).toContain('DEMO_TASKS.map')
    expect(source).toContain('DEMO_FORMATS.map')
    expect(source).toContain('aria-pressed={selected}')
    expect(source).toContain('aria-live="polite"')
    expect(source).toContain('No working proof matches those filters')
    expect(source).toContain('Show all proofs')
  })
})

import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'

describe('public site metadata', () => {
  it('uses the AI-first positioning and the correct Google Workspace name', () => {
    const html = readFileSync(new URL('../index.html', import.meta.url), 'utf8')
    expect(html).toContain('<title>InjOffice — AI-first open-source alternative to Microsoft Office and Google Workspace</title>')
    expect(html).toContain('<meta name="description" content="InjOffice is an AI-first, open-source alternative to Microsoft Office and Google Workspace for spreadsheets, documents, slides, and PDFs." />')
    expect(html).not.toContain('Google Office')
  })
})

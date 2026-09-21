import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'

describe('public site metadata', () => {
  it('describes the document libraries and supported formats', () => {
    const html = readFileSync(new URL('../index.html', import.meta.url), 'utf8')
    expect(html).toContain('<title>InjOffice — Open-source document editing libraries</title>')
    expect(html).toContain('<meta name="description" content="Open-source TypeScript and Go libraries for XLSX, DOCX, PPTX, and PDF editing, with agent APIs and browser-local processing." />')
  })
})

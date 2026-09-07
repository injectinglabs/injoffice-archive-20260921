import { describe, expect, it } from 'vitest'
import { DEMOS } from './demoRegistry'
import { filterShowcaseDemos, filterShowcaseItems, showcaseItems } from './showcaseCatalog'

describe('showcase catalog filtering', () => {
  it('finds proofs by user-language keywords as well as titles and packages', () => {
    expect(filterShowcaseDemos(DEMOS, { query: 'redact', task: 'All tasks', format: 'All formats' }).map((demo) => demo.surface)).toEqual(['pdf'])
    expect(filterShowcaseDemos(DEMOS, { query: 'native mutation download', task: 'All tasks', format: 'All formats' }).map((demo) => demo.surface)).toEqual(['sheets', 'docs', 'pptx-native'])
    expect(filterShowcaseDemos(DEMOS, { query: '@injoffice/history', task: 'All tasks', format: 'All formats' }).map((demo) => demo.surface)).toEqual(['history'])
  })

  it('combines task, format, and text filters', () => {
    const results = filterShowcaseDemos(DEMOS, { query: 'text', task: 'Review changes', format: 'DOCX' })
    expect(results.map((demo) => demo.surface)).toEqual(['docs', 'history', 'font-metrics'])
  })

  it('treats blank and whitespace-only queries as the full catalog', () => {
    expect(filterShowcaseDemos(DEMOS, { query: '   ', task: 'All tasks', format: 'All formats' })).toHaveLength(DEMOS.length)
  })

  it('exposes a distinct AI change-set proof for sheets, docs, slides, and pdf', () => {
    const items = showcaseItems(DEMOS)
    expect(items.filter((item) => item.surface === 'agent').map((item) => item.key)).toEqual([
      'agent-sheets',
      'agent-docs',
      'agent-slides',
      'agent-pdf',
    ])
    expect(filterShowcaseItems(items, { query: '', task: 'All tasks', format: 'XLSX' }).some((item) => item.key === 'agent-sheets')).toBe(true)
    expect(filterShowcaseItems(items, { query: '', task: 'All tasks', format: 'XLSX' }).some((item) => item.key === 'agent-docs')).toBe(false)
    expect(filterShowcaseItems(items, { query: '', task: 'All tasks', format: 'DOCX' }).some((item) => item.key === 'agent-docs')).toBe(true)
    expect(filterShowcaseItems(items, { query: '', task: 'All tasks', format: 'PPTX' }).some((item) => item.key === 'agent-slides')).toBe(true)
    expect(filterShowcaseItems(items, { query: '', task: 'All tasks', format: 'PDF' }).some((item) => item.key === 'agent-pdf')).toBe(true)
  })
})

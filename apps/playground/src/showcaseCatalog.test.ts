import { describe, expect, it } from 'vitest'
import { DEMOS } from './demoRegistry'
import { filterShowcaseDemos } from './showcaseCatalog'

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
})

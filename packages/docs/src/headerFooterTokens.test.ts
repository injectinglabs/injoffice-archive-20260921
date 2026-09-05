import { describe, expect, it } from 'vitest'
import { hasHeaderFooterTokens, renderHeaderFooterTemplate } from './headerFooterTokens'

describe('renderHeaderFooterTemplate', () => {
  const ctx = { page: 3, totalPages: 12, date: new Date(2026, 7, 25, 15, 4), filename: 'report.md' }

  it('substitutes &P and &N', () => {
    expect(renderHeaderFooterTemplate('Page &P of &N', ctx)).toBe('Page 3 of 12')
  })

  it('substitutes &D as a short date', () => {
    expect(renderHeaderFooterTemplate('&D', ctx)).toBe('8/25/2026')
  })

  it('substitutes &T as a 12-hour time', () => {
    expect(renderHeaderFooterTemplate('&T', ctx)).toBe('3:04 PM')
  })

  it('substitutes &F as the filename', () => {
    expect(renderHeaderFooterTemplate('&F', ctx)).toBe('report.md')
  })

  it('is case-insensitive on the token letter', () => {
    expect(renderHeaderFooterTemplate('&p/&n', ctx)).toBe('3/12')
  })

  it('leaves unrecognized &-codes untouched', () => {
    expect(renderHeaderFooterTemplate('&Z stays &Z', ctx)).toBe('&Z stays &Z')
  })

  it('handles a template with no tokens', () => {
    expect(renderHeaderFooterTemplate('Confidential', ctx)).toBe('Confidential')
  })

  it('empty template renders empty', () => {
    expect(renderHeaderFooterTemplate('', ctx)).toBe('')
  })

  it('defaults filename to empty string when absent', () => {
    expect(renderHeaderFooterTemplate('&F', { page: 1, totalPages: 1 })).toBe('')
  })

  it('defaults date to now when absent (does not throw, produces a date string)', () => {
    expect(renderHeaderFooterTemplate('&D', { page: 1, totalPages: 1 })).toMatch(/^\d{1,2}\/\d{1,2}\/\d{4}$/)
  })
})

describe('hasHeaderFooterTokens', () => {
  it('detects a token', () => {
    expect(hasHeaderFooterTokens('Page &P')).toBe(true)
  })
  it('reports false for plain text', () => {
    expect(hasHeaderFooterTokens('Confidential')).toBe(false)
  })
  it('is safe to call repeatedly (no stateful lastIndex leak)', () => {
    expect(hasHeaderFooterTokens('&P')).toBe(true)
    expect(hasHeaderFooterTokens('&P')).toBe(true)
    expect(hasHeaderFooterTokens('&P')).toBe(true)
  })
})

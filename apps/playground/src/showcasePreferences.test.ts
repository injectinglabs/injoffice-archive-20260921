import { describe, expect, it } from 'vitest'
import { persistShowcasePreferences, readShowcasePreferences } from './showcasePreferences'

function memoryStorage(value: string | null = null) {
  return { getItem: () => value, setItem: (_key: string, next: string) => { value = next } }
}

describe('showcase preferences', () => {
  it('preserves search, task, and format across catalogue remounts', () => {
    const storage = memoryStorage()
    const filters = { query: 'native', task: 'Review changes', format: 'XLSX' } as const
    persistShowcasePreferences(filters, storage)
    expect(readShowcasePreferences(storage)).toEqual(filters)
  })

  it('persists reset filters', () => {
    const storage = memoryStorage()
    persistShowcasePreferences({ query: 'native', task: 'Edit files', format: 'XLSX' }, storage)
    const defaults = { query: '', task: 'All tasks', format: 'All formats' } as const
    persistShowcasePreferences(defaults, storage)
    expect(readShowcasePreferences(storage)).toEqual(defaults)
  })

  it('falls back for corrupt, outdated, or unavailable storage', () => {
    const defaults = { query: '', task: 'All tasks', format: 'All formats' } as const
    expect(readShowcasePreferences(memoryStorage('not json'))).toEqual(defaults)
    expect(readShowcasePreferences(memoryStorage('{"query":42,"task":"old","format":"old"}'))).toEqual(defaults)
    const blocked = { getItem: () => { throw new Error('disabled') }, setItem: () => { throw new Error('full') } }
    expect(readShowcasePreferences(blocked)).toEqual(defaults)
    expect(() => persistShowcasePreferences(defaults, blocked)).not.toThrow()
  })
})

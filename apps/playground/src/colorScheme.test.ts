import { describe, expect, it } from 'vitest'
import { COLOR_SCHEME_STORAGE_KEY, parseColorScheme, resolveColorScheme } from './colorScheme'
import { univerDarkMode } from './univerColorScheme'

describe('playground color scheme', () => {
  it('accepts only light and dark stored values', () => {
    expect(parseColorScheme('light')).toBe('light')
    expect(parseColorScheme('dark')).toBe('dark')
    expect(parseColorScheme('system')).toBeNull()
    expect(parseColorScheme('')).toBeNull()
    expect(parseColorScheme(null)).toBeNull()
  })

  it('uses an explicit choice and otherwise follows the OS preference', () => {
    expect(resolveColorScheme('dark', false)).toBe('dark')
    expect(resolveColorScheme('light', true)).toBe('light')
    expect(resolveColorScheme(null, true)).toBe('dark')
    expect(resolveColorScheme(undefined, false)).toBe('light')
  })

  it('persists under a playground-specific key', () => {
    expect(COLOR_SCHEME_STORAGE_KEY).toBe('injoffice-playground-theme')
  })

  it('maps the playground scheme onto Univer darkMode', () => {
    expect(univerDarkMode('dark')).toBe(true)
    expect(univerDarkMode('light')).toBe(false)
  })
})

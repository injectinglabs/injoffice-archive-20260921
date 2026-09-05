import type { DeckTheme } from './types'

// Built-in themes. "boardroom" is the neutral default; hosts register their
// brand as a custom DeckTheme (the product injects its own).

export const boardroomTheme: DeckTheme = {
  id: 'boardroom',
  background: '#ffffff',
  surface: '#f4f6f5',
  ink: '#172426',
  muted: '#5c6b6d',
  accent: '#0f9f88',
  displayFont: "'Archivo', 'Helvetica Neue', Arial, sans-serif",
  bodyFont: "'Source Serif 4', Georgia, serif",
  monoFont: "'IBM Plex Mono', ui-monospace, monospace",
}

export const midnightTheme: DeckTheme = {
  id: 'midnight',
  background: '#131a1b',
  surface: '#1b2425',
  ink: '#e6ecea',
  muted: '#93a3a1',
  accent: '#2cc2a8',
  displayFont: "'Archivo', 'Helvetica Neue', Arial, sans-serif",
  bodyFont: "'Source Serif 4', Georgia, serif",
  monoFont: "'IBM Plex Mono', ui-monospace, monospace",
}

export const slateTheme: DeckTheme = {
  id: 'slate',
  background: '#f2f4f7',
  surface: '#e5e9ef',
  ink: '#1f2733',
  muted: '#5d6b7d',
  accent: '#3b6fe0',
  displayFont: "'Archivo', 'Helvetica Neue', Arial, sans-serif",
  bodyFont: "'Source Serif 4', Georgia, serif",
  monoFont: "'IBM Plex Mono', ui-monospace, monospace",
}

export const terraTheme: DeckTheme = {
  id: 'terra',
  background: '#f7f1e8',
  surface: '#efe5d6',
  ink: '#33261a',
  muted: '#7d6b57',
  accent: '#c05621',
  displayFont: "'Archivo', 'Helvetica Neue', Arial, sans-serif",
  bodyFont: "'Source Serif 4', Georgia, serif",
  monoFont: "'IBM Plex Mono', ui-monospace, monospace",
}

export const forestTheme: DeckTheme = {
  id: 'forest',
  background: '#0f1f18',
  surface: '#1a2b23',
  ink: '#e5eee8',
  muted: '#8fa79a',
  accent: '#4fd58b',
  displayFont: "'Archivo', 'Helvetica Neue', Arial, sans-serif",
  bodyFont: "'Source Serif 4', Georgia, serif",
  monoFont: "'IBM Plex Mono', ui-monospace, monospace",
}

export const plumTheme: DeckTheme = {
  id: 'plum',
  background: '#ffffff',
  surface: '#f4eff9',
  ink: '#241c2e',
  muted: '#6b6076',
  accent: '#7c4dbc',
  displayFont: "'Archivo', 'Helvetica Neue', Arial, sans-serif",
  bodyFont: "'Source Serif 4', Georgia, serif",
  monoFont: "'IBM Plex Mono', ui-monospace, monospace",
}

export const BUILTIN_THEMES: readonly DeckTheme[] = [boardroomTheme, midnightTheme, slateTheme, terraTheme, forestTheme, plumTheme]

/** Resolve a spec's theme field to concrete tokens (default: boardroom). */
export function resolveTheme(theme: string | DeckTheme | undefined): DeckTheme {
  if (!theme) return boardroomTheme
  if (typeof theme === 'object') return { ...boardroomTheme, ...theme }
  return BUILTIN_THEMES.find((t) => t.id === theme) ?? boardroomTheme
}

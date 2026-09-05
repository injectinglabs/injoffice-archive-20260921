export const COLOR_SCHEME_STORAGE_KEY = 'injoffice-playground-theme'
export const COLOR_SCHEMES = ['light', 'dark'] as const
export type ColorScheme = (typeof COLOR_SCHEMES)[number]

const THEME_COLOR: Record<ColorScheme, string> = {
  light: '#151b24',
  dark: '#10151c',
}

export function parseColorScheme(value: string | null | undefined): ColorScheme | null {
  return value === 'light' || value === 'dark' ? value : null
}

export function resolveColorScheme(stored: string | null | undefined, prefersDark: boolean): ColorScheme {
  return parseColorScheme(stored) ?? (prefersDark ? 'dark' : 'light')
}

export function readStoredColorScheme(): ColorScheme | null {
  try {
    return parseColorScheme(window.localStorage.getItem(COLOR_SCHEME_STORAGE_KEY))
  } catch {
    return null
  }
}

export function preferredColorScheme(): ColorScheme {
  return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light'
}

export function queryColorScheme(): ColorScheme | null {
  try {
    return parseColorScheme(new URLSearchParams(window.location.search).get('theme'))
  } catch {
    return null
  }
}

export function currentColorScheme(): ColorScheme {
  return resolveColorScheme(
    queryColorScheme() ?? readStoredColorScheme(),
    preferredColorScheme() === 'dark',
  )
}

const listeners = new Set<(scheme: ColorScheme) => void>()

export function subscribeColorScheme(listener: (scheme: ColorScheme) => void): () => void {
  listeners.add(listener)
  return () => { listeners.delete(listener) }
}

export function applyColorScheme(scheme: ColorScheme): void {
  const root = document.documentElement
  root.dataset.theme = scheme
  root.style.colorScheme = scheme
  document.querySelector('meta[name="theme-color"]')?.setAttribute('content', THEME_COLOR[scheme])
  for (const listener of listeners) listener(scheme)
}

export function persistColorScheme(scheme: ColorScheme): void {
  try {
    window.localStorage.setItem(COLOR_SCHEME_STORAGE_KEY, scheme)
  } catch {
    /* private mode */
  }
  applyColorScheme(scheme)
}

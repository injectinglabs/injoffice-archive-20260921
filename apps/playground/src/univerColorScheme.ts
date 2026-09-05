import { currentColorScheme, subscribeColorScheme, type ColorScheme } from './colorScheme'

export function univerDarkMode(scheme: ColorScheme = currentColorScheme()): boolean {
  return scheme === 'dark'
}

/** Keep a live Univer instance on the playground Light/Dark toggle. */
export function bindUniverColorScheme(api: { toggleDarkMode(isDarkMode: boolean): void }): () => void {
  api.toggleDarkMode(univerDarkMode())
  return subscribeColorScheme((scheme) => {
    api.toggleDarkMode(univerDarkMode(scheme))
  })
}

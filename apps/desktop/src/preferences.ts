export type ViewOptions = { zoom: number; navigation: boolean; focus: boolean };
export type ThemePreference = 'system' | 'light' | 'dark';
export const themePreferences: readonly ThemePreference[] = ['system', 'light', 'dark'];
export type Preferences = { version: 1; defaultZoom: number; showNavigation: boolean; theme: ThemePreference };
export const defaultPreferences: Preferences = { version: 1, defaultZoom: 100, showNavigation: true, theme: 'system' };
function isTheme(value: unknown): value is ThemePreference { return (themePreferences as readonly unknown[]).includes(value); }
export function readPreferences(): Preferences {
  try {
    const value = JSON.parse(window.localStorage.getItem('injoffice.preferences.v1') ?? 'null');
    // Blobs written before the theme preference existed carry no theme and follow the system; any other unknown value fails closed.
    const theme = value?.theme === undefined ? 'system' : value.theme;
    if (value?.version === 1 && Number.isInteger(value.defaultZoom) && value.defaultZoom >= 50 && value.defaultZoom <= 200 && typeof value.showNavigation === 'boolean' && isTheme(theme)) return { version: 1, defaultZoom: value.defaultZoom, showNavigation: value.showNavigation, theme };
  } catch { /* Missing or damaged preferences never prevent opening documents. */ }
  return { ...defaultPreferences };
}
export function writePreferences(preferences: Preferences) { window.localStorage.setItem('injoffice.preferences.v1', JSON.stringify(preferences)); }
export function initialView(preferences: Preferences): ViewOptions { return { zoom: preferences.defaultZoom, navigation: preferences.showNavigation, focus: false }; }

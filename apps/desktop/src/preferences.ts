export type ViewOptions = { zoom: number; navigation: boolean; focus: boolean };
export type Preferences = { version: 1; defaultZoom: number; showNavigation: boolean };
export const defaultPreferences: Preferences = { version: 1, defaultZoom: 100, showNavigation: true };
export function readPreferences(): Preferences {
  try {
    const value = JSON.parse(window.localStorage.getItem('injoffice.preferences.v1') ?? 'null');
    if (value?.version === 1 && Number.isInteger(value.defaultZoom) && value.defaultZoom >= 50 && value.defaultZoom <= 200 && typeof value.showNavigation === 'boolean') return { version: 1, defaultZoom: value.defaultZoom, showNavigation: value.showNavigation };
  } catch { /* Missing or damaged preferences never prevent opening documents. */ }
  return { ...defaultPreferences };
}
export function writePreferences(preferences: Preferences) { window.localStorage.setItem('injoffice.preferences.v1', JSON.stringify(preferences)); }
export function initialView(preferences: Preferences): ViewOptions { return { zoom: preferences.defaultZoom, navigation: preferences.showNavigation, focus: false }; }

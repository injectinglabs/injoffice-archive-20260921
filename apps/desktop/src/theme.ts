import type { ThemePreference } from './preferences';
type ThemeRoot = { dataset: DOMStringMap };
// styles.css follows prefers-color-scheme unless <html data-theme> pins light or dark; "system" removes the pin.
export function applyTheme(theme: ThemePreference, root: ThemeRoot | undefined = globalThis.document?.documentElement): void {
  if (!root) return;
  if (theme === 'system') delete root.dataset.theme;
  else root.dataset.theme = theme;
}

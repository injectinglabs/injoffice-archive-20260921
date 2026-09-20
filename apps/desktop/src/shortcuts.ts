// Single source of truth for the keyboard shortcuts the desktop editors honour.
// Ribbon tooltips, command hints and aria-keyshortcuts all read from here so the
// same binding is never spelled twice.
export type ShortcutId = 'undo' | 'redo' | 'bold' | 'italic' | 'underline' | 'find' | 'save' | 'saveAs' | 'open' | 'new' | 'commands' | 'apply' | 'confirm' | 'cancel'

interface Binding { key: string; mod?: boolean; shift?: boolean; other?: Partial<Binding> }

// `mod` is ⌘ on macOS and Ctrl elsewhere. `other` overrides the binding on non-mac platforms.
const bindings: Record<ShortcutId, Binding> = {
  undo: { key: 'Z', mod: true },
  redo: { key: 'Z', mod: true, shift: true, other: { key: 'Y', shift: false } },
  bold: { key: 'B', mod: true },
  italic: { key: 'I', mod: true },
  underline: { key: 'U', mod: true },
  find: { key: 'F', mod: true },
  save: { key: 'S', mod: true },
  saveAs: { key: 'S', mod: true, shift: true },
  open: { key: 'O', mod: true },
  new: { key: 'N', mod: true },
  commands: { key: 'K', mod: true },
  apply: { key: 'Enter', mod: true },
  confirm: { key: 'Enter' },
  cancel: { key: 'Escape' },
}

export type ShortcutPlatform = 'mac' | 'other'

export function shortcutPlatform(): ShortcutPlatform {
  if (typeof navigator === 'undefined') return 'mac'
  return /Mac|iPhone|iPad/.test(navigator.platform || navigator.userAgent || '') ? 'mac' : 'other'
}

const macKeys: Record<string, string> = { Enter: '↩', Escape: 'Esc' }
const otherKeys: Record<string, string> = { Escape: 'Esc' }

function binding(id: ShortcutId, platform: ShortcutPlatform): Binding {
  const base = bindings[id]
  return platform === 'other' && base.other ? { ...base, ...base.other } : base
}

/** Human-readable label, e.g. "⌘B" on macOS or "Ctrl+B" elsewhere. */
export function shortcutLabel(id: ShortcutId, platform: ShortcutPlatform = shortcutPlatform()): string {
  const { key, mod, shift } = binding(id, platform)
  if (platform === 'mac') return `${mod ? '⌘' : ''}${shift ? '⇧' : ''}${macKeys[key] ?? key}`
  return [mod && 'Ctrl', shift && 'Shift', otherKeys[key] ?? key].filter(Boolean).join('+')
}

/** WAI-ARIA `aria-keyshortcuts` value for the binding. */
export function shortcutKeys(id: ShortcutId, platform: ShortcutPlatform = shortcutPlatform()): string {
  const { key, mod, shift } = binding(id, platform)
  return [mod && (platform === 'mac' ? 'Meta' : 'Control'), shift && 'Shift', key].filter(Boolean).join('+')
}

/** Tooltip text: the label followed by the shortcut in parentheses when one exists. */
export function shortcutTooltip(label: string, id?: ShortcutId, platform: ShortcutPlatform = shortcutPlatform()): string {
  return id ? `${label} (${shortcutLabel(id, platform)})` : label
}

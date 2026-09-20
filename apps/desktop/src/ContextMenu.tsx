import { useCallback, useEffect, useLayoutEffect, useRef, useState, type MouseEvent as ReactMouseEvent, type KeyboardEvent as ReactKeyboardEvent } from 'react'
import type { FormattingPatch, FormattingValues } from './FormattingToolbar'
import './context-menu.css'

/** One entry of an Office-style context menu. Disabled entries keep their explanatory title, exactly like the toolbars. */
export type ContextMenuItem =
  | { id: string; label: string; shortcut?: string; disabled?: boolean; title?: string; checked?: boolean; run(): void }
  | { separator: true }
export type ContextMenuAnchor = { x: number; y: number; source: HTMLElement | null }

/** Shortcut hints in the workspace's "Ctrl / ⌘ K" spelling used by the command palette details in App.tsx. */
export const SHORTCUTS = {
  cut: 'Ctrl / ⌘ X', copy: 'Ctrl / ⌘ C', paste: 'Ctrl / ⌘ V', bold: 'Ctrl / ⌘ B', italic: 'Ctrl / ⌘ I', underline: 'Ctrl / ⌘ U',
  find: 'Ctrl / ⌘ F', undo: 'Ctrl / ⌘ Z', redo: 'Ctrl / ⌘ Y', apply: 'Ctrl / ⌘ Enter', cancel: 'Esc',
} as const
/** The renderer cannot read the clipboard: the host denies clipboard permissions and execCommand('paste') is unsupported. */
export const PASTE_UNAVAILABLE = 'Paste with Ctrl / ⌘ V; menu paste is not permitted by the app clipboard policy'
export const XLSX_STRUCTURE_UNAVAILABLE = 'Row and column insert/delete are not supported by the native XLSX transaction'
export const PPTX_CLIPBOARD_UNAVAILABLE = 'Slide objects have no clipboard in this editor'

const MARGIN = 8
/** Keep the whole menu inside the viewport, flipping above / left of the pointer when there is no room. */
export function clampMenuPosition(x: number, y: number, width: number, height: number, viewportWidth: number, viewportHeight: number): { x: number; y: number } {
  let left = x, top = y
  if (left + width + MARGIN > viewportWidth) left = Math.max(MARGIN, Math.min(x - width, viewportWidth - width - MARGIN))
  if (top + height + MARGIN > viewportHeight) top = Math.max(MARGIN, Math.min(y - height, viewportHeight - height - MARGIN))
  return { x: Math.max(MARGIN, left), y: Math.max(MARGIN, top) }
}
/** Next enabled entry for arrow navigation; separators and disabled entries are skipped and the walk wraps. */
export function nextMenuIndex(items: readonly ContextMenuItem[], current: number, delta: 1 | -1): number {
  if (!items.length) return -1
  let index = current
  for (let step = 0; step < items.length; step++) {
    index = (index + delta + items.length) % items.length
    const item = items[index]!
    if (!('separator' in item) && !item.disabled) return index
  }
  return current
}
function isEnabled(item: ContextMenuItem | undefined): item is Exclude<ContextMenuItem, { separator: true }> { return !!item && !('separator' in item) && !item.disabled }

/** Anchor state for an editor: `open` goes on `onContextMenu`, the anchor renders `<ContextMenu>`. */
export function useContextMenu() {
  const [anchor, setAnchor] = useState<ContextMenuAnchor>()
  const open = useCallback((event: ReactMouseEvent<HTMLElement>) => { event.preventDefault(); setAnchor({ x: event.clientX, y: event.clientY, source: event.currentTarget }) }, [])
  const close = useCallback(() => setAnchor(undefined), [])
  return { anchor, open, close }
}

export default function ContextMenu({ anchor, items, label, onClose }: { anchor: ContextMenuAnchor; items: ContextMenuItem[]; label: string; onClose(): void }) {
  const menu = useRef<HTMLDivElement>(null)
  const [position, setPosition] = useState({ x: anchor.x, y: anchor.y })
  const restore = useRef<Element | null>(null)
  const closeRef = useRef(onClose); closeRef.current = onClose
  function close() {
    const previous = restore.current as HTMLElement | null
    closeRef.current()
    if (previous && typeof previous.focus === 'function' && previous.isConnected) previous.focus({ preventScroll: true })
  }
  useLayoutEffect(() => {
    const node = menu.current
    if (!node || typeof window === 'undefined') return
    restore.current = window.document.activeElement
    const rect = node.getBoundingClientRect()
    setPosition(clampMenuPosition(anchor.x, anchor.y, rect.width, rect.height, window.innerWidth, window.innerHeight))
    node.querySelector<HTMLElement>('[role^="menuitem"]:not([aria-disabled="true"])')?.focus({ preventScroll: true })
  }, [anchor])
  useEffect(() => {
    if (typeof window === 'undefined') return
    const pointer = (event: PointerEvent) => { if (!menu.current?.contains(event.target as Node)) closeRef.current() }
    const escape = (event: KeyboardEvent) => { if (event.key === 'Escape') { event.preventDefault(); closeRef.current() } }
    const dismiss = () => closeRef.current()
    window.document.addEventListener('pointerdown', pointer, true)
    window.document.addEventListener('keydown', escape, true)
    window.document.addEventListener('scroll', dismiss, true)
    window.addEventListener('resize', dismiss); window.addEventListener('blur', dismiss)
    return () => {
      window.document.removeEventListener('pointerdown', pointer, true); window.document.removeEventListener('keydown', escape, true)
      window.document.removeEventListener('scroll', dismiss, true); window.removeEventListener('resize', dismiss); window.removeEventListener('blur', dismiss)
    }
  }, [])
  function activate(item: ContextMenuItem) {
    if (!isEnabled(item)) return
    close()
    item.run()
  }
  function keys(event: ReactKeyboardEvent<HTMLDivElement>) {
    const buttons = Array.from(menu.current?.querySelectorAll<HTMLElement>('[role^="menuitem"]') ?? [])
    const focused = buttons.findIndex(button => button === window.document.activeElement)
    const focus = (index: number) => { if (index >= 0) (menu.current?.querySelector<HTMLElement>(`[data-index="${index}"]`) ?? buttons[index])?.focus({ preventScroll: true }) }
    if (event.key === 'ArrowDown' || event.key === 'ArrowUp') { event.preventDefault(); event.stopPropagation(); focus(nextMenuIndex(items, focused < 0 ? (event.key === 'ArrowDown' ? -1 : 0) : Number(buttons[focused]!.dataset.index), event.key === 'ArrowDown' ? 1 : -1)) }
    else if (event.key === 'Home') { event.preventDefault(); event.stopPropagation(); focus(nextMenuIndex(items, -1, 1)) }
    else if (event.key === 'End') { event.preventDefault(); event.stopPropagation(); focus(nextMenuIndex(items, 0, -1)) }
    else if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); close() }
    else if (event.key === 'Tab') close()
  }
  return <div ref={menu} className="context-menu" role="menu" aria-label={label} tabIndex={-1} style={{ left: position.x, top: position.y }} onKeyDown={keys} onContextMenu={event => event.preventDefault()}>
    {items.map((item, index) => 'separator' in item ? <div key={`separator-${index}`} role="separator" className="context-menu-separator" /> :
      <button key={item.id} type="button" role={item.checked === undefined ? 'menuitem' : 'menuitemcheckbox'} aria-checked={item.checked} aria-disabled={item.disabled || undefined} data-index={index} title={item.title} tabIndex={-1}
        onClick={() => activate(item)} onPointerEnter={event => { if (!item.disabled) event.currentTarget.focus({ preventScroll: true }) }}>
        <span className="context-menu-label">{item.label}</span>{item.shortcut && <kbd className="context-menu-shortcut">{item.shortcut}</kbd>}
      </button>)}
  </div>
}

// ---- Per-editor item builders. Every entry drives a command the editor already exposes. ----

function selectionState(): { selected: boolean; inEditable: boolean } {
  if (typeof window === 'undefined') return { selected: false, inEditable: false }
  const selection = window.getSelection()
  if (!selection?.rangeCount || selection.isCollapsed) return { selected: false, inEditable: false }
  const node = selection.getRangeAt(0).commonAncestorContainer
  const element = node.nodeType === 1 ? node as Element : node.parentElement
  return { selected: true, inEditable: !!element?.closest('.office-inline-input') }
}
function command(name: 'cut' | 'copy') { if (typeof window !== 'undefined') window.document.execCommand(name) }
/** Reveals a ribbon disclosure control (Insert link / Insert table) after switching to its ribbon tab. */
function openRibbonControl(source: HTMLElement | null, selector: string) {
  if (typeof window === 'undefined') return
  window.requestAnimationFrame(() => { const button = source?.closest('.office-editor')?.querySelector<HTMLElement>(selector); if (button?.getAttribute('aria-expanded') === 'false') button.click(); button?.focus() })
}
/** Right-click on an inactive text run makes it the active segment, like a left click at the same point would. */
export function activateRunAt(event: ReactMouseEvent<HTMLElement>) {
  const run = (event.target as Element | null)?.closest?.<HTMLElement>('.office-text-run')
  if (!run || run.getAttribute('aria-disabled') === 'true') return
  run.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, clientX: event.clientX, clientY: event.clientY }))
}
export interface DocumentMenuContext {
  anchor: ContextMenuAnchor
  target: boolean
  values?: FormattingValues
  disabled: boolean
  link: boolean
  table: boolean
  onFormat(patch: FormattingPatch): void
  onFind(): void
  onRibbonTab(tab: 'Insert'): void
}
export function documentContextMenu(context: DocumentMenuContext): ContextMenuItem[] {
  const { values, disabled } = context
  const selection = selectionState()
  const character = disabled || !values || values.characterEditable === false
  const alignment = (value: string, label: string): ContextMenuItem => ({ id: `align-${value}`, label, disabled: disabled || !values, checked: values?.alignment === value, run: () => context.onFormat({ alignment: value }) })
  return [
    { id: 'cut', label: 'Cut', shortcut: SHORTCUTS.cut, disabled: disabled || !selection.inEditable, title: selection.inEditable ? undefined : 'Select text inside the segment you are editing to cut', run: () => command('cut') },
    { id: 'copy', label: 'Copy', shortcut: SHORTCUTS.copy, disabled: !selection.selected, title: selection.selected ? undefined : 'Select text to copy', run: () => command('copy') },
    { id: 'paste', label: 'Paste', shortcut: SHORTCUTS.paste, disabled: true, title: PASTE_UNAVAILABLE, run: () => {} },
    { separator: true },
    { id: 'bold', label: 'Bold', shortcut: SHORTCUTS.bold, disabled: character, checked: !!values?.bold, run: () => context.onFormat({ bold: !values?.bold }) },
    { id: 'italic', label: 'Italic', shortcut: SHORTCUTS.italic, disabled: character, checked: !!values?.italic, run: () => context.onFormat({ italic: !values?.italic }) },
    { id: 'underline', label: 'Underline', shortcut: SHORTCUTS.underline, disabled: character, checked: !!values?.underline, run: () => context.onFormat({ underline: !values?.underline }) },
    { separator: true },
    alignment('left', 'Align left'), alignment('center', 'Center'), alignment('right', 'Align right'), alignment('both', 'Justify'),
    { separator: true },
    { id: 'hyperlink', label: 'Hyperlink…', disabled: disabled || !context.link, title: context.link ? undefined : 'Place the caret in a linkable text segment', run: () => { context.onRibbonTab('Insert'); openRibbonControl(context.anchor.source, '.office-link-control > button') } },
    { id: 'table', label: 'Insert table…', disabled: disabled || !context.table, title: context.table ? undefined : 'Select a body paragraph to insert a table after it', run: () => { context.onRibbonTab('Insert'); openRibbonControl(context.anchor.source, '.document-insert-table > button') } },
    { separator: true },
    { id: 'find', label: 'Find / replace', shortcut: SHORTCUTS.find, run: context.onFind },
  ]
}

/** Cell address under a right-click on the worksheet grid, read from the cell's title ("B3" or "B3 · merged …"). */
export function contextMenuCellKey(target: EventTarget | null): string | undefined {
  const cell = (target as Element | null)?.closest?.<HTMLElement>('td[role="gridcell"]')
  const key = cell?.title.split(' ')[0]
  return key && /^[A-Z]{1,3}[1-9]\d*$/.test(key) ? key : undefined
}
export interface SpreadsheetMenuContext {
  anchor: ContextMenuAnchor
  disabled: boolean
  onClear(): void
}
/** Opens a toolbar disclosure (for example "Cell size") and focuses the field inside it. */
function revealToolbarField(source: HTMLElement | null, summary: string, selector: string) {
  if (typeof window === 'undefined') return
  const editor = source?.closest('.sheet-editor')
  const details = Array.from(editor?.querySelectorAll('details') ?? []).find(item => item.querySelector('summary')?.textContent?.trim() === summary)
  if (details) details.open = true
  window.requestAnimationFrame(() => { const field = details?.querySelector<HTMLElement>(selector); field?.focus(); (field as HTMLInputElement | null)?.select?.() })
}
/**
 * Copies the worksheet selection with the grid's own onCopy handler (copySelection TSV): the handler fills a probe
 * ClipboardEvent, and the native copy command then writes that payload, because the host denies navigator.clipboard.
 */
export function copyThroughGrid(grid: HTMLElement | null): boolean {
  if (!grid || typeof window === 'undefined' || typeof DataTransfer === 'undefined') return false
  const probe = new ClipboardEvent('copy', { clipboardData: new DataTransfer(), bubbles: true, cancelable: true })
  grid.dispatchEvent(probe)
  const text = probe.clipboardData?.getData('text/plain') ?? ''
  if (!probe.defaultPrevented || !text) return false
  const sink = (event: ClipboardEvent) => { event.clipboardData?.setData('text/plain', text); event.preventDefault() }
  const selection = window.getSelection(), previous = selection?.rangeCount ? selection.getRangeAt(0).cloneRange() : undefined
  window.document.addEventListener('copy', sink, { capture: true })
  try {
    grid.focus({ preventScroll: true })
    const range = window.document.createRange(); range.selectNodeContents(grid)
    selection?.removeAllRanges(); selection?.addRange(range)
    return window.document.execCommand('copy')
  } finally {
    window.document.removeEventListener('copy', sink, { capture: true })
    selection?.removeAllRanges(); if (previous) selection?.addRange(previous)
  }
}
export function spreadsheetContextMenu(context: SpreadsheetMenuContext): ContextMenuItem[] {
  const { disabled } = context
  const copy = () => copyThroughGrid(context.anchor.source)
  const structure = (id: string, label: string): ContextMenuItem => ({ id, label, disabled: true, title: XLSX_STRUCTURE_UNAVAILABLE, run: () => {} })
  return [
    { id: 'cut', label: 'Cut', shortcut: SHORTCUTS.cut, disabled, run: () => { copy(); context.onClear() } },
    { id: 'copy', label: 'Copy', shortcut: SHORTCUTS.copy, disabled: !context.anchor.source, run: copy },
    { id: 'paste', label: 'Paste', shortcut: SHORTCUTS.paste, disabled: true, title: PASTE_UNAVAILABLE, run: () => {} },
    { id: 'clear', label: 'Clear contents', shortcut: 'Delete', disabled, run: context.onClear },
    { separator: true },
    structure('row-insert', 'Insert rows'), structure('row-delete', 'Delete rows'), structure('column-insert', 'Insert columns'), structure('column-delete', 'Delete columns'),
    { separator: true },
    { id: 'row-height', label: 'Row height…', disabled, run: () => revealToolbarField(context.anchor.source, 'Cell size', '[aria-label="Row height in points"]') },
    { id: 'column-width', label: 'Column width…', disabled, run: () => revealToolbarField(context.anchor.source, 'Cell size', '[aria-label="Column width in characters"]') },
  ]
}

/** Right-click on an unselected slide object selects it first, like a left click would. */
export function selectObjectAt(event: ReactMouseEvent<HTMLElement>) {
  const object = (event.target as Element | null)?.closest?.<HTMLElement>('.presentation-object')
  if (!object || object.classList.contains('is-selected')) return
  object.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, clientX: event.clientX, clientY: event.clientY }))
}
export interface PresentationMenuContext {
  object: boolean
  slide: boolean
  disabled: boolean
  canDeleteSlide: boolean
  onDeleteObject(): void
  onNewSlide(): void
  onDuplicateSlide(): void
  onDeleteSlide(): void
}
export function presentationContextMenu(context: PresentationMenuContext): ContextMenuItem[] {
  const { disabled } = context
  const clipboard = (id: string, label: string, shortcut: string, title: string): ContextMenuItem => ({ id, label, shortcut, disabled: true, title, run: () => {} })
  return [
    clipboard('cut', 'Cut', SHORTCUTS.cut, PPTX_CLIPBOARD_UNAVAILABLE), clipboard('copy', 'Copy', SHORTCUTS.copy, PPTX_CLIPBOARD_UNAVAILABLE), clipboard('paste', 'Paste', SHORTCUTS.paste, PASTE_UNAVAILABLE),
    { id: 'delete-object', label: 'Delete object', disabled: disabled || !context.object, title: context.object ? undefined : 'Select an editable text box, shape or picture', run: context.onDeleteObject },
    { separator: true },
    { id: 'new-slide', label: 'New slide', disabled: disabled || !context.slide, run: context.onNewSlide },
    { id: 'duplicate-slide', label: 'Duplicate slide', disabled: disabled || !context.slide, run: context.onDuplicateSlide },
    { id: 'delete-slide', label: 'Delete slide', disabled: disabled || !context.canDeleteSlide, title: context.canDeleteSlide ? undefined : 'A presentation keeps at least one slide', run: context.onDeleteSlide },
  ]
}

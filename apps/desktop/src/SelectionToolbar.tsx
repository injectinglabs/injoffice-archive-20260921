import { useEffect, useLayoutEffect, useRef, useState } from 'react'
import { FONT_SIZES, TEXT_COLORS, stepFontSize, type FormattingPatch, type FormattingValues } from './FormattingToolbar'
import { AlignmentToggles, ColorButton, RibbonButton, RibbonCombo } from './Ribbon'
import './selection-toolbar.css'

/** Screen box of the selected text's first line; the mini toolbar hangs above it. */
export type SelectionAnchor = { left: number; top: number; bottom: number; width: number }
const GAP = 8, MARGIN = 8

/** Commands Office's mini toolbar carries that the document transaction does not take yet; shown disabled with the reason, never wired to a no-op. */
export const SELECTION_UNSUPPORTED = {
  bullets: 'Bullet list editing is not supported in this document.',
} as const

/** The selection's first line box when it is a non-empty text selection inside `within`; otherwise undefined. */
export function selectionAnchorRect(selection: Selection | null | undefined, within: string): SelectionAnchor | undefined {
  if (!selection || selection.rangeCount === 0 || selection.isCollapsed || !selection.toString()) return
  const range = selection.getRangeAt(0)
  const node = range.commonAncestorContainer
  const element = node.nodeType === 1 ? node as Element : node.parentElement
  if (!element?.closest(within)) return
  const rects = typeof range.getClientRects === 'function' ? range.getClientRects() : undefined
  const first = rects?.length ? rects[0]! : range.getBoundingClientRect()
  if (!first || (!first.width && !first.height)) return
  return { left: first.left, top: first.top, bottom: first.bottom, width: first.width }
}

/** Fixed position for the toolbar: above the selection's first line, flipping below it when the top is off screen, clamped horizontally. */
export function toolbarPosition(anchor: SelectionAnchor, size: { width: number; height: number }, viewport: { width: number; height: number }): { x: number; y: number; placement: 'above' | 'below' } {
  const above = anchor.top - size.height - GAP
  const placement = above >= MARGIN ? 'above' : 'below'
  const y = placement === 'above' ? above : Math.min(anchor.bottom + GAP, Math.max(MARGIN, viewport.height - size.height - MARGIN))
  const x = Math.max(MARGIN, Math.min(anchor.left, viewport.width - size.width - MARGIN))
  return { x, y, placement }
}

export interface SelectionToolbarProps {
  values?: FormattingValues
  disabled: boolean
  onChange(patch: FormattingPatch): void
  /** Selector of the container whose text selections show the toolbar. */
  within?: string
  /** Test hook: use this anchor instead of the live DOM selection. */
  anchor?: SelectionAnchor
  /** Apply the paragraph's bullet list. Without it the button is shown disabled with its reason, as Office's greyed commands are. */
  onBullets?(): void
  /** Whether the selected paragraph already carries a bullet list. */
  bullets?: boolean
}

/**
 * Office's mini toolbar: a small floating card of icon buttons over the selection —
 * font, size, grow/shrink, B/I/U, highlight, font colour, bullets and alignment.
 * Every control has a fixed width, so no value is ever truncated by a chevron, and
 * each one drives the same `onChange` patch path as the ribbon; nothing here formats.
 */
export default function SelectionToolbar({ values, disabled, onChange, within = '.office-document-canvas', anchor: fixedAnchor, onBullets, bullets }: SelectionToolbarProps) {
  const [anchor, setAnchor] = useState<SelectionAnchor | undefined>(fixedAnchor)
  const [dismissed, setDismissed] = useState(false)
  const [position, setPosition] = useState<{ x: number; y: number; placement: 'above' | 'below' }>()
  const toolbar = useRef<HTMLDivElement>(null)
  useEffect(() => {
    if (fixedAnchor || typeof window === 'undefined' || typeof window.addEventListener !== 'function' || typeof window.document?.addEventListener !== 'function') return
    const update = () => { const next = selectionAnchorRect(window.getSelection(), within); setAnchor(next); if (next) setDismissed(false) }
    // Click-away hides the toolbar until the next selection; clicks on the toolbar itself keep it (buttons do not collapse the selection).
    const pointer = (event: PointerEvent) => { if (!toolbar.current?.contains(event.target as Node)) setDismissed(true) }
    window.document.addEventListener('selectionchange', update)
    window.document.addEventListener('pointerdown', pointer, true)
    window.document.addEventListener('scroll', update, true)
    window.addEventListener('resize', update)
    update()
    return () => {
      window.document.removeEventListener('selectionchange', update); window.document.removeEventListener('pointerdown', pointer, true)
      window.document.removeEventListener('scroll', update, true); window.removeEventListener('resize', update)
    }
  }, [within, fixedAnchor])
  useLayoutEffect(() => {
    const node = toolbar.current
    if (!anchor || !node || typeof window === 'undefined' || typeof node.getBoundingClientRect !== 'function') { setPosition(undefined); return }
    const rect = node.getBoundingClientRect()
    setPosition(toolbarPosition(anchor, { width: rect.width, height: rect.height }, { width: window.innerWidth, height: window.innerHeight }))
  }, [anchor])
  if (!anchor || dismissed) return null
  const inactive = disabled || !values
  const characterInactive = inactive || values?.characterEditable === false
  const fonts = [...new Set([values?.font || 'Arial', 'Arial', 'Calibri', 'Cambria', 'Georgia', 'Times New Roman', 'Verdana', 'DejaVu Sans'])]
  const sizes = [...new Set([values?.size || 11, ...FONT_SIZES])].sort((a, b) => a - b)
  const grow = stepFontSize(values?.size, 1), shrink = stepFontSize(values?.size, -1)
  return <div ref={toolbar} className="selection-toolbar" role="toolbar" aria-label="Selection formatting" data-placement={position?.placement ?? 'above'} style={{ left: position?.x ?? anchor.left, top: position?.y ?? anchor.top, visibility: position ? undefined : 'hidden' }}>
    <RibbonCombo className="ribbon-combo-font" label="Font family" disabled={characterInactive} value={values?.font ?? ''} options={fonts.map(font => ({ value: font, label: font }))} onChange={font => onChange({ font })} />
    <RibbonCombo className="ribbon-combo-size" label="Font size" disabled={characterInactive} value={values?.size === undefined ? '' : String(values.size)} options={sizes.map(size => ({ value: String(size), label: String(size) }))} onChange={size => onChange({ size: Number(size) })} />
    <RibbonButton icon="growFont" label="Grow font" labelHidden disabled={characterInactive || grow === undefined} onClick={() => grow !== undefined && onChange({ size: grow })} />
    <RibbonButton icon="shrinkFont" label="Shrink font" labelHidden disabled={characterInactive || shrink === undefined} onClick={() => shrink !== undefined && onChange({ size: shrink })} />
    <span className="selection-toolbar-separator" aria-hidden="true" />
    <RibbonButton icon="bold" label="Bold" shortcut="bold" labelHidden aria-pressed={values?.bold ?? 'mixed'} disabled={characterInactive} onClick={() => onChange({ bold: !values?.bold })} />
    <RibbonButton icon="italic" label="Italic" shortcut="italic" labelHidden aria-pressed={values?.italic ?? 'mixed'} disabled={characterInactive} onClick={() => onChange({ italic: !values?.italic })} />
    <RibbonButton icon="underline" label="Underline" shortcut="underline" labelHidden aria-pressed={values?.underline ?? 'mixed'} disabled={characterInactive} onClick={() => onChange({ underline: !values?.underline })} />
    <RibbonButton icon="highlight" label="Highlight" labelHidden title={characterInactive ? "Select text that supports character formatting." : "Toggle yellow highlighting"} disabled={characterInactive} aria-pressed={values?.highlight === 'yellow'} onClick={() => onChange({ highlight: values?.highlight === 'yellow' ? 'none' : 'yellow' })} />
    <ColorButton label="Text color" icon="fontColor" disabled={characterInactive} value={values?.color?.toUpperCase()} colors={TEXT_COLORS} onChange={color => onChange({ color })} />
    <span className="selection-toolbar-separator" aria-hidden="true" />
    <RibbonButton icon="list" label="Bullets" labelHidden title={onBullets ? 'Bullets' : SELECTION_UNSUPPORTED.bullets} disabled={inactive || !onBullets} aria-pressed={bullets ?? false} onClick={() => onBullets?.()} />
    <AlignmentToggles className="selection-toolbar-alignment" kind="docx" horizontal={values?.alignment} disabled={inactive} onHorizontal={alignment => onChange({ alignment })} />
  </div>
}

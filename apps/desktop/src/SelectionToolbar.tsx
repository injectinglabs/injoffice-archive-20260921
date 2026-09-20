import { useEffect, useLayoutEffect, useRef, useState } from 'react'
import FormattingToolbar, { type FormattingPatch, type FormattingValues } from './FormattingToolbar'
import './selection-toolbar.css'

/** Screen box of the selected text's first line; the mini toolbar hangs above it. */
export type SelectionAnchor = { left: number; top: number; bottom: number; width: number }
const GAP = 8, MARGIN = 8

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
}

/**
 * Office's mini toolbar for the document editor. It renders the same FormattingToolbar as the ribbon, so every
 * control drives the same `onChange` patch path (and the staged Apply / Cancel model behind it); nothing here formats.
 */
export default function SelectionToolbar({ values, disabled, onChange, within = '.office-document-canvas', anchor: fixedAnchor }: SelectionToolbarProps) {
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
  return <div ref={toolbar} className="selection-toolbar" role="toolbar" aria-label="Selection formatting" data-placement={position?.placement ?? 'above'} style={{ left: position?.x ?? anchor.left, top: position?.y ?? anchor.top, visibility: position ? undefined : 'hidden' }}>
    <FormattingToolbar kind="docx" scopeLabel="Selected text" values={values} disabled={disabled} onChange={onChange} />
  </div>
}

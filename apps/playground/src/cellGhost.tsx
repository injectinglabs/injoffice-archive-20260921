import { useCallback, useEffect, useRef, useState, type CSSProperties, type RefObject } from 'react'
import type { FUniver } from '@univerjs/core/lib/facade'
import type {} from '@univerjs/sheets-ui/lib/facade'
import type { PeerInfo, PresenceManager, SheetSelection } from '../../../packages/collab/src/index.js'
import {
  measureUniverCellAnchor,
  type SheetViewportScroll,
} from './collabCellAnchor'
import {
  cellGhostTarget,
  cellGhostViewModel,
  type CellGhostViewModel,
} from './collab/cellGhost'

interface RemoteCellGhostProps {
  api: FUniver
  manager: PresenceManager
  editorRef: RefObject<HTMLDivElement | null>
  stageRef: RefObject<HTMLDivElement | null>
}

function editingPeer(manager: PresenceManager, activeSheet: string): PeerInfo<SheetSelection> | null {
  for (const peer of manager.peers()) {
    const selection = peer.selection as SheetSelection | null | undefined
    if (cellGhostTarget(selection, activeSheet)) return peer as PeerInfo<SheetSelection>
  }
  return null
}

function reducedMotion(): boolean {
  return typeof matchMedia === 'function' && matchMedia('(prefers-reduced-motion: reduce)').matches
}

/** React-owned visual projection of a peer's ephemeral draft. It deliberately
 * sits above Univer's canvas instead of registering another React root inside
 * Univer, so two editor instances can mount and unmount independently. */
export function RemoteCellGhost({ api, manager, editorRef, stageRef }: RemoteCellGhostProps) {
  const [view, setView] = useState<CellGhostViewModel | null>(null)
  const scrollRef = useRef<SheetViewportScroll>({ viewportScrollX: 0, viewportScrollY: 0 })

  const measure = useCallback(() => {
    const editor = editorRef.current
    const stage = stageRef.current
    const worksheet = api.getActiveWorkbook()?.getActiveSheet()
    if (!editor || !stage || !worksheet) {
      setView(null)
      return
    }

    const sheetId = worksheet.getSheetId()
    const peer = editingPeer(manager, sheetId)
    const selection = peer?.selection
    const target = cellGhostTarget(selection, sheetId)
    if (!peer || !selection || !target) {
      setView(null)
      return
    }

    try {
      const cell = measureUniverCellAnchor({
        editorRoot: editor,
        overlayRoot: stage,
        worksheet,
        row: target.row,
        column: target.column,
        scroll: scrollRef.current,
      })
      const next = cellGhostViewModel(
        peer,
        selection,
        sheetId,
        cell,
        { left: 0, top: 0, width: stage.clientWidth, height: stage.clientHeight },
        reducedMotion(),
      )
      setView(next)
    } catch {
      // The peer can move while Univer is replacing its active skeleton.
      setView(null)
    }
  }, [api, editorRef, manager, stageRef])

  useEffect(() => {
    let frame = 0
    const scheduleMeasure = () => {
      cancelAnimationFrame(frame)
      frame = requestAnimationFrame(measure)
    }
    const onScroll = api.addEvent(api.Event.Scroll, ({ scrollX, scrollY }) => {
      scrollRef.current = { viewportScrollX: scrollX, viewportScrollY: scrollY }
      scheduleMeasure()
    })
    const subscriptions = [
      onScroll,
      api.addEvent(api.Event.SheetZoomChanged, scheduleMeasure),
      api.addEvent(api.Event.SheetSkeletonChanged, scheduleMeasure),
      api.addEvent(api.Event.ActiveSheetChanged, scheduleMeasure),
    ]
    const unsubscribeManager = manager.onChange(scheduleMeasure)
    const resizeObserver = new ResizeObserver(scheduleMeasure)
    if (stageRef.current) resizeObserver.observe(stageRef.current)
    if (editorRef.current) resizeObserver.observe(editorRef.current)

    // Univer attaches its canvas after the workbook call returns. Observe that
    // finite setup window rather than polling on every animation frame.
    const mutationObserver = new MutationObserver(scheduleMeasure)
    if (editorRef.current) mutationObserver.observe(editorRef.current, { childList: true, subtree: true })
    scheduleMeasure()

    return () => {
      cancelAnimationFrame(frame)
      unsubscribeManager()
      subscriptions.forEach((subscription) => subscription.dispose())
      resizeObserver.disconnect()
      mutationObserver.disconnect()
    }
  }, [api, editorRef, manager, measure, stageRef])

  if (!view) return null
  const labelPlacement = view.rect.top < 26 ? 'below' : 'above'
  return (
    <div
      className={`collab-cell-ghost${view.animate ? '' : ' collab-cell-ghost--static'}`}
      data-label-placement={labelPlacement}
      style={{
        '--collab-color': view.color,
        left: view.rect.left,
        top: view.rect.top,
        width: view.rect.width,
        height: view.rect.height,
      } as CSSProperties}
      role={view.role}
      aria-live={view.ariaLive}
      aria-atomic={view.ariaAtomic}
      aria-label={view.ariaLabel}
    >
      <span className="collab-cell-ghost__draft">{view.draft}<i aria-hidden="true" /></span>
      <span className="collab-cell-ghost__label">{view.name} typing</span>
    </div>
  )
}

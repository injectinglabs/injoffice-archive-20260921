import type { FUniver } from '@univerjs/core/lib/facade'
// Facade surface via module augmentation (see @injoffice/charts' manager note).
import type {} from '@univerjs/sheets/lib/facade'
import type {} from '@univerjs/sheets-ui/lib/facade'
import type {} from '@univerjs/sheets-drawing-ui/lib/facade'
import type {} from '@univerjs/ui/lib/facade'
import { dropShape, publishShapeSpec, setShapeTextEditHandler } from './registry'
import { SHAPE_CATEGORIES, SHAPE_DEFAULTS } from './types'
import type { ShapeKind, ShapeSpec } from './types'
import type { FileShapeAnchor } from './fromFile'
import { toWireShapeRemove, toWireShapeUpdate } from './toFile'
import type { ShapeLifecycleWireResult, WireShapeRemove, WireShapeUpdate } from './toFile'

// ShapeManager — floating shapes on the sheet: text boxes, rectangles,
// ellipses, arrows, lines, callouts. Mounting/move/resize ride the same
// float-DOM layer as charts; text edits inside a shape flow back through the
// registry so the spec (and anything persisting it) stays truthful.

export const SHAPE_COMPONENT_KEY = 'injoffice-shape'

let seq = 0
function newShapeId(): string {
  return `shape-${Date.now().toString(36)}-${(++seq).toString(36)}`
}

// Default insert size, bucketed by category rather than per-preset — with
// 80 presets a per-kind table would be all noise, no signal (a "star8" and
// a "star10" don't need different defaults). Categories match the toolbar
// grouping in types.ts.
const CATEGORY_SIZE: Record<string, { w: number; h: number }> = {
  Basic: { w: 200, h: 130 },
  Arrows: { w: 180, h: 60 },
  Callouts: { w: 220, h: 130 },
  'Stars & banners': { w: 180, h: 180 },
  Flowchart: { w: 200, h: 110 },
  'Lines & text': { w: 200, h: 50 },
}

const KIND_SIZE: Record<ShapeKind, { w: number; h: number }> = (() => {
  const out = {} as Record<ShapeKind, { w: number; h: number }>
  for (const cat of SHAPE_CATEGORIES) {
    const size = CATEGORY_SIZE[cat.label] ?? { w: 200, h: 120 }
    for (const kind of cat.kinds) out[kind as ShapeKind] = size
  }
  return out
})()

export class ShapeManager {
  private readonly api: FUniver
  private readonly shapes = new Map<string, { spec: ShapeSpec; dispose: () => void; cellAnchor?: FileShapeAnchor; sheetId: string }>()
  private readonly changeListeners = new Set<() => void>()
  private cascade = 0

  constructor(api: FUniver) {
    this.api = api
  }

  activeSheetId(): string | undefined {
    return this.api.getActiveWorkbook()?.getActiveSheet()?.getSheetId()
  }

  start(): void {
    setShapeTextEditHandler((id, text) => {
      const s = this.shapes.get(id)
      if (s && (s.spec.text ?? '') !== text) this.updateSpec(id, { text })
    })
  }

  stop(): void {
    setShapeTextEditHandler(null)
    for (const id of [...this.shapes.keys()]) this.remove(id)
  }

  onChange(listener: () => void): () => void {
    this.changeListeners.add(listener)
    return () => this.changeListeners.delete(listener)
  }

  private emitChange(): void {
    this.changeListeners.forEach((l) => l())
  }

  /** Insert a shape of the given kind with sensible defaults. */
  create(kind: ShapeKind): ShapeSpec | null {
    const spec: ShapeSpec = { id: newShapeId(), kind, ...SHAPE_DEFAULTS[kind] }
    return this.add(spec) ? spec : null
  }

  /**
   * Mount a shape. cellAnchor places it at the FILE's real grid position (a
   * shape hydrated from ?shapes=json on load) — approximated to pixels since
   * Univer's float-DOM speaks pixels, not cells (see CELL_PX below). Without
   * one (a fresh, browser-created shape) it cascades on-screen the way it
   * always has.
   */
  add(spec: ShapeSpec, cellAnchor?: FileShapeAnchor, sheetId?: string): boolean {
    const workbook = this.api.getActiveWorkbook()
    const sheet = sheetId ? workbook?.getSheetBySheetId(sheetId) : workbook?.getActiveSheet()
    if (!sheet || this.shapes.has(spec.id)) return false
    publishShapeSpec(spec)
    const size = KIND_SIZE[spec.kind]
    let startX: number, startY: number, endX: number, endY: number
    if (cellAnchor) {
      startX = cellAnchor.FromCol * CELL_PX.w
      startY = cellAnchor.FromRow * CELL_PX.h
      endX = Math.max(startX + 40, cellAnchor.ToCol * CELL_PX.w)
      endY = Math.max(startY + 40, cellAnchor.ToRow * CELL_PX.h)
    } else {
      const offset = (this.cascade++ % 6) * 26
      startX = 160 + offset
      startY = 120 + offset
      endX = startX + size.w
      endY = startY + size.h
    }
    const mounted = sheet.addFloatDomToPosition(
      {
        componentKey: SHAPE_COMPONENT_KEY,
        initPosition: { startX, endX, startY, endY },
        data: JSON.parse(JSON.stringify({ shapeId: spec.id, spec })),
        allowTransform: true,
      },
      spec.id,
    )
    if (!mounted) {
      dropShape(spec.id)
      return false
    }
    this.shapes.set(spec.id, { spec, dispose: mounted.dispose, cellAnchor, sheetId: sheet.getSheetId() })
    this.emitChange()
    return true
  }

  remove(id: string): void {
    const s = this.shapes.get(id)
    if (!s) return
    s.dispose()
    this.shapes.delete(id)
    dropShape(id)
    this.emitChange()
  }

  updateSpec(id: string, patch: Partial<Omit<ShapeSpec, 'id' | 'kind' | 'nativeIdentity'>>): void {
    const s = this.shapes.get(id)
    if (!s) return
    s.spec = { ...s.spec, ...patch, id, kind: s.spec.kind, nativeIdentity: s.spec.nativeIdentity }
    publishShapeSpec(s.spec)
    this.emitChange()
  }

  getSpec(id: string): ShapeSpec | undefined {
    return this.shapes.get(id)?.spec
  }

  list(): ShapeSpec[] {
    return [...this.shapes.values()].map((s) => s.spec)
  }

  serialize(): ShapeSpec[] {
    return this.list()
  }

  /** Specs plus their file anchors (when known) — what the save-time OOXML
   *  writer consumes (toFile.ts), mirroring ChartManager.listWithAnchors. */
  listWithAnchors(): Array<{ spec: ShapeSpec; cellAnchor?: FileShapeAnchor; sheetId: string }> {
    return [...this.shapes.values()].map((s) => ({ spec: s.spec, cellAnchor: s.cellAnchor, sheetId: s.sheetId }))
  }

  /** Build a stable-identity native delete without disposing the UI first. */
  nativeRemoveRequest(id: string): ShapeLifecycleWireResult<WireShapeRemove> {
    const spec = this.shapes.get(id)?.spec
    if (!spec) return { skipped: [`shape ${id}: it is not managed`] }
    return toWireShapeRemove(spec)
  }

  /** Build a stable-identity native replacement for the current shape. */
  nativeUpdateRequest(id: string, sheetName: string | null): ShapeLifecycleWireResult<WireShapeUpdate> {
    const shape = this.shapes.get(id)
    if (!shape) return { skipped: [`shape ${id}: it is not managed`] }
    return toWireShapeUpdate({ spec: shape.spec, cellAnchor: shape.cellAnchor }, sheetName)
  }

  hydrate(specs: ShapeSpec[]): void {
    for (const spec of specs) this.add(spec)
  }
}

// Approximate cell size in pixels — Univer's float-DOM speaks pixels, our
// file anchors speak cells, and there is no live cell-boundary lookup wired
// here yet. Good enough to land a loaded shape roughly on its saved cell;
// the user's own drag (in-session) is exact regardless.
const CELL_PX = { w: 88, h: 20 }

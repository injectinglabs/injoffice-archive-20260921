import type { CollabTransport } from './types'

/** The minimal DeckSpec shape this sync layer needs. Kept structural so the
 * canonical collab package does not depend on @injoffice/slides. */
export interface SyncDeck {
  id: string
  title: string
  theme?: unknown
  slides: Array<{ id: string; shapeOverrides?: unknown }>
}

/** A last-writer-wins field update addressed by stable deck/slide identity. */
export interface DeckFieldOp {
  id: 'deck.set-field'
  params: {
    scope: 'deck' | 'slide'
    field: string
    value: unknown
    slideId?: string
  }
}

/** One shape-geometry override, independently addressable by its stable key. */
export interface DeckShapeOverrideOp {
  id: 'deck.set-shape-override'
  params: { slideId: string; shapeKey: string; value: unknown }
}

/** Structural operations are addressed by stable slide IDs, never array
 * positions. That makes the gateway's existing ordered op log sufficient for
 * the bounded one-slide insert/delete/move edits supported here. */
export interface DeckInsertSlideOp {
  id: 'deck.insert-slide'
  params: { slide: { id: string; [key: string]: unknown }; afterId?: string | null }
}

export interface DeckDeleteSlideOp {
  id: 'deck.delete-slide'
  params: { slideId: string }
}

export interface DeckMoveSlideOp {
  id: 'deck.move-slide'
  params: { slideId: string; afterId?: string | null }
}

export type DeckOp = DeckFieldOp | DeckShapeOverrideOp | DeckInsertSlideOp | DeckDeleteSlideOp | DeckMoveSlideOp
type DeckTransport = Pick<CollabTransport, 'opSubmit' | 'opSince'>

const DECK_FIELDS = new Set(['title', 'theme'])
const BLOCKED_SLIDE_FIELDS = new Set(['id', 'shapeOverrides', 'shapeAnimations'])

function copy<T>(value: T): T {
  return value === undefined ? value : JSON.parse(JSON.stringify(value)) as T
}

function equal(a: unknown, b: unknown): boolean {
  return JSON.stringify(a) === JSON.stringify(b)
}

/**
 * Turn a local edit into stable operations. One-slide insert/delete/move is
 * safe because it is addressed entirely by stable IDs and replayed in the
 * gateway's total op-log order. Compound structural edits still return null:
 * there is no sequence transform for pretending a bulk rewrite is safe.
 */
export function diffDeck(before: SyncDeck, after: SyncDeck): DeckOp[] | null {
  const beforeIDs = before.slides.map((slide) => slide.id)
  const afterIDs = after.slides.map((slide) => slide.id)
  if (new Set(beforeIDs).size !== beforeIDs.length || new Set(afterIDs).size !== afterIDs.length || afterIDs.some((id) => !id)) return null
  if (!sameOrder(beforeIDs, afterIDs)) {
    const structural = diffOneStructuralEdit(before, after)
    if (!structural) return null
    return [structural]
  }
  const ops: DeckOp[] = []
  const beforeFields = before as unknown as Record<string, unknown>
  const afterFields = after as unknown as Record<string, unknown>
  for (const field of DECK_FIELDS) {
    if (!equal(beforeFields[field], afterFields[field])) ops.push({ id: 'deck.set-field', params: { scope: 'deck', field, value: copy(afterFields[field]) } })
  }
  for (let i = 0; i < before.slides.length; i++) {
    const prev = before.slides[i]
    const next = after.slides[i]
    const prevFields = prev as Record<string, unknown>
    const nextFields = next as Record<string, unknown>
    for (const field of new Set([...Object.keys(prevFields), ...Object.keys(nextFields)])) {
      if (BLOCKED_SLIDE_FIELDS.has(field) || !equal(prevFields[field], nextFields[field])) {
        if (!BLOCKED_SLIDE_FIELDS.has(field)) ops.push({ id: 'deck.set-field', params: { scope: 'slide', slideId: prev.id, field, value: copy(nextFields[field]) } })
      }
    }
    const oldOverrides = (prev.shapeOverrides ?? {}) as Record<string, unknown>
    const newOverrides = (next.shapeOverrides ?? {}) as Record<string, unknown>
    for (const shapeKey of new Set([...Object.keys(oldOverrides), ...Object.keys(newOverrides)])) {
      if (!equal(oldOverrides[shapeKey], newOverrides[shapeKey])) {
        ops.push({ id: 'deck.set-shape-override', params: { slideId: prev.id, shapeKey, value: copy(newOverrides[shapeKey]) } })
      }
    }
  }
  return ops
}

function sameOrder(a: readonly string[], b: readonly string[]): boolean {
  return a.length === b.length && a.every((id, index) => id === b[index])
}

function sameIDsInOrder(ids: readonly string[], allowed: ReadonlySet<string>, expected: readonly string[]): boolean {
  return sameOrder(ids.filter((id) => allowed.has(id)), expected)
}

function diffOneStructuralEdit(before: SyncDeck, after: SyncDeck): DeckOp | null {
  const beforeIDs = before.slides.map((slide) => slide.id)
  const afterIDs = after.slides.map((slide) => slide.id)
  const beforeSet = new Set(beforeIDs)
  const afterSet = new Set(afterIDs)
  const beforeByID = new Map(before.slides.map((slide) => [slide.id, slide]))
  const afterByID = new Map(after.slides.map((slide) => [slide.id, slide]))
  // Do not silently discard a field change bundled with a structural edit.
  // Hosts submit those as separate user actions; a bulk rewrite remains
  // deliberately outside this small, transform-free protocol.
  for (const id of beforeIDs) {
    if (afterSet.has(id) && !equal(beforeByID.get(id), afterByID.get(id))) return null
  }

  // Insert exactly one fresh ID while retaining all existing slides in order.
  if (afterIDs.length === beforeIDs.length + 1 && sameIDsInOrder(afterIDs, beforeSet, beforeIDs)) {
    const index = afterIDs.findIndex((id) => !beforeSet.has(id))
    const slide = after.slides[index]
    if (!slide?.id) return null
    return { id: 'deck.insert-slide', params: { slide: copy(slide) as { id: string; [key: string]: unknown }, afterId: index === 0 ? null : afterIDs[index - 1] } }
  }

  // Delete exactly one existing ID while preserving survivor order. Do not
  // generate an operation that leaves an invalid empty deck.
  if (beforeIDs.length > 1 && afterIDs.length === beforeIDs.length - 1 && sameIDsInOrder(beforeIDs, afterSet, afterIDs)) {
    const removed = beforeIDs.find((id) => !afterSet.has(id))
    return removed ? { id: 'deck.delete-slide', params: { slideId: removed } } : null
  }

  // Reorder exactly one stable ID. Several equivalent move descriptions can
  // exist for a simple swap; use the first deterministic before-order match.
  if (afterIDs.length === beforeIDs.length && beforeIDs.every((id) => afterSet.has(id))) {
    for (const slideId of beforeIDs) {
      const without = beforeIDs.filter((id) => id !== slideId)
      const index = afterIDs.indexOf(slideId)
      const candidate = without.slice()
      candidate.splice(index, 0, slideId)
      if (sameOrder(candidate, afterIDs)) {
        return { id: 'deck.move-slide', params: { slideId, afterId: index === 0 ? null : afterIDs[index - 1] } }
      }
    }
  }
  return null
}

/** Apply only recognised, stable-key deck operations. Invalid remote input is ignored. */
export function applyDeckOps(deck: SyncDeck, ops: readonly unknown[]): SyncDeck {
  let next = deck
  for (const raw of ops) {
    const op = raw as { id?: unknown; params?: Record<string, unknown> }
    const p = op.params
    if (!p) continue
    const inserted = p.slide
    if (op.id === 'deck.insert-slide' && isSlide(inserted)) {
      const afterId = p.afterId
      if (afterId !== undefined && afterId !== null && typeof afterId !== 'string') continue
      if (afterId === inserted.id || next.slides.some((slide) => slide.id === inserted.id)) continue
      const slides = next.slides.slice()
      const anchor = afterId === null || afterId === undefined ? -1 : slides.findIndex((slide) => slide.id === afterId)
      slides.splice(anchor < 0 && afterId ? slides.length : anchor + 1, 0, copy(inserted) as typeof slides[number])
      next = { ...next, slides }
      continue
    }
    if (op.id === 'deck.delete-slide' && typeof p.slideId === 'string') {
      const index = next.slides.findIndex((slide) => slide.id === p.slideId)
      if (index >= 0 && next.slides.length > 1) {
        const slides = next.slides.slice()
        slides.splice(index, 1)
        next = { ...next, slides }
      }
      continue
    }
    if (op.id === 'deck.move-slide' && typeof p.slideId === 'string') {
      const afterId = p.afterId
      if (afterId !== undefined && afterId !== null && typeof afterId !== 'string') continue
      if (afterId === p.slideId) continue
      const index = next.slides.findIndex((slide) => slide.id === p.slideId)
      if (index < 0) continue
      const slides = next.slides.slice()
      const [slide] = slides.splice(index, 1)
      const anchor = afterId === null || afterId === undefined ? -1 : slides.findIndex((item) => item.id === afterId)
      slides.splice(anchor < 0 && afterId ? slides.length : anchor + 1, 0, slide)
      next = { ...next, slides }
      continue
    }
    if (op.id === 'deck.set-shape-override' && typeof p.slideId === 'string' && typeof p.shapeKey === 'string') {
      const index = next.slides.findIndex((slide) => slide.id === p.slideId)
      if (index < 0) continue
      const slides = next.slides.slice()
      const previous = slides[index]
      const overrides = { ...(previous.shapeOverrides as Record<string, unknown> | undefined) }
      if (p.value === undefined) delete overrides[p.shapeKey]
      else overrides[p.shapeKey] = copy(p.value)
      slides[index] = { ...previous, shapeOverrides: overrides }
      next = { ...next, slides }
    }
    if (op.id !== 'deck.set-field' || typeof p.field !== 'string') continue
    if (p.scope === 'deck' && DECK_FIELDS.has(p.field)) next = { ...next, [p.field]: copy(p.value) }
    if (p.scope === 'slide' && typeof p.slideId === 'string' && !BLOCKED_SLIDE_FIELDS.has(p.field)) {
      const index = next.slides.findIndex((slide) => slide.id === p.slideId)
      if (index < 0) continue
      const slides = next.slides.slice()
      slides[index] = { ...slides[index], [p.field]: copy(p.value) }
      next = { ...next, slides }
    }
  }
  return next
}

function isSlide(value: unknown): value is { id: string; [key: string]: unknown } {
  return Boolean(value) && typeof value === 'object' && typeof (value as { id?: unknown }).id === 'string' && Boolean((value as { id: string }).id)
}

export interface DeckSyncHooks<TDeck extends SyncDeck> {
  /** Render the confirmed deck plus all unacknowledged local operations. */
  onDeck(deck: TDeck): void
  /** The log cannot be reconciled with the saved file. */
  onResync(): void
}

/**
 * Ordered, optimistic DeckSpec sync. The confirmed base advances only with
 * server-ordered entries; the displayed deck reapplies pending local field
 * updates on top. A STALE_BASE therefore naturally rebases a local change
 * after everything that won before it, rather than losing a simultaneous
 * title or geometry edit.
 */
export class DeckSyncEngine<TDeck extends SyncDeck> {
  private base: TDeck
  private appliedSeq = 0
  private pending: DeckOp[][] = []
  private inFlight = false
  private stopped = false
  private resynced = false
  private readonly reorder = new Map<number, { ops: unknown[] }>()

  constructor(private readonly transport: DeckTransport, private readonly path: string, initial: TDeck, private readonly hooks: DeckSyncHooks<TDeck>) {
    this.base = copy(initial)
  }

  get applied(): number { return this.appliedSeq }
  get syncing(): boolean { return Boolean(this.transport.opSubmit && this.transport.opSince) && !this.stopped }

  async bootstrap(savedSeq: number): Promise<void> {
    this.appliedSeq = savedSeq
    await this.catchUp()
  }

  stop(): void { this.stopped = true; this.reorder.clear() }

  change(next: TDeck): boolean {
    const visible = this.visible()
    const ops = diffDeck(visible, next)
    if (!ops) return false
    if (ops.length === 0) return true
    this.pending.push(ops)
    this.emit()
    void this.submitLoop()
    return true
  }

  receiveEntry(seq: number, ops: unknown[]): void {
    if (this.stopped || seq <= this.appliedSeq) return
    if (seq === this.appliedSeq + 1) { this.apply(seq, ops); this.drain(); return }
    this.reorder.set(seq, { ops })
    void this.catchUp()
  }

  async resume(): Promise<void> { await this.catchUp(); void this.submitLoop() }

  private visible(): TDeck {
    return this.pending.reduce((deck, ops) => applyDeckOps(deck, ops) as TDeck, this.base)
  }

  private emit(): void { this.hooks.onDeck(this.visible()) }
  private apply(seq: number, ops: unknown[]): void { this.base = applyDeckOps(this.base, ops) as TDeck; this.appliedSeq = seq; this.emit() }
  private drain(): void { for (;;) { const entry = this.reorder.get(this.appliedSeq + 1); if (!entry) return; this.reorder.delete(this.appliedSeq + 1); this.apply(this.appliedSeq + 1, entry.ops) } }

  private async submitLoop(): Promise<void> {
    if (this.inFlight || this.stopped || !this.transport.opSubmit) return
    this.inFlight = true
    try {
      for (let tries = 0; this.pending.length && tries < 8 && !this.stopped; tries++) {
        const batch = this.pending[0]
        try {
          const seq = await this.transport.opSubmit(this.path, batch, this.appliedSeq)
          this.base = applyDeckOps(this.base, batch) as TDeck
          this.pending.shift()
          this.appliedSeq = seq
          this.emit()
        } catch (error) {
          if (error instanceof Error && error.message.includes('STALE_BASE')) { await this.catchUp(); continue }
          return
        }
      }
    } finally { this.inFlight = false }
  }

  private async catchUp(): Promise<void> {
    if (!this.transport.opSince || this.stopped) return
    try {
      const result = await this.transport.opSince(this.path, this.appliedSeq)
      if (result.reset) return this.resync()
      for (const entry of result.ops) if (entry.seq > this.appliedSeq) this.apply(entry.seq, entry.ops as unknown[])
      this.drain()
    } catch { /* reconnect/event delivery will retry */ }
  }

  private resync(): void { if (!this.resynced) { this.resynced = true; this.hooks.onResync() } }
}

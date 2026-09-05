import type { IDisposable, IRange } from '@univerjs/core'
import type { FUniver } from '@univerjs/core/lib/facade'
import type { ComponentType } from 'react'
// Side-effect type imports: the facade's sheet/UI surface (getActiveWorkbook,
// highlightRanges, Event.SelectionChanged, …) exists only as module
// augmentations — see packages/charts/src/manager.ts for the same note.
import type {} from '@univerjs/sheets/lib/facade'
import type {} from '@univerjs/sheets-ui/lib/facade'
import type {} from '@univerjs/ui/lib/facade'
import { PresenceState } from './presence'
import { SheetPresenceLabel, SheetPresencePopup, type SheetPresenceLabelProps } from './SheetPresenceLabel'
import { sameSelection, withAlpha } from './selection'
import {
  SubmitQueue,
  cellsOf,
  cloneParams,
  rebaseUnit,
  reconcileRemote,
  shouldShipMutation,
  type ExecOptions,
  type MutationInfo,
  type OpEntry,
  type OpRecord,
} from './sync'
import { structEditsOf, transformOps, transformSelection, transformSelectionThroughMove } from './transform'
import type { CollabTransport, FileChange, PeerInfo, SheetSelection } from './types'

// PresenceManager — InjOffice collaboration bound to one Univer instance and
// one file.
//
// Stage A (presence + live followers):
//   • joins the file's room over the host transport and keeps PresenceState
//   • publishes the LOCAL selection (throttled) on every SelectionChanged
//   • draws every REMOTE peer's selection on the active sheet through the
//     public facade (FWorksheet.highlightRanges → IMarkSelectionService)
//   • tells the host when the file on disk changed underneath it
//
// Stage B (cell-level co-editing), active when the host passes the
// command service AND the transport speaks the op log:
//   • every local MUTATION (not command — options don't propagate to nested
//     commands, and commands push undo entries) is captured through
//     ICommandService.onMutationExecutedForCollab, filtered (sync.ts),
//     batched per tick and shipped as one ordered submission
//   • remote entries are reconciled (cells we wrote later stay ours) and
//     replayed as leaf mutations with {fromCollab:true}, which keeps them
//     out of the local undo stack and tells Univer's own controllers
//     (active sheet, freeze, editor, repeat-last-action) not to react
//   • a joiner loads the file and replays only ops after its saved_seq; a
//     reconnect gap-fills from the last applied seq; when the ring can't
//     fill a gap or the agent replaced the file, the host reloads
//
// Univer gotchas this works around (verified against 0.25.1):
//   - the mark-selection layer is wiped by MarkSelectionRenderController on
//     every SetCellEditVisibleOperation (opening/closing the cell editor), so
//     remote highlights are re-added after that operation;
//   - addShape stamps the CURRENT sheet onto a shape regardless of the range's
//     own sheet, so only peers on the active sheet are drawn, and everything
//     is redrawn on ActiveSheetChanged;
//   - highlight() throws when there is no active sheet; wrapped.

/** The slice of Univer's ICommandService the sync layer needs (DI-resolved by the host). */
export interface CommandServiceLike {
  onMutationExecutedForCollab(listener: (info: MutationInfo, options?: ExecOptions) => void): IDisposable
  syncExecuteCommand(id: string, params?: object, options?: ExecOptions): unknown
}

export interface PresenceOptions {
  /** The file's /v1/files URL or absolute path — the room's identity. */
  path: string
  /** Display name shown to other peers. */
  name: string
  /** Throttle for local selection publishes, ms. */
  publishIntervalMs?: number
  /** Stage B: enables co-editing. `univer.__getInjector().get(ICommandService)`. */
  commandService?: CommandServiceLike
  /** Cell-anchored remote-user flag. Pass false to keep range highlights only. */
  peerLabelComponent?: ComponentType<SheetPresenceLabelProps> | false
}

export type ResyncReason = 'reset' | 'gap'

const CELL_EDIT_VISIBLE_OP = 'sheet.operation.set-cell-edit-visible'
let nextPresenceManagerId = 1

export class PresenceManager {
  readonly state = new PresenceState()
  private readonly api: FUniver
  private readonly transport: CollabTransport
  private readonly opts: PresenceOptions & { publishIntervalMs: number }
  private readonly peerLabelKey = `injoffice.sheet.presence-label.${nextPresenceManagerId++}`
  private readonly disposables: Array<() => void> = []
  private readonly peerVisuals = new Map<string, IDisposable[]>()
  private readonly peerLabelTimers = new Map<string, ReturnType<typeof setTimeout>>()
  private readonly changeListeners = new Set<() => void>()
  private readonly fileListeners = new Set<(c: FileChange) => void>()
  private readonly resyncListeners = new Set<(reason: ResyncReason) => void>()
  private readonly peerSavedListeners = new Set<(c: FileChange) => void>()
  private started = false
  private joined = false
  private lastPublished: SheetSelection | null = null
  private pendingPublish: SheetSelection | null = null
  private publishTimer: ReturnType<typeof setTimeout> | null = null
  private lastPublishAt = 0
  private localMode: NonNullable<SheetSelection['mode']> = 'selecting'

  // ---- Stage B/C state ----
  private unitId = ''
  private applied = 0
  private queue: SubmitQueue | null = null
  private lossyWarned = false
  private readonly reorder = new Map<number, OpEntry | 'self'>()
  private catchUpTimer: ReturnType<typeof setTimeout> | null = null
  private catchingUp: Promise<void> | null = null
  private resynced = false

  constructor(api: FUniver, transport: CollabTransport, opts: PresenceOptions) {
    this.api = api
    this.transport = transport
    this.opts = { publishIntervalMs: 80, peerLabelComponent: SheetPresenceLabel, ...opts }
  }

  /** Our gateway client id once joined — the collab_origin of our own writes. */
  get clientId(): string | null {
    return this.joined ? this.state.clientId || null : null
  }

  get self(): PeerInfo | null {
    return this.state.me
  }

  /** Whether cell-level co-editing is live for this file. */
  get syncing(): boolean {
    return this.queue !== null && this.joined
  }

  /** Last op-log seq applied locally (send as collab_seq on Save). */
  get appliedSeq(): number {
    return this.applied
  }

  /** Submissions not yet acknowledged by the gateway. */
  get pendingOps(): number {
    return this.queue?.pending ?? 0
  }

  peers(): PeerInfo[] {
    return this.state.peers()
  }

  onChange(cb: () => void): () => void {
    this.changeListeners.add(cb)
    return () => this.changeListeners.delete(cb)
  }

  /** The file on disk changed and this editor is NOT in sync with it — reload or banner. */
  onFileChanged(cb: (c: FileChange) => void): () => void {
    this.fileListeners.add(cb)
    return () => this.fileListeners.delete(cb)
  }

  /** A peer saved and we are already in sync — informational only. */
  onPeerSaved(cb: (c: FileChange) => void): () => void {
    this.peerSavedListeners.add(cb)
    return () => this.peerSavedListeners.delete(cb)
  }

  /** The op log can't bring this editor up to date — the host must reload the file. */
  onResync(cb: (reason: ResyncReason) => void): () => void {
    this.resyncListeners.add(cb)
    return () => this.resyncListeners.delete(cb)
  }

  async start(): Promise<void> {
    if (this.started) return
    this.started = true

    if (this.opts.peerLabelComponent) {
      const componentSub = this.api.registerComponent(this.peerLabelKey, SheetPresencePopup)
      this.disposables.push(() => componentSub.dispose())
    }

    this.disposables.push(this.transport.onEvent((ev) => this.handleEvent(ev)))
    this.disposables.push(this.transport.onReconnect(() => void this.join()))

    const selSub = this.api.addEvent(this.api.Event.SelectionChanged, (params) => {
      const sheetId = params.worksheet?.getSheetId()
      if (!sheetId) return
      this.queuePublish(toSelection(sheetId, params.selections, this.localMode))
    })
    this.disposables.push(() => selSub.dispose())

    const sheetSub = this.api.addEvent(this.api.Event.ActiveSheetChanged, () => this.redrawAll())
    this.disposables.push(() => sheetSub.dispose())

    const editChangingSub = this.api.addEvent(this.api.Event.SheetEditChanging, (params) => {
      const sheetId = params.worksheet.getSheetId()
      const draft = params.value.toPlainText().replace(/\r\n$/, '').slice(0, 512)
      this.localMode = 'editing'
      this.queuePublish({
        sheet: sheetId,
        ranges: [[params.row, params.column, params.row, params.column]],
        active: [params.row, params.column],
        mode: 'editing',
        draft,
      })
    })
    this.disposables.push(() => editChangingSub.dispose())

    const cmdSub = this.api.addEvent(this.api.Event.CommandExecuted, (e) => {
      if (e.id !== CELL_EDIT_VISIBLE_OP) return
      const visible = Boolean((e.params as { visible?: boolean } | undefined)?.visible)
      this.localMode = visible ? 'editing' : 'selecting'
      const sel = currentSelection(this.api, this.localMode)
      if (sel) this.queuePublish(sel)
      this.redrawAll()
    })
    this.disposables.push(() => cmdSub.dispose())

    const cmd = this.opts.commandService
    if (cmd && this.transport.opSubmit && this.transport.opSince) {
      this.unitId = this.api.getActiveWorkbook()?.getId() ?? ''
      // base_seq = what this client has applied: the gateway keeps the log
      // linear by rejecting a stale base; onStale catches us up (transforming
      // the queued ops) and the queue resubmits against the new head.
      this.queue = new SubmitQueue((ops) => this.transport.opSubmit!(this.opts.path, ops, this.applied), {
        onQueued: () => this.emitChange(),
        onAcked: (ops, seq) => this.acked(ops, seq),
        onStale: () => this.catchUp(this.applied),
      })
      const mutSub = cmd.onMutationExecutedForCollab((info, options) => {
        if (!this.queue || !this.joined) return
        if (!shouldShipMutation(info, options, this.unitId)) return
        this.queue.push({ id: info.id, params: cloneParams(info.params) })
      })
      this.disposables.push(() => mutSub.dispose())
    }

    await this.join()
  }

  stop(): void {
    if (!this.started) return
    this.started = false
    if (this.publishTimer) {
      clearTimeout(this.publishTimer)
      this.publishTimer = null
    }
    if (this.catchUpTimer) {
      clearTimeout(this.catchUpTimer)
      this.catchUpTimer = null
    }
    for (const d of this.disposables.splice(0)) d()
    this.clearHighlights()
    this.queue?.clear()
    this.queue = null
    this.reorder.clear()
    if (this.joined) {
      this.joined = false
      void this.transport.leave(this.opts.path).catch(() => undefined)
    }
    this.state.clear()
    this.changeListeners.clear()
    this.fileListeners.clear()
    this.resyncListeners.clear()
    this.peerSavedListeners.clear()
  }

  private async join(): Promise<void> {
    try {
      const res = await this.transport.join(this.opts.path, this.opts.name)
      if (!this.started) return
      const first = !this.joined
      this.state.applyJoin(res)
      this.joined = true
      this.lastPublished = null
      this.redrawAll()
      this.emitChange()
      // Announce where we are right away so peers don't see a nameless ghost.
      const sel = currentSelection(this.api, this.localMode)
      if (sel) this.queuePublish(sel)

      if (this.queue && res.log) {
        // First join: the file we loaded reflects saved_seq; replay the
        // rest. Reconnect: we know what we applied; fill the gap. Then
        // ship whatever we queued while offline.
        await this.catchUp(first ? res.log.saved_seq : this.applied)
        void this.queue.flush()
      }
    } catch {
      // Transport unavailable (old gateway, socket down): the editor works
      // exactly as before, just without presence. Reconnect will retry.
      this.joined = false
    }
  }

  private handleEvent(ev: Parameters<CollabTransport['onEvent']>[0] extends (e: infer E) => void ? E : never): void {
    const res = this.state.apply(ev)
    if (!res) return
    switch (res.kind) {
      case 'peers':
        this.redrawAll()
        this.emitChange()
        break
      case 'selection':
        this.drawPeer(res.peer)
        this.emitChange()
        break
      case 'file':
        void this.handleFileChanged(res.change)
        break
      case 'op':
        this.receive(res.entry)
        break
    }
  }

  private emitChange(): void {
    for (const cb of this.changeListeners) cb()
  }

  // ---- live followers ------------------------------------------------------

  private async handleFileChanged(change: FileChange): Promise<void> {
    const savedSeq = change.saved_seq ?? 0
    if (!change.reset && this.queue && savedSeq <= this.applied) {
      for (const cb of this.peerSavedListeners) cb(change)
      return
    }
    if (!change.reset && this.queue && savedSeq > this.applied) {
      // The saver may simply be ahead of us by ops still in flight.
      await this.catchUp(this.applied)
      if (savedSeq <= this.applied) {
        for (const cb of this.peerSavedListeners) cb(change)
        return
      }
    }
    for (const cb of this.fileListeners) cb(change)
  }

  // ---- local → log ---------------------------------------------------------

  private acked(_ops: OpRecord[], seq: number): void {
    if (seq === this.applied + 1) {
      this.applied = seq
      this.drain()
    } else if (seq > this.applied) {
      this.reorder.set(seq, 'self')
    }
    this.emitChange()
  }

  // ---- log → local ---------------------------------------------------------

  private receive(entry: OpEntry): void {
    if (this.state.isSelf(entry.client_id)) {
      // Our own op from a previous socket identity (replayed by the log).
      if (entry.seq > this.applied) this.reorder.set(entry.seq, 'self')
      this.drain()
      return
    }
    if (entry.seq <= this.applied) return
    if (entry.seq === this.applied + 1) {
      this.applyEntry(entry)
      this.drain()
      return
    }
    // Out of order (fan-out races) or a real gap: hold it, ask the log.
    this.reorder.set(entry.seq, entry)
    if (!this.catchUpTimer) {
      this.catchUpTimer = setTimeout(() => {
        this.catchUpTimer = null
        void this.catchUp(this.applied)
      }, 300)
    }
  }

  private drain(): void {
    for (;;) {
      const next = this.reorder.get(this.applied + 1)
      if (next === undefined) return
      this.reorder.delete(this.applied + 1)
      if (next === 'self') this.applied++
      else this.applyEntry(next)
    }
  }

  private applyEntry(entry: OpEntry): void {
    const cmd = this.opts.commandService
    if (!cmd) return
    // Every tab has its own workbook id — rebase before anything else.
    const incoming = rebaseUnit(entry.ops, this.unitId)
    const pending = this.queue?.allPending() ?? []

    // T(incoming, pending): the incoming entry is ordered BEFORE our pending
    // ops (they will be resubmitted against a later head), but our doc
    // already contains them. Cells our pending ops write stay ours (last
    // writer by final order), and the entry is shifted through our pending
    // structural edits so it lands where our doc expects it.
    let apply = incoming
    if (pending.length > 0) {
      const cells = new Set<string>()
      for (const op of pending) for (const k of cellsOf(op)) cells.add(k)
      apply = reconcileRemote({ ...entry, ops: apply }, [{ seq: null, cells }])
      const tIn = transformOps(apply, pending)
      apply = tIn.ops
      this.warnLossy(tIn.lossy)
    }

    // T(pending, incoming): everything still queued must be re-expressed on
    // top of the entry before it ships (the gateway only accepts it against
    // the new head).
    const { edits, moves, removedSheets, lossy } = structEditsOf(incoming)
    this.warnLossy(lossy)
    if (this.queue && (edits.length > 0 || moves.length > 0 || removedSheets.length > 0)) {
      this.queue.transformQueued((ops) => transformOps(ops, incoming).ops)
    }

    for (const op of apply) {
      try {
        cmd.syncExecuteCommand(op.id, op.params, { fromCollab: true })
      } catch {
        // one bad mutation must not stall the stream; the next Save re-bases everyone
      }
    }

    // Structural edits move everyone's coordinates: shift the stored remote
    // selections (and redraw) so cursors don't point at the wrong rows.
    if (edits.length > 0 || moves.length > 0) {
      let moved = false
      for (const { edit } of edits) {
        if (this.state.shiftSelections((sel) => transformSelection(sel, edit, edit.subUnitId))) moved = true
      }
      for (const move of moves) {
        if (this.state.shiftSelections((sel) => transformSelectionThroughMove(sel, move))) moved = true
      }
      if (moved) this.redrawAll()
    }

    this.applied = entry.seq
    this.emitChange()
  }

  private warnLossy(lossy: boolean): void {
    if (!lossy || this.lossyWarned) return
    this.lossyWarned = true
    // eslint-disable-next-line no-console
    console.warn('[injoffice-collab] a concurrent range reorder could not be transformed exactly; positions may drift until the next save')
  }

  /** Replay everything after `since` from the log; reload when it can't. */
  private catchUp(since: number): Promise<void> {
    if (this.catchingUp) return this.catchingUp
    this.catchingUp = (async () => {
      if (!this.transport.opSince || !this.queue) return
      try {
        const res = await this.transport.opSince(this.opts.path, since)
        if (!this.started) return
        if (res.reset) {
          this.resync('reset')
          return
        }
        for (const e of res.ops) {
          if (e.seq <= this.applied) continue
          if (this.state.isSelf(e.client_id)) this.applied = e.seq
          else this.applyEntry(e)
        }
        if (res.head > this.applied && res.ops.length === 0) this.applied = res.head
        this.drain()
        if (this.reorder.size > 0 && [...this.reorder.keys()].some((s) => s <= this.applied)) {
          for (const s of [...this.reorder.keys()]) if (s <= this.applied) this.reorder.delete(s)
        }
      } catch {
        // transport hiccup — the next event or reconnect retries
      }
    })().finally(() => {
      this.catchingUp = null
    })
    return this.catchingUp
  }

  private resync(reason: ResyncReason): void {
    if (this.resynced) return
    this.resynced = true
    this.queue?.clear()
    for (const cb of this.resyncListeners) cb(reason)
  }

  // ---- local → room (presence) --------------------------------------------

  private queuePublish(sel: SheetSelection): void {
    if (!this.joined) return
    if (sameSelection(sel, this.lastPublished)) return
    this.pendingPublish = sel
    if (this.publishTimer) return
    const wait = Math.max(0, this.opts.publishIntervalMs - (Date.now() - this.lastPublishAt))
    this.publishTimer = setTimeout(() => {
      this.publishTimer = null
      const next = this.pendingPublish
      this.pendingPublish = null
      if (!next || !this.joined) return
      this.lastPublished = next
      this.lastPublishAt = Date.now()
      void this.transport.presence(this.opts.path, next).catch(() => undefined)
    }, wait)
  }

  // ---- room → canvas -------------------------------------------------------

  private clearHighlights(): void {
    for (const timer of this.peerLabelTimers.values()) clearTimeout(timer)
    this.peerLabelTimers.clear()
    for (const disposables of this.peerVisuals.values()) for (const d of disposables) d.dispose()
    this.peerVisuals.clear()
  }

  private redrawAll(): void {
    this.clearHighlights()
    for (const p of this.state.peers()) this.drawPeer(p)
  }

  private drawPeer(peer: PeerInfo): void {
    const labelTimer = this.peerLabelTimers.get(peer.client_id)
    if (labelTimer) clearTimeout(labelTimer)
    this.peerLabelTimers.delete(peer.client_id)
    for (const d of this.peerVisuals.get(peer.client_id) ?? []) d.dispose()
    this.peerVisuals.delete(peer.client_id)
    const sel = peer.selection
    if (!sel || sel.ranges.length === 0) return
    const ws = this.api.getActiveWorkbook()?.getActiveSheet()
    if (!ws || ws.getSheetId() !== sel.sheet) return
    try {
      const ranges = sel.ranges.map(([r0, c0, r1, c1]) => ws.getRange(r0, c0, r1 - r0 + 1, c1 - c0 + 1))
      const visuals: IDisposable[] = []
      visuals.push(ws.highlightRanges(
        ranges,
        {
          stroke: peer.color,
          strokeWidth: 1.5,
          fill: withAlpha(peer.color, 0.12),
          widgets: {},
          autofillSize: 0,
        },
        null,
      ))
      const Label = this.opts.peerLabelComponent
      if (Label) {
        // Univer mounts its canvas container after createWorkbook returns.
        // Deferring the popup avoids subscribing to its rect stream before
        // that container exists when two editors join in the same frame.
        const timer = setTimeout(() => {
          this.peerLabelTimers.delete(peer.client_id)
          if (!this.started) return
          const current = this.state.peers().find((candidate) => candidate.client_id === peer.client_id)
          const currentSelection = current?.selection
          const currentSheet = this.api.getActiveWorkbook()?.getActiveSheet()
          if (!current || !currentSelection || currentSelection.ranges.length === 0 || !currentSheet || currentSheet.getSheetId() !== currentSelection.sheet) return
          const [row, column] = currentSelection.active ?? [currentSelection.ranges[0][0], currentSelection.ranges[0][1]]
          try {
            const popup = currentSheet.getRange(row, column).attachRangePopup({
              componentKey: this.peerLabelKey,
              direction: 'top-left',
              offset: [0, -3],
              hideOnInvisible: true,
              noPushMinimumGap: true,
              zIndex: 1000,
              extraProps: { peer: current, selection: currentSelection, labelComponent: Label },
            })
            if (!popup) return
            const currentVisuals = this.peerVisuals.get(peer.client_id)
            if (currentVisuals) currentVisuals.push(popup)
            else popup.dispose()
          } catch {
            // The editor may be tearing down or not have mounted its canvas yet.
          }
        }, 100)
        this.peerLabelTimers.set(peer.client_id, timer)
      }
      this.peerVisuals.set(peer.client_id, visuals)
    } catch {
      // no active sheet yet (highlight throws) — next redraw will catch up
    }
  }
}

function toSelection(sheetId: string, ranges: IRange[], mode: NonNullable<SheetSelection['mode']>): SheetSelection {
  const out: SheetSelection = {
    sheet: sheetId,
    ranges: ranges.slice(0, 64).map((r) => [r.startRow, r.startColumn, r.endRow, r.endColumn]),
    mode,
  }
  const last = ranges[ranges.length - 1]
  if (last) out.active = [last.startRow, last.startColumn]
  return out
}

function currentSelection(api: FUniver, mode: NonNullable<SheetSelection['mode']>): SheetSelection | null {
  const ws = api.getActiveWorkbook()?.getActiveSheet()
  if (!ws) return null
  const r = ws.getSelection()?.getActiveRange()?.getRange()
  if (!r) return null
  return toSelection(ws.getSheetId(), [r], mode)
}

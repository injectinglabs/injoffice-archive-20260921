import type { CollabTransport } from './types'

// DocSyncEngine (D3) — entry-level ordering for collaborative DOCUMENTS
// over the same linear op log the sheet editor uses. The engine is
// editor-agnostic: ops are opaque (ProseMirror steps in practice) and the
// host supplies three hooks. ProseMirror's collab module does the actual
// text OT (rebasing unconfirmed steps when remote entries apply); the
// engine's job is the wire discipline:
//
//   - bootstrap: the loaded file reflects saved_seq — replay (saved, head]
//   - live entries apply in seq order (reorder buffer + gap catch-up)
//   - local changes submit against base_seq = applied; a STALE_BASE
//     rejection means "catch up, let the editor rebase, resubmit"
//   - one submission in flight at a time; new keystrokes accumulate in the
//     editor's own sendable state, not in a queue here
export interface DocSyncHooks {
  /** Apply one remote entry's ops (host: receiveTransaction). May be async. */
  applyRemote(ops: unknown[], clientID: string): void | Promise<void>
  /** The editor's current unconfirmed change set, or null when clean. */
  getPending(): { ops: unknown[] } | null
  /** The submission that just got seq was accepted (host: confirm own steps). */
  onAcked(ops: unknown[]): void
  /** The log can't bring us up to date — reload the document. */
  onResync(): void
}

export class DocSyncEngine {
  private readonly transport: CollabTransport
  private readonly path: string
  private readonly hooks: DocSyncHooks
  private appliedSeq = 0
  private inFlight = false
  private submitScheduled = false
  private catchingUp: Promise<void> | null = null
  private readonly reorder = new Map<number, { ops: unknown[]; clientID: string }>()
  private stopped = false
  private resynced = false

  constructor(
    transport: CollabTransport,
    path: string,
    hooks: DocSyncHooks,
    /** Skip catch-up apply for this id; those ops are already in the local editor. */
    private readonly selfClientId: string | null = null,
  ) {
    this.transport = transport
    this.path = path
    this.hooks = hooks
  }

  /** Last entry seq applied — Save sends this as collab_seq. */
  get applied(): number {
    return this.appliedSeq
  }

  /** Replay everything after the file's saved_seq. Call once after load. */
  async bootstrap(savedSeq: number): Promise<void> {
    this.appliedSeq = savedSeq
    await this.catchUp()
  }

  stop(): void {
    this.stopped = true
    this.reorder.clear()
  }

  /** Feed one live collab.op entry (self-echoes already filtered upstream). */
  receiveEntry(seq: number, ops: unknown[], clientID: string): Promise<void> {
    if (this.stopped || seq <= this.appliedSeq) return Promise.resolve()
    if (seq === this.appliedSeq + 1) {
      return this.apply(seq, ops, clientID).then(() => this.drain())
    }
    this.reorder.set(seq, { ops, clientID })
    return this.catchUp()
  }

  /** The editor changed — submit soon (coalesced; one in flight). */
  submitSoon(): void {
    if (this.stopped || this.submitScheduled) return
    this.submitScheduled = true
    queueMicrotask(() => {
      this.submitScheduled = false
      void this.submitLoop()
    })
  }

  /** After a reconnect: catch up, then push whatever is pending. */
  async resume(): Promise<void> {
    await this.catchUp()
    this.submitSoon()
  }

  private async apply(seq: number, ops: unknown[], clientID: string): Promise<void> {
    if (this.selfClientId && clientID === this.selfClientId) {
      if (seq > this.appliedSeq) this.appliedSeq = seq
      return
    }
    try {
      await this.hooks.applyRemote(ops, clientID)
    } catch {
      // A step that fails to apply means our doc diverged — reload.
      this.resync()
      return
    }
    if (seq > this.appliedSeq) this.appliedSeq = seq
  }

  private async drain(): Promise<void> {
    for (;;) {
      const next = this.reorder.get(this.appliedSeq + 1)
      if (!next) return
      this.reorder.delete(this.appliedSeq + 1)
      await this.apply(this.appliedSeq + 1, next.ops, next.clientID)
    }
  }

  private async submitLoop(): Promise<void> {
    if (this.inFlight || this.stopped || !this.transport.opSubmit) return
    this.inFlight = true
    try {
      for (let attempt = 0; attempt < 8; attempt++) {
        const pending = this.hooks.getPending()
        if (!pending || pending.ops.length === 0) return
        try {
          const seq = await this.transport.opSubmit(this.path, pending.ops as never, this.appliedSeq)
          if (seq > this.appliedSeq) this.appliedSeq = seq
          this.hooks.onAcked(pending.ops)
          await this.drain()
          // Loop again: keystrokes may have accumulated while in flight.
        } catch (e) {
          if (e instanceof Error && e.message.includes('STALE_BASE')) {
            await this.catchUp()
            continue // the editor rebased; resubmit the fresh sendable set
          }
          return // transport down — resume() retries after reconnect
        }
      }
    } finally {
      this.inFlight = false
    }
  }

  private catchUp(): Promise<void> {
    if (this.catchingUp) return this.catchingUp
    this.catchingUp = (async () => {
      if (!this.transport.opSince || this.stopped) return
      try {
        const res = await this.transport.opSince(this.path, this.appliedSeq)
        if (this.stopped) return
        if (res.reset) {
          this.resync()
          return
        }
        for (const e of res.ops) {
          if (e.seq <= this.appliedSeq) continue
          await this.apply(e.seq, e.ops as unknown[], e.client_id)
        }
        await this.drain()
      } catch {
        // transport hiccup — the next event or reconnect retries
      }
    })().finally(() => {
      this.catchingUp = null
    })
    return this.catchingUp
  }

  private resync(): void {
    if (this.resynced) return
    this.resynced = true
    this.hooks.onResync()
  }
}

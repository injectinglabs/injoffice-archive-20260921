import {
  CalculationError,
  assertCalculationResultCurrent,
  validateCalculationResult,
  type CalculationResult,
  type SpillItem,
} from './calculation'

export const FORMULA_COLLABORATION_PROTOCOL = 'injoffice.formula-results.v1' as const

export interface FormulaWorkbookIdentity {
  workbookId: string
  revision: string
  fingerprint: string
}

export interface FormulaCellAddress {
  sheetId: string
  row: number
  column: number
}

export interface FormulaCellOccupancy extends FormulaCellAddress {
  kind: 'empty' | 'formula-anchor' | 'spill' | 'value'
  /** Present for spill cells and optionally their formula anchor. */
  owner?: FormulaCellAddress
  /** A local edit newer than the authoritative calculation input. */
  dirty?: boolean
}

export interface CollaborativeFormulaWrite extends FormulaCellAddress {
  anchor: FormulaCellAddress
  value: SpillItem
}

export interface FormulaApplyConflict extends FormulaCellAddress {
  code: 'dirty' | 'occupied' | 'overlap' | 'missing-anchor' | 'occupancy-missing'
  message: string
}

export type FormulaApplyPlan =
  | { status: 'ready'; writes: CollaborativeFormulaWrite[] }
  | { status: 'conflict'; conflicts: FormulaApplyConflict[] }

export interface FormulaResultEnvelope {
  protocol: typeof FORMULA_COLLABORATION_PROTOCOL
  room: string
  sequence: number
  authorityId: string
  authorityEpoch: string
  result: CalculationResult
}

export interface FormulaCollaborationTarget {
  currentIdentity(): FormulaWorkbookIdentity
  inspect(addresses: readonly FormulaCellAddress[]): FormulaCellOccupancy[] | Promise<FormulaCellOccupancy[]>
  applyDerived(writes: readonly CollaborativeFormulaWrite[], envelope: Readonly<FormulaResultEnvelope>): void | Promise<void>
}

export interface FormulaCollaborationTransport {
  publish(envelope: Readonly<FormulaResultEnvelope>, signal: AbortSignal): Promise<void>
  subscribe(handler: (envelope: unknown) => void): () => void
}

export type FormulaCollaborationEvent =
  | { type: 'published' | 'applied'; sequence: number; jobId: string }
  | { type: 'stale' | 'conflict'; sequence: number; jobId: string; detail: unknown }
  | { type: 'gap'; expected: number; received: number }
  | { type: 'rejected'; detail: unknown }

export interface FormulaCollaborationOptions {
  room: string
  authorityId: string
  authorityEpoch: string
  role: 'authority' | 'follower'
  initialSequence?: number
  transport: FormulaCollaborationTransport
  target: FormulaCollaborationTarget
  onEvent?: (event: FormulaCollaborationEvent) => void
  onRecalculationRequired?: (reason: 'stale' | 'spill-conflict', envelope: FormulaResultEnvelope) => void
  onResyncRequired?: (expectedSequence: number, receivedSequence: number) => void
}

export type FormulaReceiveStatus = 'applied' | 'stale' | 'conflict' | 'duplicate' | 'gap' | 'rejected' | 'blocked'

/** Server-ordered distribution of derived results. Formula edits remain normal
 * workbook operations; only a negotiated authority publishes derived values. */
export class FormulaCollaborationSession {
  private sequence: number
  private blocked = false
  private disposed = false
  private unsubscribe?: () => void
  private applyTail: Promise<void> = Promise.resolve()
  private publishTail: Promise<void> = Promise.resolve()

  constructor(private readonly options: FormulaCollaborationOptions) {
    for (const [name, value] of [['room', options.room], ['authorityId', options.authorityId], ['authorityEpoch', options.authorityEpoch]] as const) {
      if (!validId(value)) throw new TypeError(`${name} must be a bounded stable identifier`)
    }
    this.sequence = options.initialSequence ?? 0
    if (!Number.isSafeInteger(this.sequence) || this.sequence < 0) throw new TypeError('initialSequence must be a non-negative safe integer')
  }

  start(): void {
    if (this.disposed) throw new Error('Formula collaboration session is disposed')
    if (this.unsubscribe) return
    this.unsubscribe = this.options.transport.subscribe((envelope) => {
      void this.receive(envelope)
    })
  }

  get state(): { sequence: number; blocked: boolean; disposed: boolean } {
    return { sequence: this.sequence, blocked: this.blocked, disposed: this.disposed }
  }

  publish(resultInput: unknown, signal = new AbortController().signal): Promise<FormulaResultEnvelope> {
    const result = this.publishTail.then(() => this.publishOne(resultInput, signal))
    this.publishTail = result.then(() => undefined, () => undefined)
    return result
  }

  private async publishOne(resultInput: unknown, signal: AbortSignal): Promise<FormulaResultEnvelope> {
    if (this.disposed) throw new Error('Formula collaboration session is disposed')
    if (this.blocked) throw new Error('Formula collaboration session requires resynchronization')
    if (this.options.role !== 'authority') throw new Error('Only the negotiated formula authority can publish results')
    const result = validateCalculationResult(resultInput)
    assertCalculationResultCurrent(result, this.options.target.currentIdentity())
    const envelope: FormulaResultEnvelope = {
      protocol: FORMULA_COLLABORATION_PROTOCOL,
      room: this.options.room,
      sequence: this.sequence + 1,
      authorityId: this.options.authorityId,
      authorityEpoch: this.options.authorityEpoch,
      result,
    }
    await this.options.transport.publish(clone(envelope), signal)
    this.sequence = envelope.sequence
    this.emit({ type: 'published', sequence: envelope.sequence, jobId: result.jobId })
    return clone(envelope)
  }

  receive(input: unknown): Promise<FormulaReceiveStatus> {
    const result = this.applyTail.then(() => this.applyEnvelope(input))
    this.applyTail = result.then(() => undefined, () => undefined)
    return result
  }

  private async applyEnvelope(input: unknown): Promise<FormulaReceiveStatus> {
    if (this.disposed || this.blocked) return 'blocked'
    let envelope: FormulaResultEnvelope
    try { envelope = validateEnvelope(input, this.options) } catch (error) {
      this.emit({ type: 'rejected', detail: error })
      return 'rejected'
    }
    if (envelope.sequence <= this.sequence) return 'duplicate'
    if (envelope.sequence !== this.sequence + 1) {
      this.blocked = true
      this.emit({ type: 'gap', expected: this.sequence + 1, received: envelope.sequence })
      this.options.onResyncRequired?.(this.sequence + 1, envelope.sequence)
      return 'gap'
    }

    const identity = this.options.target.currentIdentity()
    try { assertCalculationResultCurrent(envelope.result, identity) } catch (error) {
      this.sequence = envelope.sequence
      this.emit({ type: 'stale', sequence: envelope.sequence, jobId: envelope.result.jobId, detail: error })
      this.options.onRecalculationRequired?.('stale', clone(envelope))
      return 'stale'
    }
    const addresses = resultAddresses(envelope.result)
    let plan: FormulaApplyPlan
    try {
      const occupancy = await this.options.target.inspect(addresses)
      plan = planCollaborativeFormulaApply(envelope.result, identity, occupancy)
    } catch (error) {
      this.blocked = true
      this.emit({ type: 'rejected', detail: error })
      this.options.onResyncRequired?.(this.sequence + 1, envelope.sequence)
      return 'blocked'
    }
    if (plan.status === 'conflict') {
      this.sequence = envelope.sequence
      this.emit({ type: 'conflict', sequence: envelope.sequence, jobId: envelope.result.jobId, detail: plan.conflicts })
      this.options.onRecalculationRequired?.('spill-conflict', clone(envelope))
      return 'conflict'
    }
    try {
      await this.options.target.applyDerived(plan.writes, clone(envelope))
    } catch (error) {
      this.blocked = true
      this.emit({ type: 'rejected', detail: error })
      this.options.onResyncRequired?.(this.sequence + 1, envelope.sequence)
      return 'blocked'
    }
    this.sequence = envelope.sequence
    this.emit({ type: 'applied', sequence: envelope.sequence, jobId: envelope.result.jobId })
    return 'applied'
  }

  resync(sequence: number, authorityEpoch = this.options.authorityEpoch): boolean {
    if (this.disposed || authorityEpoch !== this.options.authorityEpoch || !Number.isSafeInteger(sequence) || sequence < this.sequence) return false
    this.sequence = sequence
    this.blocked = false
    return true
  }

  dispose(): void {
    if (this.disposed) return
    this.disposed = true
    this.unsubscribe?.()
    this.unsubscribe = undefined
  }

  private emit(event: FormulaCollaborationEvent): void {
    try { this.options.onEvent?.(event) } catch { /* observers do not change protocol state */ }
  }
}

export function planCollaborativeFormulaApply(
  resultInput: unknown,
  current: FormulaWorkbookIdentity,
  occupancyInput: readonly FormulaCellOccupancy[],
): FormulaApplyPlan {
  const result = validateCalculationResult(resultInput)
  assertCalculationResultCurrent(result, current)
  const addresses = resultAddresses(result)
  const occupancy = new Map<string, FormulaCellOccupancy>()
  for (const item of occupancyInput) {
    const key = addressKey(item)
    if (occupancy.has(key)) throw new CalculationError('INVALID_RESULT', `duplicate occupancy for ${key}`)
    occupancy.set(key, item)
  }
  const conflicts: FormulaApplyConflict[] = []
  const writes: CollaborativeFormulaWrite[] = []
  const claimed = new Map<string, string>()
  for (const cell of result.cells) {
    const anchor = addressOf(cell)
    const values: ReadonlyArray<ReadonlyArray<SpillItem>> = cell.result.kind === 'spill'
      ? cell.result.values
      : [[cell.result.kind === 'value' ? cell.result.value : cell.result]]
    for (let row = 0; row < values.length; row++) for (let column = 0; column < values[row]!.length; column++) {
      const address = { sheetId: cell.sheetId, row: cell.row + row, column: cell.column + column }
      const key = addressKey(address)
      const anchorKey = addressKey(anchor)
      const previous = claimed.get(key)
      if (previous && previous !== anchorKey) {
        conflicts.push({ ...address, code: 'overlap', message: 'Authoritative result spills overlap' })
        continue
      }
      claimed.set(key, anchorKey)
      const present = occupancy.get(key)
      if (!present) conflicts.push({ ...address, code: 'occupancy-missing', message: 'Host did not return occupancy for a result cell' })
      else if (present.dirty) conflicts.push({ ...address, code: 'dirty', message: 'Cell has a newer local edit' })
      else if (row === 0 && column === 0 && present.kind !== 'formula-anchor') conflicts.push({ ...address, code: 'missing-anchor', message: 'Formula anchor no longer exists' })
      else if ((row !== 0 || column !== 0) && present.kind !== 'empty'
        && !(present.kind === 'spill' && present.owner && addressKey(present.owner) === anchorKey)) {
        conflicts.push({ ...address, code: 'occupied', message: 'Spill destination contains unrelated content' })
      }
      writes.push({ ...address, anchor, value: clone(values[row]![column]!) })
    }
  }
  return conflicts.length ? { status: 'conflict', conflicts } : { status: 'ready', writes }
}

function validateEnvelope(input: unknown, options: FormulaCollaborationOptions): FormulaResultEnvelope {
  if (!input || typeof input !== 'object') throw new TypeError('Formula result envelope must be an object')
  const envelope = input as Partial<FormulaResultEnvelope>
  if (envelope.protocol !== FORMULA_COLLABORATION_PROTOCOL || envelope.room !== options.room
    || envelope.authorityId !== options.authorityId || envelope.authorityEpoch !== options.authorityEpoch
    || !Number.isSafeInteger(envelope.sequence) || (envelope.sequence as number) < 1) throw new TypeError('Formula result envelope identity or sequence is invalid')
  return clone({ ...envelope, result: validateCalculationResult(envelope.result) } as FormulaResultEnvelope)
}

function resultAddresses(result: CalculationResult): FormulaCellAddress[] {
  const addresses: FormulaCellAddress[] = []
  const seen = new Set<string>()
  for (const cell of result.cells) {
    const height = cell.result.kind === 'spill' ? cell.result.values.length : 1
    const width = cell.result.kind === 'spill' ? cell.result.values[0]!.length : 1
    for (let row = 0; row < height; row++) for (let column = 0; column < width; column++) {
      const address = { sheetId: cell.sheetId, row: cell.row + row, column: cell.column + column }
      const key = addressKey(address)
      if (!seen.has(key)) { seen.add(key); addresses.push(address) }
    }
  }
  return addresses
}

function addressOf(value: FormulaCellAddress): FormulaCellAddress { return { sheetId: value.sheetId, row: value.row, column: value.column } }
function addressKey(value: FormulaCellAddress): string { return JSON.stringify([value.sheetId, value.row, value.column]) }
function validId(value: string): boolean { return typeof value === 'string' && /^[A-Za-z0-9][A-Za-z0-9._:/-]{0,255}$/.test(value) }
function clone<T>(value: T): T { return structuredClone(value) }

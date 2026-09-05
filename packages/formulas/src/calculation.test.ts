import { describe, expect, it, vi } from 'vitest'
import {
  CALCULATION_PROTOCOL_VERSION,
  CalculationError,
  CalculationManager,
  assertCalculationResultCurrent,
  validateCalculationResult,
  type CalculationEngine,
  type CalculationEngineContext,
  type CalculationEngineOutput,
  type CalculationRequest,
} from './calculation'

interface Snapshot { cells: Record<string, string | number> }

const baseRequest = (overrides: Partial<CalculationRequest<Snapshot>> = {}): CalculationRequest<Snapshot> => ({
  jobId: 'job-1',
  workbookId: 'book-1',
  sourceRevision: 'revision-7',
  sourceFingerprint: 'sha256:source-seven',
  snapshot: { cells: { A1: 2, A2: 3 } },
  formulas: [
    { sheetId: 'sheet-1', row: 0, column: 1, formula: '=SUM(A1:A2)' },
    { sheetId: 'sheet-1', row: 0, column: 2, formula: '=SEQUENCE(2,2)' },
  ],
  metadata: { locale: 'en-US' },
  ...overrides,
})

function engine(calculate?: CalculationEngine<Snapshot>['calculate']): CalculationEngine<Snapshot> {
  return {
    id: 'test-engine',
    version: '1.2.3',
    deterministic: true,
    calculate: calculate ?? (() => ({
      cells: [
        { sheetId: 'sheet-1', row: 0, column: 1, result: { kind: 'value', value: 5 } },
        { sheetId: 'sheet-1', row: 0, column: 2, result: { kind: 'spill', values: [[1, 2], [3, { kind: 'error', code: '#N/A' }]] } },
        { sheetId: 'sheet-1', row: 5, column: 0, result: { kind: 'error', code: '#DIV/0!', message: 'division by zero' } },
      ],
      diagnostics: [{ level: 'info', code: 'CACHE_HIT', message: 'dependency graph reused' }],
    })),
  }
}

describe('CalculationManager', () => {
  it('publishes a deterministic, revision-bound result protocol', async () => {
    const manager = new CalculationManager(engine())
    const result = await manager.submit(baseRequest()).result

    expect(result).toMatchObject({
      protocolVersion: CALCULATION_PROTOCOL_VERSION,
      jobId: 'job-1',
      workbookId: 'book-1',
      sourceRevision: 'revision-7',
      sourceFingerprint: 'sha256:source-seven',
      engineFingerprint: manager.engineFingerprint,
    })
    expect(result.cells.map((cell) => cell.result.kind)).toEqual(['value', 'spill', 'error'])
    expect(result.diagnostics).toEqual([{ level: 'info', code: 'CACHE_HIT', message: 'dependency graph reused' }])
  })

  it('canonicalizes formulas and metadata for order-independent request fingerprints', async () => {
    const first = new CalculationManager(engine())
    const second = new CalculationManager(engine())
    const request = baseRequest()
    const reversed = baseRequest({
      jobId: 'job-2',
      formulas: [...request.formulas].reverse(),
      metadata: { locale: 'en-US' },
    })

    const [one, two] = await Promise.all([first.submit(request).result, second.submit(reversed).result])
    expect(one.requestFingerprint).toBe(two.requestFingerprint)
  })

  it('clones the snapshot and presents formulas in canonical frozen order', async () => {
    let received: CalculationEngineContext<Snapshot> | undefined
    const snapshot = { cells: { A1: 2 } }
    const manager = new CalculationManager(engine((context) => {
      received = context
      return { cells: [] }
    }))
    const job = manager.submit(baseRequest({
      snapshot,
      formulas: [
        { sheetId: 'z', row: 2, column: 0, formula: '=1' },
        { sheetId: 'a', row: 1, column: 0, formula: '=2' },
      ],
    }))
    snapshot.cells.A1 = 99
    await job.result

    expect(received?.snapshot.cells.A1).toBe(2)
    expect(received?.formulas.map((cell) => cell.sheetId)).toEqual(['a', 'z'])
    expect(Object.isFrozen(received?.formulas)).toBe(true)
    expect(Object.isFrozen(received?.formulas[0])).toBe(true)
  })

  it('emits a complete lifecycle and retains inspectable terminal state', async () => {
    const manager = new CalculationManager(engine())
    const events: string[] = []
    manager.onEvent((event) => events.push(event.type))
    const job = manager.submit(baseRequest())
    await job.result

    expect(events).toEqual(['queued', 'started', 'completed'])
    expect(job.getState().state).toBe('completed')
    expect(manager.getJob('job-1')?.state).toBe('completed')
  })

  it('fails closed when the caller expects another engine contract', () => {
    const selectedEngine = engine()
    const calculate = vi.spyOn(selectedEngine, 'calculate')
    const manager = new CalculationManager(selectedEngine)
    expect(() => manager.submit(baseRequest({ expectedEngineFingerprint: 'fnv1a32:wrong' }))).toThrowError(expect.objectContaining({ code: 'ENGINE_MISMATCH' }))
    expect(calculate).not.toHaveBeenCalled()
  })

  it('rejects duplicate active or retained job identities', async () => {
    const manager = new CalculationManager(engine())
    const job = manager.submit(baseRequest())
    expect(() => manager.submit(baseRequest())).toThrowError(expect.objectContaining({ code: 'DUPLICATE_JOB' }))
    await job.result
    expect(() => manager.submit(baseRequest())).toThrowError(expect.objectContaining({ code: 'DUPLICATE_JOB' }))
  })

  it('validates per-job timeouts before reserving the job identity', async () => {
    const manager = new CalculationManager(engine())
    expect(() => manager.submit(baseRequest(), { timeoutMs: 0 })).toThrowError(expect.objectContaining({ code: 'INVALID_REQUEST' }))
    await expect(manager.submit(baseRequest()).result).resolves.toMatchObject({ jobId: 'job-1' })
  })

  it('isolates calculation semantics from throwing lifecycle observers', async () => {
    const manager = new CalculationManager(engine())
    manager.onEvent(() => { throw new Error('observer failed') })
    await expect(manager.submit(baseRequest()).result).resolves.toMatchObject({ jobId: 'job-1' })
  })

  it('cancels promptly even when an engine ignores its signal', async () => {
    const manager = new CalculationManager(engine(() => new Promise<CalculationEngineOutput>(() => {})), { timeoutMs: 5_000 })
    const events: string[] = []
    manager.onEvent((event) => events.push(event.type))
    const job = manager.submit(baseRequest())
    expect(job.cancel('superseded')).toBe(true)

    await expect(job.result).rejects.toMatchObject({ code: 'ABORTED' })
    expect(job.getState().state).toBe('canceled')
    expect(events).toEqual(['queued', 'started', 'canceled'])
    expect(job.cancel()).toBe(false)
  })

  it('supports caller AbortSignals, including already-aborted signals', async () => {
    const controller = new AbortController()
    controller.abort('disconnected')
    const manager = new CalculationManager(engine())
    const job = manager.submit(baseRequest(), { signal: controller.signal })

    await expect(job.result).rejects.toMatchObject({ code: 'ABORTED' })
    expect(job.getState().state).toBe('canceled')
  })

  it('times out promptly even when an engine never settles', async () => {
    const manager = new CalculationManager(engine(() => new Promise<CalculationEngineOutput>(() => {})), { timeoutMs: 10 })
    const job = manager.submit(baseRequest())

    await expect(job.result).rejects.toMatchObject({ code: 'TIMEOUT' })
    expect(job.getState().state).toBe('timed-out')
  })

  it('wraps engine failures without leaking them as protocol results', async () => {
    const manager = new CalculationManager(engine(() => { throw new Error('secret backend detail') }))
    const job = manager.submit(baseRequest())

    await expect(job.result).rejects.toMatchObject({ code: 'ENGINE_FAILED' })
    expect(job.getState().state).toBe('failed')
  })

  it.each([
    {
      name: 'duplicate output cells',
      output: { cells: [
        { sheetId: 's', row: 0, column: 0, result: { kind: 'value', value: 1 } },
        { sheetId: 's', row: 0, column: 0, result: { kind: 'value', value: 2 } },
      ] } as CalculationEngineOutput,
    },
    {
      name: 'ragged spill arrays',
      output: { cells: [{ sheetId: 's', row: 0, column: 0, result: { kind: 'spill', values: [[1, 2], [3]] } }] } as CalculationEngineOutput,
    },
    {
      name: 'non-finite values',
      output: { cells: [{ sheetId: 's', row: 0, column: 0, result: { kind: 'value', value: Number.NaN } }] } as CalculationEngineOutput,
    },
    {
      name: 'out-of-bounds spills',
      output: { cells: [{ sheetId: 's', row: 1_048_575, column: 0, result: { kind: 'spill', values: [[1], [2]] } }] } as CalculationEngineOutput,
    },
  ])('rejects invalid engine output: $name', async ({ output }) => {
    const manager = new CalculationManager(engine(() => output))
    await expect(manager.submit(baseRequest()).result).rejects.toMatchObject({ code: 'INVALID_RESULT' })
  })

  it.each([
    { formulas: [{ sheetId: 's', row: 0, column: 0, formula: 'SUM(A1)' }] },
    { formulas: [
      { sheetId: 's', row: 0, column: 0, formula: '=1' },
      { sheetId: 's', row: 0, column: 0, formula: '=2' },
    ] },
    { sourceRevision: '' },
  ])('rejects invalid requests before execution: %#', (override) => {
    const calculate = vi.fn(() => ({ cells: [] }))
    const manager = new CalculationManager(engine(calculate))
    expect(() => manager.submit(baseRequest(override))).toThrowError(CalculationError)
    expect(calculate).not.toHaveBeenCalled()
  })
})

describe('assertCalculationResultCurrent', () => {
  it('accepts exact workbook/revision/fingerprint identity and rejects stale results', async () => {
    const result = await new CalculationManager(engine()).submit(baseRequest()).result
    expect(() => assertCalculationResultCurrent(result, {
      workbookId: 'book-1', revision: 'revision-7', fingerprint: 'sha256:source-seven',
    })).not.toThrow()
    expect(() => assertCalculationResultCurrent(result, {
      workbookId: 'book-1', revision: 'revision-8', fingerprint: 'sha256:source-eight',
    })).toThrowError(expect.objectContaining({ code: 'STALE_RESULT' }))
  })

  it('validates and detaches results received across a transport boundary', async () => {
    const result = await new CalculationManager(engine()).submit(baseRequest()).result
    const checked = validateCalculationResult(result)
    expect(checked).toEqual(result)
    ;(checked.cells[0] as any).row = 99
    expect(result.cells[0]!.row).not.toBe(99)
    expect(() => validateCalculationResult({ ...result, protocolVersion: 'future' })).toThrowError(expect.objectContaining({ code: 'INVALID_RESULT' }))
    expect(() => validateCalculationResult({ ...result, cells: [{ ...result.cells[0], row: -1 }] })).toThrowError(expect.objectContaining({ code: 'INVALID_RESULT' }))
  })
})

import { describe, expect, it, vi } from 'vitest'
import { RangePreprocessError, RangePreprocessPipeline, fingerprintRangePreprocessStages, isRangePreprocessManifest, type RangePreprocessContext, type RangePreprocessStage } from './preprocess'

const context: RangePreprocessContext = {
  range: { sheetId: 'sheet-1', startRow: 2, startColumn: 3, rowCount: 2, columnCount: 2 },
  revision: 'rev-7',
}

describe('RangePreprocessPipeline', () => {
  it('runs stages by stable order and returns an isolated result and trace', async () => {
    const pipeline = new RangePreprocessPipeline()
    pipeline.register({ id: 'trim', version: '1', order: 20, deterministic: true, process: (grid) => grid.map((row) => row.map((cell) => typeof cell === 'string' ? cell.trim() : cell)) })
    pipeline.register({ id: 'null-zero', version: '2', order: 10, deterministic: true, process: (grid) => grid.map((row) => row.map((cell) => cell === null ? 0 : cell)) })
    const events: string[] = []
    const result = await pipeline.run([[' a ', null], [' b ', 2]], context, { onEvent: (event) => events.push(event.type) })
    expect(result.grid).toEqual([['a', 0], ['b', 2]])
    expect(result.stages.map((stage) => stage.id)).toEqual(['null-zero', 'trim'])
    expect(result.fingerprint).toMatch(/^sha256:[0-9a-f]{64}$/)
    expect(events).toEqual(['started', 'stage-started', 'stage-completed', 'stage-started', 'stage-completed', 'completed'])
  })

  it('has a registration-order-independent collaboration fingerprint', () => {
    const a = new RangePreprocessPipeline()
    const b = new RangePreprocessPipeline()
    const stages = [
      { id: 'alpha', version: '1', order: 10, deterministic: true, process: (grid: never) => grid },
      { id: 'beta', version: '3', order: 10, deterministic: true, process: (grid: never) => grid },
    ]
    a.register(stages[1] as never); a.register(stages[0] as never)
    b.register(stages[0] as never); b.register(stages[1] as never)
    expect(a.fingerprint()).toBe(b.fingerprint())
    expect(fingerprintRangePreprocessStages([
      { id: 'alpha', version: '1', order: 10, deterministic: true },
      { id: 'beta', version: '3', order: 10, deterministic: true },
    ])).toBe('sha256:25187b6e4390a60dee61470cfb40cf8a07deed2a16946e2fb0e3890e49691c84')
  })

  it('snapshots registered contracts and disposes them after caller mutation', async () => {
    const pipeline = new RangePreprocessPipeline()
    const stage: RangePreprocessStage = { id: 'stable', version: '1', order: 5, deterministic: true, process: (grid) => grid }
    const remove = pipeline.register(stage)
    const manifest = pipeline.manifest()
    stage.id = 'changed'
    stage.version = '2'
    stage.order = 99
    stage.deterministic = false
    stage.process = () => { throw new Error('mutated callback must not run') }

    expect(pipeline.manifest()).toEqual(manifest)
    await expect(pipeline.run([['a', null], ['b', 2]], context)).resolves.toMatchObject({ fingerprint: manifest.fingerprint })
    remove()
    expect(pipeline.list()).toEqual([])
  })

  it('runs the verified implementation snapshot when lifecycle callbacks replace a stage', async () => {
    const pipeline = new RangePreprocessPipeline()
    const original = vi.fn((grid: Parameters<RangePreprocessStage['process']>[0]) => grid)
    const replacement = vi.fn((grid: Parameters<RangePreprocessStage['process']>[0]) => grid)
    const remove = pipeline.register({ id: 'stable', version: '1', order: 0, deterministic: true, process: original })
    const manifest = pipeline.manifest()

    const result = await pipeline.run([['a', null], ['b', 2]], context, {
      expectedManifest: manifest,
      onEvent: (event) => {
        if (event.type !== 'started') return
        remove()
        pipeline.register({ id: 'stable', version: '2', order: 0, deterministic: true, process: replacement })
      },
    })

    expect(result.fingerprint).toBe(manifest.fingerprint)
    expect(original).toHaveBeenCalledOnce()
    expect(replacement).not.toHaveBeenCalled()
    expect(pipeline.manifest().fingerprint).not.toBe(manifest.fingerprint)
  })

  it('creates a strict portable manifest and verifies it before execution', async () => {
    const pipeline = new RangePreprocessPipeline()
    pipeline.register({ id: 'trim', version: '1', order: 20, deterministic: true, process: (grid) => grid })
    pipeline.register({ id: 'normalize', version: '3', order: 10, deterministic: true, process: (grid) => grid })
    const manifest = JSON.parse(JSON.stringify(pipeline.manifest()))
    expect(isRangePreprocessManifest(manifest)).toBe(true)
    await expect(pipeline.run([['a', null], ['b', 2]], context, { expectedManifest: manifest })).resolves.toMatchObject({ fingerprint: manifest.fingerprint })

    const drifted = new RangePreprocessPipeline()
    drifted.register({ id: 'trim', version: '2', order: 20, deterministic: true, process: (grid) => grid })
    drifted.register({ id: 'normalize', version: '3', order: 10, deterministic: true, process: (grid) => grid })
    expect(() => drifted.assertCompatibleManifest(manifest)).toThrowError(expect.objectContaining({ code: 'FINGERPRINT_MISMATCH' }))
    expect(isRangePreprocessManifest({ ...manifest, extra: true })).toBe(false)
    expect(isRangePreprocessManifest({ ...manifest, stages: [...manifest.stages].reverse() })).toBe(false)
    expect(isRangePreprocessManifest({ ...manifest, fingerprint: `sha256:${'0'.repeat(64)}` })).toBe(false)
  })

  it('rejects peer contract drift and nondeterministic collaboration stages', async () => {
    const pipeline = new RangePreprocessPipeline()
    pipeline.register({ id: 'random', version: '1', order: 0, deterministic: false, process: (grid) => grid })
    await expect(pipeline.run([['a', null], ['b', 2]], { ...context, mode: 'collaborative' })).rejects.toMatchObject({ code: 'NON_DETERMINISTIC_STAGE', stageId: 'random' })
    await expect(pipeline.run([['a', null], ['b', 2]], context, { expectedFingerprint: `sha256:${'0'.repeat(64)}` })).rejects.toMatchObject({ code: 'FINGERPRINT_MISMATCH' })
  })

  it('cancels before execution and removes its abort listener', async () => {
    const pipeline = new RangePreprocessPipeline()
    const process = vi.fn((grid) => grid)
    pipeline.register({ id: 'never', version: '1', order: 0, deterministic: true, process })
    const controller = new AbortController()
    controller.abort('stop')
    await expect(pipeline.run([['a', null], ['b', 2]], context, { signal: controller.signal })).rejects.toMatchObject({ code: 'ABORTED' })
    expect(process).not.toHaveBeenCalled()
  })

  it('wraps stage failures with stable error metadata and events', async () => {
    const pipeline = new RangePreprocessPipeline()
    pipeline.register({ id: 'explode', version: '1', order: 0, deterministic: true, process: () => { throw new Error('secret host detail') } })
    const events: string[] = []
    await expect(pipeline.run([['a', null], ['b', 2]], context, { onEvent: (event) => events.push(event.type) })).rejects.toEqual(expect.objectContaining({ code: 'STAGE_FAILED', stageId: 'explode' }))
    expect(events).toContain('stage-failed')
  })

  it('rejects duplicate or malformed stages, grids, contexts, and outputs', async () => {
    const pipeline = new RangePreprocessPipeline()
    const remove = pipeline.register({ id: 'valid-stage', version: '1', order: 0, deterministic: true, process: () => [['too', 'wide', 'now']] })
    expect(() => pipeline.register({ id: 'valid-stage', version: '2', order: 1, deterministic: true, process: (grid) => grid })).toThrowError(RangePreprocessError)
    await expect(pipeline.run([['a'], ['b']], context)).rejects.toMatchObject({ code: 'INVALID_GRID' })
    remove()
    await expect(pipeline.run([['a', null], ['b', 2]], { ...context, revision: '' })).rejects.toMatchObject({ code: 'INVALID_CONTEXT' })
  })
})

import { describe, expect, it, vi } from 'vitest'
import { defaultPrintLayout, defaultPrintRender } from './manager'
import { PrintPreviewManager, type PrintPreviewOutput } from './preview'

const snapshot = () => ({
  layout: { ...defaultPrintLayout('sheet-1'), subUnitIds: [{ id: 'sheet-1', range: { startRow: 0, startColumn: 0, endRow: 20, endColumn: 5 } }] },
  render: defaultPrintRender(),
  dialogOpen: true,
})

describe('PrintPreviewManager', () => {
  it('renders a canonical selected-page session while preserving the cell-range target', async () => {
    const renderPreview = vi.fn((request) => ({
      totalPages: 4,
      pages: request.selection.kind === 'pages'
        ? request.selection.pages.map((number: number) => ({ number, widthPoints: 612, heightPoints: 792, payload: { bitmap: `page-${number}` } }))
        : [],
    }))
    const manager = new PrintPreviewManager({ renderPreview })
    const events: string[] = []
    manager.onEvent((event) => events.push(event.type))
    const session = manager.start('preview-1', snapshot(), { selection: { kind: 'pages', pages: [3, 1] } })
    const document = await session.result

    expect(document.selection).toEqual({ kind: 'pages', pages: [1, 3] })
    expect(document.pages.map((page) => page.number)).toEqual([1, 3])
    expect(document.snapshot.layout.subUnitIds).toEqual([{ id: 'sheet-1', range: { startRow: 0, startColumn: 0, endRow: 20, endColumn: 5 } }])
    expect(session.getState().state).toBe('ready')
    expect(events).toEqual(['started', 'ready'])
  })

  it('supports inclusive page ranges', async () => {
    const manager = new PrintPreviewManager({
      renderPreview: () => ({
        totalPages: 5,
        pages: [2, 3, 4].map((number) => ({ number, widthPoints: 595, heightPoints: 842, payload: number })),
      }),
    })
    const result = await manager.start('preview-range', snapshot(), { selection: { kind: 'range', startPage: 2, endPage: 4 } }).result
    expect(result.pages.map((page) => page.number)).toEqual([2, 3, 4])
  })

  it('cancels promptly when a host ignores AbortSignal', async () => {
    const manager = new PrintPreviewManager({ renderPreview: () => new Promise<PrintPreviewOutput>(() => {}) }, { timeoutMs: 5_000 })
    const session = manager.start('preview-cancel', snapshot())
    expect(session.cancel('superseded')).toBe(true)
    await expect(session.result).rejects.toMatchObject({ code: 'ABORTED' })
    expect(session.getState().state).toBe('canceled')
    expect(session.cancel()).toBe(false)
  })

  it('times out a non-cooperative host', async () => {
    const manager = new PrintPreviewManager({ renderPreview: () => new Promise<PrintPreviewOutput>(() => {}) }, { timeoutMs: 10 })
    const session = manager.start('preview-timeout', snapshot())
    await expect(session.result).rejects.toMatchObject({ code: 'TIMEOUT' })
    expect(session.getState().state).toBe('timed-out')
  })

  it.each([
    { totalPages: 2, pages: [{ number: 3, widthPoints: 10, heightPoints: 10, payload: null }] },
    { totalPages: 2, pages: [{ number: 1, widthPoints: 0, heightPoints: 10, payload: null }, { number: 2, widthPoints: 10, heightPoints: 10, payload: null }] },
    { totalPages: 2, pages: [{ number: 1, widthPoints: 10, heightPoints: 10, payload: null }] },
  ])('rejects invalid host page output: %#', async (output) => {
    const manager = new PrintPreviewManager({ renderPreview: () => output })
    await expect(manager.start('preview-invalid', snapshot()).result).rejects.toMatchObject({ code: 'INVALID_OUTPUT' })
  })

  it('isolates rendering from throwing lifecycle observers and wraps host failures', async () => {
    const ready = new PrintPreviewManager({ renderPreview: () => ({ totalPages: 0, pages: [] }) })
    ready.onEvent(() => { throw new Error('observer failure') })
    await expect(ready.start('preview-empty', snapshot()).result).resolves.toMatchObject({ totalPages: 0 })

    const failed = new PrintPreviewManager({ renderPreview: () => { throw new Error('renderer failure') } })
    await expect(failed.start('preview-failed', snapshot()).result).rejects.toMatchObject({ code: 'HOST_FAILED' })
  })

  it('rejects duplicate session ids and invalid page selections before invoking the host', async () => {
    const renderPreview = vi.fn(() => ({ totalPages: 0, pages: [] }))
    const manager = new PrintPreviewManager({ renderPreview })
    const first = manager.start('same-session', snapshot())
    expect(() => manager.start('same-session', snapshot())).toThrowError(expect.objectContaining({ code: 'DUPLICATE_SESSION' }))
    await first.result
    expect(() => manager.start('bad-pages', snapshot(), { selection: { kind: 'pages', pages: [1, 1] } })).toThrowError(expect.objectContaining({ code: 'INVALID_REQUEST' }))
    expect(() => manager.start('bad-selection', snapshot(), { selection: null as never })).toThrowError(expect.objectContaining({ code: 'INVALID_REQUEST' }))
    expect(renderPreview).toHaveBeenCalledTimes(1)
  })
})

import { describe, expect, it, vi } from 'vitest'
import { createBrowserCollabHub, describePresence } from './browserCollabTransport'

describe('browser collaboration transport', () => {
  it('joins peers, broadcasts presence, and reports departures', async () => {
    const hub = createBrowserCollabHub()
    const mira = hub.connect({ userId: 'mira', color: '#2855d9' })
    const noah = hub.connect({ userId: 'noah', color: '#b34c36' })
    const miraEvent = vi.fn()
    const noahEvent = vi.fn()
    mira.onEvent(miraEvent)
    noah.onEvent(noahEvent)

    const first = await mira.join('demo', 'Mira')
    const second = await noah.join('demo', 'Noah')
    expect(first.peers).toEqual([])
    expect(second.peers).toHaveLength(1)
    expect(miraEvent).toHaveBeenCalledWith(expect.objectContaining({ event: 'collab.peer.joined' }))

    await mira.presence('demo', { sheet: 's1', ranges: [[1, 1, 1, 1]], active: [1, 1], mode: 'editing' })
    expect(noahEvent).toHaveBeenCalledWith(expect.objectContaining({
      event: 'collab.presence',
      payload: expect.objectContaining({ selection: expect.objectContaining({ mode: 'editing' }) }),
    }))
    expect(hub.getSnapshot().activity[0].detail).toBe('started editing a cell')

    await mira.presence('demo', { sheet: 's1', ranges: [[1, 1, 1, 1]], active: [1, 1], mode: 'editing', draft: 'Ready' })
    expect(hub.getSnapshot().activity[0].detail).toBe('typing “Ready”')

    await noah.leave('demo')
    expect(miraEvent).toHaveBeenCalledWith(expect.objectContaining({ event: 'collab.peer.left' }))
  })

  it('orders operations, rejects stale bases, and supports catch-up', async () => {
    const hub = createBrowserCollabHub()
    const mira = hub.connect({ userId: 'mira', color: '#2855d9' })
    const noah = hub.connect({ userId: 'noah', color: '#b34c36' })
    await mira.join('demo', 'Mira')
    await noah.join('demo', 'Noah')

    const seq = await mira.opSubmit!('demo', [{ id: 'sheet.mutation.set-range-values', params: { value: 'Ready' } }], 0)
    expect(seq).toBe(1)
    await expect(noah.opSubmit!('demo', [], 0)).rejects.toThrow('STALE_BASE')

    const catchUp = await noah.opSince!('demo', 0)
    expect(catchUp).toMatchObject({ head: 1, reset: false })
    expect(catchUp.ops).toHaveLength(1)
    expect(hub.getSnapshot()).toMatchObject({ peers: 2, head: 1 })
  })

  it('describes document, slide, and PDF presence without assuming a sheet', async () => {
    expect(describePresence({ from: 3, to: 8 })).toBe('selected 3–8')
    expect(describePresence({ from: 4, to: 4 })).toBe('caret at 4')
    expect(describePresence({ slideId: 's2' })).toBe('viewing slide s2')
    expect(describePresence({ page: 2 })).toBe('on page 2')

    const hub = createBrowserCollabHub()
    const mira = hub.connect({ userId: 'mira', color: '#2855d9' })
    await mira.join('doc', 'Mira')
    await mira.presence('doc', { from: 1, to: 1 })
    expect(hub.getSnapshot().activity[0].detail).toBe('caret at 1')
  })
})

import { describe, expect, it } from 'vitest'
import { applyDeckOps, DeckSyncEngine, diffDeck, type SyncDeck } from './deckSync'

type TestSlide = { id: string; title?: string; kind?: string; bullets?: string[]; shapeOverrides?: Record<string, unknown> }
type TestDeck = SyncDeck & { slides: TestSlide[] }

const deck = (): TestDeck => ({
  id: 'd1', title: 'Quarterly review', theme: 'slate',
  slides: [
    { id: 'intro', kind: 'title', title: 'Hello', shapeOverrides: { title: { x: 10, y: 20 } } },
    { id: 'results', kind: 'bullets', title: 'Results', bullets: ['Up'] },
  ],
})

describe('DeckSpec operation model', () => {
  it('diffs stable slide fields and independent shape overrides', () => {
    const before = deck()
    const after = structuredClone(before)
    after.slides[1].title = 'Q3 results'
    after.slides[0].shapeOverrides = { title: { x: 99, y: 20 } }
    const ops = diffDeck(before, after)
    expect(ops).toEqual([
      { id: 'deck.set-shape-override', params: { slideId: 'intro', shapeKey: 'title', value: { x: 99, y: 20 } } },
      { id: 'deck.set-field', params: { scope: 'slide', slideId: 'results', field: 'title', value: 'Q3 results' } },
    ])
    expect(applyDeckOps(before, ops ?? [])).toEqual(after)
  })

  it('diffs one stable-ID slide move without relying on array positions', () => {
    const next = deck()
    next.slides.reverse()
    expect(diffDeck(deck(), next)).toEqual([
      { id: 'deck.move-slide', params: { slideId: 'intro', afterId: 'results' } },
    ])
    expect(applyDeckOps(deck(), diffDeck(deck(), next) ?? [])).toEqual(next)
  })

  it('diffs one insert and one delete, preserving a complete inserted slide payload', () => {
    const inserted = deck()
    inserted.slides.splice(1, 0, { id: 'agenda', kind: 'bullets', title: 'Agenda', bullets: ['One', 'Two'] })
    expect(diffDeck(deck(), inserted)).toEqual([
      { id: 'deck.insert-slide', params: { slide: inserted.slides[1], afterId: 'intro' } },
    ])
    expect(applyDeckOps(deck(), diffDeck(deck(), inserted) ?? [])).toEqual(inserted)

    const deleted = deck()
    deleted.slides.splice(0, 1)
    expect(diffDeck(deck(), deleted)).toEqual([{ id: 'deck.delete-slide', params: { slideId: 'intro' } }])
    expect(applyDeckOps(deck(), diffDeck(deck(), deleted) ?? [])).toEqual(deleted)
  })

  it('fails closed for compound structural rewrites and invalid remote operations', () => {
    const rewritten = deck()
    rewritten.slides = [{ id: 'new-1', kind: 'title' }, { id: 'new-2', kind: 'closing' }]
    expect(diffDeck(deck(), rewritten)).toBeNull()

    const changedWhileMoving = deck()
    changedWhileMoving.slides.reverse()
    changedWhileMoving.slides[0].title = 'Changed too'
    expect(diffDeck(deck(), changedWhileMoving)).toBeNull()

    const malformed = applyDeckOps(deck(), [
      { id: 'deck.insert-slide', params: { slide: { id: 'intro', kind: 'title' } } },
      { id: 'deck.delete-slide', params: { slideId: 'intro' } },
      { id: 'deck.delete-slide', params: { slideId: 'results' } },
      { id: 'deck.move-slide', params: { slideId: 'missing', afterId: null } },
    ])
    // Duplicate insert is ignored; delete will never make an empty deck.
    expect(malformed.slides.map((slide) => slide.id)).toEqual(['results'])
  })

  it('converges ordered structural entries even when an anchor was deleted first', () => {
    const ops = [
      { id: 'deck.delete-slide', params: { slideId: 'intro' } },
      { id: 'deck.insert-slide', params: { slide: { id: 'agenda', kind: 'bullets' }, afterId: 'intro' } },
      { id: 'deck.move-slide', params: { slideId: 'results', afterId: 'agenda' } },
    ]
    const a = applyDeckOps(deck(), ops)
    const b = applyDeckOps(deck(), ops)
    expect(a).toEqual(b)
    // A missing anchor intentionally means append, a deterministic safe
    // fallback shared by all collaborators replaying the ordered log.
    expect(a.slides.map((slide) => slide.id)).toEqual(['agenda', 'results'])
  })

  it('rebases a local edit over an ordered remote entry after STALE_BASE', async () => {
    let remote = deck()
    const sent: unknown[][] = []
    let first = true
    const transport = {
      opSubmit: async (_path: string, ops: unknown[]) => {
        sent.push(ops)
        if (first) { first = false; throw new Error('STALE_BASE') }
        return 2
      },
      opSince: async (_path: string, since: number) => ({
        reset: false,
        head: 1,
        ops: since === 0 && !first ? [{ room: 'r', seq: 1, client_id: 'other', ops: [{ id: 'deck.set-field', params: { scope: 'slide', slideId: 'results', field: 'bullets', value: ['Remote'] } }] }] : [],
      }),
    }
    const shown: TestDeck[] = []
    const engine = new DeckSyncEngine(transport, '/deck.json', deck(), { onDeck: (value) => shown.push(value), onResync: () => expect.unreachable() })
    await engine.bootstrap(0)
    const next = structuredClone(remote)
    next.slides[1].title = 'Local title'
    expect(engine.change(next)).toBe(true)
    await new Promise((resolve) => setTimeout(resolve, 0))
    expect(sent).toHaveLength(2)
    expect(shown.at(-1)?.slides[1]).toMatchObject({ title: 'Local title', bullets: ['Remote'] })
    remote = shown.at(-1) ?? remote
    expect(engine.applied).toBe(2)
  })
})

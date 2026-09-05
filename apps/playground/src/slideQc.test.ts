import { describe, expect, it } from 'vitest'
import { deckFromOutline, resetSlideIds } from '../../../packages/slides/src/outline'
import { playgroundDeckAudit, updateSlideTransition } from './slideQc'

describe('playground deck QC', () => {
  it('audits a short outline as passing and a long title as overflowing', () => {
    resetSlideIds()
    const ok = playgroundDeckAudit(deckFromOutline('InjOffice', '# Title\nA short deck.'))
    expect(ok.issues).toEqual([])
    expect(ok.report).toContain('Passed')

    resetSlideIds()
    const overflow = playgroundDeckAudit(deckFromOutline('InjOffice', `# ${'Overflow '.repeat(80)}\n${'- bullet '.repeat(40)}`))
    expect(overflow.issues.length).toBeGreaterThan(0)
    expect(overflow.report).toContain('issue')
  })

  it('updates only the selected slide with an exact supported transition', () => {
    resetSlideIds()
    const deck = deckFromOutline('InjOffice', '# One\nFirst\n# Two\nSecond')
    const pushed = updateSlideTransition(deck, 1, 'push', 'right')
    expect(pushed.slides[0]?.transition).toBeUndefined()
    expect(pushed.slides[1]?.transition).toEqual({ kind: 'push', direction: 'right' })
    expect(updateSlideTransition(pushed, 1, 'fade').slides[1]?.transition).toEqual({ kind: 'fade' })
    expect(updateSlideTransition(pushed, 1, 'none').slides[1]?.transition).toBeUndefined()
    expect(deck.slides[1]?.transition).toBeUndefined()
  })
})

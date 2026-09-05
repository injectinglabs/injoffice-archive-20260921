import { describe, expect, it } from 'vitest'
import { SheetPresenceLabel, SheetPresencePopup } from './SheetPresenceLabel'

const peer = {
  client_id: 'peer-1',
  user_id: 'mira',
  name: 'Mira',
  color: '#2855d9',
  joined_at: 1,
}

describe('SheetPresencePopup', () => {
  it('fails closed when Univer renders it without a popup payload', () => {
    expect(SheetPresencePopup({})).toBeNull()
    expect(SheetPresencePopup({ popup: {} })).toBeNull()
  })

  it('adapts Univer popup details to the public label component props', () => {
    const selection = { sheet: 's1', ranges: [[1, 1, 1, 1]], mode: 'editing' as const }
    const element = SheetPresencePopup({
      popup: { extraProps: { peer, selection, labelComponent: SheetPresenceLabel } },
    })
    expect(element).toMatchObject({ props: { peer, selection } })
  })
})

import { describe, expect, it } from 'vitest'
import { PresenceState } from './presence'
import { initials, parseSelection, rangeA1, sameSelection, selectionA1, withAlpha } from './selection'
import type { JoinResult, PeerInfo } from './types'

const peer = (id: string, name: string, joined: number, selection?: unknown): PeerInfo =>
  ({ client_id: id, user_id: 'u-' + id, name, color: '#0fa98f', joined_at: joined, selection: selection as never })

const join = (peers: PeerInfo[]): JoinResult => ({
  room: 'room1',
  self: peer('me', 'Me', 10),
  peers,
  file: { mtime: 1, size: 2 },
})

describe('parseSelection', () => {
  it('accepts a well-formed selection and normalizes range order', () => {
    expect(parseSelection({ sheet: 's1', ranges: [[5, 3, 1, 0]], active: [1, 0], mode: 'editing', draft: 'hello' })).toEqual({
      sheet: 's1',
      ranges: [[1, 0, 5, 3]],
      active: [1, 0],
      mode: 'editing',
      draft: 'hello',
    })
  })
  it('rejects garbage', () => {
    expect(parseSelection(null)).toBeNull()
    expect(parseSelection({ sheet: '', ranges: [] })).toBeNull()
    expect(parseSelection({ sheet: 's', ranges: [[1, 2, 3]] })).toBeNull()
    expect(parseSelection({ sheet: 's', ranges: [[-1, 0, 0, 0]] })).toBeNull()
    expect(parseSelection({ sheet: 's', ranges: [['a', 0, 0, 0]] })).toBeNull()
    expect(parseSelection({ sheet: 's', ranges: Array.from({ length: 65 }, () => [0, 0, 0, 0]) })).toBeNull()
  })
  it('drops a malformed active cell but keeps the ranges', () => {
    expect(parseSelection({ sheet: 's', ranges: [[0, 0, 0, 0]], active: [1] })).toEqual({ sheet: 's', ranges: [[0, 0, 0, 0]] })
    expect(parseSelection({ sheet: 's', ranges: [[0, 0, 0, 0]], mode: 'typing' })).toEqual({ sheet: 's', ranges: [[0, 0, 0, 0]] })
    expect(parseSelection({ sheet: 's', ranges: [[0, 0, 0, 0]], mode: 'selecting', draft: 'not editing' })).toEqual({ sheet: 's', ranges: [[0, 0, 0, 0]], mode: 'selecting' })
    expect(parseSelection({ sheet: 's', ranges: [[0, 0, 0, 0]], mode: 'editing', draft: 'x'.repeat(513) })).toEqual({ sheet: 's', ranges: [[0, 0, 0, 0]], mode: 'editing' })
  })
})

describe('selection helpers', () => {
  it('compares selections structurally', () => {
    const a = { sheet: 's', ranges: [[0, 0, 1, 1]], active: [0, 0] as [number, number] }
    expect(sameSelection(a, { ...a, ranges: [[0, 0, 1, 1]] })).toBe(true)
    expect(sameSelection(a, { ...a, ranges: [[0, 0, 1, 2]] })).toBe(false)
    expect(sameSelection(a, { ...a, mode: 'editing' })).toBe(false)
    expect(sameSelection({ ...a, mode: 'editing', draft: 'one' }, { ...a, mode: 'editing', draft: 'two' })).toBe(false)
    expect(sameSelection(a, { sheet: 's', ranges: [[0, 0, 1, 1]] })).toBe(false)
    expect(sameSelection(null, null)).toBe(true)
    expect(sameSelection(a, null)).toBe(false)
  })
  it('renders A1 labels', () => {
    expect(rangeA1([0, 0, 0, 0])).toBe('A1')
    expect(rangeA1([1, 1, 4, 3])).toBe('B2:D5')
    expect(rangeA1([0, 26, 0, 27])).toBe('AA1:AB1')
    expect(selectionA1({ sheet: 's', ranges: [[0, 0, 0, 0], [2, 2, 8, 2]] })).toBe('A1, C3:C9')
    expect(selectionA1(null)).toBe('')
  })
  it('initials and alpha', () => {
    expect(initials('Ann Lee')).toBe('AL')
    expect(initials('nick')).toBe('NI')
    expect(initials('  ')).toBe('?')
    expect(withAlpha('#0fa98f', 0.2)).toBe('rgba(15, 169, 143, 0.2)')
    expect(withAlpha('red', 0.2)).toBe('red')
  })
})

describe('PresenceState', () => {
  it('seeds remote peers from join, excluding self, parsing selections', () => {
    const st = new PresenceState()
    st.applyJoin(join([peer('a', 'Ann', 1, { sheet: 's1', ranges: [[0, 0, 1, 1]] }), peer('me', 'Me', 10), peer('b', 'Bob', 2, 'junk')]))
    const peers = st.peers()
    expect(peers.map((p) => p.client_id)).toEqual(['a', 'b'])
    expect(peers[0].selection).toEqual({ sheet: 's1', ranges: [[0, 0, 1, 1]] })
    expect(peers[1].selection).toBeNull()
    expect(st.clientId).toBe('me')
  })

  it('folds events in and ignores other rooms / self echoes', () => {
    const st = new PresenceState()
    st.applyJoin(join([]))
    expect(st.apply({ event: 'collab.peer.joined', payload: { room: 'other', peer: peer('x', 'X', 3) } })).toBeNull()
    expect(st.apply({ event: 'collab.peer.joined', payload: { room: 'room1', peer: peer('me', 'Me', 3) } })).toBeNull()
    expect(st.apply({ event: 'collab.peer.joined', payload: { room: 'room1', peer: peer('c', 'Cy', 3) } })).toEqual({ kind: 'peers' })
    const moved = st.apply({ event: 'collab.presence', payload: { room: 'room1', client_id: 'c', selection: { sheet: 's2', ranges: [[3, 3, 3, 3]] } } })
    expect(moved?.kind).toBe('selection')
    expect(st.peers()[0].selection).toEqual({ sheet: 's2', ranges: [[3, 3, 3, 3]] })
    expect(st.apply({ event: 'collab.presence', payload: { room: 'room1', client_id: 'ghost', selection: {} } })).toBeNull()
    expect(st.nameOf('c')).toBe('Cy')
    expect(st.apply({ event: 'collab.peer.left', payload: { room: 'room1', client_id: 'c' } })).toEqual({ kind: 'peers' })
    expect(st.apply({ event: 'collab.peer.left', payload: { room: 'room1', client_id: 'c' } })).toBeNull()
    expect(st.peers()).toEqual([])
  })

  it('surfaces file changes from others but never its own save', () => {
    const st = new PresenceState()
    st.applyJoin(join([]))
    const base = { room: 'room1', path: '/x.xlsx', author: 'user' as const, action: 'saved', mtime: 1, size: 1 }
    expect(st.apply({ event: 'collab.file.changed', payload: { ...base, origin: 'me' } })).toBeNull()
    expect(st.apply({ event: 'collab.file.changed', payload: { ...base, origin: 'a' } })?.kind).toBe('file')
    expect(st.apply({ event: 'collab.file.changed', payload: { ...base, author: 'agent', action: 'delivered' } })?.kind).toBe('file')
  })

  it('re-join replaces the remote set', () => {
    const st = new PresenceState()
    st.applyJoin(join([peer('a', 'Ann', 1)]))
    st.applyJoin(join([peer('b', 'Bob', 2)]))
    expect(st.peers().map((p) => p.name)).toEqual(['Bob'])
    st.clear()
    expect(st.peers()).toEqual([])
    expect(st.apply({ event: 'collab.peer.joined', payload: { room: 'room1', peer: peer('z', 'Z', 1) } })).toBeNull()
  })
})

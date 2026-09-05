import { describe, expect, it } from 'vitest'
import { decodePdfCollabOperation, encodePdfCollabOperation } from './collab.js'

describe('PDF collaboration operation contract', () => {
  it('accepts supported annotation and form operation envelopes', () => {
    for (const kind of ['annotation.markup', 'annotation.note', 'annotation.delete', 'form.value']) {
      expect(decodePdfCollabOperation({ kind, value: { id: 'saved-object' } }).ok).toBe(true)
    }
  })

  it('preserves accepted opaque envelopes without widening their kinds', () => {
    const operation = { kind: 'form.value', value: { name: 'customer.taxId', kind: 'text', value: 'updated' } } as const
    expect(encodePdfCollabOperation(operation)).toEqual(operation)
  })

  it('rejects missing, primitive, array, and structural/content operations', () => {
    expect(decodePdfCollabOperation(null).ok).toBe(false)
    expect(decodePdfCollabOperation({ kind: 'form.value' }).ok).toBe(false)
    expect(decodePdfCollabOperation({ kind: 'form.value', value: [] })).toEqual({ ok: false, reason: 'operation value must be an object' })
    expect(decodePdfCollabOperation({ kind: 'page.rotate', value: {} })).toEqual({ ok: false, reason: 'unsupported PDF collaboration operation: page.rotate' })
    expect(() => encodePdfCollabOperation({ kind: 'image.delete', value: {} } as never)).toThrow('unsupported PDF collaboration operation')
  })
})

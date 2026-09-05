import { describe, expect, it, vi } from 'vitest'
import {
  CONNECTOR_NATIVE_EXTENSION_NAMESPACE,
  CONNECTOR_NATIVE_EXTENSION_VERSION,
  ConnectorNativePersistence,
  ConnectorNativePersistenceError,
  fromNativeConnectorDefinitions,
  hydrateConnectorManagerFromNative,
  persistConnectorManagerToNative,
  toNativeConnectorDefinitions,
  type ConnectorNativePersistenceCodec,
} from './native'
import type { ConnectorSpec } from './types'

function spec(overrides: Partial<ConnectorSpec> = {}): ConnectorSpec {
  return {
    id: 'weather', name: 'Weather', source: { kind: 'http', url: '/gateway/weather', format: 'json', path: 'items' },
    target: { sheetId: 'sheet-1', startRow: 2, startColumn: 1 }, refresh: 'interval',
    schedule: { intervalMs: 60_000 }, cache: { mode: 'memory', ttlMs: 30_000 },
    schema: { columns: [{ index: 0, type: 'string' }, { index: 1, type: 'number', nullable: true }] },
    ...overrides,
  }
}

function codec(overrides: Partial<ConnectorNativePersistenceCodec> = {}): ConnectorNativePersistenceCodec {
  return {
    readConnectorDefinitions: vi.fn(async () => [spec()]),
    setConnectorDefinitions: vi.fn(async () => new Uint8Array([9, 8, 7])),
    ...overrides,
  }
}

describe('native connector bridge', () => {
  it('exposes a stable extension identity and copies valid wire definitions', () => {
    expect(CONNECTOR_NATIVE_EXTENSION_VERSION).toBe(1)
    expect(CONNECTOR_NATIVE_EXTENSION_NAMESPACE).toContain('/xlsx/connectors/2026')
    const original = spec()
    const wire = toNativeConnectorDefinitions([original])
    wire[0].name = 'copy'
    expect(original.name).toBe('Weather')
    expect(fromNativeConnectorDefinitions([original])).toEqual([original])
  })

  it.each([
    ['query secrets', { source: { kind: 'http', url: '/gateway/weather?token=secret', format: 'json' } }],
    ['userinfo', { source: { kind: 'http', url: 'https://user:pass@example.test/data', format: 'json' } }],
    ['row bounds', { target: { sheetId: 'sheet-1', startRow: 1_048_576, startColumn: 0 } }],
    ['column bounds', { target: { sheetId: 'sheet-1', startRow: 0, startColumn: 16_384 } }],
  ] as const)('rejects %s from the persisted subset', (_name, patch) => {
    expect(() => toNativeConnectorDefinitions([spec(patch as Partial<ConnectorSpec>)]))
      .toThrow(ConnectorNativePersistenceError)
  })

  it('rejects duplicate IDs and unknown fields from native responses', () => {
    expect(() => fromNativeConnectorDefinitions([spec(), spec()])).toThrow(/unique connector snapshot/)
    expect(() => fromNativeConnectorDefinitions([{ ...spec(), token: 'secret' }])).toThrow(/credential-free/)
  })

  it('isolates workbook bytes and codec results on read and write', async () => {
    const original = new Uint8Array([1, 2, 3])
    const adapter = codec({
      readConnectorDefinitions: vi.fn(async (bytes) => {
        bytes.fill(0)
        return [spec()]
      }),
      setConnectorDefinitions: vi.fn(async (bytes, definitions) => {
        bytes.fill(0)
        ;(definitions[0] as ConnectorSpec).name = 'codec mutation'
        return new Uint8Array([4, 5, 6])
      }),
    })
    const persistence = new ConnectorNativePersistence(adapter)
    await expect(persistence.read(original)).resolves.toEqual([spec()])
    const output = await persistence.write(original, [spec()])
    output.fill(7)
    expect(original).toEqual(new Uint8Array([1, 2, 3]))
    expect(await persistence.write(original, [spec()])).toEqual(new Uint8Array([4, 5, 6]))
  })

  it('hydrates definitions without on-open refresh state and persists manager state', async () => {
    const persistence = new ConnectorNativePersistence(codec())
    const manager = {
      serialize: vi.fn(() => [spec({ refresh: 'manual', schedule: undefined })]),
      restoreState: vi.fn(() => true),
    }
    await expect(hydrateConnectorManagerFromNative(manager, persistence, new Uint8Array([1])))
      .resolves.toEqual([spec()])
    expect(manager.restoreState).toHaveBeenCalledWith({
      version: 1,
      connectors: [{ spec: spec(), status: {}, lastRows: 0, lastColumns: 0 }],
    })
    await expect(persistConnectorManagerToNative(manager, persistence, new Uint8Array([1])))
      .resolves.toEqual(new Uint8Array([9, 8, 7]))
  })

  it('fails closed when manager restore or codec responses are invalid', async () => {
    const invalid = new ConnectorNativePersistence(codec({ readConnectorDefinitions: vi.fn(async () => [{ ...spec(), token: 'x' }]) }))
    await expect(invalid.read(new Uint8Array([1]))).rejects.toMatchObject({ code: 'invalid-response' })

    const valid = new ConnectorNativePersistence(codec())
    await expect(hydrateConnectorManagerFromNative(
      { serialize: () => [], restoreState: () => false }, valid, new Uint8Array([1]),
    )).rejects.toMatchObject({ code: 'restore-failed' })
  })

  it('checks cancellation before and after native work', async () => {
    const before = new AbortController()
    before.abort('closed')
    const adapter = codec()
    const persistence = new ConnectorNativePersistence(adapter)
    await expect(persistence.read(new Uint8Array([1]), { signal: before.signal })).rejects.toMatchObject({ code: 'canceled' })
    expect(adapter.readConnectorDefinitions).not.toHaveBeenCalled()

    const after = new AbortController()
    const late = new ConnectorNativePersistence(codec({ readConnectorDefinitions: vi.fn(async () => {
      after.abort('stale')
      return [spec()]
    }) }))
    await expect(late.read(new Uint8Array([1]), { signal: after.signal })).rejects.toMatchObject({ code: 'canceled' })
  })
})

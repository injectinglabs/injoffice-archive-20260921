import { describe, expect, it } from 'vitest'
import {
  DOCX_PAGE_PAINT_WORKER_PROTOCOL,
  DOCX_PAGE_PAINT_WORKER_VERSION,
  dispatchNativeDocxPagePaintWorkerRequestV1,
  encodeNativeDocxPagePaintWorkerResponseV1,
} from './protocol.js'

describe('native DOCX page-paint worker protocol', () => {
  it('refuses undeclared or malformed host-size policy objects', async () => {
    for (const font_size_policy of [{ kind: 'word-default', half_points: 22 }, { kind: 'host-default-size-v1', half_points: 24 }, { kind: 'host-default-size-v1', half_points: 22, extra: true }]) {
      const response = await dispatchNativeDocxPagePaintWorkerRequestV1({ protocol: DOCX_PAGE_PAINT_WORKER_PROTOCOL, version: 1, id: 'size:malformed', op: 'render-approximate', input: { prepare: {}, eligibility: {}, font_size_policy } })
      expect(response).toMatchObject({ ok: false, error: { code: 'COMPILATION_REFUSED' } })
    }
  })
  it('keeps automatic-border approximation opt-in and rejects caller policy/font overrides', async () => {
    for (const input of [{}, { prepare: {}, eligibility: {} }, { prepare: {}, legacy_eligibility: {}, policy: 'auto-is-black' }, { prepare: {}, host_font_manifest_path: '/caller/fonts.json' }, { prepare: { font_manifest: {} } }]) {
      const response = await dispatchNativeDocxPagePaintWorkerRequestV1({ protocol: DOCX_PAGE_PAINT_WORKER_PROTOCOL, version: 1, id: 'automatic:malformed', op: 'render-auto-borders', input })
      expect(response).toMatchObject({ ok: false, error: { code: 'COMPILATION_REFUSED' } })
      expect(response).not.toHaveProperty('result')
    }
  })
  it('refuses substitution policy injection and unsupported approximate combinators', async () => {
    for(const input of [{prepare:{},policy:{}},{prepare:{},eligibility:{}},{prepare:{},font_size_policy:{kind:'host-default-size-v1',half_points:22}},{prepare:{},host_font_manifest_path:'/caller/fonts.json'}]){
      const response=await dispatchNativeDocxPagePaintWorkerRequestV1({protocol:DOCX_PAGE_PAINT_WORKER_PROTOCOL,version:1,id:'fonts:malformed',op:'render-font-substitution',input})
      expect(response).toMatchObject({ok:false,error:{code:'COMPILATION_REFUSED'}})
      expect(response).not.toHaveProperty('result')
    }
  })
  it('requires an explicit exact approximate envelope and rejects caller font overrides', async () => {
    for (const input of [{}, { prepare: {}, eligibility: {}, host_font_manifest_path: '/caller/fonts.json' }, { prepare: { font_manifest: {} }, eligibility: {} }]) {
      const response = await dispatchNativeDocxPagePaintWorkerRequestV1({ protocol: DOCX_PAGE_PAINT_WORKER_PROTOCOL, version: 1, id: 'approximate:malformed', op: 'render-approximate', input })
      expect(response).toMatchObject({ ok: false, error: { code: 'COMPILATION_REFUSED' } })
      expect(response).not.toHaveProperty('result')
    }
  })
  it('answers deterministic lifecycle probes in a bounded frame', async () => {
    const request = { protocol: DOCX_PAGE_PAINT_WORKER_PROTOCOL, version: DOCX_PAGE_PAINT_WORKER_VERSION, id: 'probe:1', op: 'ping' }
    const first = await dispatchNativeDocxPagePaintWorkerRequestV1(request)
    const second = await dispatchNativeDocxPagePaintWorkerRequestV1(request)
    expect(second).toEqual(first)
    expect(first).toMatchObject({ id: 'probe:1', ok: true, result: { status: 'ready' } })
    const frame = encodeNativeDocxPagePaintWorkerResponseV1(first)
    expect(frame.readUInt32BE(0)).toBe(frame.byteLength - 4)
  })

  it('refuses malformed envelopes and non-canonical font bytes without partial results', async () => {
    const invalid = await dispatchNativeDocxPagePaintWorkerRequestV1({ id: 'bad request', op: 'ping' })
    expect(invalid).toMatchObject({ id: 'invalid', ok: false, error: { code: 'COMPILATION_REFUSED' } })
    expect(invalid).not.toHaveProperty('result')
    const malformed = await dispatchNativeDocxPagePaintWorkerRequestV1({
      protocol: DOCX_PAGE_PAINT_WORKER_PROTOCOL, version: DOCX_PAGE_PAINT_WORKER_VERSION, id: 'prepare:1', op: 'prepare',
      input: { font_assets: [{ face_id: 'face:test', content_digest: `sha256:${'a'.repeat(64)}`, bytes_base64: 'A===' }], media_assets: [] },
    })
    expect(malformed).toMatchObject({ ok: false, error: { code: 'COMPILATION_REFUSED' } })
    expect(malformed).not.toHaveProperty('result')

    const substitution = await dispatchNativeDocxPagePaintWorkerRequestV1({
      protocol: DOCX_PAGE_PAINT_WORKER_PROTOCOL, version: DOCX_PAGE_PAINT_WORKER_VERSION, id: 'prepare:exploit-87', op: 'prepare',
      input: {
        protocol: 'injoffice.docx.page-paint-compiler', version: 1, source_revision: 'git:test', outline_provider: { provider_id: 'outline:test', provider_revision: 'v1' },
        document: {}, resolved_layout: {}, pagination_settings: {}, font_inventory_json: '{}',
        font_manifest: { version: 1, manifestId: 'caller:unrelated', revision: 'rev:stale', faces: [{ faceId: 'face:test', family: 'Same Family', weight: 400, style: 'normal', stretch: 100, source: { kind: 'document', resourceId: 'font:caller', contentDigest: `sha256:${'a'.repeat(64)}` } }], fallbackChains: [] },
        font_assets: [{ face_id: 'face:test', face_slot: 'embedRegular', resource_id: 'font:caller', content_digest: `sha256:${'a'.repeat(64)}`, collection_index: null, bytes_base64: 'AA==' }],
      },
    })
    expect(substitution).toMatchObject({ id: 'prepare:exploit-87', ok: false, error: { code: 'COMPILATION_REFUSED' } })
    expect(substitution).not.toHaveProperty('result')
  })

  it('rejects non-canonical prepare objects and media counts before touching attacker entries', async () => {
    const base = {
      protocol: 'injoffice.docx.page-paint-compiler', version: 1, source_revision: 'source:1',
      outline_provider: {}, document: {}, resolved_layout: {}, pagination_settings: {}, font_inventory_json: '{}', font_assets: [], media_assets: [],
    }
    const unknown = await dispatchNativeDocxPagePaintWorkerRequestV1({
      protocol: DOCX_PAGE_PAINT_WORKER_PROTOCOL, version: DOCX_PAGE_PAINT_WORKER_VERSION, id: 'prepare:unknown', op: 'prepare', input: { ...base, caller_override: true },
    })
    expect(unknown).toMatchObject({ ok: false, error: { code: 'COMPILATION_REFUSED', message: expect.stringContaining('exact canonical compiler input') } })
    for (const field of ['host_font_manifest_path', 'fonts']) {
      const pathOverride = await dispatchNativeDocxPagePaintWorkerRequestV1({
        protocol: DOCX_PAGE_PAINT_WORKER_PROTOCOL, version: DOCX_PAGE_PAINT_WORKER_VERSION,
        id: 'prepare:path-override', op: 'prepare', input: { ...base, [field]: '/caller/chosen/fonts.json' },
      })
      expect(pathOverride).toMatchObject({ ok: false, error: { code: 'COMPILATION_REFUSED', message: expect.stringContaining('exact canonical compiler input') } })
    }

    let touched = false
    const hostile: Record<string, unknown> = {}
    Object.defineProperty(hostile, 'bytes_base64', { enumerable: true, get() { touched = true; throw new Error('must not materialize') } })
    const overCount = await dispatchNativeDocxPagePaintWorkerRequestV1({
      protocol: DOCX_PAGE_PAINT_WORKER_PROTOCOL, version: DOCX_PAGE_PAINT_WORKER_VERSION, id: 'prepare:bounded', op: 'prepare',
      input: { ...base, media_assets: Array.from({ length: 257 }, () => hostile) },
    })
    expect(overCount).toMatchObject({ ok: false, error: { code: 'COMPILATION_REFUSED', message: expect.stringContaining('exceeds 256 entries') } })
    expect(touched).toBe(false)
  })
})

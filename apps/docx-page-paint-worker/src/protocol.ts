import {
  DOCX_INLINE_IMAGE_LIMITS,
  completeNativeDocxPagePaintV1,
  prepareNativeDocxPagePaintV1,
  renderNativeDocxApproximatePagePreviewV1,
  renderNativeDocxAutomaticBorderPreviewV1,
  validNativeDocxHostDefaultSizePolicyV1,
  type NativeDocxAuthoritativeFontAssetV1,
  type NativeDocxAuthoritativeMediaAssetV1,
  type NativeDocxPagePaintCompleteInputV1,
  type NativeDocxPagePaintPrepareInputV1,
} from '@injoffice/docs/native-page-paint-compiler'
import { createHarfBuzzTextShaperV1, createHarfBuzzOutlineProviderV1, type HarfBuzzTextShaperV1 } from '@injoffice/font-metrics/harfbuzz'
import {loadHostFonts} from './hostFonts.js'

export const DOCX_PAGE_PAINT_WORKER_PROTOCOL = 'injoffice.docx.page-paint-worker'
export const DOCX_PAGE_PAINT_WORKER_VERSION = 1 as const
export const DOCX_PAGE_PAINT_WORKER_MAX_FRAME_BYTES = 192 * 1024 * 1024
export const DOCX_PAGE_PAINT_WORKER_MAX_OUTPUT_BYTES = 64 * 1024 * 1024

const REQUEST_ID = /^[A-Za-z0-9][A-Za-z0-9._:-]{0,127}$/
const BASE64 = /^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/
let runtimeSourceRevision: string | undefined
let runtimeShaper: HarfBuzzTextShaperV1 | undefined

interface PrepareAssetWire extends Omit<NativeDocxAuthoritativeFontAssetV1, 'bytes'> {
  bytes_base64: string
}

interface PrepareMediaAssetWire extends Omit<NativeDocxAuthoritativeMediaAssetV1, 'bytes'> {
  bytes_base64: string
}

export interface NativeDocxPagePaintWorkerRequestV1 {
  protocol: typeof DOCX_PAGE_PAINT_WORKER_PROTOCOL
  version: typeof DOCX_PAGE_PAINT_WORKER_VERSION
  id: string
  op: 'prepare' | 'complete' | 'render' | 'render-approximate' | 'render-auto-borders' | 'ping'
  input?: unknown
}

export type NativeDocxPagePaintWorkerResponseV1 = {
  protocol: typeof DOCX_PAGE_PAINT_WORKER_PROTOCOL
  version: typeof DOCX_PAGE_PAINT_WORKER_VERSION
  id: string
} & ({ ok: true; result: unknown } | { ok: false; error: { code: string; message: string } })

function record(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
}

function exactKeys(value: Record<string, unknown>, keys: readonly string[]): boolean {
  const allowed = new Set(keys)
  return Object.keys(value).every((key) => allowed.has(key))
}

function exactFieldSet(value: Record<string, unknown>, keys: readonly string[]): boolean {
  return Object.keys(value).length === keys.length && exactKeys(value, keys)
}

function boundedMessage(value: unknown): string {
  try {
    const candidate = value instanceof Error ? value.message : String(value)
    return candidate.replace(/[\u0000\r\n]/g, ' ').slice(0, 2_048) || 'native page-paint worker rejected the request'
  } catch {
    return 'native page-paint worker rejected an unreadable request'
  }
}

function decodeBase64(value: unknown, label = 'font', maxBytes = DOCX_PAGE_PAINT_WORKER_MAX_FRAME_BYTES): Uint8Array {
  if (typeof value !== 'string' || value.length === 0 || value.length > Math.ceil(maxBytes / 3) * 4 || !BASE64.test(value)) throw new TypeError(`${label} bytes must be canonical bounded base64`)
  const bytes = Buffer.from(value, 'base64')
  if (bytes.byteLength > maxBytes || bytes.toString('base64') !== value) throw new TypeError(`${label} bytes use a non-canonical or oversized base64 encoding`)
  return Uint8Array.from(bytes)
}

function prepareInput(value: unknown): NativeDocxPagePaintPrepareInputV1 {
  const fields = ['protocol', 'version', 'source_revision', 'outline_provider', 'document', 'resolved_layout', 'pagination_settings', 'font_inventory_json', 'font_assets', 'media_assets'] as const
  if (!record(value) || !exactFieldSet(value, fields) || !Array.isArray(value.font_assets) || !Array.isArray(value.media_assets)) throw new TypeError('prepare input must be the exact canonical compiler input with authoritative font and media asset arrays')
  if (value.media_assets.length > DOCX_INLINE_IMAGE_LIMITS.maxAssets) throw new RangeError(`media asset wire inventory exceeds ${DOCX_INLINE_IMAGE_LIMITS.maxAssets} entries`)
  const assets = value.font_assets.map((asset): NativeDocxAuthoritativeFontAssetV1 => {
    if (!record(asset) || Object.keys(asset).sort().join(',') !== 'bytes_base64,collection_index,content_digest,face_id,face_slot,resource_id') throw new TypeError('font asset wire object contains unknown, duplicate, or incomplete fields')
    return {
      face_id: asset.face_id as string,
      face_slot: asset.face_slot as string,
      resource_id: asset.resource_id as string,
      content_digest: asset.content_digest as `sha256:${string}`,
      collection_index: asset.collection_index as number | null,
      bytes: decodeBase64(asset.bytes_base64),
    }
  })
  let mediaBytes = 0
  const media = value.media_assets.map((asset): NativeDocxAuthoritativeMediaAssetV1 => {
    if (!record(asset) || !exactFieldSet(asset, ['part_name', 'content_type', 'content_digest', 'bytes_base64'])) throw new TypeError('media asset wire object contains unknown or missing fields')
    const bytes = decodeBase64(asset.bytes_base64, 'media', DOCX_INLINE_IMAGE_LIMITS.maxAssetBytes)
    mediaBytes += bytes.byteLength
    if (!Number.isSafeInteger(mediaBytes) || mediaBytes > DOCX_INLINE_IMAGE_LIMITS.maxTotalBytes) throw new RangeError('media asset wire bytes exceed the cumulative page-paint budget')
    return {
      part_name: asset.part_name as string,
      content_type: asset.content_type as string,
      content_digest: asset.content_digest as `sha256:${string}`,
      bytes,
    }
  })
  return {
    protocol: value.protocol as NativeDocxPagePaintPrepareInputV1['protocol'],
    version: value.version as NativeDocxPagePaintPrepareInputV1['version'],
    source_revision: value.source_revision as string,
    outline_provider: value.outline_provider as NativeDocxPagePaintPrepareInputV1['outline_provider'],
    document: value.document,
    resolved_layout: value.resolved_layout,
    pagination_settings: value.pagination_settings,
    font_inventory_json: value.font_inventory_json as string,
    font_assets: assets,
    media_assets: media,
  }
}

function workerShaper(sourceRevision: string): HarfBuzzTextShaperV1 {
  if (runtimeSourceRevision !== undefined && runtimeSourceRevision !== sourceRevision) throw new TypeError('worker source revision is immutable for one process generation')
  runtimeSourceRevision ??= sourceRevision
  runtimeShaper ??= createHarfBuzzTextShaperV1({ sourceRevision })
  return runtimeShaper
}

export async function dispatchNativeDocxPagePaintWorkerRequestV1(value: unknown, hostFontManifestPath?: string): Promise<NativeDocxPagePaintWorkerResponseV1> {
  const candidateID = record(value) && typeof value.id === 'string' && REQUEST_ID.test(value.id) ? value.id : 'invalid'
  const base = { protocol: DOCX_PAGE_PAINT_WORKER_PROTOCOL, version: DOCX_PAGE_PAINT_WORKER_VERSION, id: candidateID } as const
  try {
    if (!record(value) || !exactKeys(value, ['protocol', 'version', 'id', 'op', 'input']) || value.protocol !== DOCX_PAGE_PAINT_WORKER_PROTOCOL || value.version !== DOCX_PAGE_PAINT_WORKER_VERSION || candidateID === 'invalid') throw new TypeError('worker request envelope is invalid')
    if (value.op === 'ping') return { ...base, ok: true, result: { status: 'ready' } }
    if (value.op === 'render-approximate' || value.op === 'render-auto-borders') {
      const automatic = value.op === 'render-auto-borders'
      if (!record(value.input)) throw new TypeError('approximate render requires an input object')
      const fields = automatic ? ('legacy_eligibility' in value.input ? ['prepare', 'legacy_eligibility'] : ['prepare']) : ['prepare', 'eligibility']
      if ('font_size_policy' in value.input) fields.push('font_size_policy')
      if (!exactFieldSet(value.input, fields)) throw new TypeError('approximate render requires exact prepare and eligibility fields')
      const fontSizePolicy = value.input.font_size_policy
      if (fontSizePolicy !== undefined && !validNativeDocxHostDefaultSizePolicyV1(fontSizePolicy)) throw new TypeError('Host default size policy is invalid')
      const input = prepareInput(value.input.prepare)
      if (input.outline_provider.provider_id !== 'injoffice.harfbuzz-outline' || input.outline_provider.provider_revision !== 'v1') throw new TypeError('approximate render requires the pinned outline provider')
      const fonts = hostFontManifestPath ? await loadHostFonts(input, hostFontManifestPath) : undefined
      const providers = new Map<string, ReturnType<typeof createHarfBuzzOutlineProviderV1>>()
      const outlineProvider: Parameters<typeof renderNativeDocxAutomaticBorderPreviewV1>[1] = {
        providerId: 'injoffice.harfbuzz-outline', providerRevision: 'v1',
        getGlyphOutline(request) {
          const asset = input.font_assets.find(asset => asset.face_id === request.face.face_id && asset.content_digest === request.face.content_digest && (asset.collection_index ?? undefined) === request.face.collection_index)
          const host = fonts?.resources.get(request.face.face_id)
          const resource = asset ?? (host && host.face.contentDigest === request.face.content_digest && host.face.collectionIndex === request.face.collection_index ? { face_id: host.face.faceId, bytes: host.bytes, content_digest: request.face.content_digest, collection_index: host.face.collectionIndex ?? null } : undefined)
          if (!resource) throw new TypeError('approximate outline face does not exact-join authoritative bytes')
          let provider = providers.get(resource.face_id)
          if (!provider) {
            provider = createHarfBuzzOutlineProviderV1({ bytes: resource.bytes, contentDigest: resource.content_digest, ...(resource.collection_index === null ? {} : { collectionIndex: resource.collection_index }) })
            providers.set(resource.face_id, provider)
          }
          const outline = provider.outline(request.glyph_id)
          return outline.path.length ? { status: 'outlined' as const, ...request, ...outline } : { status: 'empty' as const, ...request, units_per_em: outline.units_per_em }
        },
      }
      const runtime = { createShaper: workerShaper, fonts, fontSizePolicy }
      const result = automatic
        ? await renderNativeDocxAutomaticBorderPreviewV1(input, outlineProvider, runtime, value.input.legacy_eligibility)
        : await renderNativeDocxApproximatePagePreviewV1(input, value.input.eligibility, outlineProvider, runtime)
      return { ...base, ok: true, result }
    }
    if (value.op === 'prepare') {
      const input = prepareInput(value.input)
      const fonts = hostFontManifestPath ? await loadHostFonts(input, hostFontManifestPath) : undefined
      return { ...base, ok: true, result: await prepareNativeDocxPagePaintV1(input, { createShaper: workerShaper, fonts }) }
    }
    if (value.op === 'render') {
      const input = prepareInput(value.input)
      if (input.outline_provider.provider_id !== 'injoffice.harfbuzz-outline' || input.outline_provider.provider_revision !== 'v1') throw new TypeError('render requires the pinned outline provider')
      const fonts = hostFontManifestPath ? await loadHostFonts(input, hostFontManifestPath) : undefined
      const prepared = await prepareNativeDocxPagePaintV1(input, { createShaper: workerShaper, fonts })
      const providers = new Map<string, ReturnType<typeof createHarfBuzzOutlineProviderV1>>()
      const results = prepared.outline_requests.map((request) => {
        const asset = input.font_assets.find((asset) => asset.face_id === request.face.face_id && asset.content_digest === request.face.content_digest && (asset.collection_index ?? undefined) === request.face.collection_index)
        const host = fonts?.resources.get(request.face.face_id)
        const resource = asset ?? (host && host.face.contentDigest === request.face.content_digest && host.face.collectionIndex === request.face.collection_index ? {face_id:host.face.faceId,bytes:host.bytes,content_digest:request.face.content_digest,collection_index:host.face.collectionIndex??null} : undefined)
        if (!resource) throw new TypeError('outline face does not exact-join an authoritative font')
        let provider = providers.get(resource.face_id)
        if (!provider) {
          provider = createHarfBuzzOutlineProviderV1({ bytes: resource.bytes, contentDigest: resource.content_digest, ...(resource.collection_index === null ? {} : { collectionIndex: resource.collection_index }) })
          providers.set(resource.face_id, provider)
        }
        const outline = provider.outline(request.glyph_id)
        return outline.path.length ? { status: 'outlined' as const, ...request, ...outline } : { status: 'empty' as const, ...request, units_per_em: outline.units_per_em }
      })
      const completed = await completeNativeDocxPagePaintV1({ prepared, outline_results: results })
      return { ...base, ok: true, result: { page_paint_output: completed.page_paint_output, canonical_output_sha256: completed.canonical_output_sha256, canonical_output_validated: completed.canonical_output_validated } }
    }
    if (value.op === 'complete') return { ...base, ok: true, result: await completeNativeDocxPagePaintV1(value.input as NativeDocxPagePaintCompleteInputV1) }
    throw new TypeError('worker operation is unsupported')
  } catch (error) {
    return { ...base, ok: false, error: { code: 'COMPILATION_REFUSED', message: boundedMessage(error) } }
  }
}

export function encodeNativeDocxPagePaintWorkerResponseV1(response: NativeDocxPagePaintWorkerResponseV1): Buffer {
  const payload = Buffer.from(JSON.stringify(response), 'utf8')
  if (payload.byteLength === 0 || payload.byteLength > DOCX_PAGE_PAINT_WORKER_MAX_OUTPUT_BYTES) throw new RangeError('worker response exceeds the bounded output frame')
  const frame = Buffer.allocUnsafe(4 + payload.byteLength)
  frame.writeUInt32BE(payload.byteLength, 0)
  payload.copy(frame, 4)
  return frame
}

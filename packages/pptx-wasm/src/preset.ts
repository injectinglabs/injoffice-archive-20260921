import { validateNativePptx, type NativeEvaluatedGeometry } from '@injoffice/pptx-native'

export interface PptxPresetGeometryRequest {
  name: string
  widthEmu: number
  heightEmu: number
  adjustments?: Readonly<Record<string, number>>
}

function object(value: unknown, allowed?: readonly string[]): Record<string, unknown> {
  if (!value || typeof value !== 'object' || ![Object.prototype, null].includes(Object.getPrototypeOf(value))) throw new TypeError('Preset request must contain plain objects.')
  const result: Record<string, unknown> = Object.create(null)
  const keys = Reflect.ownKeys(value)
  if (keys.length > 1024) throw new RangeError('Preset adjustment budget exceeded.')
  for (const key of keys) {
    if (typeof key !== 'string' || (allowed && !allowed.includes(key))) throw new TypeError('Unknown preset request field.')
    const descriptor = Object.getOwnPropertyDescriptor(value, key)!
    if (!('value' in descriptor) || !descriptor.enumerable) throw new TypeError('Preset request accessors are unsupported.')
    result[key] = descriptor.value
  }
  return result
}

export function snapshotPresetRequest(input: PptxPresetGeometryRequest): string {
  const request = object(input, ['name', 'widthEmu', 'heightEmu', 'adjustments'])
  if (typeof request.name !== 'string' || !/^[A-Za-z][A-Za-z0-9]{0,63}$/.test(request.name)) throw new TypeError('Invalid preset name.')
  for (const key of ['widthEmu', 'heightEmu']) if (!Number.isSafeInteger(request[key]) || (request[key] as number) <= 0) throw new RangeError('Preset dimensions must be positive safe integers.')
  const adjustments = request.adjustments === undefined ? Object.create(null) as Record<string, unknown> : object(request.adjustments)
  for (const [key, value] of Object.entries(adjustments)) {
    if (!/^[A-Za-z_][A-Za-z0-9_]{0,127}$/.test(key) || !Number.isSafeInteger(value) || Object.is(value, -0)) throw new TypeError('Invalid preset adjustment.')
  }
  const payload = JSON.stringify({ name: request.name, widthEmu: request.widthEmu, heightEmu: request.heightEmu, adjustments })
  if (new TextEncoder().encode(payload).byteLength > 65536) throw new RangeError('Preset evaluation request budget exceeded.')
  return payload
}

export function decodePresetGeometry(json: string, payload: string): NativeEvaluatedGeometry {
  if (json.length > 2 * 1024 * 1024) throw new RangeError('Preset evaluation response budget exceeded.')
  const result = object(JSON.parse(json), ['protocol', 'request', 'geometry'])
  if (result.protocol !== 'pptx-preset-evaluation-v1') throw new TypeError('Invalid preset evaluation protocol.')
  const request = JSON.parse(payload) as PptxPresetGeometryRequest
  // Go sorts map keys; compare canonical requests independently of key order.
  const responseRequest = JSON.parse(snapshotPresetRequest(result.request as PptxPresetGeometryRequest)) as PptxPresetGeometryRequest
  if (request.name !== responseRequest.name || request.widthEmu !== responseRequest.widthEmu || request.heightEmu !== responseRequest.heightEmu || Object.keys(request.adjustments!).length !== Object.keys(responseRequest.adjustments!).length || Object.entries(request.adjustments!).some(([key, value]) => responseRequest.adjustments![key] !== value)) throw new TypeError('Preset evaluation response does not match its request.')
  const compatibility = { status: 'preserveOnly', diagnostics: [{ severity: 'warning', code: 'pptx.preset-catalog-preview', message: 'Evaluated geometry is preview-only until a serializer is qualified.' }] }
  const deck = { contractVersion: 'pptx-native/v1', documentId: 'preset-evaluation', origin: 'authored', size: { cx: request.widthEmu, cy: request.heightEmu }, assets: [], compatibility,
    slides: [{ id: 'slide', provenance: 'authored', passthrough: [], compatibility, elements: [{ kind: 'shape', id: 'shape', provenance: 'authored', transform: { x: 0, y: 0, cx: request.widthEmu, cy: request.heightEmu }, paragraphs: [], passthrough: [], compatibility, geometry: result.geometry }] }] }
  const validated = validateNativePptx(deck)
  if (!validated.ok) throw new TypeError(`Invalid preset geometry: ${validated.issues.map(issue => issue.message).join('; ')}`)
  return result.geometry as NativeEvaluatedGeometry
}

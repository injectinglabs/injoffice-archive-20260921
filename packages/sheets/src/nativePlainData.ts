export class NativePlainDataError extends TypeError {
  readonly path: string

  constructor(path: string, message: string) {
    super(message)
    this.name = 'NativePlainDataError'
    this.path = path
  }
}

export interface NativePlainDataLimits {
  readonly maxDepth: number
  readonly maxNodes: number
}

/**
 * Takes one descriptor-only snapshot of JSON-shaped input. Accessors are never
 * invoked and the source is never revisited after its descriptor has been
 * captured. Reflection is Proxy-observable, so every reflective operation is
 * caught and converted to one bounded structured error; no clone API is treated
 * as a callback-free Proxy oracle.
 */
export function snapshotNativePlainData(input: unknown, limits: NativePlainDataLimits): unknown {
  let nodes = 0
  const active = new Set<object>()

  const snapshot = (value: unknown, path: string, depth: number): unknown => {
    nodes++
    if (nodes > limits.maxNodes) throw new NativePlainDataError(path, `plain-data snapshot exceeds ${limits.maxNodes} nodes`)
    if (depth > limits.maxDepth) throw new NativePlainDataError(path, `plain-data snapshot exceeds depth ${limits.maxDepth}`)
    if (value === null || typeof value === 'string' || typeof value === 'boolean' || typeof value === 'number') return value
    if (typeof value !== 'object') throw new NativePlainDataError(path, 'value is not canonical JSON plain data')
    if (active.has(value)) throw new NativePlainDataError(path, 'cyclic input is not canonical JSON plain data')
    let prototype: object | null
    let array: boolean
    let keys: ReadonlyArray<string | symbol>
    let descriptors: PropertyDescriptorMap
    try { array = Array.isArray(value) }
    catch { throw new NativePlainDataError(path, 'array classification trap failed') }
    try { prototype = Object.getPrototypeOf(value) }
    catch { throw new NativePlainDataError(path, 'prototype reflection trap failed') }
    if (array ? prototype !== Array.prototype : prototype !== Object.prototype && prototype !== null) {
      throw new NativePlainDataError(path, 'value has a non-plain prototype')
    }
    try { keys = Array.from(Reflect.ownKeys(value)) }
    catch { throw new NativePlainDataError(path, 'own-key reflection trap failed') }
    if (keys.some((key) => typeof key !== 'string')) throw new NativePlainDataError(path, 'symbol fields are not canonical JSON plain data')
    try { descriptors = Object.getOwnPropertyDescriptors(value) }
    catch { throw new NativePlainDataError(path, 'property-descriptor reflection trap failed') }
    active.add(value)
    try {
      if (array) {
        const lengthDescriptor = descriptors.length
        if (!lengthDescriptor || !('value' in lengthDescriptor) || !Number.isSafeInteger(lengthDescriptor.value) || Object.is(lengthDescriptor.value, -0)) {
          throw new NativePlainDataError(path, 'array length is not canonical')
        }
        const length = lengthDescriptor.value as number
        const actual = keys.filter((key): key is string => typeof key === 'string' && key !== 'length')
        if (actual.length !== length) throw new NativePlainDataError(path, 'sparse or decorated arrays are not canonical JSON plain data')
        const result: unknown[] = new Array(length)
        for (let index = 0; index < length; index++) {
          const name = String(index)
          const descriptor = descriptors[name]
          if (!descriptor || !('value' in descriptor) || !descriptor.enumerable) throw new NativePlainDataError(`${path}/${name}`, 'array item must be an enumerable data property')
          result[index] = snapshot(descriptor.value, `${path}/${name}`, depth + 1)
        }
        return result
      }
      const result: Record<string, unknown> = Object.create(null) as Record<string, unknown>
      for (const key of keys as string[]) {
        const descriptor = descriptors[key]
        if (!descriptor || !('value' in descriptor) || !descriptor.enumerable) throw new NativePlainDataError(`${path}/${pointerToken(key)}`, 'object field must be an enumerable data property')
        const child = snapshot(descriptor.value, `${path}/${pointerToken(key)}`, depth + 1)
        try {
          Object.defineProperty(result, key, { value: child, enumerable: true, configurable: true, writable: true })
        } catch {
          throw new NativePlainDataError(`${path}/${pointerToken(key)}`, 'plain-data property could not be defined')
        }
      }
      return result
    } finally {
      active.delete(value)
    }
  }
  const result = snapshot(input, '', 0)
  // A clone rejection is only an additional fail-closed signal for exotic
  // inputs; it is not used as the snapshot and is not described as a
  // callback-free or complete Proxy detector.
  if (typeof structuredClone === 'function') {
    try { structuredClone(input) }
    catch { throw new NativePlainDataError('', 'input is not cloneable canonical JSON plain data') }
  }
  return result
}

function pointerToken(value: string): string { return value.replace(/~/g, '~0').replace(/\//g, '~1') }

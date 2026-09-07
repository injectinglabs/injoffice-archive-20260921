import { AgentToolsError } from './errors'
import type { JsonValue } from './types'

const encoder = new TextEncoder()
const MAX_JSON_DEPTH = 100

export function cloneJson<T>(value: T, label: string): T {
  const seen = new Set<object>()
  const visit = (current: unknown, path: string, depth = 0): JsonValue => {
    if (depth > MAX_JSON_DEPTH) invalid(label, path, `nesting may not exceed ${MAX_JSON_DEPTH} levels`)
    if (current === null || typeof current === 'string' || typeof current === 'boolean') return current
    if (typeof current === 'number') {
      if (!Number.isFinite(current)) invalid(label, path, 'numbers must be finite')
      return Object.is(current, -0) ? 0 : current
    }
    if (Array.isArray(current)) {
      if (seen.has(current)) invalid(label, path, 'cyclic values are not supported')
      const keys = Reflect.ownKeys(current)
      if (keys.some((key) => typeof key === 'symbol')) invalid(label, path, 'symbol properties are not valid JSON')
      const stringKeys = keys.filter((key): key is string => typeof key === 'string' && key !== 'length')
      if (stringKeys.length !== current.length) invalid(label, path, 'arrays must be dense and may not have extra properties')
      seen.add(current)
      const result: JsonValue[] = []
      for (let index = 0; index < current.length; index += 1) {
        const descriptor = Object.getOwnPropertyDescriptor(current, String(index))
        if (!descriptor || !descriptor.enumerable || !('value' in descriptor)) invalid(label, `${path}/${index}`, 'array entries must be enumerable data properties')
        if (descriptor.value === undefined) invalid(label, `${path}/${index}`, 'undefined is not valid JSON')
        result.push(visit(descriptor.value, `${path}/${index}`, depth + 1))
      }
      seen.delete(current)
      return result
    }
    if (typeof current === 'object' && current !== null) {
      const prototype = Object.getPrototypeOf(current)
      if (prototype !== Object.prototype && prototype !== null) invalid(label, path, 'values must use plain JSON objects')
      if (seen.has(current)) invalid(label, path, 'cyclic values are not supported')
      const ownKeys = Reflect.ownKeys(current)
      if (ownKeys.some((key) => typeof key === 'symbol')) invalid(label, path, 'symbol properties are not valid JSON')
      if (ownKeys.some((key) => typeof key === 'string' && !Object.prototype.propertyIsEnumerable.call(current, key))) invalid(label, path, 'object properties must be enumerable')
      seen.add(current)
      const result: Record<string, JsonValue> = {}
      for (const key of Object.keys(current as object).sort()) {
        const descriptor = Object.getOwnPropertyDescriptor(current, key)
        if (!descriptor || !('value' in descriptor)) invalid(label, `${path}/${escapePointer(key)}`, 'accessor properties are not valid JSON data')
        const entry = descriptor.value
        if (entry === undefined) invalid(label, `${path}/${escapePointer(key)}`, 'undefined is not valid JSON')
        Object.defineProperty(result, key, {
          configurable: true,
          enumerable: true,
          writable: true,
          value: visit(entry, `${path}/${escapePointer(key)}`, depth + 1),
        })
      }
      seen.delete(current)
      return result
    }
    invalid(label, path, `${typeof current} is not valid JSON`)
  }
  return visit(value, '') as T
}

export function freezeJson<T>(value: T): Readonly<T> {
  if (value && typeof value === 'object' && !Object.isFrozen(value)) {
    for (const entry of Object.values(value as object)) freezeJson(entry)
    Object.freeze(value)
  }
  return value
}

export function immutableJson<T>(value: T, label: string): Readonly<T> {
  return freezeJson(cloneJson(value, label))
}

export function canonicalJson(value: unknown): string {
  return JSON.stringify(cloneJson(value, 'value'))
}

export function jsonBytes(value: unknown): number {
  return encoder.encode(canonicalJson(value)).byteLength
}

/** Stable non-cryptographic content identifier. Never use as a security digest. */
export function stableId(value: unknown): string {
  const bytes = encoder.encode(canonicalJson(value))
  let hash = 0xcbf29ce484222325n
  for (const byte of bytes) {
    hash ^= BigInt(byte)
    hash = BigInt.asUintN(64, hash * 0x100000001b3n)
  }
  return hash.toString(16).padStart(16, '0')
}

function escapePointer(value: string): string { return value.replace(/~/g, '~0').replace(/\//g, '~1') }

function invalid(label: string, path: string, reason: string): never {
  throw new AgentToolsError('INVALID_ARGUMENT', `${label}${path || '/'}: ${reason}`, false, undefined, { path })
}

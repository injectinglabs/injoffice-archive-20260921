import { XLSX_NATIVE_V2_RESOURCE_LIMITS, XLSX_NATIVE_V2_SCHEMA } from './nativeContractV2.generated.js'
import type { NativeWorkbookV2 } from './nativeContractV2.generated.js'
import { NativePlainDataError, snapshotNativePlainData } from './nativePlainData.js'

export interface NativeWorkbookV2ValidationIssue { readonly code: string; readonly path: string; readonly message: string }
export type ValidateNativeWorkbookV2Result =
  | { readonly ok: true; readonly value: NativeWorkbookV2 }
  | { readonly ok: false; readonly issues: ReadonlyArray<NativeWorkbookV2ValidationIssue> }

type Schema = Record<string, unknown>
type Budget = { nodes: number; exhausted: boolean }

export function validateNativeWorkbookShape(input: unknown): ValidateNativeWorkbookV2Result {
  const issues: NativeWorkbookV2ValidationIssue[] = []
  let snapshot: unknown
  try {
    snapshot = snapshotNativePlainData(input, { maxDepth: XLSX_NATIVE_V2_RESOURCE_LIMITS.maxJsonDepth, maxNodes: XLSX_NATIVE_V2_RESOURCE_LIMITS.maxJsonTokens })
  } catch (error) {
    if (error instanceof NativePlainDataError) return { ok: false, issues: [{ code: 'INVALID_PLAIN_DATA', path: error.path || '/', message: error.message }] }
    throw error
  }
  validate(snapshot, XLSX_NATIVE_V2_SCHEMA as unknown as Schema, '', issues, 0, { nodes: 0, exhausted: false })
  return issues.length === 0 ? { ok: true, value: snapshot as NativeWorkbookV2 } : { ok: false, issues }
}

function validate(value: unknown, schema: Schema, path: string, issues: NativeWorkbookV2ValidationIssue[], depth: number, budget: Budget): void {
  if (budget.exhausted || issues.length >= XLSX_NATIVE_V2_RESOURCE_LIMITS.maxIssues) return
  budget.nodes++
  if (budget.nodes > XLSX_NATIVE_V2_RESOURCE_LIMITS.maxJsonTokens) {
    budget.exhausted = true
    return issue(issues, 'LIMIT_EXCEEDED', path, `contract exceeds ${XLSX_NATIVE_V2_RESOURCE_LIMITS.maxJsonTokens} JSON nodes`)
  }
  if (depth > XLSX_NATIVE_V2_RESOURCE_LIMITS.maxJsonDepth) return issue(issues, 'LIMIT_EXCEEDED', path, `contract nesting exceeds ${XLSX_NATIVE_V2_RESOURCE_LIMITS.maxJsonDepth}`)
  if (typeof schema.$ref === 'string') {
    const prefix = '#/$defs/'
    const target = schema.$ref.startsWith(prefix) ? (XLSX_NATIVE_V2_SCHEMA.$defs as unknown as Record<string, Schema>)[schema.$ref.slice(prefix.length)] : undefined
    return target ? validate(value, target, path, issues, depth, budget) : issue(issues, 'SCHEMA_ERROR', path, `unsupported schema reference ${schema.$ref}`)
  }
  if ('const' in schema && value !== schema.const) return issue(issues, 'INVALID_VALUE', path, `must equal ${JSON.stringify(schema.const)}`)
  if (Array.isArray(schema.enum) && !schema.enum.includes(value)) return issue(issues, 'INVALID_VALUE', path, `must be one of ${schema.enum.join(', ')}`)
  switch (schema.type) {
    case 'object': {
      if (!isRecord(value)) return issue(issues, 'INVALID_TYPE', path, 'must be an object')
      const properties = (schema.properties ?? {}) as Record<string, Schema>
      for (const name of (schema.required ?? []) as string[]) if (!Object.prototype.hasOwnProperty.call(value, name)) issue(issues, 'REQUIRED', pointer(path, name), 'is required')
      if (schema.additionalProperties === false) for (const name of Object.keys(value)) if (!Object.prototype.hasOwnProperty.call(properties, name)) issue(issues, 'UNKNOWN_FIELD', pointer(path, name), 'is not allowed')
      for (const [name, child] of Object.entries(properties)) if (Object.prototype.hasOwnProperty.call(value, name)) validate(value[name], child, pointer(path, name), issues, depth + 1, budget)
      return
    }
    case 'array': {
      if (!Array.isArray(value)) return issue(issues, 'INVALID_TYPE', path, 'must be an array')
      if (typeof schema.minItems === 'number' && value.length < schema.minItems) issue(issues, 'REQUIRED', path, `must contain at least ${schema.minItems} item(s)`)
      if (typeof schema.maxItems === 'number' && value.length > schema.maxItems) return issue(issues, 'LIMIT_EXCEEDED', path, `must contain at most ${schema.maxItems} item(s)`)
      if (isRecord(schema.items)) for (let index = 0; index < value.length; index++) validate(value[index], schema.items as Schema, pointer(path, String(index)), issues, depth + 1, budget)
      return
    }
    case 'string': {
      if (typeof value !== 'string') return issue(issues, 'INVALID_TYPE', path, 'must be a string')
      if (!hasPairedSurrogates(value)) return issue(issues, 'INVALID_TEXT', path, 'must contain valid Unicode scalar values')
      const scalarLength = unicodeScalarLength(value)
      if (typeof schema.minLength === 'number' && scalarLength < schema.minLength) issue(issues, 'REQUIRED', path, 'is too short')
      if (typeof schema.maxLength === 'number' && scalarLength > schema.maxLength) return issue(issues, 'LIMIT_EXCEEDED', path, 'is too long')
      if (typeof schema['x-maxUtf16Length'] === 'number' && value.length > schema['x-maxUtf16Length']) return issue(issues, 'LIMIT_EXCEEDED', path, 'exceeds its OOXML UTF-16 code-unit limit')
      if (typeof schema.pattern === 'string' && !new RegExp(schema.pattern).test(value)) issue(issues, 'INVALID_VALUE', path, 'has an invalid format')
      return
    }
    case 'integer':
      if (!Number.isSafeInteger(value) || Object.is(value, -0)) return issue(issues, 'INVALID_TYPE', path, 'must be a safe non-negative-zero integer')
      return validateNumberBounds(value as number, schema, path, issues)
    case 'number':
      if (typeof value !== 'number' || !Number.isFinite(value) || Object.is(value, -0)) return issue(issues, 'INVALID_NUMBER', path, 'must be a finite non-negative-zero number')
      return validateNumberBounds(value, schema, path, issues)
    case 'boolean':
      if (typeof value !== 'boolean') issue(issues, 'INVALID_TYPE', path, 'must be a boolean')
  }
}

function validateNumberBounds(value: number, schema: Schema, path: string, issues: NativeWorkbookV2ValidationIssue[]): void {
  if (typeof schema.minimum === 'number' && value < schema.minimum) issue(issues, 'OUT_OF_RANGE', path, `must be at least ${schema.minimum}`)
  if (typeof schema.maximum === 'number' && value > schema.maximum) issue(issues, 'OUT_OF_RANGE', path, `must be at most ${schema.maximum}`)
  if (typeof schema.exclusiveMinimum === 'number' && value <= schema.exclusiveMinimum) issue(issues, 'OUT_OF_RANGE', path, `must be greater than ${schema.exclusiveMinimum}`)
}

function issue(issues: NativeWorkbookV2ValidationIssue[], code: string, path: string, message: string): void {
  if (issues.length < XLSX_NATIVE_V2_RESOURCE_LIMITS.maxIssues) issues.push({ code, path: path || '/', message })
}

function isRecord(value: unknown): value is Record<string, unknown> { return value !== null && typeof value === 'object' && !Array.isArray(value) }
function pointer(path: string, name: string): string { return `${path}/${name.replace(/~/g, '~0').replace(/\//g, '~1')}` }
function unicodeScalarLength(value: string): number { return [...value].length }
function hasPairedSurrogates(value: string): boolean {
  for (let index = 0; index < value.length; index++) {
    const unit = value.charCodeAt(index)
    if (unit >= 0xd800 && unit <= 0xdbff) {
      const low = value.charCodeAt(++index)
      if (!(low >= 0xdc00 && low <= 0xdfff)) return false
    } else if (unit >= 0xdc00 && unit <= 0xdfff) return false
  }
  return true
}

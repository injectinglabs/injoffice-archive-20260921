import { DOCX_NATIVE_LIMITS, type NativeDocxIssueCode, type NativeDocxValidationIssue } from './nativeContract.js'

export const DOCX_PAGINATION_SETTINGS_PROTOCOL = 'injoffice.docx.pagination-settings'
export const DOCX_PAGINATION_SETTINGS_VERSION = 1 as const
export const DOCX_DEFAULT_TAB_STOP_TWIPS = 720
export const DOCX_PAGINATION_SETTINGS_DIAGNOSTIC_CODES = Object.freeze([
  'UNKNOWN_SETTINGS_ELEMENT',
  'DUPLICATE_SETTINGS_PROPERTY',
  'INVALID_DEFAULT_TAB_STOP',
  'INVALID_SETTINGS_ON_OFF',
  'MIRROR_MARGINS_UNSUPPORTED',
  'GUTTER_AT_TOP_UNSUPPORTED',
  'CHARACTER_SPACING_CONTROL_UNSUPPORTED',
  'COMPATIBILITY_SETTING_UNSUPPORTED',
  'PAGINATION_SETTING_UNSUPPORTED',
  'INVALID_SETTINGS_STRUCTURE',
] as const)

export type NativeDocxPaginationSettingsDiagnosticCode = typeof DOCX_PAGINATION_SETTINGS_DIAGNOSTIC_CODES[number]

export interface NativeDocxPaginationSettingsDiagnosticV1 {
  code: NativeDocxPaginationSettingsDiagnosticCode
  severity: 'unsupported'
  part_name: string
  path: string
  preservation: 'preserve-verbatim'
  message: string
}

/** Extractor-owned attestation for settings.xml values that affect shaping or pagination. */
export interface NativeDocxPaginationSettingsV1 {
  protocol: typeof DOCX_PAGINATION_SETTINGS_PROTOCOL
  version: typeof DOCX_PAGINATION_SETTINGS_VERSION
  document_id: string
  revision: string
  package_sha256: string
  main_part: string
  relationships_part?: string
  relationships_sha256?: string
  relationship_id?: string
  settings_part?: string
  settings_sha256?: string
  profile: 'absent-default' | 'word-modern-default' | 'unsupported'
  default_tab_stop_twips: number
  mirror_margins: boolean
  gutter_at_top: boolean
  even_and_odd_headers: boolean
  compatibility_mode?: 15
  diagnostics: NativeDocxPaginationSettingsDiagnosticV1[]
}

export type DecodeNativeDocxPaginationSettingsResult =
  | { ok: true; value: NativeDocxPaginationSettingsV1 }
  | { ok: false; issues: NativeDocxValidationIssue[] }

export const DOCX_PAGINATION_SETTINGS_V1_BINDING_FIELDS = {
  DiagnosticV1: ['code', 'severity', 'part_name', 'path', 'preservation', 'message'],
  SettingsV1: ['protocol', 'version', 'document_id', 'revision', 'package_sha256', 'main_part', 'relationships_part', 'relationships_sha256', 'relationship_id', 'settings_part', 'settings_sha256', 'profile', 'default_tab_stop_twips', 'mirror_margins', 'gutter_at_top', 'even_and_odd_headers', 'compatibility_mode', 'diagnostics'],
} as const
const ROOT_FIELDS = DOCX_PAGINATION_SETTINGS_V1_BINDING_FIELDS.SettingsV1
const DIAGNOSTIC_FIELDS = DOCX_PAGINATION_SETTINGS_V1_BINDING_FIELDS.DiagnosticV1
// Exact parity with docxpatch.nativeIDPattern. OPC part names use a separate
// grammar; durable native ids and relationship ids never admit '/'.
const ID = /^[A-Za-z0-9][A-Za-z0-9._:-]{0,255}$/
const SHA256 = /^sha256:[0-9a-f]{64}$/
const PART_SEGMENT = /^(?:[A-Za-z0-9._~!$&'()*+,;=@-]|%[0-9A-F]{2})+$/

function add(issues: NativeDocxValidationIssue[], code: NativeDocxIssueCode, path: string, message: string): void {
  if (issues.length < DOCX_NATIVE_LIMITS.maxIssues) issues.push({ code, path, message })
}

function record(value: unknown, path: string, fields: readonly string[], issues: NativeDocxValidationIssue[]): Record<string, unknown> | undefined {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) {
    add(issues, value === undefined ? 'REQUIRED' : 'INVALID_TYPE', path, 'must be an object')
    return undefined
  }
  const result = value as Record<string, unknown>
  const allowed = new Set(fields)
  for (const key of Object.keys(result).sort()) if (!allowed.has(key)) add(issues, 'UNKNOWN_FIELD', `${path}/${key.replace(/~/g, '~0').replace(/\//g, '~1')}`, `unknown field ${JSON.stringify(key)}`)
  for (const key of fields) if (!['relationships_part', 'relationships_sha256', 'relationship_id', 'settings_part', 'settings_sha256', 'compatibility_mode'].includes(key) && !(key in result)) add(issues, 'REQUIRED', `${path}/${key}`, 'field is required')
  return result
}

function stringValue(value: unknown, path: string, issues: NativeDocxValidationIssue[], pattern = ID, max = 256): string | undefined {
  if (typeof value !== 'string') {
    add(issues, value === undefined ? 'REQUIRED' : 'INVALID_TYPE', path, 'must be a string')
    return undefined
  }
  if (value.length === 0 || value.length > max || /[\u0000\r\n]/.test(value) || !pattern.test(value)) add(issues, 'INVALID_VALUE', path, 'contains an invalid or unbounded string')
  return value
}

function partName(value: unknown, path: string, issues: NativeDocxValidationIssue[]): string | undefined {
  const part = stringValue(value, path, issues, /[^\u0000\r\n]+/, 4096)
  if (!part) return undefined
  const valid = !part.startsWith('/') && !part.includes('\\') && !part.includes('//') && part.split('/').every((segment) => {
    if (!PART_SEGMENT.test(segment) || segment === '.' || segment === '..') return false
    try {
      const decoded = decodeURIComponent(segment)
      return decoded !== '.' && decoded !== '..' && !decoded.endsWith('.') && !/[\\/?#%]/.test(decoded) && ![...decoded].some((character) => {
        const code = character.codePointAt(0) ?? 0
        return code < 0x20 || code === 0x7f
      })
    } catch { return false }
  })
  if (!valid) add(issues, 'INVALID_VALUE', path, 'must be a canonical OPC part name')
  return part
}

function bool(value: unknown, path: string, issues: NativeDocxValidationIssue[]): boolean | undefined {
  if (typeof value !== 'boolean') {
    add(issues, value === undefined ? 'REQUIRED' : 'INVALID_TYPE', path, 'must be a boolean')
    return undefined
  }
  return value
}

function decodeNativeDocxPaginationSettingsUnsafe(value: unknown): DecodeNativeDocxPaginationSettingsResult {
  const issues: NativeDocxValidationIssue[] = []
  const root = record(value, '', ROOT_FIELDS, issues)
  if (!root) return { ok: false, issues }
  if (root.protocol !== DOCX_PAGINATION_SETTINGS_PROTOCOL) add(issues, 'UNSUPPORTED_PROTOCOL', '/protocol', `must equal ${DOCX_PAGINATION_SETTINGS_PROTOCOL}`)
  if (root.version !== DOCX_PAGINATION_SETTINGS_VERSION) add(issues, 'UNSUPPORTED_VERSION', '/version', `must equal ${DOCX_PAGINATION_SETTINGS_VERSION}`)
  const documentID = stringValue(root.document_id, '/document_id', issues)
  const revision = stringValue(root.revision, '/revision', issues)
  const packageSHA = stringValue(root.package_sha256, '/package_sha256', issues, SHA256, 71)
  const mainPart = partName(root.main_part, '/main_part', issues)
  const relationshipsPart = root.relationships_part === undefined ? undefined : partName(root.relationships_part, '/relationships_part', issues)
  const relationshipsSHA = root.relationships_sha256 === undefined ? undefined : stringValue(root.relationships_sha256, '/relationships_sha256', issues, SHA256, 71)
  const relationshipID = root.relationship_id === undefined ? undefined : stringValue(root.relationship_id, '/relationship_id', issues)
  const relationshipClosure = [relationshipsPart, relationshipsSHA, relationshipID].filter((entry) => entry !== undefined).length
  if (relationshipClosure !== 0 && relationshipClosure !== 3) add(issues, 'INVALID_UNION', '/relationships_part', 'relationships part, SHA-256, and relationship id must appear together')
  const settingsPart = root.settings_part === undefined ? undefined : partName(root.settings_part, '/settings_part', issues)
  const settingsSHA = root.settings_sha256 === undefined ? undefined : stringValue(root.settings_sha256, '/settings_sha256', issues, SHA256, 71)
  if ((settingsPart === undefined) !== (settingsSHA === undefined)) add(issues, 'INVALID_UNION', '/settings_part', 'settings part and SHA-256 must appear together')
  const profile = typeof root.profile === 'string' && ['absent-default', 'word-modern-default', 'unsupported'].includes(root.profile) ? root.profile as NativeDocxPaginationSettingsV1['profile'] : undefined
  if (!profile) add(issues, root.profile === undefined ? 'REQUIRED' : 'INVALID_VALUE', '/profile', 'must be absent-default, word-modern-default, or unsupported')
  const tabStop = Number.isSafeInteger(root.default_tab_stop_twips) && !Object.is(root.default_tab_stop_twips, -0) && (root.default_tab_stop_twips as number) > 0 && (root.default_tab_stop_twips as number) <= 1_000_000_000 ? root.default_tab_stop_twips as number : undefined
  if (tabStop === undefined) add(issues, root.default_tab_stop_twips === undefined ? 'REQUIRED' : 'OUT_OF_RANGE', '/default_tab_stop_twips', 'must be a positive bounded safe integer')
  const mirror = bool(root.mirror_margins, '/mirror_margins', issues)
  const gutterAtTop = bool(root.gutter_at_top, '/gutter_at_top', issues)
  const evenOdd = bool(root.even_and_odd_headers, '/even_and_odd_headers', issues)
  const compatibilityMode = root.compatibility_mode === undefined ? undefined : root.compatibility_mode === 15 ? 15 as const : undefined
  if (root.compatibility_mode !== undefined && compatibilityMode === undefined) add(issues, 'INVALID_VALUE', '/compatibility_mode', 'only modern compatibility mode 15 is attested')
  const diagnostics: NativeDocxPaginationSettingsDiagnosticV1[] = []
  if (!Array.isArray(root.diagnostics)) add(issues, root.diagnostics === undefined ? 'REQUIRED' : 'INVALID_TYPE', '/diagnostics', 'must be an array')
  else {
    if (root.diagnostics.length > DOCX_NATIVE_LIMITS.maxIssues) add(issues, 'LIMIT_EXCEEDED', '/diagnostics', `must contain at most ${DOCX_NATIVE_LIMITS.maxIssues} entries`)
    root.diagnostics.slice(0, DOCX_NATIVE_LIMITS.maxIssues).forEach((value, index) => {
      const path = `/diagnostics/${index}`
      const entry = record(value, path, DIAGNOSTIC_FIELDS, issues)
      if (!entry) return
      const code = typeof entry.code === 'string' && (DOCX_PAGINATION_SETTINGS_DIAGNOSTIC_CODES as readonly string[]).includes(entry.code) ? entry.code as NativeDocxPaginationSettingsDiagnosticCode : undefined
      if (!code) add(issues, entry.code === undefined ? 'REQUIRED' : 'INVALID_VALUE', `${path}/code`, 'must be an enumerated pagination-settings diagnostic code')
      if (entry.severity !== 'unsupported') add(issues, entry.severity === undefined ? 'REQUIRED' : 'INVALID_VALUE', `${path}/severity`, 'must equal unsupported')
      const diagnosticPart = partName(entry.part_name, `${path}/part_name`, issues)
      const diagnosticPath = stringValue(entry.path, `${path}/path`, issues, /[^\u0000\r\n]+/, 4096)
      if (entry.preservation !== 'preserve-verbatim') add(issues, entry.preservation === undefined ? 'REQUIRED' : 'INVALID_VALUE', `${path}/preservation`, 'must equal preserve-verbatim')
      const message = stringValue(entry.message, `${path}/message`, issues, /[^\u0000]+/, 4096)
      if (code && diagnosticPart && diagnosticPath && message && entry.severity === 'unsupported' && entry.preservation === 'preserve-verbatim') diagnostics.push({ code, severity: 'unsupported', part_name: diagnosticPart, path: diagnosticPath, preservation: 'preserve-verbatim', message })
    })
  }
  if (profile === 'unsupported' && diagnostics.length === 0) add(issues, 'REQUIRED', '/diagnostics', 'unsupported profile requires at least one diagnostic')
  if (profile && profile !== 'unsupported' && diagnostics.length !== 0) add(issues, 'INVALID_UNION', '/diagnostics', 'supported profiles cannot carry diagnostics')
  if (profile === 'absent-default' && (settingsPart !== undefined || relationshipClosure !== 0 || compatibilityMode !== undefined || tabStop !== DOCX_DEFAULT_TAB_STOP_TWIPS || mirror !== false || gutterAtTop !== false || evenOdd !== false)) add(issues, 'INVALID_UNION', '/profile', 'absent-default must carry the exact Word defaults and no settings relationship closure')
  if (profile === 'word-modern-default' && (!settingsPart || compatibilityMode !== 15)) add(issues, 'REQUIRED', '/settings_part', 'word-modern-default must bind an extracted settings part and explicit compatibility mode 15')
  if (profile === 'unsupported' && !settingsPart) add(issues, 'REQUIRED', '/settings_part', 'unsupported settings must bind the preserved source part that caused refusal')
  if (settingsPart && relationshipClosure !== 3) add(issues, 'REQUIRED', '/relationships_part', 'an extracted settings part must bind its exact owning relationship closure')
  if (profile && profile !== 'unsupported' && (mirror || gutterAtTop)) add(issues, 'INVALID_UNION', '/profile', 'supported settings profiles cannot enable mirror margins or top gutter')
  if (issues.length > 0 || !documentID || !revision || !packageSHA || !mainPart || !profile || tabStop === undefined || mirror === undefined || gutterAtTop === undefined || evenOdd === undefined) return { ok: false, issues }
  return { ok: true, value: {
    protocol: DOCX_PAGINATION_SETTINGS_PROTOCOL, version: DOCX_PAGINATION_SETTINGS_VERSION,
    document_id: documentID, revision, package_sha256: packageSHA, main_part: mainPart,
    ...(relationshipsPart ? { relationships_part: relationshipsPart, relationships_sha256: relationshipsSHA!, relationship_id: relationshipID! } : {}),
    ...(settingsPart ? { settings_part: settingsPart, settings_sha256: settingsSHA! } : {}),
    profile, default_tab_stop_twips: tabStop, mirror_margins: mirror, gutter_at_top: gutterAtTop,
    even_and_odd_headers: evenOdd, ...(compatibilityMode ? { compatibility_mode: compatibilityMode } : {}), diagnostics,
  } }
}

export function decodeNativeDocxPaginationSettings(value: unknown): DecodeNativeDocxPaginationSettingsResult {
  try {
    return decodeNativeDocxPaginationSettingsUnsafe(value)
  } catch {
    return { ok: false, issues: [{ code: 'INVALID_VALUE', path: '', message: 'pagination settings could not be safely inspected' }] }
  }
}

import { asciiLower, hasAsciiEdgeWhitespace, validateFontManifest, type NativeFontManifest } from '@injoffice/font-metrics/layout'
import { isUnicode13WhiteSpace } from '@injoffice/font-metrics/unicode13'
import { sha256 } from '@noble/hashes/sha2.js'
import { bytesToHex } from '@noble/hashes/utils.js'

export const NATIVE_DOCX_FONT_INVENTORY_PROTOCOL = 'injoffice.docx.font-inventory'
export const NATIVE_DOCX_FONT_INVENTORY_VERSION = 1 as const
export const NATIVE_DOCX_FONT_INVENTORY_MAX_JSON_BYTES = 8 * 1024 * 1024

const ID = /^[A-Za-z0-9][A-Za-z0-9._:-]{0,255}$/
const SHA256 = /^sha256:[0-9a-f]{64}$/
const FONT_KEY = /^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$/
const PART_SEGMENT = /^(?:[A-Za-z0-9._~!$&'()*+,;=@-]|%[0-9A-F]{2})+$/
const OBFUSCATED_FONT = 'application/vnd.openxmlformats-officedocument.obfuscatedFont'
const TRANSITIONAL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
const STRICT = 'http://purl.oclc.org/ooxml/officeDocument/relationships/'

export interface NativeDOCXFontInventoryV1 {
  protocol: typeof NATIVE_DOCX_FONT_INVENTORY_PROTOCOL
  version: typeof NATIVE_DOCX_FONT_INVENTORY_VERSION
  document_id: string
  revision: string
  package_sha256: string
  main_part: string
  main_sha256: string
  inventory_sha256: string
  font_table?: NativeDOCXFontTableBindingV1
  families: NativeDOCXFontFamilyV1[]
  references: NativeDOCXFontReferenceV1[]
  native_text_manifest?: NativeFontManifest
  native_text_manifest_sha256?: string
}

export interface NativeDOCXFontTableBindingV1 {
  part_name: string
  sha256: string
  main_relationships_part: string
  main_relationships_sha256: string
  relationship_id: string
  relationship_type: string
  relationship_target: string
  font_relationships_part?: string
  font_relationships_sha256?: string
}

export interface NativeDOCXFontFamilyV1 {
  family_id: string
  name: string
  alt_name?: string
  faces: NativeDOCXFontFaceV1[]
}

export interface NativeDOCXFontReferenceV1 { family: string; weight: number; style: 'normal' | 'italic'; scope_ids: string[] }

export interface NativeDOCXFontFaceV1 {
  face_id: string
  family: string
  alt_name?: string
  weight: 400 | 700
  style: 'normal' | 'italic'
  stretch: 100
  source: NativeDOCXEmbeddedFontSourceV1
}

export interface NativeDOCXEmbeddedFontSourceV1 {
  kind: 'document'
  face_slot: 'embedRegular' | 'embedBold' | 'embedItalic' | 'embedBoldItalic'
  font_table_part: string
  font_table_path: string
  relationships_part: string
  relationships_sha256: string
  relationship_id: string
  relationship_type: string
  relationship_target: string
  asset_part: string
  asset_content_type: string
  stored_byte_length: number
  stored_sha256: string
  content_sha256: string
  resource_id: string
  collection_index?: number
  obfuscation: { algorithm: 'ecma-376-font-obfuscation'; font_key: string; subsetted?: boolean }
  licensing: {
    embedding_origin: 'document-package'
    rights_source: 'sfnt-os2-fstype'
    rights_status: 'verified'
    embedding_rights: 'installable' | 'preview-print' | 'editable'
    no_subsetting: boolean
    allowed_scope: 'document-only'
  }
}

function record(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value) && (Object.getPrototypeOf(value) === Object.prototype || Object.getPrototypeOf(value) === null)
}

function exact(value: unknown, keys: readonly string[], label: string): asserts value is Record<string, unknown> {
  if (!record(value)) throw new TypeError(`${label} must be a plain object`)
  const actual = Object.keys(value).sort()
  const expected = [...keys].sort()
  if (actual.length !== expected.length || actual.some((key, index) => key !== expected[index])) throw new TypeError(`${label} contains unknown, duplicate, or incomplete attestation fields`)
}

function exactOptional(value: unknown, required: readonly string[], optional: readonly string[], label: string): asserts value is Record<string, unknown> {
  if (!record(value)) throw new TypeError(`${label} must be a plain object`)
  const keys = Object.keys(value)
  const allowed = new Set([...required, ...optional])
  if (required.some((key) => !Object.hasOwn(value, key)) || keys.some((key) => !allowed.has(key))) throw new TypeError(`${label} contains unknown or incomplete attestation fields`)
}

function string(value: unknown, label: string, pattern?: RegExp): asserts value is string {
  if (typeof value !== 'string' || value.length === 0 || value.length > 8_192 || pattern && !pattern.test(value)) throw new TypeError(`${label} is invalid`)
}

function boundedText(value: unknown, label: string, max = 256): asserts value is string {
  const first = typeof value === 'string' && value.length > 0 ? value.codePointAt(0) : undefined
  const last = typeof value === 'string' && value.length > 0 ? [...value].at(-1)?.codePointAt(0) : undefined
  if (typeof value !== 'string' || value.length === 0 || new TextEncoder().encode(value).byteLength > max || first === undefined || last === undefined || isUnicode13WhiteSpace(first) || isUnicode13WhiteSpace(last) || /[\u0000\r\n]/.test(value)) throw new TypeError(`${label} is invalid`)
}

function asciiFold(value: string): string { return asciiLower(value) }

function compareGoStrings(left: string, right: string): number {
  const leftBytes = new TextEncoder().encode(left), rightBytes = new TextEncoder().encode(right)
  for (let index = 0; index < Math.min(leftBytes.length, rightBytes.length); index += 1) if (leftBytes[index] !== rightBytes[index]) return leftBytes[index]! - rightBytes[index]!
  return leftBytes.length - rightBytes.length
}

function validPart(value: unknown, label: string): asserts value is string {
  string(value, label)
  if (value.startsWith('/') || value.includes('\\') || value.includes('//') || !value.split('/').every((segment) => {
    if (segment === '.' || segment === '..' || !PART_SEGMENT.test(segment)) return false
    try {
      const decoded = decodeURIComponent(segment)
      return decoded !== '.' && decoded !== '..' && !decoded.endsWith('.') && !/[\\/?#%]/.test(decoded) && !/[\u0000-\u001f\u007f]/.test(decoded)
    } catch { return false }
  })) throw new TypeError(`${label} is not a canonical OPC part name`)
}

function canonicalPart(value: string): string { return asciiLower(value.split('/').map((segment) => decodeURIComponent(segment)).join('/')) }

function relationshipOwner(value: string): string | undefined {
  if (asciiFold(value) === '_rels/.rels') return ''
  const segments = value.split('/')
  if (segments.length < 2 || asciiFold(segments.at(-2)!) !== '_rels' || !asciiFold(segments.at(-1)!).endsWith('.rels')) return undefined
  const ownerName = segments.at(-1)!.slice(0, -5)
  const owner = [...segments.slice(0, -2), ownerName].join('/')
  try { validPart(owner, 'relationship owner'); return owner } catch { return undefined }
}

function resolveTarget(owner: string, target: string): string {
  if (!target || target.startsWith('/') || /[\\?#]/.test(target) || hasAsciiEdgeWhitespace(target)) throw new TypeError('relationship target is unsafe')
  const parts = owner.includes('/') ? owner.split('/').slice(0, -1) : []
  for (const raw of target.split('/')) {
    if (raw === '') throw new TypeError('relationship target contains an empty path segment')
    if (raw === '.') continue
    if (raw === '..') { if (parts.length === 0) throw new TypeError('relationship target escapes the package'); parts.pop(); continue }
    validPart(raw, 'relationship target segment')
    parts.push(raw)
  }
  const result = parts.join('/')
  validPart(result, 'relationship target')
  return result
}

function canonicalValue(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(canonicalValue)
  if (record(value)) return Object.fromEntries(Object.keys(value).sort().map((key) => [key, canonicalValue(value[key])]))
  return value
}

function jsonWithGoEscapes(value: unknown, html: boolean): string {
  let encoded = JSON.stringify(value).replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029')
  if (html) encoded = encoded.replace(/&/g, '\\u0026').replace(/</g, '\\u003c').replace(/>/g, '\\u003e')
  return encoded
}

export function nativeDOCXCanonicalWireSHA256V1(value: unknown): string {
  return `sha256:${bytesToHex(sha256(new TextEncoder().encode(jsonWithGoEscapes(canonicalValue(value), false))))}`
}

function stableToken(value: string): string { return bytesToHex(sha256(new TextEncoder().encode(value))).slice(0, 24) }

function orderedInventory(value: NativeDOCXFontInventoryV1): unknown {
  const orderedSource = (source: NativeDOCXEmbeddedFontSourceV1) => ({
    kind: source.kind, face_slot: source.face_slot, font_table_part: source.font_table_part, font_table_path: source.font_table_path,
    relationships_part: source.relationships_part, relationships_sha256: source.relationships_sha256, relationship_id: source.relationship_id,
    relationship_type: source.relationship_type, relationship_target: source.relationship_target, asset_part: source.asset_part,
    asset_content_type: source.asset_content_type, stored_byte_length: source.stored_byte_length, stored_sha256: source.stored_sha256,
    content_sha256: source.content_sha256, resource_id: source.resource_id,
    ...(source.collection_index !== undefined ? { collection_index: source.collection_index } : {}),
    obfuscation: { algorithm: source.obfuscation.algorithm, font_key: source.obfuscation.font_key, ...(source.obfuscation.subsetted !== undefined ? { subsetted: source.obfuscation.subsetted } : {}) },
    licensing: { embedding_origin: source.licensing.embedding_origin, rights_source: source.licensing.rights_source, rights_status: source.licensing.rights_status, embedding_rights: source.licensing.embedding_rights, no_subsetting: source.licensing.no_subsetting, allowed_scope: source.licensing.allowed_scope },
  })
  const manifest = value.native_text_manifest && {
    version: value.native_text_manifest.version, manifestId: value.native_text_manifest.manifestId, revision: value.native_text_manifest.revision,
    faces: value.native_text_manifest.faces.map((face) => ({
      faceId: face.faceId, family: face.family, ...(face.aliases !== undefined && face.aliases.length > 0 ? { aliases: [...face.aliases] } : {}),
      weight: face.weight, style: face.style, stretch: face.stretch,
      source: { kind: face.source.kind, resourceId: face.source.resourceId, ...(face.source.contentDigest !== undefined ? { contentDigest: face.source.contentDigest } : {}), ...(face.source.collectionIndex !== undefined ? { collectionIndex: face.source.collectionIndex } : {}) },
    })),
    fallbackChains: value.native_text_manifest.fallbackChains.map((chain) => ({ chainId: chain.chainId, faceIds: [...chain.faceIds], ...(chain.scripts !== undefined ? { scripts: [...chain.scripts] } : {}), ...(chain.languages !== undefined ? { languages: [...chain.languages] } : {}) })),
  }
  return {
    protocol: value.protocol, version: value.version, document_id: value.document_id, revision: value.revision,
    package_sha256: value.package_sha256, main_part: value.main_part, main_sha256: value.main_sha256, inventory_sha256: value.inventory_sha256,
    ...(value.font_table ? { font_table: {
      part_name: value.font_table.part_name, sha256: value.font_table.sha256, main_relationships_part: value.font_table.main_relationships_part,
      main_relationships_sha256: value.font_table.main_relationships_sha256, relationship_id: value.font_table.relationship_id,
      relationship_type: value.font_table.relationship_type, relationship_target: value.font_table.relationship_target,
      ...(value.font_table.font_relationships_part !== undefined ? { font_relationships_part: value.font_table.font_relationships_part } : {}),
      ...(value.font_table.font_relationships_sha256 !== undefined ? { font_relationships_sha256: value.font_table.font_relationships_sha256 } : {}),
    } } : {}),
    families: value.families.map((family) => ({ family_id: family.family_id, name: family.name, ...(family.alt_name !== undefined ? { alt_name: family.alt_name } : {}), faces: family.faces.map((face) => ({ face_id: face.face_id, family: face.family, ...(face.alt_name !== undefined ? { alt_name: face.alt_name } : {}), weight: face.weight, style: face.style, stretch: face.stretch, source: orderedSource(face.source) })) })),
    references: value.references.map((reference) => ({ family: reference.family, weight: reference.weight, style: reference.style, scope_ids: [...reference.scope_ids] })),
    ...(manifest ? { native_text_manifest: manifest } : {}),
    ...(value.native_text_manifest_sha256 !== undefined ? { native_text_manifest_sha256: value.native_text_manifest_sha256 } : {}),
  }
}

export function encodeNativeDOCXFontInventoryV1(value: NativeDOCXFontInventoryV1): string {
  return jsonWithGoEscapes(orderedInventory(value), true)
}

function validateInventory(value: unknown): NativeDOCXFontInventoryV1 {
  exactOptional(value, ['protocol', 'version', 'document_id', 'revision', 'package_sha256', 'main_part', 'main_sha256', 'inventory_sha256', 'families', 'references'], ['font_table', 'native_text_manifest', 'native_text_manifest_sha256'], 'font inventory')
  if (value.protocol !== NATIVE_DOCX_FONT_INVENTORY_PROTOCOL || value.version !== NATIVE_DOCX_FONT_INVENTORY_VERSION) throw new TypeError('font inventory protocol/version is unsupported')
  string(value.document_id, 'document_id', ID); string(value.revision, 'revision', ID)
  string(value.package_sha256, 'package_sha256', SHA256); string(value.main_sha256, 'main_sha256', SHA256); string(value.inventory_sha256, 'inventory_sha256', SHA256)
  if (value.revision !== `rev:${value.package_sha256.slice(7, 39)}`) throw new TypeError('font inventory revision does not bind the complete package digest')
  validPart(value.main_part, 'main_part')
  if (!Array.isArray(value.families) || !Array.isArray(value.references) || value.families.length > 10_000 || value.references.length > 10_000) throw new TypeError('font inventory collections are invalid')

  let fontRelationshipType = ''
  if (value.font_table !== undefined) {
    exactOptional(value.font_table, ['part_name', 'sha256', 'main_relationships_part', 'main_relationships_sha256', 'relationship_id', 'relationship_type', 'relationship_target'], ['font_relationships_part', 'font_relationships_sha256'], 'font_table')
    validPart(value.font_table.part_name, 'font_table.part_name'); validPart(value.font_table.main_relationships_part, 'font_table.main_relationships_part')
    if (relationshipOwner(value.font_table.main_relationships_part) !== value.main_part) throw new TypeError('main relationship part does not exactly belong to the main part')
    string(value.font_table.sha256, 'font_table.sha256', SHA256); string(value.font_table.main_relationships_sha256, 'font_table.main_relationships_sha256', SHA256); string(value.font_table.relationship_id, 'font_table.relationship_id', ID); string(value.font_table.relationship_target, 'font_table.relationship_target')
    if (value.font_table.relationship_type === `${TRANSITIONAL}fontTable`) fontRelationshipType = `${TRANSITIONAL}font`
    else if (value.font_table.relationship_type === `${STRICT}fontTable`) fontRelationshipType = `${STRICT}font`
    else throw new TypeError('font-table relationship type is invalid')
    if ((value.font_table.font_relationships_part === undefined) !== (value.font_table.font_relationships_sha256 === undefined)) throw new TypeError('font relationship-part binding is incomplete')
    if (value.font_table.font_relationships_part !== undefined) { validPart(value.font_table.font_relationships_part, 'font_table.font_relationships_part'); string(value.font_table.font_relationships_sha256, 'font_table.font_relationships_sha256', SHA256); if (relationshipOwner(value.font_table.font_relationships_part) !== value.font_table.part_name) throw new TypeError('font relationship part does not exactly belong to the font table') }
    if (resolveTarget(value.main_part, value.font_table.relationship_target) !== value.font_table.part_name) throw new TypeError('font-table relationship target is not exact')
  } else if (value.families.length > 0) throw new TypeError('font families require a font-table binding')

  const faces = new Map<string, NativeDOCXFontFaceV1>()
  const familyIDs = new Set<string>(), familyNames = new Set<string>(), resources = new Set<string>(), assetParts = new Set<string>(), relationshipIDs = new Set<string>()
  let priorFamily = ''
  for (const [familyIndex, familyValue] of value.families.entries()) {
    exactOptional(familyValue, ['family_id', 'name', 'faces'], ['alt_name'], `families[${familyIndex}]`)
    string(familyValue.family_id, 'family_id', ID); boundedText(familyValue.name, 'family.name')
    const familyKey = asciiFold(familyValue.name)
    if (familyIDs.has(familyValue.family_id) || familyNames.has(familyKey) || familyIndex > 0 && compareGoStrings(priorFamily, familyKey) >= 0) throw new TypeError('font families are duplicate or not canonically ordered')
    familyIDs.add(familyValue.family_id); familyNames.add(familyKey); priorFamily = familyKey
    if (familyValue.alt_name !== undefined) { boundedText(familyValue.alt_name, 'family.alt_name'); if (asciiFold(familyValue.alt_name) === familyKey) throw new TypeError('font family alias is invalid') }
    if (!Array.isArray(familyValue.faces) || familyValue.faces.length > 10_000) throw new TypeError('font face collection is invalid')
    let priorFace = ''
    for (const [faceIndex, faceValue] of familyValue.faces.entries()) {
      exactOptional(faceValue, ['face_id', 'family', 'weight', 'style', 'stretch', 'source'], ['alt_name'], `families[${familyIndex}].faces[${faceIndex}]`)
      string(faceValue.face_id, 'face.face_id', ID)
      if (faceValue.family !== familyValue.name || faceValue.alt_name !== familyValue.alt_name || ![400, 700].includes(faceValue.weight as number) || !['normal', 'italic'].includes(faceValue.style as string) || faceValue.stretch !== 100 || faces.has(faceValue.face_id)) throw new TypeError('font face identity is invalid or duplicate')
      const faceOrder = `${String(faceValue.weight).padStart(4, '0')}\u0000${faceValue.style}`
      if (faceIndex > 0 && priorFace >= faceOrder) throw new TypeError('font faces are not canonically ordered')
      priorFace = faceOrder
      const source = faceValue.source
      exactOptional(source, ['kind', 'face_slot', 'font_table_part', 'font_table_path', 'relationships_part', 'relationships_sha256', 'relationship_id', 'relationship_type', 'relationship_target', 'asset_part', 'asset_content_type', 'stored_byte_length', 'stored_sha256', 'content_sha256', 'resource_id', 'obfuscation', 'licensing'], ['collection_index'], 'font face source')
      const expectedSlot = new Map([['400:normal', 'embedRegular'], ['700:normal', 'embedBold'], ['400:italic', 'embedItalic'], ['700:italic', 'embedBoldItalic']]).get(`${faceValue.weight}:${faceValue.style}`)
      if (source.kind !== 'document' || source.face_slot !== expectedSlot || source.collection_index !== undefined) throw new TypeError('font face slot or collection index is invalid')
      if (!value.font_table || source.font_table_part !== value.font_table.part_name || source.relationships_part !== value.font_table.font_relationships_part || source.relationships_sha256 !== value.font_table.font_relationships_sha256 || source.relationship_type !== fontRelationshipType) throw new TypeError('font face relationship is not inventory-backed')
      validPart(source.font_table_part, 'source.font_table_part'); validPart(source.relationships_part, 'source.relationships_part'); validPart(source.asset_part, 'source.asset_part')
      string(source.font_table_path, 'source.font_table_path'); string(source.relationship_id, 'source.relationship_id', ID); string(source.relationship_target, 'source.relationship_target')
      string(source.relationships_sha256, 'source.relationships_sha256', SHA256); string(source.stored_sha256, 'source.stored_sha256', SHA256); string(source.content_sha256, 'source.content_sha256', SHA256)
      string(source.asset_content_type, 'source.asset_content_type'); string(source.resource_id, 'source.resource_id', ID)
      if (typeof source.stored_byte_length !== 'number' || !Number.isSafeInteger(source.stored_byte_length) || source.stored_byte_length < 32 || asciiFold(source.asset_content_type) !== asciiFold(OBFUSCATED_FONT) || source.resource_id !== `font:${source.content_sha256}` || resolveTarget(source.font_table_part, source.relationship_target) !== source.asset_part) throw new TypeError('font resource binding is incomplete')
      exactOptional(source.obfuscation, ['algorithm', 'font_key'], ['subsetted'], 'font obfuscation')
      string(source.obfuscation.font_key, 'font obfuscation key')
      if (source.obfuscation.algorithm !== 'ecma-376-font-obfuscation' || !FONT_KEY.test(source.obfuscation.font_key) || source.obfuscation.font_key === '{00000000-0000-0000-0000-000000000000}' || source.obfuscation.subsetted !== undefined && typeof source.obfuscation.subsetted !== 'boolean') throw new TypeError('font obfuscation binding is invalid')
      exact(source.licensing, ['embedding_origin', 'rights_source', 'rights_status', 'embedding_rights', 'no_subsetting', 'allowed_scope'], 'font licensing')
      if (source.licensing.embedding_origin !== 'document-package' || source.licensing.rights_source !== 'sfnt-os2-fstype' || source.licensing.rights_status !== 'verified' || !['installable', 'preview-print', 'editable'].includes(source.licensing.embedding_rights as string) || typeof source.licensing.no_subsetting !== 'boolean' || source.licensing.allowed_scope !== 'document-only' || source.licensing.no_subsetting && source.obfuscation.subsetted === true) throw new TypeError('font licensing attestation is invalid')
      const assetKey = canonicalPart(source.asset_part)
      if (resources.has(source.resource_id) || assetParts.has(assetKey) || relationshipIDs.has(source.relationship_id)) throw new TypeError('font resources are duplicate')
      resources.add(source.resource_id); assetParts.add(assetKey); relationshipIDs.add(source.relationship_id)
      faces.set(faceValue.face_id, faceValue as unknown as NativeDOCXFontFaceV1)
      if (faces.size > 4_096) throw new TypeError('font face collection exceeds the canonical Go v1 bound')
    }
  }

  let priorReference = ''
  const referenceKeys = new Set<string>()
  for (const [index, reference] of value.references.entries()) {
    exact(reference, ['family', 'weight', 'style', 'scope_ids'], `references[${index}]`)
    boundedText(reference.family, 'reference.family')
    if (![400, 700].includes(reference.weight as number) || !['normal', 'italic'].includes(reference.style as string) || !Array.isArray(reference.scope_ids) || reference.scope_ids.length === 0 || reference.scope_ids.length > 10_000) throw new TypeError('font reference is invalid')
    const key = `${asciiFold(reference.family)}\u0000${String(reference.weight).padStart(4, '0')}\u0000${reference.style}`
    if (referenceKeys.has(key) || index > 0 && compareGoStrings(priorReference, key) >= 0) throw new TypeError('font references are duplicate or not canonically ordered')
    referenceKeys.add(key); priorReference = key
    const scopes = new Set<string>()
    for (const [scopeIndex, scope] of reference.scope_ids.entries()) { string(scope, 'reference.scope_id', ID); if (scopes.has(scope) || scopeIndex > 0 && reference.scope_ids[scopeIndex - 1] >= scope) throw new TypeError('reference scopes are duplicate or not canonically ordered'); scopes.add(scope) }
  }

  if (value.native_text_manifest === undefined) {
    if (faces.size > 0 || value.native_text_manifest_sha256 !== undefined) throw new TypeError('native text manifest is required for embedded faces')
  } else {
    string(value.native_text_manifest_sha256, 'native_text_manifest_sha256', SHA256)
    const decoded = validateFontManifest(value.native_text_manifest)
    if (!decoded.ok || decoded.value.manifestId !== `docx.fonts.${stableToken(value.document_id)}` || decoded.value.revision !== value.revision || decoded.value.fallbackChains.length !== 0 || decoded.value.faces.length !== faces.size || value.native_text_manifest_sha256 !== nativeDOCXCanonicalWireSHA256V1(decoded.value)) throw new TypeError('native text manifest identity or digest is invalid')
    let priorFaceID = ''
    for (const manifestFace of decoded.value.faces) {
      const backing = faces.get(manifestFace.faceId)
      const aliases = backing?.alt_name === undefined ? [] : [backing.alt_name]
      if (!backing || manifestFace.faceId <= priorFaceID || JSON.stringify(manifestFace.aliases ?? []) !== JSON.stringify(aliases) || manifestFace.family !== backing.family || manifestFace.weight !== backing.weight || manifestFace.style !== backing.style || manifestFace.stretch !== backing.stretch || manifestFace.source.kind !== 'document' || manifestFace.source.resourceId !== backing.source.resource_id || manifestFace.source.contentDigest !== backing.source.content_sha256 || manifestFace.source.collectionIndex !== backing.source.collection_index) throw new TypeError('native text manifest face is not exactly inventory-backed')
      priorFaceID = manifestFace.faceId
    }
  }
  const typed = value as unknown as NativeDOCXFontInventoryV1
  const digestInput = structuredClone(typed)
  digestInput.inventory_sha256 = ''
  if (typed.inventory_sha256 !== nativeDOCXCanonicalWireSHA256V1(digestInput)) throw new TypeError('inventory_sha256 does not attest the complete canonical inventory')
  return typed
}

export function decodeNativeDOCXFontInventoryV1(json: string): NativeDOCXFontInventoryV1 {
  if (typeof json !== 'string' || new TextEncoder().encode(json).byteLength === 0 || new TextEncoder().encode(json).byteLength > NATIVE_DOCX_FONT_INVENTORY_MAX_JSON_BYTES) throw new TypeError('font inventory JSON is empty or exceeds its bound')
  let parsed: unknown
  try { parsed = JSON.parse(json) } catch { throw new TypeError('font inventory JSON is malformed') }
  const value = validateInventory(parsed)
  if (encodeNativeDOCXFontInventoryV1(value) !== json) throw new TypeError('font inventory JSON must be the exact canonical Go v1 encoding (duplicate fields and alternate encodings are refused)')
  return structuredClone(value)
}

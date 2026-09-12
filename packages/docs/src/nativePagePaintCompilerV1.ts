/**
 * Canonical server-side join for the qualified native DOCX page-paint slice.
 *
 * This module composes an authoritative content-addressed font resolver, the
 * pinned HarfBuzz shaper, deterministic pagination, and page-paint. It has no
 * browser, DOM, HTML, CSS, canvas, LibreOffice, or platform-font dependency.
 */

import {
  NATIVE_TEXT_LAYOUT_VERSION,
  validateFontManifest,
  type FontResource,
  type FontResolutionRequest,
  type NativeFontManifest,
  type NativeFontResolver,
  type NativeTextRefusal,
  type ResolvedFontFace,
  type TextRunInput,
} from '@injoffice/font-metrics/layout'
import { hasNativeDocxPageFieldsV1, nativeDocxPageFieldDocumentV1, DOCX_PAGE_FIELD_LIMITS, type NativeDocxPageFieldVariantV1 } from './nativePageFieldsV1.js'
import { solveNativeDocxLayoutFixedPointV1 } from './nativeLayoutFixedPointV1.js'
import { canonicalWireSha256 } from './nativePagePaintWireV1.js'
import {decodeExplicitFontPolicyV1,selectExplicitFontV1,type ExplicitFontPolicyV1} from '@injoffice/font-metrics/layout'
import {qualifyNativeDocxFontSubstitutionsV1,isQualifiedNativeDocxFontDiagnosticV1} from './nativeFontSubstitutionEvidenceV1.js'
import {compileNativeDocxFontSubstitutionPreviewV1} from './nativePagePaintV1.js'
export {decodeNativeDocxFontSubstitutionPreviewV1,DOCX_FONT_SUBSTITUTION_PREVIEW_PROTOCOL,DOCX_FONT_SUBSTITUTION_WARNING} from './nativeFontSubstitutionPreviewV1.js'
export type {NativeDocxFontSubstitutionPreviewV1} from './nativeFontSubstitutionPreviewV1.js'
import { nativeDocxPageNumberV1 } from './nativePageNumbersV1.js'
import { nativeDocxBodyPageFieldRunsV1, nativeDocxBodyPageFieldDocumentV1, nativeDocxBodyPageFieldValuesV1 } from './nativeBodyPageFieldsV1.js'
import {
  createHarfBuzzTextShaperV1,
  inspectHarfBuzzFontMetricsV1,
  isCanonicalHarfBuzzTextShaperV1,
  type HarfBuzzTextShaperV1,
} from '@injoffice/font-metrics/harfbuzz'
import { sha256 } from '@noble/hashes/sha2.js'
import { bytesToHex } from '@noble/hashes/utils.js'
import {
  DOCX_MAX_TWIPS_FOR_MILLIPOINTS,
  decodeNativeDocxDocument,
  type NativeDocxDocumentV1,
  type NativeDocxValidationIssue,
} from './nativeContract.js'
import { decodeNativeDocxResolvedLayout, type NativeDocxResolvedLayoutInputV1 } from './nativeResolvedLayout.js'
import { shapeNativeDocxLinesWithParagraphWidthsV1 } from './nativeShapingLines.js'
import type { NativeDocxLineIntervalPlanV1 } from './nativeShapingLines.js'
import { hasNativeSquareWrapV1, deriveNativeSquareWrapPlanV1 } from './nativeSquareWrapV1.js'
import { qualifyNativeDocxTablesV1, nativeDocxTableProjectionSha256V1, nativeDocxTableGeometryV1 } from './nativeTablePagePaintV1.js'
import {qualifyApproximateLegacyTables} from './nativeLegacyTableOriginV1.js'
import { asciiLowerNative, compareNativeCodeUnits } from './nativeDeterminism.js'
import { decodeNativeDOCXFontInventoryV1, type NativeDOCXFontInventoryV1 } from './nativeFontInventoryV1.js'
export { decodeNativeDOCXFontInventoryV1 } from './nativeFontInventoryV1.js'
import {
  DOCX_PAGINATION_REQUEST_PROTOCOL,
  DOCX_PAGINATION_REQUEST_VERSION,
  decodeNativeDocxPaginationRequestV1,
  paginateNativeDocxV1,
  paginateNativeDocxApproximateLegacyV1,
  validateNativeDocxApproximatePaginatedLayoutSourceV1,
  validateNativeDocxPaginatedLayoutSourceV1,
  type NativeDocxPaginationRequestV1,
} from './nativePaginationV1.js'
import { decodeNativeDocxPaginationSettings, type NativeDocxPaginationSettingsV1 } from './nativePaginationSettings.js'
import {
  DOCX_PAGE_PAINT_LIMITS,
  DOCX_PAGE_PAINT_REQUEST_PROTOCOL,
  DOCX_PAGE_PAINT_REQUEST_VERSION,
  compileNativeDocxPagePaintV1,
  compileNativeDocxApproximatePagePreviewV1,
  compileNativeDocxApproximateComputedPagePreviewV1,
  decodeNativeDocxApproximateComputedPagePaintV1,
  decodeNativeDocxPagePaintForRequestV1,
  decodeNativeDocxPagePaintRequestV1,
  nativeDocxPagePaintFontManifestSha256V1,
  nativeDocxPagePaintMediaAssetsSha256V1,
  nativeDocxPagePaintOutputSha256V1,
  nativeDocxPagePaintRequestSha256V1,
  nativeDocxPagePaintPaginatedLayoutSha256V1,
  nativeDocxPagePaintShapedLinesSha256V1,
  validateNativeDocxPagePaintForRequestV1,
  type NativeDocxContentAddressedFaceV1,
  type NativeDocxGlyphOutlineRequestV1,
  type NativeDocxGlyphOutlineResultV1,
  type NativeDocxPagePaintRequestV1,
  type NativeDocxPagePaintV1,
} from './nativePagePaintV1.js'
import { layoutNativeDocxHeadersFootersV1 } from './nativeHeaderFooterLayoutV1.js'
import {
  prepareNativeDocxPagePaintMediaAssetsV1,
  type NativeDocxAuthoritativeMediaAssetV1,
} from './nativeImagePagePaintV1.js'
export { DOCX_INLINE_IMAGE_LIMITS } from './nativeImagePagePaintV1.js'
export type { NativeDocxAuthoritativeMediaAssetV1 } from './nativeImagePagePaintV1.js'
import { qualifyNativeDocxSectionColumnsV1 } from './nativeSectionColumnsV1.js'

export const DOCX_PAGE_PAINT_COMPILER_PROTOCOL = 'injoffice.docx.page-paint-compiler'
export const DOCX_PAGE_PAINT_COMPILER_VERSION = 1 as const
export const DOCX_EMBEDDED_FONT_RESOLVER_ID = 'injoffice.docx.embedded-fonts'

const PROVIDER_ID = /^[A-Za-z0-9][A-Za-z0-9._:/-]{0,127}$/
const SHA256 = /^sha256:[0-9a-f]{64}$/
const MAX_FONT_BYTES = 64 * 1024 * 1024
const MAX_UNIQUE_FONT_BYTES = 128 * 1024 * 1024

export interface NativeDocxAuthoritativeFontAssetV1 {
  face_id: string
  face_slot: string
  resource_id: string
  content_digest: `sha256:${string}`
  collection_index: number | null
  bytes: Uint8Array
}

export interface NativeDocxPagePaintPrepareInputV1 {
  protocol: typeof DOCX_PAGE_PAINT_COMPILER_PROTOCOL
  version: typeof DOCX_PAGE_PAINT_COMPILER_VERSION
  source_revision: string
  outline_provider: { provider_id: string; provider_revision: string }
  document: unknown
  resolved_layout: unknown
  pagination_settings: unknown
  font_inventory_json: string
  font_assets: readonly NativeDocxAuthoritativeFontAssetV1[]
  media_assets: readonly NativeDocxAuthoritativeMediaAssetV1[]
}

/** Explicit host-owned font resources for read-only layout. No system lookup or
 * substitution is performed. Embedded document faces cannot be overridden. */
export interface NativeDocxHostFontsV1 {
  manifest: NativeFontManifest
  resolver: NativeFontResolver
  /** Only the dedicated approximate font renderer accepts this explicit policy. */
  substitutionPolicy?:ExplicitFontPolicyV1
}

export interface NativeDocxPagePaintPreparedV1 {
  protocol: typeof DOCX_PAGE_PAINT_COMPILER_PROTOCOL
  version: typeof DOCX_PAGE_PAINT_COMPILER_VERSION
  page_paint_request: NativeDocxPagePaintRequestV1
  request_sha256: string
  outline_requests: NativeDocxGlyphOutlineRequestV1[]
  providers: {
    resolver_id: string
    resolver_revision: string
    shaper_id: string
    shaper_revision: string
    bidi_id: string
    bidi_revision: string
    bidi_unicode_version: string
    unicode13_revision: string
    outline_id: string
    outline_revision: string
  }
}

export interface NativeDocxPagePaintCompleteInputV1 {
  prepared: NativeDocxPagePaintPreparedV1
  outline_results: readonly NativeDocxGlyphOutlineResultV1[]
}

export type { NativeDocxApproximationEligibilityV1, NativeDocxApproximatePagePreviewV1 } from './nativeApproximationV1.js'
export { DOCX_APPROXIMATE_PREVIEW_PROTOCOL, DOCX_APPROXIMATE_PREVIEW_POLICY, decodeNativeDocxApproximatePagePreviewV1 } from './nativeApproximationV1.js'
import { decodeNativeDocxApproximationEligibilityV1, decodeNativeDocxApproximatePagePreviewV1 } from './nativeApproximationV1.js'
import { projectNativeDocxAbsentFontSizesV1, DOCX_ABSENT_FONT_SIZE_WARNING, type NativeDocxHostDefaultSizePolicyV1, type NativeDocxApproximatedFontSizeV1 } from './nativeAbsentFontSizeV1.js'
export type { NativeDocxHostDefaultSizePolicyV1, NativeDocxAbsentFontSizeV1, NativeDocxApproximatedFontSizeV1 } from './nativeAbsentFontSizeV1.js'
export { validNativeDocxHostDefaultSizePolicyV1 } from './nativeAbsentFontSizeV1.js'
export interface NativeDocxApproximateRuntimeV1 {
  createShaper?: (sourceRevision: string) => HarfBuzzTextShaperV1
  fonts?: NativeDocxHostFontsV1
  fontSizePolicy?: NativeDocxHostDefaultSizePolicyV1
}
import { projectNativeDocxAutomaticBordersV1, decodeNativeDocxAutomaticBorderPreviewV1, DOCX_AUTO_BORDER_PREVIEW_PROTOCOL, type NativeDocxAutomaticBorderPreviewV1 } from './nativeAutomaticBorderPreviewV1.js'
import { DOCX_AUTO_BORDER_POLICY, DOCX_AUTO_BORDER_WARNING } from './nativeAutomaticBorderEvidenceV1.js'
export { decodeNativeDocxAutomaticBorderPreviewV1, DOCX_AUTO_BORDER_PREVIEW_PROTOCOL } from './nativeAutomaticBorderPreviewV1.js'
export type { NativeDocxAutomaticBorderPreviewV1 } from './nativeAutomaticBorderPreviewV1.js'
export { DOCX_AUTO_BORDER_POLICY, DOCX_AUTO_BORDER_WARNING } from './nativeAutomaticBorderEvidenceV1.js'
export type { NativeDocxAutomaticBorderEvidenceV1 } from './nativeAutomaticBorderEvidenceV1.js'

/** Opt-in consumer contrast policy. Only the final visibly approximate envelope
 * escapes; neither transformed source nor strict prepared artifacts are public. */
export async function renderNativeDocxAutomaticBorderPreviewV1(input: NativeDocxPagePaintPrepareInputV1, outlineProvider: import('./nativePagePaintV1.js').NativeDocxGlyphOutlineProviderV1, runtime?: NativeDocxApproximateRuntimeV1, legacyEligibility?: unknown): Promise<NativeDocxAutomaticBorderPreviewV1> {
  if (runtime?.fontSizePolicy !== undefined && legacyEligibility === undefined) throw new TypeError('Host default size requires independently eligible current-layout approximation')
  const projection = projectNativeDocxAutomaticBordersV1(input.document, input.resolved_layout)
  const projectedInput = { ...input, document: projection.document, resolved_layout: projection.resolved }
  let paint: Pick<NativeDocxPagePaintV1, 'status' | 'pages' | 'resources' | 'diagnostics'>
  let provenance: NativeDocxPagePaintV1['provenance']
  let eligibility: import('./nativeApproximationV1.js').NativeDocxApproximationEligibilityV1 | undefined
  let approximatedFontSizes: NativeDocxApproximatedFontSizeV1[] | undefined
  const reasons = [DOCX_AUTO_BORDER_WARNING as string]
  if (legacyEligibility !== undefined) {
    const legacy = await renderNativeDocxApproximatePagePreviewV1(projectedInput, legacyEligibility, outlineProvider, runtime)
    eligibility = decodeNativeDocxApproximationEligibilityV1(legacyEligibility, legacy.rendering_provenance.pagination_settings)
    paint = legacy; provenance = legacy.rendering_provenance
    approximatedFontSizes = legacy.approximated_font_sizes
    reasons.push(...legacy.reasons)
  } else {
    const prepared = await prepareNativeDocxPagePaintV1(projectedInput, runtime)
    const painted = await compileNativeDocxPagePaintV1(prepared.page_paint_request, outlineProvider)
    if (!painted.ok) throw new TypeError('Automatic-border painting failed validation')
    paint = painted.value; provenance = painted.value.provenance
  }
  const result: NativeDocxAutomaticBorderPreviewV1 = {
    protocol: DOCX_AUTO_BORDER_PREVIEW_PROTOCOL, version: 1, fidelity: 'approximate', read_only: true,
    policy: DOCX_AUTO_BORDER_POLICY, page_background_rgb: 'FFFFFF', source: projection.source,
    approximated_render_properties: projection.facts, source_diagnostics: projection.source_diagnostics, reasons,
    ...(eligibility ? { legacy_eligibility: eligibility,table_border_layout_policy:'collapsed-horizontal-border-reservation-v1' as const } : {}),
    ...(approximatedFontSizes ? { approximated_font_sizes: approximatedFontSizes } : {}),
    status: paint.status, pages: paint.pages, resources: paint.resources, diagnostics: paint.diagnostics, rendering_provenance: provenance,
  }
  const validated = decodeNativeDocxAutomaticBorderPreviewV1(result)
  if (!validated.ok) throw new TypeError('Automatic-border output failed distinct approximate validation')
  return validated.value
}

/** Explicit read-only legacy-settings preview. The strict preparation and
 * original settings remain intact; only the separate result is approximate. */
export async function renderNativeDocxApproximatePagePreviewV1(input: NativeDocxPagePaintPrepareInputV1, eligibilityValue: unknown, outlineProvider: import('./nativePagePaintV1.js').NativeDocxGlyphOutlineProviderV1, runtime?: NativeDocxApproximateRuntimeV1): Promise<import('./nativeApproximationV1.js').NativeDocxApproximatePagePreviewV1> {
  const settings = decodeNativeDocxPaginationSettings(input.pagination_settings)
  if (!settings.ok) failIssues('Approximate settings are invalid', settings.issues)
  const eligibility = decodeNativeDocxApproximationEligibilityV1(eligibilityValue, settings.value)
  let applied: NativeDocxApproximatedFontSizeV1[] = []
  if (runtime?.fontSizePolicy !== undefined) {
    if (eligibility.status !== 'eligible') throw new TypeError('Host size policy requires independently eligible approximate settings')
    const projected = projectNativeDocxAbsentFontSizesV1(input.document, input.resolved_layout, eligibility.absent_font_sizes ?? [], runtime.fontSizePolicy)
    input = { ...input, resolved_layout: projected.resolved }
    applied = projected.applied
  }
  const result = await renderNativeDocxApproximatePagePreviewInternalV1(input, eligibility, outlineProvider, runtime)
  if (applied.length > 0) {
    result.approximated_font_sizes = applied
    result.reasons.push(DOCX_ABSENT_FONT_SIZE_WARNING)
  }
  const validated = decodeNativeDocxApproximatePagePreviewV1(result)
  if (!validated.ok) throw new TypeError('Approximate output omitted its source absence or explicit host-size policy')
  return validated.value
}

async function renderNativeDocxApproximatePagePreviewInternalV1(input: NativeDocxPagePaintPrepareInputV1, eligibility: unknown, outlineProvider: import('./nativePagePaintV1.js').NativeDocxGlyphOutlineProviderV1, runtime?: NativeDocxApproximateRuntimeV1): Promise<import('./nativeApproximationV1.js').NativeDocxApproximatePagePreviewV1> {
  const document = decodeNativeDocxDocument(input.document)
  if (!document.ok) failIssues('native document is invalid', document.issues)
  const bodyFields = nativeDocxBodyPageFieldRunsV1(document.value)
  const probe = bodyFields.length ? nativeDocxBodyPageFieldDocumentV1(document.value, Object.fromEntries(bodyFields.map(run => [run.id, '1']))) : document.value
  if (bodyFields.length || hasNativeSquareWrapV1(document.value) || hasNativeDocxPageFieldsV1(probe)) {
    const prepared = await prepareNativeDocxPagePaintInternalV1(input, runtime, eligibility)
    return compileNativeDocxApproximateComputedPagePreviewV1(prepared.page_paint_request, eligibility, outlineProvider)
  }
  const prepared = await prepareNativeDocxPagePaintV1(input, runtime)
  return compileNativeDocxApproximatePagePreviewV1(prepared.page_paint_request, eligibility, outlineProvider)
}

export interface NativeDocxPagePaintCompletedV1 {
  page_paint_request: NativeDocxPagePaintRequestV1
  page_paint_output: NativeDocxPagePaintV1
  outline_results: NativeDocxGlyphOutlineResultV1[]
  canonical_request_sha256: string
  canonical_output_sha256: string
  canonical_request_validated: true
  canonical_output_validated: true
  outline_coverage_complete: true
}

function refusal(code: NativeTextRefusal['decisions'][number]['code'], message: string, attemptedFaceIds: readonly string[] = []): NativeTextRefusal {
  return Object.freeze({
    status: 'refused' as const,
    decisions: Object.freeze([{ code, message, recoverable: false }]),
    attemptedFaceIds: Object.freeze([...attemptedFaceIds]),
  })
}

function sameFace(left: ResolvedFontFace, right: ResolvedFontFace): boolean {
  return left.faceId === right.faceId && left.family === right.family && left.postscriptName === right.postscriptName
    && left.weight === right.weight && left.style === right.style && left.stretch === right.stretch
    && left.sourceKind === right.sourceKind && left.resourceId === right.resourceId
    && left.contentDigest === right.contentDigest && left.collectionIndex === right.collectionIndex
    && left.resolution === right.resolution && left.matchedFamily === right.matchedFamily
    && left.fallbackChainId === right.fallbackChainId
}

function asciiEqual(left: string, right: string): boolean {
  return left.length === right.length && asciiLowerNative(left) === asciiLowerNative(right)
}

function compareUTF8(left: string, right: string): number {
  const leftBytes = new TextEncoder().encode(left), rightBytes = new TextEncoder().encode(right)
  for (let index = 0; index < Math.min(leftBytes.length, rightBytes.length); index += 1) if (leftBytes[index] !== rightBytes[index]) return leftBytes[index]! - rightBytes[index]!
  return leftBytes.length - rightBytes.length
}

function verifyAttestedFontBytes(backing: NativeDOCXFontInventoryV1['families'][number]['faces'][number], bytes: Uint8Array): void {
  if (bytes.byteLength !== backing.source.stored_byte_length || bytes.byteLength < 32) throw new TypeError('font bytes do not match the attested stored byte length')
  const key = backing.source.obfuscation.font_key.replace(/[{}-]/g, '')
  if (key.length !== 32) throw new TypeError('font obfuscation key is invalid')
  const stored = Uint8Array.from(bytes)
  for (let index = 0; index < 32; index += 1) stored[index] ^= Number.parseInt(key.slice((15 - index % 16) * 2, (16 - index % 16) * 2), 16)
  if (`sha256:${bytesToHex(sha256(stored))}` !== backing.source.stored_sha256) throw new TypeError('font bytes, obfuscation key, and stored package digest do not exact-join')
  if (bytes.byteLength < 12) throw new TypeError('font bytes have a malformed sfnt envelope')
  const tables = (bytes[4]! << 8) | bytes[5]!
  let fsType: number | undefined
  for (let index = 0; index < tables; index += 1) {
    const base = 12 + 16 * index
    if (base + 16 > bytes.byteLength) throw new TypeError('font bytes have a malformed sfnt directory')
    if (String.fromCharCode(...bytes.subarray(base, base + 4)) !== 'OS/2') continue
    const offset = bytes[base + 8]! * 0x1000000 + (bytes[base + 9]! << 16) + (bytes[base + 10]! << 8) + bytes[base + 11]!
    const length = bytes[base + 12]! * 0x1000000 + (bytes[base + 13]! << 16) + (bytes[base + 14]! << 8) + bytes[base + 15]!
    if (length < 10 || offset < 0 || offset + length > bytes.byteLength) throw new TypeError('font bytes have a malformed OS/2 licensing table')
    fsType = (bytes[offset + 8]! << 8) | bytes[offset + 9]!
    break
  }
  const embeddingLevel = fsType === undefined ? -1 : fsType & 0x000e
  if (fsType === undefined || fsType & 0xfcf1 || ![0, 2, 4, 8].includes(embeddingLevel) || embeddingLevel === 2 || fsType & 0x0200) throw new TypeError('font bytes do not carry qualified embedding permissions')
  const rights = fsType & 0x0004 ? 'preview-print' : fsType & 0x0008 ? 'editable' : 'installable'
  const noSubsetting = (fsType & 0x0100) !== 0
  if (backing.source.licensing.embedding_rights !== rights || backing.source.licensing.no_subsetting !== noSubsetting || noSubsetting && backing.source.obfuscation.subsetted === true) throw new TypeError('font license flags do not exact-join the supplied SFNT bytes')
}

/** Creates an exact resolver over already-authorized, content-addressed bytes. */
export function createNativeDocxEmbeddedFontResolverV1(inventory: NativeDOCXFontInventoryV1, assetsValue: readonly NativeDocxAuthoritativeFontAssetV1[]): NativeFontResolver {
  if (!inventory.native_text_manifest || !inventory.native_text_manifest_sha256) throw new TypeError('font inventory has no authoritative native text manifest')
  const decoded = validateFontManifest(inventory.native_text_manifest)
  if (!decoded.ok) throw new TypeError('font inventory native text manifest is invalid')
  const resolverRevision = inventory.inventory_sha256
  if (!PROVIDER_ID.test(resolverRevision)) throw new TypeError('resolver revision must be a bounded stable provider revision')
  const manifest = decoded.value
  if (!Array.isArray(assetsValue) || assetsValue.length !== manifest.faces.length) throw new TypeError('authoritative font assets must exactly cover the manifest faces')
  const inventoryFaces = new Map(inventory.families.flatMap((family) => family.faces).map((face) => [face.face_id, face]))
  const resources = new Map<string, FontResource>()
  const resourceIDs = new Set<string>()
  let totalBytes = 0
  for (const asset of assetsValue) {
    if (!asset || typeof asset !== 'object' || Object.keys(asset).sort().join(',') !== 'bytes,collection_index,content_digest,face_id,face_slot,resource_id' || !(asset.bytes instanceof Uint8Array) || !PROVIDER_ID.test(asset.face_id) || !PROVIDER_ID.test(asset.resource_id) || !SHA256.test(asset.content_digest) || asset.collection_index !== null && (!Number.isSafeInteger(asset.collection_index) || asset.collection_index < 0 || asset.collection_index > 65_535)) throw new TypeError('authoritative font asset is malformed')
    const face = manifest.faces.find((candidate) => candidate.faceId === asset.face_id)
    const backing = inventoryFaces.get(asset.face_id)
    const collectionIndex = asset.collection_index ?? undefined
    if (!face || !backing || resources.has(face.faceId) || resourceIDs.has(asset.resource_id) || face.source.kind !== 'document' || face.source.resourceId !== asset.resource_id || face.source.contentDigest !== asset.content_digest || face.source.collectionIndex !== collectionIndex || backing.source.face_slot !== asset.face_slot || backing.source.resource_id !== asset.resource_id || backing.source.content_sha256 !== asset.content_digest || backing.source.collection_index !== collectionIndex) throw new TypeError('authoritative font asset does not exact-join one attested document resource and face slot')
    if (asset.bytes.byteLength === 0 || asset.bytes.byteLength > MAX_FONT_BYTES || totalBytes + asset.bytes.byteLength > MAX_UNIQUE_FONT_BYTES) throw new RangeError('authoritative font assets exceed the bounded compiler budget')
    const owned = Uint8Array.from(asset.bytes)
    verifyAttestedFontBytes(backing, owned)
    const metrics = inspectHarfBuzzFontMetricsV1({ bytes: owned, contentDigest: asset.content_digest, ...(collectionIndex !== undefined ? { collectionIndex } : {}) })
    const resolved: ResolvedFontFace = Object.freeze({
      faceId: face.faceId,
      family: face.family,
      ...(face.postscriptName !== undefined ? { postscriptName: face.postscriptName } : {}),
      weight: face.weight,
      style: face.style,
      stretch: face.stretch,
      sourceKind: face.source.kind,
      resourceId: face.source.resourceId,
      contentDigest: asset.content_digest,
      ...(collectionIndex !== undefined ? { collectionIndex } : {}),
      resolution: 'exact',
      matchedFamily: face.family,
    })
    resources.set(face.faceId, Object.freeze({ face: resolved, bytes: owned, metrics }))
    resourceIDs.add(asset.resource_id)
    totalBytes += owned.byteLength
  }

  const manifestDigest = nativeDocxPagePaintFontManifestSha256V1(manifest)
  const resolve = (request: FontResolutionRequest) => {
    try {
      if (nativeDocxPagePaintFontManifestSha256V1(request.manifest) !== manifestDigest) return refusal('invalid-contract', 'resolver request manifest does not match the authoritative manifest')
      const attempted: string[] = []
      for (const family of request.run.font.families) {
        const matches = manifest.faces.filter((face) => {
          const names = [face.family, ...(face.aliases ?? [])]
          return names.some((name) => asciiEqual(name, family)) && face.weight === request.run.font.weight
            && face.style === request.run.font.style && face.stretch === request.run.font.stretch
            && (request.run.font.postscriptName === undefined || face.postscriptName === request.run.font.postscriptName)
        })
        attempted.push(...matches.map((face) => face.faceId))
        if (matches.length > 1) return refusal('invalid-contract', 'more than one authoritative face matches the exact authored request', attempted)
        if (matches.length === 1) {
          const resource = resources.get(matches[0]!.faceId)!
          const selected = Object.freeze({ ...resource.face, matchedFamily: family })
          return Object.freeze({ status: 'resolved' as const, face: selected, attemptedFaceIds: Object.freeze([...attempted]), decisions: Object.freeze([]) })
        }
      }
      return refusal('font-not-found', 'no exact document-embedded face matches the authored request', attempted)
    } catch {
      return refusal('provider-failure', 'authoritative font resolver rejected hostile or unreadable input')
    }
  }
  const load = (face: ResolvedFontFace) => {
    try {
      const resource = resources.get(face.faceId)
      if (!resource || !sameFace(resource.face, { ...face, matchedFamily: resource.face.matchedFamily })) return refusal('font-bytes-unavailable', 'resolved face does not exact-join an authoritative font resource', [face.faceId])
      return Object.freeze({ face: Object.freeze({ ...face }), bytes: Uint8Array.from(resource.bytes), metrics: Object.freeze({ ...resource.metrics }) })
    } catch {
      return refusal('provider-failure', 'authoritative font resource could not be snapshotted')
    }
  }
  return Object.freeze({ providerId: DOCX_EMBEDDED_FONT_RESOLVER_ID, providerRevision: resolverRevision, resolve, load })
}

function failIssues(label: string, issues: readonly NativeDocxValidationIssue[]): never {
  throw new TypeError(`${label}: ${issues.map((entry) => `${entry.path} ${entry.code}: ${entry.message}`).join('; ')}`)
}

function twips(value: number): number {
  if (!Number.isSafeInteger(value) || Object.is(value, -0) || Math.abs(value) > DOCX_MAX_TWIPS_FOR_MILLIPOINTS) throw new RangeError('DOCX twip geometry exceeds safe integer milli-points')
  return value * 50
}

function shapingDimensions(document: NativeDocxDocumentV1, settings: NativeDocxPaginationSettingsV1): { width: number; tab: number } {
  if (document.sections.length === 0) throw new TypeError('native document has no section geometry')
  const geometries = document.sections.map(qualifyNativeDocxSectionColumnsV1)
  if (geometries.some((entry) => !entry.ok)) throw new TypeError('native section column geometry is not exactly representable')
  const widths = geometries.flatMap((entry) => entry.ok ? entry.value.columns.map((column) => column.width_millipoints) : [])
  const width = widths[0] ?? 0
  if (widths.some((candidate) => candidate !== width)) throw new TypeError('native sections require more than one shaping width')
  const tab = twips(settings.default_tab_stop_twips)
  if (width <= 0 || tab <= 0) throw new RangeError('native page width or default tab stop is non-positive')
  return { width, tab }
}

function shapingParagraphWidths(
  document: NativeDocxDocumentV1,
  tableWidths: ReadonlyMap<string, number>,
): Map<string, number> {
  const widths = new Map(tableWidths)
  const relatedStories = [...document.headers, ...document.footers]
  if (relatedStories.length === 0) return widths
  const bodyWidths = document.sections.map((section) => {
    const qualified = qualifyNativeDocxSectionColumnsV1(section)
    if (!qualified.ok) throw new TypeError(`native section ${section.id} has no exact header/footer body geometry`)
    return qualified.value.body_width_millipoints
  })
  const bodyWidth = bodyWidths[0]
  if (!bodyWidth || bodyWidths.some((candidate) => candidate !== bodyWidth)) throw new TypeError('header/footer stories require one exact shared section body width')
  for (const story of relatedStories) for (const block of story.blocks) {
    if (block.paragraph) widths.set(block.paragraph.id, bodyWidth)
    if (block.table) for (const row of block.table.rows) for (const cell of row.cells) for (const paragraph of cell.paragraphs) widths.set(paragraph.id, bodyWidth)
  }
  return widths
}

function resolvedFontReferences(resolved: NativeDocxResolvedLayoutInputV1): NativeDOCXFontInventoryV1['references'] {
  const references = new Map<string, { family: string; weight: 400 | 700; style: 'normal' | 'italic'; scopes: Set<string> }>()
  const add = (properties: { font_family?: string; bold?: boolean; italic?: boolean }, scope: string) => {
    if (!properties.font_family) return
    const weight = properties.bold === true ? 700 : 400
    const style = properties.italic === true ? 'italic' : 'normal'
    const key = `${asciiLowerNative(properties.font_family)}\u0000${weight}\u0000${style}`
    let entry = references.get(key)
    if (!entry) { entry = { family: properties.font_family, weight, style, scopes: new Set() }; references.set(key, entry) }
    entry.scopes.add(scope)
  }
  for (const run of resolved.runs) add(run.properties, run.run_id)
  for (const paragraph of resolved.paragraphs) {
    add(paragraph.paragraph_mark_properties, paragraph.paragraph_id)
    if (paragraph.numbering) add(paragraph.numbering.marker_properties, paragraph.paragraph_id)
  }
  return [...references.values()].sort((left, right) => {
    const leftKey = `${asciiLowerNative(left.family)}\u0000${String(left.weight).padStart(4, '0')}\u0000${left.style}`
    const rightKey = `${asciiLowerNative(right.family)}\u0000${String(right.weight).padStart(4, '0')}\u0000${right.style}`
    return compareUTF8(leftKey, rightKey)
  }).map((entry) => ({ family: entry.family, weight: entry.weight, style: entry.style, scope_ids: [...entry.scopes].sort() }))
}

async function attestResolvedFontReferencesBeforeBidi(
  resolver: NativeFontResolver,
  manifest: NativeFontManifest,
  references: NativeDOCXFontInventoryV1['references'],
  requireExactMatching = false,
): Promise<void> {
  const loadedFaces = new Set<string>()
  let totalBytes = 0
  for (const reference of references) {
    const run: TextRunInput = Object.freeze({
      version: NATIVE_TEXT_LAYOUT_VERSION,
      text: '',
      fontSizeMilliPoints: 1_000,
      font: Object.freeze({ families: Object.freeze([reference.family]), weight: reference.weight, style: reference.style, stretch: 100 }),
      script: 'Zyyy',
      language: 'und',
      direction: 'ltr',
    })
    const resolution = await resolver.resolve(Object.freeze({ manifest, run }))
    if (!resolution || typeof resolution !== 'object' || 'status' in resolution && resolution.status === 'refused' || !('face' in resolution)) throw new TypeError('every authored font reference must resolve through the attested font provider before bidi')
    const face = resolution.face
    if (requireExactMatching && face.resolution !== 'exact') throw new TypeError('Preserved font matching metadata requires exact supplied font faces; substitution is not implemented')
    const manifestFace = manifest.faces.find((candidate) => candidate.faceId === face.faceId)
    if (!manifestFace || face.sourceKind === 'system' || !face.contentDigest || manifestFace.source.contentDigest !== face.contentDigest || face.weight !== reference.weight || face.style !== reference.style || face.stretch !== 100) throw new TypeError('pre-bidi font resolution does not exact-join one content-addressed manifest face')
    if (loadedFaces.has(face.faceId)) continue
    const resource = await resolver.load(face)
    if (!resource || typeof resource !== 'object' || 'status' in resource && resource.status === 'refused' || !('bytes' in resource) || !(resource.bytes instanceof Uint8Array) || !('face' in resource) || !sameFace(resource.face, face)) throw new TypeError('pre-bidi font load does not exact-join the attested content-addressed resource')
    if (resource.bytes.byteLength > MAX_FONT_BYTES || totalBytes + resource.bytes.byteLength > MAX_UNIQUE_FONT_BYTES) throw new RangeError('font resources exceed the bounded compiler budget')
    if (`sha256:${bytesToHex(sha256(resource.bytes))}` !== face.contentDigest) throw new TypeError('pre-bidi font load does not exact-join the attested content-addressed resource')
    totalBytes += resource.bytes.byteLength
    loadedFaces.add(face.faceId)
  }
}

function validateInventoryPackagePartJoins(inventory: NativeDOCXFontInventoryV1, document: NativeDocxDocumentV1, settings: NativeDocxPaginationSettingsV1): void {
  const binding = inventory.font_table
  if (!binding) throw new TypeError('font inventory has no font-table package binding')
  if (binding.main_relationships_part !== settings.relationships_part || binding.main_relationships_sha256 !== settings.relationships_sha256) throw new TypeError('font inventory main relationship part/hash does not exact-join pagination settings')
  const parts = new Map(document.passthrough_parts.map((part) => [part.part_name, part]))
  const join = (partName: string, sha: string, contentType?: string, byteLength?: number) => {
    const part = parts.get(partName)
    if (!part || part.sha256 !== sha || contentType !== undefined && !asciiEqual(part.content_type, contentType) || byteLength !== undefined && part.byte_length !== byteLength) throw new TypeError(`font inventory package part ${partName} does not exact-join the native document passthrough inventory`)
  }
  join(binding.main_relationships_part, binding.main_relationships_sha256, 'application/vnd.openxmlformats-package.relationships+xml')
  join(binding.part_name, binding.sha256, 'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml')
  if (binding.font_relationships_part !== undefined || binding.font_relationships_sha256 !== undefined) {
    if (!binding.font_relationships_part || !binding.font_relationships_sha256) throw new TypeError('font inventory has incomplete font relationship-part evidence')
    join(binding.font_relationships_part, binding.font_relationships_sha256, 'application/vnd.openxmlformats-package.relationships+xml')
  } else if (inventory.families.some((family) => family.faces.length > 0)) throw new TypeError('embedded fonts require relationship-part evidence')
  for (const family of inventory.families) for (const face of family.faces) join(face.source.asset_part, face.source.stored_sha256, face.source.asset_content_type, face.source.stored_byte_length)
}

function outlineKey(request: NativeDocxGlyphOutlineRequestV1): string {
  return `${request.face.face_id}\u0000${request.face.content_digest}\u0000${request.face.collection_index ?? -1}\u0000${String(request.glyph_id).padStart(10, '0')}`
}

export function collectNativeDocxPagePaintOutlineRequestsV1(requestValue: unknown): NativeDocxGlyphOutlineRequestV1[] {
  const decoded = decodeNativeDocxPagePaintRequestV1(requestValue)
  if (!decoded.ok) failIssues('page-paint request is invalid', decoded.issues)
  const request = decoded.value
  const headerFooter = layoutNativeDocxHeadersFootersV1({
    document: request.pagination_request.document,
    resolved_layout: request.pagination_request.resolved_layout,
    shaped_lines: request.pagination_request.shaped_lines,
    pagination_settings: request.pagination_request.pagination_settings,
    paginated_layout: request.paginated_layout,
    page_field_variants: request.page_field_variants,
  })
  if (request.paginated_layout.status !== 'paginated' || headerFooter.status !== 'placed') return []
  const faces = new Map(request.font_manifest.faces.map((face) => [face.faceId, face]))
  const shaped = new Map(request.pagination_request.shaped_lines.paragraphs.map((paragraph) => [paragraph.paragraph_id, paragraph]))
  const variantByPlacedID = new Map(headerFooter.pages.flatMap((page) => page.lines.map((line) => [line.id, request.page_field_variants?.find((variant) => variant.page_id === page.page_id)?.shaped_lines] as const)))
  const unique = new Map<string, NativeDocxGlyphOutlineRequestV1>()
  const placedLines = [
    ...request.paginated_layout.pages.flatMap((page) => [...page.lines, ...(page.note_stories ?? []).flatMap((story) => story.lines)]),
    ...headerFooter.pages.flatMap((page) => page.lines),
  ]
  for (const placed of placedLines) {
    const line = (variantByPlacedID.get(placed.id)?.paragraphs.find((paragraph) => paragraph.paragraph_id === placed.paragraph_id) ?? shaped.get(placed.paragraph_id))?.lines[placed.source_line_ordinal]
    if (!line || line.id !== placed.line_id) throw new TypeError('placed line does not exact-join its shaped paragraph projection')
    for (const fragment of line.fragments) for (const glyph of fragment.glyphs) {
      const face = fragment.face_id ? faces.get(fragment.face_id) : undefined
      if (!face || !face.source.contentDigest || face.source.kind === 'system') throw new TypeError('painted glyph does not exact-join a content-addressed non-system manifest face')
      const outline: NativeDocxGlyphOutlineRequestV1 = {
        face: { face_id: face.faceId, content_digest: face.source.contentDigest, ...(face.source.collectionIndex !== undefined ? { collection_index: face.source.collectionIndex } : {}) },
        glyph_id: glyph.glyph_id,
      }
      if (outline.glyph_id <= 0) throw new TypeError('painted glyph uses the missing-glyph sentinel')
      unique.set(outlineKey(outline), outline)
      if (unique.size > DOCX_PAGE_PAINT_LIMITS.maxUniqueGlyphOutlines) throw new RangeError('page-paint outline requests exceed the compiler budget')
    }
  }
  return [...unique.entries()].sort(([left], [right]) => compareNativeCodeUnits(left, right)).map(([, value]) => value)
}

export async function prepareNativeDocxPagePaintV1(input: NativeDocxPagePaintPrepareInputV1, runtime?: { createShaper?: (sourceRevision: string) => HarfBuzzTextShaperV1; fonts?: NativeDocxHostFontsV1 }): Promise<NativeDocxPagePaintPreparedV1> {
  if(runtime?.fonts?.substitutionPolicy!==undefined)throw new TypeError('Font substitution requires the dedicated approximate renderer, not strict preparation')
  return prepareNativeDocxPagePaintInternalV1(input, runtime)
}

/** Original extraction/inventory is retained; only this separate read-only
 * renderer qualifies explicit source-bound substitution diagnostics. */
export async function renderNativeDocxFontSubstitutionPreviewV1(input:NativeDocxPagePaintPrepareInputV1,outlineProvider:import('./nativePagePaintV1.js').NativeDocxGlyphOutlineProviderV1,runtime:{createShaper?:(sourceRevision:string)=>HarfBuzzTextShaperV1;fonts:NativeDocxHostFontsV1}):Promise<import('./nativeFontSubstitutionPreviewV1.js').NativeDocxFontSubstitutionPreviewV1>{
  if(!runtime?.fonts?.substitutionPolicy)throw new TypeError('Font preview requires an explicit supplied-font policy')
  const policy=decodeExplicitFontPolicyV1(runtime.fonts.substitutionPolicy)
  const settings=decodeNativeDocxPaginationSettings(input.pagination_settings),document=decodeNativeDocxDocument(input.document)
  if(!settings.ok||!document.ok||settings.value.profile!=='word-modern-default'||hasNativeDocxPageFieldsV1(document.value)||nativeDocxBodyPageFieldRunsV1(document.value).length||hasNativeSquareWrapV1(document.value))throw new TypeError('Font-only preview excludes legacy settings, fields and wrap policy combinations')
  const prepared=await prepareNativeDocxPagePaintInternalV1(input,runtime,undefined,policy)
  return compileNativeDocxFontSubstitutionPreviewV1(prepared.page_paint_request,policy,outlineProvider)
}

// Approximate preparation is private: its computed placement is never returned
// as a public strict prepared artifact, and retains the original settings.
async function prepareNativeDocxPagePaintInternalV1(input: NativeDocxPagePaintPrepareInputV1, runtime?: { createShaper?: (sourceRevision: string) => HarfBuzzTextShaperV1; fonts?: NativeDocxHostFontsV1 }, approximateEligibility?: unknown,fontPolicy?:ExplicitFontPolicyV1): Promise<NativeDocxPagePaintPreparedV1> {
  if (!input || typeof input !== 'object' || Object.keys(input).sort().join(',') !== 'document,font_assets,font_inventory_json,media_assets,outline_provider,pagination_settings,protocol,resolved_layout,source_revision,version' || input.protocol !== DOCX_PAGE_PAINT_COMPILER_PROTOCOL || input.version !== DOCX_PAGE_PAINT_COMPILER_VERSION || !PROVIDER_ID.test(input.source_revision) || !input.outline_provider || typeof input.outline_provider !== 'object' || Object.keys(input.outline_provider).sort().join(',') !== 'provider_id,provider_revision' || !PROVIDER_ID.test(input.outline_provider.provider_id) || !PROVIDER_ID.test(input.outline_provider.provider_revision)) throw new TypeError('native page-paint compiler input identity is invalid')
  const document = decodeNativeDocxDocument(input.document)
  if (!document.ok) failIssues('native document is invalid', document.issues)
  const resolved = decodeNativeDocxResolvedLayout(input.resolved_layout)
  if (!resolved.ok) failIssues('resolved layout is invalid', resolved.issues)
  const settings = decodeNativeDocxPaginationSettings(input.pagination_settings)
  if (!settings.ok) failIssues('pagination settings are invalid', settings.issues)
  if (approximateEligibility !== undefined && decodeNativeDocxApproximationEligibilityV1(approximateEligibility, settings.value).status !== 'eligible') throw new TypeError('Source settings are ineligible for approximate body-field layout')
  const inventory = decodeNativeDOCXFontInventoryV1(input.font_inventory_json)
  if (inventory.document_id !== document.value.document_id || inventory.revision !== document.value.revision || inventory.package_sha256 !== document.value.source.package_sha256 || inventory.main_part !== document.value.source.main_part
    || inventory.document_id !== resolved.value.document_id || inventory.revision !== resolved.value.revision || inventory.main_part !== resolved.value.source_parts.main_part
    || inventory.font_table?.part_name !== resolved.value.source_parts.font_table_part
    || inventory.document_id !== settings.value.document_id || inventory.revision !== settings.value.revision || inventory.package_sha256 !== settings.value.package_sha256 || inventory.main_part !== settings.value.main_part) throw new TypeError('font inventory does not exact-join the document, resolved layout, pagination settings, package, main part, and source revision')
  if (JSON.stringify(inventory.references) !== JSON.stringify(resolvedFontReferences(resolved.value))) throw new TypeError('font inventory references do not exactly and completely cover the resolved document scopes')
  validateInventoryPackagePartJoins(inventory, document.value, settings.value)
  if (!inventory.native_text_manifest && !runtime?.fonts) throw new TypeError('document has no embedded fonts; configure explicit host fonts for native preview')
  const manifest = validateFontManifest(runtime?.fonts?.manifest ?? inventory.native_text_manifest)
  if (!manifest.ok) throw new TypeError('font inventory native text manifest is invalid')
  if (runtime?.fonts) {
    if (manifest.value.revision !== document.value.revision || manifest.value.fallbackChains.length !== 0 || !PROVIDER_ID.test(runtime.fonts.resolver.providerId) || !PROVIDER_ID.test(runtime.fonts.resolver.providerRevision)) throw new TypeError('host font provider must bind this revision without fallback chains')
    if (inventory.native_text_manifest) createNativeDocxEmbeddedFontResolverV1(inventory, input.font_assets)
    else if (input.font_assets.length !== 0) throw new TypeError('document font assets lack an embedded manifest')
    const embedded = inventory.native_text_manifest?.faces ?? []
    for (const face of embedded) {
      if (!manifest.value.faces.some((candidate) => canonicalWireSha256(candidate) === canonicalWireSha256(face))) throw new TypeError('host fonts cannot override an embedded document face')
    }
    for (const face of manifest.value.faces) {
      if (face.source.kind === 'document' ? !embedded.some((candidate) => canonicalWireSha256(candidate) === canonicalWireSha256(face)) : face.source.kind !== 'host' || !face.source.contentDigest) throw new TypeError('host manifest contains an unattested or non-host face')
    }
  }
  for (const reference of inventory.references) {
    const matches = manifest.value.faces.filter((face) => face.weight === reference.weight && face.style === reference.style && face.stretch === 100 && [face.family, ...(face.aliases ?? [])].some((name) => asciiEqual(name, reference.family)))
    if (matches.length !== 1) {
      const selected=fontPolicy?selectExplicitFontV1(manifest.value,{version:1,text:'',fontSizeMilliPoints:1000,font:{families:[reference.family],weight:reference.weight,style:reference.style,stretch:100},script:'Zyyy',language:'und',direction:'ltr'},fontPolicy):null
      if(!selected||selected.face.resolution!=='substitute')throw new TypeError('every authored font reference must have exactly one attested manifest face and resource before shaping')
    }
  }
  const references = resolvedFontReferences(resolved.value)
  const resolver = runtime?.fonts?.resolver ?? createNativeDocxEmbeddedFontResolverV1(inventory, input.font_assets)
  const mediaAssets = prepareNativeDocxPagePaintMediaAssetsV1(document.value, input.media_assets)
  const shaper = runtime?.createShaper?.(input.source_revision) ?? createHarfBuzzTextShaperV1({ sourceRevision: input.source_revision })
  if (!isCanonicalHarfBuzzTextShaperV1(shaper, input.source_revision)) throw new TypeError('HarfBuzz shaper provenance does not attest the exact pinned runtime and requested engine source revision')
  await attestResolvedFontReferencesBeforeBidi(resolver, manifest.value, references, fontPolicy===undefined||resolved.value.diagnostics.some((entry) => entry.code === 'FONT_MATCHING_METADATA_PRESERVED'))
  const dimensions = shapingDimensions(document.value, settings.value)
  const bodyFields = nativeDocxBodyPageFieldRunsV1(document.value)
  const initialBodyFieldValues = Object.fromEntries(bodyFields.map(run => [run.id, '1']))
  let layoutFragmentWork = 0
  let measuredTables: import('./nativeShapingLines.js').NativeDocxShapedLinesV1 | undefined
  if (document.value.body.blocks.some(block => block.table && nativeDocxTableGeometryV1(block.table, resolved.value).layout === 'autofit')) {
    // The probe uses the same source, attested fonts and canonical shaper. A
    // generous bounded width yields intrinsic advances; final qualification
    // independently derives the same min/max from the final wrapped clusters.
    const probeWidths = new Map<string, number>()
    for (const block of document.value.body.blocks) for (const row of block.table?.rows ?? []) for (const cell of row.cells) for (const paragraph of cell.paragraphs) probeWidths.set(paragraph.id, 1_000_000_000)
    // Body fields outside tables need visible seed text even during the
    // intrinsic table probe; field values do not influence table content.
    const probeDocument = bodyFields.length ? nativeDocxBodyPageFieldDocumentV1(document.value, initialBodyFieldValues) : document.value
    const measured = await shapeNativeDocxLinesWithParagraphWidthsV1({ protocol: 'injoffice.docx.shaping-request', version: 1, document: probeDocument, resolved_layout: resolved.value, font_manifest: manifest.value, available_width_millipoints: dimensions.width, tab_interval_millipoints: dimensions.tab }, { resolver, shaper }, shapingParagraphWidths(document.value, probeWidths))
    if (!measured.ok) failIssues('autofit measurement failed validation', measured.issues)
    // Page controls are consumed by the final source-bound paginator, not by
    // intrinsic text measurement. Every other diagnostic remains a refusal.
    if(!fontPolicy&&measured.value.font_substitutions?.length)throw new TypeError('Strict preparation does not accept substituted fonts')
    const measuredFonts=fontPolicy?qualifyNativeDocxFontSubstitutionsV1(measured.value,resolved.value,manifest.value,fontPolicy,document.value):[]
    const measurementIssues = measured.value.diagnostics.filter(diagnostic => !isQualifiedNativeDocxFontDiagnosticV1(diagnostic,measuredFonts)&&(diagnostic.code !== 'page-control-deferred' || diagnostic.severity !== 'deferred'))
    if (measurementIssues.length) throw new TypeError(`Content autofit measurement refused unqualified source text or exceeded its budget: ${measurementIssues.map(diagnostic => diagnostic.code).join(', ')}`)
    measuredTables = measured.value
    layoutFragmentWork += measured.value.paragraphs.reduce((n, paragraph) => n + paragraph.lines.reduce((m, line) => m + line.fragments.length, 0), 0)
    if (bodyFields.length && layoutFragmentWork > DOCX_PAGE_FIELD_LIMITS.maxFragments) throw new RangeError('Body-field layout probe exceeds cumulative shaping fragment budget')
  }
  const qualifiedTables = approximateEligibility===undefined?qualifyNativeDocxTablesV1(document.value,resolved.value,measuredTables):qualifyApproximateLegacyTables(document.value,resolved.value,measuredTables,decodeNativeDocxApproximationEligibilityV1(approximateEligibility,settings.value))
  if (measuredTables && qualifiedTables.status !== 'qualified') throw new TypeError('Content autofit refused unsupported source geometry or unsatisfied intrinsic widths')
  if (document.value.body.blocks.some((block) => block.table !== undefined) && document.value.sections.some((section) => section.page.columns > 1)) {
    throw new TypeError('native page-paint compiler refuses table content when any section uses multi-column flow')
  }
  const paragraphWidths = shapingParagraphWidths(document.value, qualifiedTables.paragraph_widths)
  const squareWrapPresent = hasNativeSquareWrapV1(document.value)
  const solved = await solveNativeDocxLayoutFixedPointV1({ bodyFieldValues: initialBodyFieldValues, wrapPlan: {} as NativeDocxLineIntervalPlanV1 }, async state => {
  const fieldDocument = bodyFields.length ? nativeDocxBodyPageFieldDocumentV1(document.value,state.bodyFieldValues) : document.value
  const shaped = await shapeNativeDocxLinesWithParagraphWidthsV1({
    protocol: 'injoffice.docx.shaping-request', version: 1,
    document: fieldDocument, resolved_layout: resolved.value, font_manifest: manifest.value,
    available_width_millipoints: dimensions.width, tab_interval_millipoints: dimensions.tab,
  }, { resolver, shaper }, paragraphWidths, state.wrapPlan)
  if (!shaped.ok) failIssues('native shaping failed validation', shaped.issues)
  if(!fontPolicy&&shaped.value.font_substitutions?.length)throw new TypeError('Strict preparation does not accept substituted fonts')
  if(fontPolicy)qualifyNativeDocxFontSubstitutionsV1(shaped.value,resolved.value,manifest.value,fontPolicy,document.value)
  layoutFragmentWork += shaped.value.paragraphs.reduce((n, paragraph) => n + paragraph.lines.reduce((m, line) => m + line.fragments.length, 0), 0)
  if ((bodyFields.length || squareWrapPresent) && layoutFragmentWork > DOCX_PAGE_FIELD_LIMITS.maxFragments) throw new RangeError('Field/wrap layout solve exceeds cumulative shaping fragment budget')
  const paginationRequest: NativeDocxPaginationRequestV1 = {
    protocol: DOCX_PAGINATION_REQUEST_PROTOCOL,
    version: DOCX_PAGINATION_REQUEST_VERSION,
    document: fieldDocument,
    resolved_layout: resolved.value,
    shaped_lines: shaped.value,
    pagination_settings: settings.value,
  }
  const decodedPagination = decodeNativeDocxPaginationRequestV1(paginationRequest)
  if (!decodedPagination.ok) failIssues('pagination request is invalid', decodedPagination.issues)
  const paginated = approximateEligibility === undefined ? paginateNativeDocxV1(decodedPagination.value) : { ok: true as const, value: paginateNativeDocxApproximateLegacyV1(decodedPagination.value, approximateEligibility).layout }
  if (!paginated.ok) failIssues('native pagination failed validation', paginated.issues)
  const paginationIssues = approximateEligibility === undefined ? validateNativeDocxPaginatedLayoutSourceV1(paginated.value, decodedPagination.value) : validateNativeDocxApproximatePaginatedLayoutSourceV1(paginated.value, decodedPagination.value, approximateEligibility)
  if (paginationIssues.length > 0) failIssues('native pagination source join failed', paginationIssues)
  const wrapPlan = squareWrapPresent ? deriveNativeSquareWrapPlanV1(fieldDocument, resolved.value, shaped.value, paginated.value) : {}
  return { next: { bodyFieldValues: bodyFields.length ? nativeDocxBodyPageFieldValuesV1(document.value,decodedPagination.value,paginated.value) : {}, wrapPlan }, result: { shaped, decodedPagination, paginated } }
  })
  const { shaped, decodedPagination, paginated } = solved.result
  const hasPageFields = hasNativeDocxPageFieldsV1(decodedPagination.value.document)
  let pageFieldVariants: NativeDocxPageFieldVariantV1[] | undefined
  if (hasPageFields && paginated.value.status === 'paginated') {
    if (paginated.value.pages.length > DOCX_PAGE_FIELD_LIMITS.maxPages) throw new RangeError('Page-field expansion exceeds bounded page count')
    pageFieldVariants = []
    let fragments = layoutFragmentWork
    for (const page of paginated.value.pages) {
      const fieldDocument = nativeDocxPageFieldDocumentV1(decodedPagination.value.document, page.ordinal, paginated.value.pages.length, nativeDocxPageNumberV1(decodedPagination.value.document, paginated.value, page.ordinal))
      const variant = await shapeNativeDocxLinesWithParagraphWidthsV1({ protocol: 'injoffice.docx.shaping-request', version: 1, document: fieldDocument, resolved_layout: resolved.value, font_manifest: manifest.value, available_width_millipoints: dimensions.width, tab_interval_millipoints: dimensions.tab }, { resolver, shaper }, paragraphWidths, solved.state.wrapPlan)
      if (!variant.ok) failIssues('page-field shaping failed', variant.issues)
      fragments += variant.value.paragraphs.reduce((n, paragraph) => n + paragraph.lines.reduce((m, line) => m + line.fragments.length, 0), 0)
      if (fragments > DOCX_PAGE_FIELD_LIMITS.maxFragments) throw new RangeError('Page-field expansion exceeds cumulative fragment budget')
      pageFieldVariants.push({ page_id: page.id, shaped_lines: variant.value })
    }
  }
  // Width probes may be unwrapped, and fixed tables have no intrinsic probe.
  // Policies depending on final line geometry must hash the final projection
  // independently reproduced by pagination/paint validation.
  const finalQualifiedTables = approximateEligibility === undefined ? qualifiedTables : qualifyApproximateLegacyTables(decodedPagination.value.document,resolved.value,shaped.value,decodeNativeDocxApproximationEligibilityV1(approximateEligibility,settings.value))
  const pagePaintRequest: NativeDocxPagePaintRequestV1 = {
    protocol: DOCX_PAGE_PAINT_REQUEST_PROTOCOL,
    version: DOCX_PAGE_PAINT_REQUEST_VERSION,
    pagination_request: decodedPagination.value,
    paginated_layout: paginated.value,
    font_manifest: manifest.value,
    media_assets: mediaAssets,
    ...(pageFieldVariants ? { page_field_variants: pageFieldVariants } : {}),
    ...(bodyFields.length ? { body_field_source: document.value } : {}),
    integrity: {
      ...(bodyFields.length ? { body_field_source_sha256: canonicalWireSha256(document.value) } : {}),
      font_manifest_sha256: nativeDocxPagePaintFontManifestSha256V1(manifest.value),
      shaped_lines_sha256: nativeDocxPagePaintShapedLinesSha256V1(shaped.value, pageFieldVariants),
      table_projection_sha256: finalQualifiedTables.status === 'qualified' ? finalQualifiedTables.sha256 : nativeDocxTableProjectionSha256V1([]),
      media_assets_sha256: nativeDocxPagePaintMediaAssetsSha256V1(mediaAssets),
      paginated_layout_sha256: nativeDocxPagePaintPaginatedLayoutSha256V1(paginated.value),
    },
    outline_provider: { ...input.outline_provider },
  }
  const decodedRequest = approximateEligibility === undefined ? decodeNativeDocxPagePaintRequestV1(pagePaintRequest) : { ok: true as const, value: decodeNativeDocxApproximateComputedPagePaintV1(pagePaintRequest, approximateEligibility).request }
  if (!decodedRequest.ok) failIssues('page-paint request is invalid', decodedRequest.issues)
  return Object.freeze({
    protocol: DOCX_PAGE_PAINT_COMPILER_PROTOCOL,
    version: DOCX_PAGE_PAINT_COMPILER_VERSION,
    page_paint_request: decodedRequest.value,
    request_sha256: approximateEligibility === undefined ? nativeDocxPagePaintRequestSha256V1(decodedRequest.value) : canonicalWireSha256(decodedRequest.value),
    outline_requests: approximateEligibility === undefined ? collectNativeDocxPagePaintOutlineRequestsV1(decodedRequest.value) : [],
    providers: Object.freeze({
      resolver_id: resolver.providerId,
      resolver_revision: resolver.providerRevision,
      shaper_id: shaper.providerId,
      shaper_revision: shaper.providerRevision,
      bidi_id: shaped.value.providers.bidi_id,
      bidi_revision: shaped.value.providers.bidi_revision,
      bidi_unicode_version: shaped.value.providers.bidi_unicode_version,
      unicode13_revision: shaped.value.providers.unicode13_revision,
      outline_id: input.outline_provider.provider_id,
      outline_revision: input.outline_provider.provider_revision,
    }),
  })
}

export async function completeNativeDocxPagePaintV1(input: NativeDocxPagePaintCompleteInputV1): Promise<NativeDocxPagePaintCompletedV1> {
  if (!input?.prepared || input.prepared.protocol !== DOCX_PAGE_PAINT_COMPILER_PROTOCOL || input.prepared.version !== DOCX_PAGE_PAINT_COMPILER_VERSION) throw new TypeError('prepared compiler envelope is invalid')
  const request = decodeNativeDocxPagePaintRequestV1(input?.prepared?.page_paint_request)
  if (!request.ok) failIssues('prepared page-paint request is invalid', request.issues)
  if (input.prepared.request_sha256 !== nativeDocxPagePaintRequestSha256V1(request.value)) throw new TypeError('prepared canonical request hash changed before completion')
  const required = collectNativeDocxPagePaintOutlineRequestsV1(request.value)
  const shapedProviders = request.value.pagination_request.shaped_lines.providers
  const preparedProviders = input.prepared.providers
  if (preparedProviders.resolver_id !== shapedProviders.resolver_id || preparedProviders.resolver_revision !== shapedProviders.resolver_revision
    || preparedProviders.shaper_id !== shapedProviders.shaper_id || preparedProviders.shaper_revision !== shapedProviders.shaper_revision
    || preparedProviders.bidi_id !== shapedProviders.bidi_id || preparedProviders.bidi_revision !== shapedProviders.bidi_revision || preparedProviders.bidi_unicode_version !== shapedProviders.bidi_unicode_version || preparedProviders.unicode13_revision !== shapedProviders.unicode13_revision
    || preparedProviders.outline_id !== request.value.outline_provider.provider_id || preparedProviders.outline_revision !== request.value.outline_provider.provider_revision) throw new TypeError('prepared provider provenance does not exact-join the canonical request')
  if (JSON.stringify(input.prepared.outline_requests) !== JSON.stringify(required)) throw new TypeError('prepared outline request inventory changed before completion')
  if (!Array.isArray(input.outline_results) || input.outline_results.length !== required.length) throw new TypeError('outline results must exactly cover every unique page-paint glyph identity')
  const results = new Map<string, NativeDocxGlyphOutlineResultV1>()
  for (const result of input.outline_results) {
    if (!result || typeof result !== 'object' || !('face' in result) || !('glyph_id' in result)) throw new TypeError('outline result is malformed')
    const key = outlineKey({ face: result.face, glyph_id: result.glyph_id })
    if (results.has(key)) throw new TypeError('outline results contain a duplicate exact glyph identity')
    results.set(key, structuredClone(result))
  }
  if (required.some((entry) => !results.has(outlineKey(entry)))) throw new TypeError('outline results do not exact-join the prepared glyph identities')
  const provider = Object.freeze({
    providerId: request.value.outline_provider.provider_id,
    providerRevision: request.value.outline_provider.provider_revision,
    getGlyphOutline: (outline: Readonly<NativeDocxGlyphOutlineRequestV1>) => {
      const result = results.get(outlineKey(outline))
      if (!result) throw new Error('page-paint requested an unprepared glyph identity')
      return structuredClone(result)
    },
  })
  const compiled = await compileNativeDocxPagePaintV1(request.value, provider)
  if (!compiled.ok) failIssues('page-paint output failed compilation', compiled.issues)
  const joined = decodeNativeDocxPagePaintForRequestV1(compiled.value, request.value, { provider_id: provider.providerId, provider_revision: provider.providerRevision })
  if (!joined.ok) failIssues('page-paint output source join failed', joined.issues)
  const authoritative = await validateNativeDocxPagePaintForRequestV1(joined.value, request.value, provider)
  if (!authoritative.ok) failIssues('page-paint deterministic replay failed', authoritative.issues)
  return Object.freeze({
    page_paint_request: request.value,
    page_paint_output: authoritative.value,
    outline_results: required.map((entry) => structuredClone(results.get(outlineKey(entry))!)),
    canonical_request_sha256: nativeDocxPagePaintRequestSha256V1(request.value),
    canonical_output_sha256: nativeDocxPagePaintOutputSha256V1(authoritative.value),
    canonical_request_validated: true,
    canonical_output_validated: true,
    outline_coverage_complete: true,
  })
}

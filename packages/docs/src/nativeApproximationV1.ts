import type { NativeDocxPaginationSettingsV1 } from './nativePaginationSettings.js'
import type { NativeDocxPagePaintV1 } from './nativePagePaintV1.js'
import { preflightWire, decodeNativeDocxPagePaintV1, DOCX_PAGE_PAINT_PROTOCOL, DOCX_PAGE_PAINT_VERSION } from './nativePagePaintWireV1.js'
import type { NativeDocxValidationIssue } from './nativeContract.js'
import { nativeApproximationSettingReason, validNativeDocxApproximatedSettingV1, type NativeDocxApproximatedSettingV1 } from './nativeApproximationSettingsV1.js'
import { DOCX_ABSENT_FONT_SIZE_WARNING, validNativeDocxAbsentFontSizesV1, validNativeDocxApproximatedFontSizesV1, type NativeDocxAbsentFontSizeV1, type NativeDocxApproximatedFontSizeV1 } from './nativeAbsentFontSizeV1.js'

export const DOCX_APPROXIMATE_PREVIEW_PROTOCOL = 'injoffice.docx.approximate-page-preview' as const
export const DOCX_APPROXIMATE_PREVIEW_POLICY = 'current-layout-approximate-v1' as const
export const DOCX_APPROXIMATE_PREVIEW_WARNING = 'Approximate read-only preview: current InjOffice layout, not Microsoft Word compatibility-mode fidelity.' as const
export const DOCX_APPROXIMATE_LINE_BOX_WARNING = 'Current-layout policy places natural ascent at the top of expanded line boxes, leaving extra leading below the text; compressed line boxes remain unsupported.' as const
export interface NativeDocxApproximationEligibilityV1 {
  protocol: 'injoffice.docx.approximation-eligibility'
  version: 1
  document_id: string
  revision: string
  package_sha256: string
  settings_sha256: string | null
  status: 'eligible' | 'ineligible'
  legacy_compatibility_mode: 12 | 14 | null
  reasons: string[]
  approximated_settings?: NativeDocxApproximatedSettingV1[]
  absent_font_sizes?: NativeDocxAbsentFontSizeV1[]
}

export interface NativeDocxApproximatePagePreviewV1 {
  protocol: typeof DOCX_APPROXIMATE_PREVIEW_PROTOCOL
  version: 1
  fidelity: 'approximate'
  policy: typeof DOCX_APPROXIMATE_PREVIEW_POLICY
  read_only: true
  status: 'painted' | 'refused'
  source: Pick<NativeDocxApproximationEligibilityV1, 'document_id' | 'revision' | 'package_sha256' | 'settings_sha256'>
  reasons: string[]
  source_settings_diagnostics: NativeDocxPaginationSettingsV1['diagnostics']
  approximated_settings?: NativeDocxApproximatedSettingV1[]
  source_absent_font_sizes?: NativeDocxAbsentFontSizeV1[]
  approximated_font_sizes?: NativeDocxApproximatedFontSizeV1[]
  rendering_provenance: NativeDocxPagePaintV1['provenance']
  diagnostics: NativeDocxPagePaintV1['diagnostics']
  resources: NativeDocxPagePaintV1['resources']
  pages: NativeDocxPagePaintV1['pages']
}

/** Eligibility is produced from original package bytes by the native extractor,
 * not inferred from a generic unsupported diagnostic code. */
export function decodeNativeDocxApproximationEligibilityV1(value: unknown, settings: NativeDocxPaginationSettingsV1): NativeDocxApproximationEligibilityV1 {
  const candidate = value as Partial<NativeDocxApproximationEligibilityV1> | null
  const issues = preflightWire(value, 'approximation eligibility', 20_000, 1000).filter(issue => !(issue.code === 'INVALID_VALUE' && ((issue.path === '/settings_sha256' && candidate?.settings_sha256 === null) || (issue.path === '/legacy_compatibility_mode' && candidate?.legacy_compatibility_mode === null))))
  if (issues.length) throw new TypeError('invalid approximation eligibility wire')
  const input = structuredClone(value) as NativeDocxApproximationEligibilityV1
  if (!input || typeof input !== 'object' || Object.keys(input).filter(key => key !== 'approximated_settings' && key !== 'absent_font_sizes').sort().join(',') !== 'document_id,legacy_compatibility_mode,package_sha256,protocol,reasons,revision,settings_sha256,status,version'
    || input.protocol !== 'injoffice.docx.approximation-eligibility' || input.version !== 1
    || !['eligible', 'ineligible'].includes(input.status) || ![null, 12, 14].includes(input.legacy_compatibility_mode)
    || !Array.isArray(input.reasons) || input.reasons.length > 256 || input.reasons.some(reason => typeof reason !== 'string' || reason.length > 8192)
    || input.document_id !== settings.document_id || input.revision !== settings.revision || input.package_sha256 !== settings.package_sha256 || input.settings_sha256 !== (settings.settings_sha256 ?? null)) throw new TypeError('approximation eligibility does not exact-join original settings')
  const facts = input.approximated_settings ?? []
  if (input.absent_font_sizes !== undefined && !validNativeDocxAbsentFontSizesV1(input.absent_font_sizes, input.package_sha256)) throw new TypeError('Invalid source-absent font-size evidence')
  if (!Array.isArray(facts) || facts.length > 8 || facts.some(fact => !validNativeDocxApproximatedSettingV1(fact)) || new Set(facts.map(fact => fact.kind)).size !== facts.length || new Set(facts.map(fact => fact.path)).size !== facts.length) throw new TypeError('invalid approximated settings source facts')
  if (input.status === 'eligible' && (input.legacy_compatibility_mode === null || settings.profile === 'word-modern-default' || settings.compatibility_mode !== undefined || !coveredSettingsDiagnostics(settings, facts) || facts.some(fact => !input.reasons.includes(nativeApproximationSettingReason(fact))) || input.reasons.length === 0)) throw new TypeError('approximation eligibility conflicts with strict settings facts')
  return input
}

export function approximatePagePreviewEnvelope(settings: NativeDocxPaginationSettingsV1, eligibility: NativeDocxApproximationEligibilityV1, paint: NativeDocxPagePaintV1): NativeDocxApproximatePagePreviewV1 {
  return {
    protocol: DOCX_APPROXIMATE_PREVIEW_PROTOCOL, version: 1, fidelity: 'approximate', policy: DOCX_APPROXIMATE_PREVIEW_POLICY, read_only: true,
    status: paint.status,
    source: { document_id: settings.document_id, revision: settings.revision, package_sha256: settings.package_sha256, settings_sha256: settings.settings_sha256 ?? null },
    reasons: [...eligibility.reasons, DOCX_APPROXIMATE_PREVIEW_WARNING, DOCX_APPROXIMATE_LINE_BOX_WARNING],
    source_settings_diagnostics: structuredClone(settings.diagnostics),
    ...(eligibility.approximated_settings ? { approximated_settings: structuredClone(eligibility.approximated_settings) } : {}),
    ...(eligibility.absent_font_sizes ? { source_absent_font_sizes: structuredClone(eligibility.absent_font_sizes) } : {}),
    rendering_provenance: paint.provenance,
    diagnostics: paint.diagnostics, resources: paint.resources, pages: paint.pages,
  }
}

/** Browser-safe structural validation of the distinct approximate artifact.
 * This does not turn its pages into a strict source-qualified V1 projection. */
export function decodeNativeDocxApproximatePagePreviewV1(value: unknown): { ok: true; value: NativeDocxApproximatePagePreviewV1 } | { ok: false; issues: NativeDocxValidationIssue[] } {
  const invalid = (message: string): { ok: false; issues: NativeDocxValidationIssue[] } => ({ ok: false, issues: [{ code: 'INVALID_VALUE', path: '', message }] })
  try {
    const candidate = value as Partial<NativeDocxApproximatePagePreviewV1> | null
    const issues = preflightWire(value, 'approximate page preview').filter(issue => !(issue.code === 'INVALID_VALUE' && issue.path === '/source/settings_sha256' && candidate?.source?.settings_sha256 === null))
    if (issues.length) return { ok: false, issues }
    const input = structuredClone(value) as NativeDocxApproximatePagePreviewV1
    if (!input || typeof input !== 'object' || Object.keys(input).filter(key => !['approximated_settings', 'source_absent_font_sizes', 'approximated_font_sizes'].includes(key)).sort().join(',') !== 'diagnostics,fidelity,pages,policy,protocol,read_only,reasons,rendering_provenance,resources,source,source_settings_diagnostics,status,version'
      || input.protocol !== DOCX_APPROXIMATE_PREVIEW_PROTOCOL || input.version !== 1 || input.fidelity !== 'approximate' || input.policy !== DOCX_APPROXIMATE_PREVIEW_POLICY || input.read_only !== true
      || !Array.isArray(input.reasons) || input.reasons.length < 1 || input.reasons.length > 259 || !input.reasons.includes(DOCX_APPROXIMATE_PREVIEW_WARNING) || !input.reasons.includes(DOCX_APPROXIMATE_LINE_BOX_WARNING) || input.reasons.some(reason => typeof reason !== 'string' || reason.length > 8192)) return invalid('invalid approximate envelope or missing fidelity warning')
    const paint = decodeNativeDocxPagePaintV1({ protocol: DOCX_PAGE_PAINT_PROTOCOL, version: DOCX_PAGE_PAINT_VERSION, status: input.status, provenance: input.rendering_provenance, diagnostics: input.diagnostics, resources: input.resources, pages: input.pages })
    if (!paint.ok) return paint
    const settings = paint.value.provenance.pagination_settings
    const absent = input.source_absent_font_sizes ?? []
    if (!validNativeDocxAbsentFontSizesV1(absent, settings.package_sha256)) return invalid('Invalid source font-size omissions')
    if (input.approximated_font_sizes !== undefined) {
      if (!validNativeDocxApproximatedFontSizesV1(input.approximated_font_sizes, absent, settings.package_sha256) || !input.reasons.includes(DOCX_ABSENT_FONT_SIZE_WARNING)) return invalid('Missing explicit host size policy, source evidence or warning')
    } else if (input.status === 'painted' && absent.length > 0) return invalid('Painted source-absent sizes require declared host policy')
    const facts = input.approximated_settings ?? []
    if (!Array.isArray(facts) || facts.length > 8 || facts.some(fact => !validNativeDocxApproximatedSettingV1(fact)) || new Set(facts.map(fact => fact.kind)).size !== facts.length || new Set(facts.map(fact => fact.path)).size !== facts.length) return invalid('invalid retained approximate settings facts')
    if (facts.some(fact => !input.reasons.includes(nativeApproximationSettingReason(fact))) || (input.status === 'painted' && !coveredSettingsDiagnostics(settings, facts))) return invalid('missing approximate setting coverage or warning')
    if (paint.value.provenance.document_id !== settings.document_id || paint.value.provenance.revision !== settings.revision || paint.value.provenance.package_sha256 !== settings.package_sha256) return invalid('approximate rendering identity does not match retained source settings')
    if (!input.source || Object.keys(input.source).sort().join(',') !== 'document_id,package_sha256,revision,settings_sha256' || input.source.document_id !== settings.document_id || input.source.revision !== settings.revision || input.source.package_sha256 !== settings.package_sha256 || input.source.settings_sha256 !== (settings.settings_sha256 ?? null) || JSON.stringify(input.source_settings_diagnostics) !== JSON.stringify(settings.diagnostics)) return invalid('approximate source facts and retained settings diagnostics do not exact-join')
    return { ok: true, value: input }
  } catch { return invalid('approximate output could not be safely inspected') }
}

function coveredSettingsDiagnostics(settings: NativeDocxPaginationSettingsV1, facts: NativeDocxApproximatedSettingV1[]): boolean {
  return settings.diagnostics.every(reason => {
    if (reason.code === 'COMPATIBILITY_SETTING_UNSUPPORTED' && ['/w:settings[1]', '/w:settings[1]/w:compat[1]', '/w:settings[1]/w:compat[1]/w:compatSetting[1]'].includes(reason.path)) return true
    return facts.some(fact => fact.path === reason.path && reason.code === (fact.kind === 'mathPr' ? 'UNKNOWN_SETTINGS_ELEMENT' : fact.path.includes('/w:compat[1]/') ? 'COMPATIBILITY_SETTING_UNSUPPORTED' : 'PAGINATION_SETTING_UNSUPPORTED'))
  }) && facts.every(fact => settings.diagnostics.some(reason => reason.path === fact.path))
}

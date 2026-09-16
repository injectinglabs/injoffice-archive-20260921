/** Exact source settings supported only by the opt-in current-layout policy.
 * Values are retained facts, not proof these Word semantics were implemented.
 * Grouped kinds (`authoringSettings`, `duplicateSettings`, `repeatedCompatSettings`) key `values` by each
 * member's settings.xml path and retain that element's own attributes; the fact
 * `path` is one of those members, the first one strict diagnosed. */
export interface NativeDocxApproximatedSettingV1 {
  kind: 'themeFontLang' | 'decimalSymbol' | 'listSeparator' | 'shapeDefaults' | 'mathPr'
    | NativeDocxApproximatedCompatSettingFlag | NativeDocxApproximatedLegacyCompatOption
    | 'compatibilityMode' | 'autoHyphenation' | 'characterSpacingControl' | 'authoringSettings'
    | 'duplicateSettings' | 'repeatedCompatSettings'
  path: string
  values: Record<string, string>
}

/** Microsoft compatSetting flags recorded as current-layout facts (MS-DOCX 2.6). */
export type NativeDocxApproximatedCompatSettingFlag = 'overrideTableStyleFontSizeAndJustification' | 'enableOpenTypeFeatures' | 'doNotFlipMirrorIndents' | 'differentiateMultirowTableHeaders' | 'useWord2013TrackBottomHyphenation' | 'allowHyphenationAtTrackBottom' | 'allowTextAfterFloatingTableBreak'

/** ECMA-376 Part 1 17.15.3 and Part 4 Transitional w:compat option leaves. */
export type NativeDocxApproximatedLegacyCompatOption = typeof DOCX_APPROXIMATE_LEGACY_COMPAT_OPTIONS[number]

export const DOCX_APPROXIMATE_LEGACY_COMPAT_OPTIONS = [
  'adjustLineHeightInTable', 'alignTablesRowByRow', 'allowSpaceOfSameStyleInTable', 'applyBreakingRules', 'autofitToFirstFixedWidthCell', 'autoSpaceLikeWord95',
  'balanceSingleByteDoubleByteWidth', 'cachedColBalance', 'convMailMergeEsc', 'displayHangulFixedWidth', 'doNotAutofitConstrainedTables', 'doNotBreakConstrainedForcedTable',
  'doNotBreakWrappedTables', 'doNotExpandShiftReturn', 'doNotLeaveBackslashAlone', 'doNotSnapToGridInCell', 'doNotSuppressIndentation', 'doNotSuppressParagraphBorders',
  'doNotUseEastAsianBreakRules', 'doNotUseHTMLParagraphAutoSpacing', 'doNotUseIndentAsNumberingTabStop', 'doNotVertAlignCellWithSp', 'doNotVertAlignInTxbx', 'doNotWrapTextWithPunct',
  'footnoteLayoutLikeWW8', 'forgetLastTabAlignment', 'growAutofit', 'layoutRawTableWidth', 'layoutTableRowsApart', 'lineWrapLikeWord6', 'mwSmallCaps', 'noExtraLineSpacing',
  'noLeading', 'noSpaceRaiseLower', 'noTabHangInd', 'printBodyTextBeforeHeader', 'printColBlack', 'selectFldWithFirstOrLastChar', 'shapeLayoutLikeWW8', 'showBreaksInFrames',
  'spaceForUL', 'spacingInWholePoints', 'splitPgBreakAndParaMark', 'subFontBySize', 'suppressBottomSpacing', 'suppressSpacingAtTopOfPage', 'suppressSpBfAfterPgBrk', 'suppressTopSpacing',
  'suppressTopSpacingWP', 'swapBordersFacingPages', 'truncateFontHeightsLikeWP6', 'uiCompat97To2003', 'ulTrailSpace', 'underlineTabInNumList', 'useAltKinsokuLineBreakRules',
  'useAnsiKerningPairs', 'useFELayout', 'useNormalStyleForList', 'usePrinterMetrics', 'useSingleBorderforContiguousCells', 'useWord2002TableStyleRules', 'useWord97LineBreakRules',
  'wpJustification', 'wpSpaceWidth', 'wrapTrailSpaces',
] as const

/** ECMA-376 17.15.1 settings.xml children that are editor, proofing, template,
 * grid-snapping, protection or save-time state; none feeds native shaping or
 * pagination. Mirrors the Go extractor allowlist exactly. */
export const DOCX_APPROXIMATE_AUTHORING_SETTINGS = [
  'activeWritingStyle', 'attachedSchema', 'attachedTemplate', 'captions', 'clickAndTypeStyle', 'defaultTableStyle', 'displayBackgroundShape',
  'displayHorizontalDrawingGridEvery', 'displayVerticalDrawingGridEvery', 'docVars', 'documentProtection', 'doNotDemarcateInvalidXml', 'doNotDisplayPageBoundaries',
  'doNotEmbedSmartTags', 'doNotUseMarginsForDrawingGridOrigin', 'drawingGridHorizontalOrigin', 'drawingGridHorizontalSpacing', 'drawingGridVerticalOrigin',
  'drawingGridVerticalSpacing', 'embedSystemFonts', 'embedTrueTypeFonts', 'forceUpgrade', 'formsDesign', 'hdrShapeDefaults', 'hideGrammaticalErrors', 'hideSpellingErrors',
  'linkStyles', 'mailMerge', 'noPunctuationKerning', 'printFormsData', 'printPostScriptOverText', 'readModeInkLockDown', 'removeDateAndTime', 'removePersonalInformation',
  'rsids', 'saveFormsData', 'saveSubsetFonts', 'saveThroughXslt', 'saveXmlDataOnly', 'schemaLibrary', 'showEnvelope', 'showXMLTags', 'smartTagType', 'stylePaneFormatFilter',
  'summaryLength', 'useXSLTWhenSaving', 'view', 'writeProtection', 'alwaysMergeEmptyNamespace',
] as const

const COMPAT_SETTING_FLAGS: ReadonlySet<string> = new Set<NativeDocxApproximatedCompatSettingFlag>(['overrideTableStyleFontSizeAndJustification', 'enableOpenTypeFeatures', 'doNotFlipMirrorIndents', 'differentiateMultirowTableHeaders', 'useWord2013TrackBottomHyphenation', 'allowHyphenationAtTrackBottom', 'allowTextAfterFloatingTableBreak'])
const LEGACY_COMPAT_OPTIONS: ReadonlySet<string> = new Set(DOCX_APPROXIMATE_LEGACY_COMPAT_OPTIONS)
const AUTHORING_SETTINGS: ReadonlySet<string> = new Set(DOCX_APPROXIMATE_AUTHORING_SETTINGS)
const ON_OFF = ['true', 'false', 'on', 'off', '0', '1']
const LANGUAGE_TAG = /^[A-Za-z]{2,3}(-[A-Za-z0-9]{1,8})*$/
const SETTINGS_ROOT = '/w:settings[1]'
const TOP_LEVEL_MEMBER = /^\/w:settings\[1\]\/w:([A-Za-z][A-Za-z0-9]*)\[([1-9][0-9]*)\]$/
const NESTED_MEMBER = /^\/w:settings\[1\](\/[A-Za-z][A-Za-z0-9]*:[A-Za-z][A-Za-z0-9]*\[[1-9][0-9]*\])+$/
const COMPAT_SETTING_MEMBER = /^\/w:settings\[1\]\/w:compat\[1\]\/w:compatSetting\[[1-9][0-9]*\]$/
const MAX_GROUP_MEMBERS = 64
const MAX_SUMMARY_LENGTH = 512

/** Byte-identical to the Go extractor's nativeApproximationSettingReason. */
export function nativeApproximationSettingReason(fact: NativeDocxApproximatedSettingV1): string {
  const members = Object.keys(fact.values).length
  switch (true) {
    case fact.kind === 'autoHyphenation':
      return `Current-layout approximation records autoHyphenation at ${fact.path} as not applied; automatic hyphenation is not performed and Word line breaks may differ`
    case fact.kind === 'compatibilityMode':
      return `Current-layout approximation records the repeated or non-leading compatibilityMode attestation at ${fact.path}; its agreeing value is the disclosed legacy mode and Word layout may differ`
    case fact.kind === 'characterSpacingControl':
      return `Current-layout approximation records characterSpacingControl at ${fact.path} as not applied; East Asian punctuation and kana advances are not compressed and Word line breaks may differ`
    case fact.kind === 'repeatedCompatSettings':
      return `Current-layout approximation records ${members} repeated compatSetting attestations anchored at ${fact.path} as not applied; none of them is an input to current layout, so their disagreement cannot change this preview`
    case fact.kind === 'themeFontLang':
      return `Current-layout approximation records themeFontLang at ${fact.path} as not applied; language-driven theme font selection is not performed and Word font choice may differ`
    case fact.kind === 'authoringSettings':
      return `Current-layout approximation records ${members} authoring-only settings anchored at ${fact.path} as not applied; current layout does not consume them, so Word editing, proofing, grid, template, and display behavior may differ`
    case fact.kind === 'duplicateSettings':
      return `Current-layout approximation records ${members} duplicate settings anchored at ${fact.path} as not applied; each repeats its first occurrence or is authoring-only, and the first occurrence is used`
    case LEGACY_COMPAT_OPTIONS.has(fact.kind):
      return `Current-layout approximation records legacy compatibility option ${fact.kind} at ${fact.path} as not applied; Word compatibility layout is not emulated`
    default:
      return `Current-layout approximation disregards ${fact.kind} at ${fact.path}; source values are retained and Word layout may differ`
  }
}

export function validNativeDocxApproximatedSettingV1(value: unknown): value is NativeDocxApproximatedSettingV1 {
  if (!value || typeof value !== 'object' || Object.keys(value).sort().join(',') !== 'kind,path,values') return false
  const fact = value as NativeDocxApproximatedSettingV1
  if (typeof fact.kind !== 'string' || typeof fact.path !== 'string' || fact.path.length > 4096) return false
  if (!fact.values || typeof fact.values !== 'object' || Array.isArray(fact.values) || Object.values(fact.values).some(value => typeof value !== 'string')) return false
  const keys = Object.keys(fact.values).sort().join(',')
  const source = fact.values
  if (COMPAT_SETTING_FLAGS.has(fact.kind)) {
    if (!/^\/w:settings\[1\]\/w:compat\[1\]\/w:compatSetting\[[1-9]\]$/.test(fact.path) || keys !== 'val') return false
    // Recorded, never applied: the approximate tier emulates neither attested
    // behaviour, so the value is disclosed rather than used to qualify a source.
    return source.val === '0' || source.val === '1'
  }
  if (LEGACY_COMPAT_OPTIONS.has(fact.kind)) {
    return fact.path === `${SETTINGS_ROOT}/w:compat[1]/w:${fact.kind}[1]` && (keys === '' || (keys === 'val' && ON_OFF.includes(source.val)))
  }
  switch (fact.kind) {
    case 'compatibilityMode':
      return /^\/w:settings\[1\]\/w:compat\[1\]\/w:compatSetting\[[2-9]\]$/.test(fact.path) && keys === 'val' && ['12', '14', '15'].includes(source.val)
    case 'autoHyphenation':
      return fact.path === `${SETTINGS_ROOT}/w:autoHyphenation[1]` && Object.entries(source).every(([key, value]) =>
        (key === 'val' && ON_OFF.includes(value)) || (key === 'doNotHyphenateCaps' && ON_OFF.includes(value)) || ((key === 'hyphenationZone' || key === 'consecutiveHyphenLimit') && /^[0-9]{1,10}$/.test(value)))
    case 'characterSpacingControl':
      return fact.path === `${SETTINGS_ROOT}/w:characterSpacingControl[1]` && keys === 'val' && ['compressPunctuation', 'compressPunctuationAndJapaneseKana'].includes(source.val)
    case 'repeatedCompatSettings':
      return validGroupedFact(fact, path => COMPAT_SETTING_MEMBER.test(path), summary => {
        const split = summary.indexOf('=')
        return split > 0 && COMPAT_SETTING_FLAGS.has(summary.slice(0, split)) && ['0', '1'].includes(summary.slice(split + 1))
      })
    case 'authoringSettings':
      return validGroupedFact(fact, path => { const match = TOP_LEVEL_MEMBER.exec(path); return match !== null && AUTHORING_SETTINGS.has(match[1]!) })
    case 'duplicateSettings':
      return validGroupedFact(fact, path => NESTED_MEMBER.test(path) && !/\[1\]$/.test(path) && !/\/w:compat\[[0-9]+\]$/.test(path))
  }
  // Native XML paths use the canonical hashed prefix for the exact Math namespace,
  // independently of the prefix chosen in the package's source XML.
  if (fact.path !== `${SETTINGS_ROOT}/${fact.kind === 'mathPr' ? 'nsf4b2b884' : 'w'}:${fact.kind}[1]`) return false
  switch (fact.kind) {
    case 'themeFontLang':
      return typeof source.val === 'string' && validLanguageSlot(source.val, false) && Object.keys(source).every(key => key === 'val' || (['eastAsia', 'bidi'].includes(key) && validLanguageSlot(source[key]!, true)))
    case 'decimalSymbol': return keys === 'val' && ['.', ','].includes(source.val)
    case 'listSeparator': return keys === 'val' && [',', ';'].includes(source.val)
    case 'shapeDefaults': return keys === 'idmap,spidmax' && source.idmap === '1' && /^[1-9][0-9]*$/.test(source.spidmax) && Number(source.spidmax) <= 2147483647
    case 'mathPr': {
      const exact = { mathFont: 'Cambria Math', brkBin: 'before', brkBinSub: '--', lMargin: '0', rMargin: '0', defJc: 'centerGroup', wrapIndent: '1440', intLim: 'subSup', naryLim: 'undOvr' }
      return keys === [...Object.keys(exact), 'smallFrac', 'dispDef'].sort().join(',')
        && Object.entries(exact).every(([key, value]) => source[key] === value)
        && ON_OFF.includes(source.smallFrac)
        && ON_OFF.includes(source.dispDef)
    }
    default: return false
  }
}

/** A well-formed BCP 47-shaped tag; script slots may be empty or Word's x-none. */
function validLanguageSlot(value: string, scriptSlot: boolean): boolean {
  if (scriptSlot && (value === '' || value === 'x-none')) return true
  return value.length <= 35 && LANGUAGE_TAG.test(value)
}

/** Grouped facts anchor at one member path and retain bounded attribute summaries. */
function validGroupedFact(fact: NativeDocxApproximatedSettingV1, memberPath: (path: string) => boolean, memberSummary?: (summary: string) => boolean): boolean {
  const members = Object.entries(fact.values)
  if (members.length === 0 || members.length > MAX_GROUP_MEMBERS || !(fact.path in fact.values)) return false
  return members.every(([path, summary]) => path.length <= 4096 && memberPath(path) && summary.length <= MAX_SUMMARY_LENGTH && (memberSummary === undefined || memberSummary(summary)))
}

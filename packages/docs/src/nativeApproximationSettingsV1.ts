/** Exact source settings supported only by the opt-in current-layout policy.
 * Values are retained facts, not proof these Word semantics were implemented. */
export interface NativeDocxApproximatedSettingV1 {
  kind: 'themeFontLang' | 'decimalSymbol' | 'listSeparator' | 'shapeDefaults' | 'mathPr' | 'overrideTableStyleFontSizeAndJustification' | 'enableOpenTypeFeatures' | 'doNotFlipMirrorIndents'
  path: string
  values: Record<string, string>
}

export function nativeApproximationSettingReason(fact: NativeDocxApproximatedSettingV1): string {
  return `Current-layout approximation disregards ${fact.kind} at ${fact.path}; source values are retained and Word layout may differ`
}

export function validNativeDocxApproximatedSettingV1(value: unknown): value is NativeDocxApproximatedSettingV1 {
  if (!value || typeof value !== 'object' || Object.keys(value).sort().join(',') !== 'kind,path,values') return false
  const fact = value as NativeDocxApproximatedSettingV1
  if (!fact.values || typeof fact.values !== 'object' || Array.isArray(fact.values) || Object.values(fact.values).some(value => typeof value !== 'string')) return false
  const keys = Object.keys(fact.values).sort().join(',')
  const source = fact.values
  if (['overrideTableStyleFontSizeAndJustification', 'enableOpenTypeFeatures', 'doNotFlipMirrorIndents'].includes(fact.kind)) {
    return /^\/w:settings\[1\]\/w:compat\[1\]\/w:compatSetting\[[1-4]\]$/.test(fact.path) && keys === 'val' && source.val === '1'
  }
  // Native XML paths use the canonical hashed prefix for the exact Math namespace,
  // independently of the prefix chosen in the package's source XML.
  if (fact.path !== `/w:settings[1]/${fact.kind === 'mathPr' ? 'nsf4b2b884' : 'w'}:${fact.kind}[1]`) return false
  switch (fact.kind) {
    case 'themeFontLang':
      return ['en-US', 'fr-FR'].includes(source.val) && Object.keys(source).every(key => key === 'val' || (['eastAsia', 'bidi'].includes(key) && source[key] === 'x-none'))
    case 'decimalSymbol': return keys === 'val' && ['.', ','].includes(source.val)
    case 'listSeparator': return keys === 'val' && [',', ';'].includes(source.val)
    case 'shapeDefaults': return keys === 'idmap,spidmax' && source.idmap === '1' && /^[1-9][0-9]*$/.test(source.spidmax) && Number(source.spidmax) <= 2147483647
    case 'mathPr': {
      const exact = { mathFont: 'Cambria Math', brkBin: 'before', brkBinSub: '--', lMargin: '0', rMargin: '0', defJc: 'centerGroup', wrapIndent: '1440', intLim: 'subSup', naryLim: 'undOvr' }
      return keys === [...Object.keys(exact), 'smallFrac', 'dispDef'].sort().join(',')
        && Object.entries(exact).every(([key, value]) => source[key] === value)
        && ['true', 'false', 'on', 'off', '0', '1'].includes(source.smallFrac)
        && ['true', 'false', 'on', 'off', '0', '1'].includes(source.dispDef)
    }
    default: return false
  }
}

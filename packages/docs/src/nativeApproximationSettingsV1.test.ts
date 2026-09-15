import { describe, expect, it } from 'vitest'
import { DOCX_APPROXIMATE_AUTHORING_SETTINGS, DOCX_APPROXIMATE_LEGACY_COMPAT_OPTIONS, nativeApproximationSettingReason, validNativeDocxApproximatedSettingV1 } from './nativeApproximationSettingsV1.js'

describe('bounded current-layout settings facts', () => {
  it('retains exact known locale facts and rejects unknown or malformed alternatives', () => {
    const fact = { kind: 'themeFontLang', path: '/w:settings[1]/w:themeFontLang[1]', values: { val: 'fr-FR', eastAsia: 'x-none' } }
    expect(validNativeDocxApproximatedSettingV1(fact)).toBe(true)
    expect(validNativeDocxApproximatedSettingV1({ ...fact, values: { val: 'en-CA', eastAsia: '', bidi: '' } })).toBe(true)
    expect(validNativeDocxApproximatedSettingV1({ ...fact, values: { val: 'en-US', eastAsia: 'ja-JP' } })).toBe(true)
    for (const invalid of [{ ...fact, values: { val: 'en_US' } }, { ...fact, values: { val: '' } }, { ...fact, path: '/w:settings[1]/w:themeFontLang[2]' }, { ...fact, values: { val: 'fr-FR', unknown: '1' } }, { ...fact, values: { eastAsia: 'x-none' } }, { ...fact, ignored: true }]) expect(validNativeDocxApproximatedSettingV1(invalid)).toBe(false)
  })
  it('bounds compatibility values and shape identifiers without guessing', () => {
    expect(validNativeDocxApproximatedSettingV1({ kind: 'enableOpenTypeFeatures', path: '/w:settings[1]/w:compat[1]/w:compatSetting[2]', values: { val: '1' } })).toBe(true)
    expect(validNativeDocxApproximatedSettingV1({ kind: 'differentiateMultirowTableHeaders', path: '/w:settings[1]/w:compat[1]/w:compatSetting[5]', values: { val: '1' } })).toBe(true)
    expect(validNativeDocxApproximatedSettingV1({ kind: 'useWord2013TrackBottomHyphenation', path: '/w:settings[1]/w:compat[1]/w:compatSetting[6]', values: { val: '0' } })).toBe(true)
    expect(validNativeDocxApproximatedSettingV1({ kind: 'enableOpenTypeFeatures', path: '/w:settings[1]/w:compat[1]/w:compatSetting[2]', values: { val: '0' } })).toBe(false)
    expect(validNativeDocxApproximatedSettingV1({ kind: 'useWord2013TrackBottomHyphenation', path: '/w:settings[1]/w:compat[1]/w:compatSetting[6]', values: { val: 'yes' } })).toBe(false)
    for (const spidmax of ['01026', '-1', '2147483648', 'oops']) expect(validNativeDocxApproximatedSettingV1({ kind: 'shapeDefaults', path: '/w:settings[1]/w:shapeDefaults[1]', values: { spidmax, idmap: '1' } })).toBe(false)
    expect(validNativeDocxApproximatedSettingV1({ kind: 'shapeDefaults', path: '/w:settings[1]/w:shapeDefaults[1]', values: { spidmax: '1026', idmap: '1' } })).toBe(true)
  })
  it('requires exact complete math defaults and the namespace-derived source path', () => {
    const fact = { kind: 'mathPr', path: '/w:settings[1]/nsf4b2b884:mathPr[1]', values: { mathFont: 'Cambria Math', brkBin: 'before', brkBinSub: '--', lMargin: '0', rMargin: '0', defJc: 'centerGroup', wrapIndent: '1440', intLim: 'subSup', naryLim: 'undOvr', smallFrac: 'off', dispDef: 'true' } }
    expect(validNativeDocxApproximatedSettingV1(fact)).toBe(true)
    expect(validNativeDocxApproximatedSettingV1({ ...fact, path: '/w:settings[1]/m:mathPr[1]' })).toBe(false)
    expect(validNativeDocxApproximatedSettingV1({ ...fact, values: { ...fact.values, wrapIndent: '999' } })).toBe(false)
    expect(validNativeDocxApproximatedSettingV1({ ...fact, values: { ...fact.values, smallFrac: 'maybe' } })).toBe(false)
  })
  it('records active hyphenation and repeated or non-leading compatibility modes at their diagnosed paths', () => {
    const hyphenation = { kind: 'autoHyphenation' as const, path: '/w:settings[1]/w:autoHyphenation[1]', values: { val: 'true', hyphenationZone: '360', consecutiveHyphenLimit: '2', doNotHyphenateCaps: '1' } }
    expect(validNativeDocxApproximatedSettingV1(hyphenation)).toBe(true)
    expect(validNativeDocxApproximatedSettingV1({ ...hyphenation, values: {} })).toBe(true)
    expect(validNativeDocxApproximatedSettingV1({ ...hyphenation, values: { val: 'sometimes' } })).toBe(false)
    expect(validNativeDocxApproximatedSettingV1({ ...hyphenation, values: { hyphenationZone: '-1' } })).toBe(false)
    expect(validNativeDocxApproximatedSettingV1({ ...hyphenation, path: '/w:settings[1]/w:autoHyphenation[2]' })).toBe(false)
    expect(nativeApproximationSettingReason(hyphenation)).toContain('automatic hyphenation is not performed')
    const mode = { kind: 'compatibilityMode', path: '/w:settings[1]/w:compat[1]/w:compatSetting[3]', values: { val: '14' } }
    expect(validNativeDocxApproximatedSettingV1(mode)).toBe(true)
    expect(validNativeDocxApproximatedSettingV1({ ...mode, path: '/w:settings[1]/w:compat[1]/w:compatSetting[1]' })).toBe(false)
    expect(validNativeDocxApproximatedSettingV1({ ...mode, values: { val: '13' } })).toBe(false)
  })
  it('records ECMA-376 legacy compat options as not applied and rejects unknown compat markup', () => {
    expect(DOCX_APPROXIMATE_LEGACY_COMPAT_OPTIONS).toContain('useFELayout')
    const option = { kind: 'useFELayout' as const, path: '/w:settings[1]/w:compat[1]/w:useFELayout[1]', values: {} }
    expect(validNativeDocxApproximatedSettingV1(option)).toBe(true)
    expect(validNativeDocxApproximatedSettingV1({ ...option, values: { val: 'off' } })).toBe(true)
    expect(validNativeDocxApproximatedSettingV1({ ...option, values: { val: 'maybe' } })).toBe(false)
    expect(validNativeDocxApproximatedSettingV1({ ...option, path: '/w:settings[1]/w:compat[1]/w:useFELayout[2]' })).toBe(false)
    expect(validNativeDocxApproximatedSettingV1({ kind: 'notACompatOption', path: '/w:settings[1]/w:compat[1]/w:notACompatOption[1]', values: {} })).toBe(false)
    expect(nativeApproximationSettingReason(option)).toBe('Current-layout approximation records legacy compatibility option useFELayout at /w:settings[1]/w:compat[1]/w:useFELayout[1] as not applied; Word compatibility layout is not emulated')
  })
  it('groups authoring-only settings and agreeing duplicates by member path anchored at a member', () => {
    expect(DOCX_APPROXIMATE_AUTHORING_SETTINGS).toContain('attachedTemplate')
    const authoring = { kind: 'authoringSettings' as const, path: '/w:settings[1]/w:removePersonalInformation[1]', values: { '/w:settings[1]/w:removePersonalInformation[1]': '', '/w:settings[1]/w:attachedTemplate[1]': 'r:id="rId1"', '/w:settings[1]/w:activeWritingStyle[1]': 'w:appName="MSWord" w:lang="en-US"' } }
    expect(validNativeDocxApproximatedSettingV1(authoring)).toBe(true)
    expect(nativeApproximationSettingReason(authoring)).toBe('Current-layout approximation records 3 authoring-only settings anchored at /w:settings[1]/w:removePersonalInformation[1] as not applied; they are editor state that current layout does not consume')
    expect(validNativeDocxApproximatedSettingV1({ ...authoring, path: '/w:settings[1]/w:linkStyles[1]' })).toBe(false)
    expect(validNativeDocxApproximatedSettingV1({ ...authoring, values: { ...authoring.values, '/w:settings[1]/w:mirrorMargins[1]': '' } })).toBe(false)
    expect(validNativeDocxApproximatedSettingV1({ ...authoring, values: { ...authoring.values, '/w:settings[1]/w:compat[1]/w:useFELayout[1]': '' } })).toBe(false)
    expect(validNativeDocxApproximatedSettingV1({ ...authoring, values: {} })).toBe(false)
    expect(validNativeDocxApproximatedSettingV1({ ...authoring, values: { ...authoring.values, '/w:settings[1]/w:attachedTemplate[1]': 'x'.repeat(513) } })).toBe(false)
    const duplicates = { kind: 'duplicateSettings' as const, path: '/w:settings[1]/w:activeWritingStyle[2]', values: { '/w:settings[1]/w:activeWritingStyle[2]': 'w:appName="MSWord"', '/w:settings[1]/w:rsids[1]/w:rsidRoot[2]': 'w:val="00C1654B"' } }
    expect(validNativeDocxApproximatedSettingV1(duplicates)).toBe(true)
    expect(nativeApproximationSettingReason(duplicates)).toContain('records 2 duplicate settings anchored at /w:settings[1]/w:activeWritingStyle[2]')
    expect(validNativeDocxApproximatedSettingV1({ ...duplicates, values: { '/w:settings[1]/w:activeWritingStyle[1]': '' }, path: '/w:settings[1]/w:activeWritingStyle[1]' })).toBe(false)
    expect(validNativeDocxApproximatedSettingV1({ ...duplicates, values: { '/w:settings[1]/w:compat[2]': '' }, path: '/w:settings[1]/w:compat[2]' })).toBe(false)
  })
})

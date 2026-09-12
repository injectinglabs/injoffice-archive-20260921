import { describe, expect, it } from 'vitest'
import { validNativeDocxApproximatedSettingV1 } from './nativeApproximationSettingsV1.js'

describe('bounded current-layout settings facts', () => {
  it('retains exact known locale facts and rejects unknown or malformed alternatives', () => {
    const fact = { kind: 'themeFontLang', path: '/w:settings[1]/w:themeFontLang[1]', values: { val: 'fr-FR', eastAsia: 'x-none' } }
    expect(validNativeDocxApproximatedSettingV1(fact)).toBe(true)
    for (const invalid of [{ ...fact, values: { val: 'ar-SA' } }, { ...fact, values: { val: 'en_US' } }, { ...fact, path: '/w:settings[1]/w:themeFontLang[2]' }, { ...fact, values: { val: 'fr-FR', unknown: '1' } }, { ...fact, ignored: true }]) expect(validNativeDocxApproximatedSettingV1(invalid)).toBe(false)
  })
  it('bounds compatibility values and shape identifiers without guessing', () => {
    expect(validNativeDocxApproximatedSettingV1({ kind: 'enableOpenTypeFeatures', path: '/w:settings[1]/w:compat[1]/w:compatSetting[2]', values: { val: '1' } })).toBe(true)
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
})

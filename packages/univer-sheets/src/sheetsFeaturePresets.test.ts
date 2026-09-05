import { beforeAll, describe, expect, it, vi } from 'vitest'

type PresetsModule = typeof import('./sheetsFeaturePresets')
let presetsModule: PresetsModule

beforeAll(async () => {
  vi.stubGlobal('Path2D', class Path2D {})
  presetsModule = await import('./sheetsFeaturePresets')
})

function pluginNames(features?: import('./sheetsFeaturePresets').SheetsFeatureConfig) {
  const bundle = presetsModule.createSheetsPresetBundle({ container: 'test', features })
  return bundle.presets.flatMap((preset) => preset.plugins.map((registration) => {
    const plugin = Array.isArray(registration) ? registration[0] : registration
    return plugin.pluginName
  }))
}

describe('complete Univer OSS Sheets feature bundle', () => {
  it('enables every optional feature by default', () => {
    const { DEFAULT_SHEETS_FEATURES, SHEETS_FEATURE_KEYS } = presetsModule
    expect(Object.keys(DEFAULT_SHEETS_FEATURES)).toEqual(SHEETS_FEATURE_KEYS)
    expect(Object.values(DEFAULT_SHEETS_FEATURES).every(Boolean)).toBe(true)
    expect(pluginNames()).toEqual(expect.arrayContaining([
      'SHEET_CONDITIONAL_FORMATTING_PLUGIN',
      'SHEET_DATA_VALIDATION_PLUGIN',
      'SHEET_DRAWING_PLUGIN',
      'SHEET_FILTER_PLUGIN',
      'SHEET_FIND_REPLACE_PLUGIN',
      'SHEET_HYPER_LINK_PLUGIN',
      'SHEET_NOTE_PLUGIN',
      'SHEET_SORT_PLUGIN',
      'SHEET_TABLE_PLUGIN',
      'SHEET_THREAD_COMMENT_BASE_PLUGIN',
    ]))
  })

  it('omits disabled features and their UI plugins', () => {
    const names = pluginNames(Object.fromEntries(
      presetsModule.SHEETS_FEATURE_KEYS.map((key) => [key, false]),
    ))

    expect(names).toContain('SHEET_PLUGIN')
    expect(names.some((name) => /CONDITIONAL|VALIDATION|DRAWING|FILTER|FIND_REPLACE|HYPER_LINK|NOTE|SORT|TABLE|COMMENT/.test(name))).toBe(false)
  })
})

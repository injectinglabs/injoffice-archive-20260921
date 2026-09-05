import { mergeLocales } from '@univerjs/core'
import { UniverSheetsConditionalFormattingPreset } from '@univerjs/preset-sheets-conditional-formatting'
import conditionalFormattingEnUS from '@univerjs/preset-sheets-conditional-formatting/locales/en-US'
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import sheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US'
import { UniverSheetsDataValidationPreset } from '@univerjs/preset-sheets-data-validation'
import dataValidationEnUS from '@univerjs/preset-sheets-data-validation/locales/en-US'
import { UniverSheetsDrawingPreset } from '@univerjs/preset-sheets-drawing'
import sheetsDrawingEnUS from '@univerjs/preset-sheets-drawing/locales/en-US'
import { UniverSheetsFilterPreset } from '@univerjs/preset-sheets-filter'
import filterEnUS from '@univerjs/preset-sheets-filter/locales/en-US'
import { UniverSheetsFindReplacePreset } from '@univerjs/preset-sheets-find-replace'
import findReplaceEnUS from '@univerjs/preset-sheets-find-replace/locales/en-US'
import { UniverSheetsHyperLinkPreset } from '@univerjs/preset-sheets-hyper-link'
import hyperLinkEnUS from '@univerjs/preset-sheets-hyper-link/locales/en-US'
import { UniverSheetsNotePreset } from '@univerjs/preset-sheets-note'
import noteEnUS from '@univerjs/preset-sheets-note/locales/en-US'
import { UniverSheetsSortPreset } from '@univerjs/preset-sheets-sort'
import sortEnUS from '@univerjs/preset-sheets-sort/locales/en-US'
import { UniverSheetsTablePreset } from '@univerjs/preset-sheets-table'
import tableEnUS from '@univerjs/preset-sheets-table/locales/en-US'
import { UniverSheetsThreadCommentPreset } from '@univerjs/preset-sheets-thread-comment'
import threadCommentEnUS from '@univerjs/preset-sheets-thread-comment/locales/en-US'
import type { UniverOssPreset } from './createOssUniver'

export const SHEETS_FEATURE_KEYS = [
  'conditionalFormatting',
  'dataValidation',
  'drawing',
  'filter',
  'findReplace',
  'hyperlinks',
  'notes',
  'sort',
  'tables',
  'threadComments',
] as const

export type SheetsFeatureKey = (typeof SHEETS_FEATURE_KEYS)[number]
export type SheetsFeatureConfig = Partial<Record<SheetsFeatureKey, boolean>>

export const DEFAULT_SHEETS_FEATURES: Readonly<Record<SheetsFeatureKey, boolean>> =
  Object.freeze(Object.fromEntries(SHEETS_FEATURE_KEYS.map((key) => [key, true])) as Record<SheetsFeatureKey, boolean>)

type CorePresetConfig = NonNullable<Parameters<typeof UniverSheetsCorePreset>[0]>

export interface SheetsPresetBundleOptions {
  container: CorePresetConfig['container']
  workerURL?: CorePresetConfig['workerURL']
  features?: SheetsFeatureConfig
}

/**
 * Compose the complete Apache-2.0 Univer Sheets surface used by InjOffice.
 *
 * Core editing and formulas are always present. Every optional feature is on
 * by default and can be omitted independently; an omitted plugin also removes
 * its ribbon/context-menu contribution instead of merely disabling the button.
 */
export function createSheetsPresetBundle(options: SheetsPresetBundleOptions): {
  features: Readonly<Record<SheetsFeatureKey, boolean>>
  locale: ReturnType<typeof mergeLocales>
  presets: UniverOssPreset[]
} {
  const features = Object.freeze({ ...DEFAULT_SHEETS_FEATURES, ...options.features })
  const presets: UniverOssPreset[] = [UniverSheetsCorePreset({
    container: options.container,
    ...(options.workerURL === undefined ? {} : { workerURL: options.workerURL }),
  })]
  const locales = [sheetsCoreEnUS]

  const include = (enabled: boolean, preset: UniverOssPreset, locale: Record<string, unknown>) => {
    if (!enabled) return
    presets.push(preset)
    locales.push(locale)
  }

  include(features.conditionalFormatting, UniverSheetsConditionalFormattingPreset(), conditionalFormattingEnUS)
  include(features.dataValidation, UniverSheetsDataValidationPreset(), dataValidationEnUS)
  include(features.drawing, UniverSheetsDrawingPreset(), sheetsDrawingEnUS)
  include(features.filter, UniverSheetsFilterPreset(), filterEnUS)
  include(features.findReplace, UniverSheetsFindReplacePreset(), findReplaceEnUS)
  include(features.hyperlinks, UniverSheetsHyperLinkPreset(), hyperLinkEnUS)
  include(features.notes, UniverSheetsNotePreset(), noteEnUS)
  include(features.sort, UniverSheetsSortPreset(), sortEnUS)
  include(features.tables, UniverSheetsTablePreset(), tableEnUS)
  include(features.threadComments, UniverSheetsThreadCommentPreset(), threadCommentEnUS)

  return { features, locale: mergeLocales(locales), presets }
}

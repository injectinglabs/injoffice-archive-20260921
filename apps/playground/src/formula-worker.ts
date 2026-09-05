import { LocaleType, mergeLocales } from '@univerjs/core'
import { UniverSheetsCoreWorkerPreset } from '@univerjs/preset-sheets-core/worker'
import sheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US'
import { createOssUniver } from '@injoffice/univer-sheets'

// Univer 0.25.1 formula worker: same core preset as the playground shell,
// minus UI. Paired with workerURL on UniverSheetsCorePreset in main.tsx.
createOssUniver({
  locale: LocaleType.EN_US,
  locales: { [LocaleType.EN_US]: mergeLocales(sheetsCoreEnUS) },
  presets: [UniverSheetsCoreWorkerPreset()],
})

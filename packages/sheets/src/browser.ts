// Browser-safe native XLSX contract boundary. The root entry also exports the
// Node-qualified HarfBuzz paint path, which is intentionally not evaluated by
// browser-local extract/apply clients.
export {
  WORKBOOK_MUTATION_PROTOCOL,
  WORKBOOK_MUTATION_VERSION,
  WorkbookMutationValidationError,
  decodeWorkbookMutationBatch,
} from './mutationProtocol.js'

export type {
  StyleDelta,
  SupportedWorkbookMutation,
  WorkbookMutationBatchV1,
} from './mutationProtocol.js'

export { validateNativeWorkbookV1 } from './nativeValidation.js'
export type { NativeWorkbookV1 } from './nativeContract.generated.js'

export {
  NativeWorkbookV2ValidationError,
  assertNativeWorkbookV2,
  decodeNativeWorkbookV2,
  validateNativeWorkbookV2,
} from './nativeValidationV2.js'

export type {
  NativeWorkbookCellV2,
  NativeWorkbookSheetV2,
  NativeWorkbookV2,
} from './nativeContractV2.generated.js'
export { formatNativeSheetCellDisplayV2 } from './nativeCellDisplayV2.js'
export type { NativeSheetCellDisplayFormatResultV2 } from './nativeCellDisplayV2.js'
export { decodeNativeWorkbookObjectsV1, layoutNativeCachedChartV1 } from './nativeObjectsPreviewV1.js'
export { nativeTableFillPreview, nativeTableHeaderTextPreview } from './nativeTableFillPreview.js'
export { nativeTableTotalsTextPreview } from './nativeTableTotalsTextPreview.js'
export { nativeStoredRowPreviewV1 } from './nativeStoredRowsPreviewV1.js'
export type { NativeStoredRowGeometryV1, NativeStoredRowV1 } from './nativeStoredRowsPreviewV1.js'
export { nativeTableNumberFormatPreview,formatNativeAccountingTextPreview } from './nativeTableNumberFormatPreview.js'
export { nativeTableBorderPreview } from './nativeTableBorderPreview.js'
export type { NativeTableBorderEdgesV1,NativeTableBorderSideV1 } from './nativeTableBorderPreview.js'
export type { NativeTableNumberDisplayV1 } from './nativeTableNumberFormatPreview.js'
export type { NativeWorkbookObjectsV1, NativeTablePreviewV1, NativeTableNumberFormatV1, NativeTableBorderPreviewV1, NativeChartPreviewV1, NativeChartSeriesPreviewV1, NativeCachedChartMarkV1 } from './nativeObjectsPreviewV1.js'

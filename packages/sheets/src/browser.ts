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

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
  SupportedWorkbookMutation,
  WorkbookMutationBatchV1,
} from './mutationProtocol.js'

export {
  NativeWorkbookV2ValidationError,
  assertNativeWorkbookV2,
  decodeNativeWorkbookV2,
} from './nativeValidationV2.js'

export type {
  NativeWorkbookCellV2,
  NativeWorkbookSheetV2,
  NativeWorkbookV2,
} from './nativeContractV2.generated.js'

export { OutlineManager } from './manager'
export { OutlineCommandController, OutlineHandle } from './commands'
export type { OutlineUndoRecord, OutlineUndoSink } from './commands'
export {
  OUTLINE_COLLABORATION_PROTOCOL,
  OutlineCollaborationError,
  OutlineCollaborationSession,
  applyOutlineCollaborationOperation,
  fingerprintOutlineSnapshot,
} from './collaboration'
export type {
  OutlineCollaborationAction,
  OutlineCollaborationApplyContext,
  OutlineCollaborationAuthorizationContext,
  OutlineCollaborationCommandTarget,
  OutlineCollaborationEntry,
  OutlineCollaborationErrorCode,
  OutlineCollaborationEvent,
  OutlineCollaborationOperation,
  OutlineCollaborationOptions,
  OutlineCollaborationResync,
  OutlineCollaborationTransport,
} from './collaboration'
export { OutlineStore } from './store'
export {
  XLSX_MAX_OUTLINE_DEPTH,
  createNativeXlsxOutlineWrite,
  decodeNativeXlsxOutlineSnapshot,
} from './native'
export { transformOutline, transformOutlines } from './transform'
export { applyOutlineLevel, collapseToOutlineLevel, layoutOutlineGutter, outlineGroupDepth } from './gutter'
export type { OutlineGutterControl, OutlineGutterModel } from './gutter'
export {
  MAX_OUTLINE_DEPTH,
  MAX_SHEET_COLUMN_INDEX,
  MAX_SHEET_ROW_INDEX,
  validateOutlineGroup,
  validateOutlineSnapshot,
} from './validation'
export type {
  OutlineAxis,
  OutlineGroup,
  OutlineIssue,
  OutlineManagerOptions,
  OutlineResult,
  OutlineStructuralEdit,
  OutlineVisibilityAdapter,
} from './types'
export type {
  NativeXlsxOutlineBand,
  NativeXlsxOutlineBandWire,
  NativeXlsxOutlineDecodeResult,
  NativeXlsxOutlineGroupWire,
  NativeXlsxOutlineSnapshot,
  NativeXlsxOutlineSnapshotWire,
  NativeXlsxOutlineWriteWire,
} from './native'

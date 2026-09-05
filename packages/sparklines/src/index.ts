export { extractSparklineValues } from './extract'
export { compileSparklineGeometry, numericExtent, sparklineGeometryToSvg } from './geometry'
export type { SparklineCompileInput } from './geometry'
export { SparklineManager } from './manager'
export type { CreateSparklineInput, SparklineManagerOptions } from './manager'
export { SparklineCommandController, SparklineHandle } from './commands'
export {
  SPARKLINE_COLLABORATION_PROTOCOL,
  SparklineCollaborationError,
  SparklineCollaborationSession,
  applySparklineCollaborationOperation,
  fingerprintSparkline,
} from './collaboration'
export type { SparklineUndoRecord, SparklineUndoSink } from './commands'
export type {
  SparklineCollaborationEntry,
  SparklineCollaborationErrorCode,
  SparklineCollaborationEvent,
  SparklineCollaborationOperation,
  SparklineCollaborationOptions,
  SparklineCollaborationTransport,
} from './collaboration'
export { sparklinesFromFile } from './fromFile'
export type { FileSparklineInfo, SparklineFileContext, SparklineFileResult } from './fromFile'
export { toWireSparklines } from './toFile'
export type { SparklineWireContext, SparklineWireResult, WireSparklineGroup, WireSparklineItem } from './toFile'
export { SPARKLINE_TYPES } from './types'
export type {
  EmptyCellBehavior, SparklineBar, SparklineCellRef, SparklineColors,
  SparklineGeometry, SparklineGroup, SparklineOptions, SparklinePath,
  SparklinePoint, SparklineRangeRef, SparklineSnapshotV1, SparklineSpec,
  SparklineType, SparklineValueReader, SparklineViewport,
} from './types'
export { validateSparkline, validateSparklineSnapshot } from './validation'
export type { SparklineValidationIssue } from './validation'

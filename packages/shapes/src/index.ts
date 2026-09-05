export { ShapeManager, SHAPE_COMPONENT_KEY } from './manager'
export { ShapeCommandController, ShapeHandle, isShapeSnapshot } from './commands'
export type { ShapeSnapshotEntry, ShapeSnapshotV1, ShapeUndoRecord, ShapeUndoSink } from './commands'
export {
  SHAPE_COLLABORATION_PROTOCOL,
  ShapeCollaborationError,
  ShapeCollaborationSession,
  applyShapeCollaborationOperation,
  fingerprintShape,
} from './collaboration'
export type {
  ShapeCollaborationEntry,
  ShapeCollaborationErrorCode,
  ShapeCollaborationEvent,
  ShapeCollaborationOperation,
  ShapeCollaborationOptions,
  ShapeCollaborationPatch,
  ShapeCollaborationTransport,
} from './collaboration'
export { ShapeFloat, ShapePreview, shapePreviewFidelity } from './ShapeFloat'
export type { ShapePreviewProps, ShapePreviewFidelity } from './ShapeFloat'
export {
  arcPath, calloutPath, diamondPath, directionalArrowPath, doubleArrowPath, ellipseCalloutPath,
  lineCalloutPath, parallelogramPath, pillPath, plusPath, regularPolygonPath, rightTrianglePath,
  roundRectPath, starPath, trapezoidPath, trianglePath,
} from './geometry'
export { SHAPE_KINDS, SHAPE_CATEGORIES, SHAPE_DEFAULTS, hasText, shapeLabel } from './types'
export type { ShapeCategory } from './types'
export { ShapePanel } from './ShapePanel'
export { specFromFileShape, specsFromFileShapes } from './fromFile'
export type { FileShapeInfo, FileShapeAnchor, FileShapeConversion } from './fromFile'
export { toWireShapes, toWireShapeRemove, toWireShapeUpdate } from './toFile'
export type { WireShape, ToWireShapeInput, ToWireShapeResult, WireShapeRemove, WireShapeUpdate, ShapeLifecycleWireResult } from './toFile'
export type { ShapeSpec, ShapeKind, NativeShapeIdentity } from './types'

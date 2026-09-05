export { PresenceManager, type PresenceOptions, type CommandServiceLike, type ResyncReason } from './manager'
export { PresenceState } from './presence'
export { DocPresenceManager } from './docManager'
export type { DocSelection } from './docManager'
export { DocSyncEngine } from './docSync'
export type { DocSyncHooks } from './docSync'
export { DeckPresenceManager } from './deckManager'
export { DeckSyncEngine, applyDeckOps, diffDeck } from './deckSync'
export type { DeckOp, DeckFieldOp, DeckShapeOverrideOp, DeckSyncHooks, SyncDeck } from './deckSync'
export { PresenceStack, FileChangedBanner, describeFileChange, type PresenceSource } from './PresenceStack'
export { SheetPresenceLabel, type SheetPresenceLabelComponent, type SheetPresenceLabelProps } from './SheetPresenceLabel'
export { parseSelection, sameSelection, selectionA1, rangeA1, colName, initials, withAlpha } from './selection'
export { shouldShipMutation, reconcileRemote, rebaseUnit, cellsOf, withoutCells, cloneParams, pruneOwned, SubmitQueue, MOVE_RANGE, REORDER_RANGE, SET_RANGE_VALUES } from './sync'
export type { OpRecord, OpEntry, LogState, MutationInfo, ExecOptions, LocalOwned } from './sync'
export { parseStruct, parseMove, mapIndex, mapSpan, mapIndexThroughMove, mapSpanThroughMove, transformOp, transformOps, transformSelection, transformSelectionThroughMove, structEditsOf } from './transform'
export type { StructEdit, MoveEdit, Axis, TransformResult } from './transform'
export { CollaborativeUndoManager } from './undo'
export type { CollaborativeUndoEntry, UndoBlockCode, UndoBlockReason, UndoPlan, UndoRebaseOptions, UndoRebaseReport } from './undo'
export { CollaborativeUndoJournalCoordinator } from './undoJournal'
export type {
  CollaborativeUndoExecutionContext,
  CollaborativeUndoExecutionOrigin,
  CollaborativeUndoJournalBlockCode,
  CollaborativeUndoJournalHost,
  CollaborativeUndoJournalOptions,
  CollaborativeUndoJournalPlanResult,
  CollaborativeUndoJournalReceipt,
  CollaborativeUndoJournalState,
  CollaborativeUndoRemoteEntry,
  CollaborativeUndoRemoteResult,
} from './undoJournal'
export { COLLAB_EVENTS } from './types'
export type { CollabEvent, CollabTransport, DeckSelection, FileChange, JoinResult, LogStateWire, PeerInfo, SheetSelection } from './types'
export { PdfPresenceManager } from './pdfManager'
export type { PdfSelection } from './pdfManager'
export {
  COLLAB_CAPABILITIES,
  CollabPermissionManager,
  PermissionDeniedError,
  validatePermissionSnapshot,
} from './permissions'
export type {
  CollabCapability,
  CollabPermissionRule,
  CollabPermissionSnapshot,
  CollabRole,
  CollabRoleChange,
  PermissionDecision,
  PermissionRejectCode,
  PermissionUpdateResult,
} from './permissions'
export { LiveShareSession, validateLiveShareEvent, validateLiveShareViewport } from './liveShare'
export type {
  LiveShareEvent,
  LiveShareIntent,
  LiveShareMode,
  LiveShareOptions,
  LiveShareRejectReason,
  LiveShareState,
  LiveShareTransport,
  LiveShareViewport,
  LiveShareViewportAdapter,
} from './liveShare'
export {
  DurableOutboundJournal,
  OUTBOUND_JOURNAL_PROTOCOL,
  OutboundJournalCancelledError,
  OutboundJournalOpenError,
  createJournalSubmitter,
} from './offlineJournal'
export type {
  JournalRebasedEntry,
  JournalResyncRequest,
  JournalResyncResult,
  JournalSubmitAck,
  JournalSubmitRequest,
  OutboundJournalEntry,
  OutboundJournalEnqueueOptions,
  OutboundJournalHooks,
  OutboundJournalOptions,
  OutboundJournalReceipt,
  OutboundJournalState,
  OutboundJournalStatus,
  OutboundJournalStorage,
  OutboundJournalTransport,
  OutboundJournalSnapshotV1,
  OutboundRetryPolicy,
  RetryScheduler,
} from './offlineJournal'

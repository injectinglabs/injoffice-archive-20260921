export { ConnectorManager, isConnectorManagerSnapshot, isConnectorRangeSnapshot } from './manager'
export type {
  ConnectorManagerOptions,
  ConnectorAuthoritativeApplyOptions,
  ConnectorManagerSnapshotV1,
  ConnectorMountedSnapshot,
  ConnectorRangeSnapshot,
  ConnectorRefreshTransaction,
} from './manager'
export { ConnectorCommandController, ConnectorHandle, isCredentialFreeConnectorSpec } from './commands'
export type {
  ConnectorCommandControllerOptions,
  ConnectorCommandSnapshotV1,
  ConnectorRestoreOptions,
  ConnectorUndoRecord,
  ConnectorUndoSink,
} from './commands'
export {
  CONNECTOR_COLLABORATION_PROTOCOL,
  ConnectorCollaborationError,
  ConnectorCollaborationSession,
  ConnectorCollaborationTransportError,
  applyConnectorCollaborationOperation,
  fingerprintConnector,
  fingerprintConnectorRange,
  verifyConnectorCollaborationCellPatch,
} from './collaboration'
export type {
  ConnectorCollaborationApplyContext,
  ConnectorCollaborationCellPatch,
  ConnectorCollaborationEntry,
  ConnectorCollaborationErrorCode,
  ConnectorCollaborationEvent,
  ConnectorCollaborationOperation,
  ConnectorCollaborationOptions,
  ConnectorCollaborationRange,
  ConnectorCollaborationResync,
  ConnectorCollaborationTransport,
  ConnectorCollaborationTransportErrorCode,
  ConnectorCollaborativeRefreshRequest,
  ConnectorLifecycleCollaborationOperation,
  ConnectorRefreshCollaborationOperation,
} from './collaboration'
export { DataPanel } from './DataPanel'
export { jsonToGrid, csvToGrid, resolvePath } from './normalize'
export {
  RANGE_PREPROCESS_MANIFEST_VERSION,
  RangePreprocessError,
  RangePreprocessPipeline,
  fingerprintRangePreprocessStages,
  isRangePreprocessManifest,
} from './preprocess'
export {
  CONNECTOR_NATIVE_EXTENSION_NAMESPACE,
  CONNECTOR_NATIVE_EXTENSION_VERSION,
  ConnectorNativePersistence,
  ConnectorNativePersistenceError,
  fromNativeConnectorDefinitions,
  hydrateConnectorManagerFromNative,
  persistConnectorManagerToNative,
  toNativeConnectorDefinitions,
} from './native'
export type {
  ConnectorNativeManager,
  ConnectorNativePersistenceCodec,
  ConnectorNativePersistenceErrorCode,
  NativeConnectorDefinitionV1,
} from './native'
export type {
  ConnectorAuthorizationRequest,
  ConnectorAuthorizer,
  ConnectorCachePolicy,
  ConnectorCellType,
  ConnectorColumnSchema,
  ConnectorRefreshOptions,
  ConnectorRefreshResult,
  ConnectorSchedule,
  ConnectorScheduler,
  ConnectorSchema,
  ConnectorSpec,
  ConnectorSource,
  ConnectorStatus,
  RefreshPolicy,
  SourceFetcher,
  SourceFetchContext,
} from './types'
export type {
  RangeGrid,
  RangePreprocessContext,
  RangePreprocessErrorCode,
  RangePreprocessEvent,
  RangePreprocessManifestV1,
  RangePreprocessResult,
  RangePreprocessRunOptions,
  RangePreprocessStage,
  RangePreprocessStageContext,
  RangePreprocessStageDescriptor,
  RangeValue,
  RangeWindow,
} from './preprocess'

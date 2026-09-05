// Data-connector model for @injoffice/connectors (Phase 6).
//
// The core idea is the BOUND RANGE: a named region of the workbook attached
// to a live source with a refresh policy — a living connection, not a
// one-shot paste. ConnectorSpec is plain JSON (the ChartSpec contract
// pattern): agents create them over RPC, snapshots persist them, and the
// engine NEVER holds credentials — refresh data arrives through a
// host-provided fetcher, which in the product is a gateway endpoint running
// under the platform's existing per-user integration layer.

/** Where the data comes from. A discriminated union that grows: 'http' ships
 *  first; integration-backed kinds (composio toolkits, enrichment-DAG runs,
 *  other spreadsheets) follow as the gateway fetcher learns them. */
export type ConnectorSource =
  | {
      kind: 'http'
      url: string
      format: 'json' | 'csv'
      /** Dot-path into the JSON payload to the array/object to tabulate
       *  (e.g. "data.items"). Ignored for csv. */
      path?: string
    }

export type RefreshPolicy = 'manual' | 'onOpen' | 'interval'

export type ConnectorCellType = 'string' | 'number' | 'boolean' | 'any'

export interface ConnectorColumnSchema {
  index: number
  type: ConnectorCellType
  nullable?: boolean
}

export interface ConnectorSchema {
  columns: ConnectorColumnSchema[]
  allowAdditionalColumns?: boolean
}

export interface ConnectorCachePolicy {
  mode: 'none' | 'memory'
  /** A zero TTL keeps a value only for concurrent refresh coalescing. */
  ttlMs: number
}

export interface ConnectorSchedule {
  intervalMs: number
}

export interface ConnectorSpec {
  id: string
  name: string
  source: ConnectorSource
  /** Top-left corner where fetched data lands. */
  target: { sheetId: string; startRow: number; startColumn: number }
  refresh: RefreshPolicy
  schedule?: ConnectorSchedule
  cache?: ConnectorCachePolicy
  schema?: ConnectorSchema
}

export interface ConnectorStatus {
  lastRefreshTs?: number
  lastError?: string
  rowCount?: number
  columnCount?: number
  refreshing?: boolean
  nextRefreshTs?: number
  cacheHit?: boolean
  lastErrorCode?: 'ABORTED' | 'AUTHORIZATION_DENIED' | 'FETCH_FAILED' | 'INVALID_DATA' | 'PREPROCESS_FAILED' | 'TARGET_MISSING'
}

export interface SourceFetchContext {
  connectorId: string
  reason: 'manual' | 'onOpen' | 'interval'
  signal: AbortSignal
}

/** The host-provided fetcher: source in, values grid out. Credentials remain
 *  host-side. Implementations should observe the abort signal. */
export type SourceFetcher = (source: ConnectorSource, context: Readonly<SourceFetchContext>) => Promise<unknown[][]>

export interface ConnectorAuthorizationRequest {
  action: 'refresh'
  connector: Readonly<ConnectorSpec>
  reason: SourceFetchContext['reason']
}

export type ConnectorAuthorizer = (request: Readonly<ConnectorAuthorizationRequest>) => boolean | Promise<boolean>

export interface ConnectorScheduler {
  setInterval(callback: () => void, intervalMs: number): unknown
  clearInterval(handle: unknown): void
}

export interface ConnectorRefreshOptions {
  reason?: SourceFetchContext['reason']
  signal?: AbortSignal
  force?: boolean
  revision?: string
  mode?: 'local' | 'collaborative'
  expectedPreprocessFingerprint?: string
}

export type ConnectorRefreshResult = 'applied' | 'canceled' | 'failed' | 'missing' | 'skipped'

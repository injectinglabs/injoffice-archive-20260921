export type OutlineAxis = 'row' | 'column'

/** A zero-based, inclusive row or column outline group. */
export interface OutlineGroup {
  id: string
  sheetId: string
  axis: OutlineAxis
  start: number
  end: number
  collapsed: boolean
}

export interface OutlineIssue {
  code: 'INVALID_ID' | 'INVALID_SHEET' | 'INVALID_AXIS' | 'INVALID_RANGE' | 'INVALID_COLLAPSED' | 'DUPLICATE' | 'CROSSING' | 'MAX_DEPTH' | 'AUTHORITY_REQUIRED' | 'APPLY_FAILED'
  message: string
  groupId?: string
}

export type OutlineResult<T> =
  | { ok: true; value: T }
  | { ok: false; issues: OutlineIssue[] }

export interface OutlineStructuralEdit {
  kind: 'insert' | 'remove'
  axis: OutlineAxis
  sheetId: string
  start: number
  count: number
}

export interface OutlineVisibilityAdapter {
  hide(sheetId: string, axis: OutlineAxis, start: number, count: number): void
  show(sheetId: string, axis: OutlineAxis, start: number, count: number): void
}

export interface OutlineManagerOptions {
  /** Collaboration mode rejects ordinary local mutations. An acknowledged
   * collaboration session applies complete snapshots through its internal
   * authoritative path. */
  mutationAuthority?: 'local' | 'collaboration'
}

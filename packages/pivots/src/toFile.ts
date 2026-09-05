// Save-time OOXML serialization (P10): turn PivotSpecs into the wire shape
// the gateway's xlsxpatch.AddPivot consumes. The baked grid is already in
// the cells (the manager writes real values); this adds the NATIVE pivot
// parts (pivotCache + pivotTable, refreshOnLoad=1) so Excel opens a live,
// editable pivot instead of a frozen block.

import { assessPivotRepresentability } from './representability'
import type { CellRangeRef, NativePivotIdentity, PivotFieldMember, PivotSpec } from './types'

/** Mirror of the gateway's wirePivot (files_apply_specs.go). */
export interface WirePivot {
  sourceSheetName: string
  sourceRef: string
  fields: string[]
  targetSheetName: string
  targetCellRef: string
  rowFields?: string[]
  colFields?: string[]
  dataFields: Array<{ field: string; agg: string }>
  pageFields?: Array<{ field: string; selectedItem?: string }>
  memberFilters?: Array<{ field: string; excludedItems: string[] }>
  sorts?: Array<{ field: string; direction: 'ascending' | 'descending' }>
  fieldMembers?: Array<{ field: string; items: PivotFieldMember[] }>
  grandTotals?: boolean
  name?: string
}

export function colName(c: number): string {
  let s = ''
  let n = c + 1
  while (n > 0) {
    const rem = (n - 1) % 26
    s = String.fromCharCode(65 + rem) + s
    n = Math.floor((n - 1) / 26)
  }
  return s
}

function a1(row: number, col: number): string {
  return `${colName(col)}${row + 1}`
}

function rangeA1(r: CellRangeRef): string {
  return `${a1(r.startRow, r.startColumn)}:${a1(r.endRow, r.endColumn)}`
}

export interface PivotWireContext {
  sheetNameOf: (sheetId: string) => string | null
  /** The source block's header row values (field names, in column order). */
  headerRowOf: (source: CellRangeRef) => unknown[]
  /** Complete distinct member inventory for fields carrying native state. */
  fieldMembersOf?: (source: CellRangeRef, field: string) => readonly PivotFieldMember[] | null
}

export interface PivotWireResult {
  pivots: WirePivot[]
  skipped: string[]
}

export interface WirePivotRemove {
  operation: 'remove'
  identity: NativePivotIdentity
}

export interface WirePivotUpdate {
  operation: 'update'
  identity: NativePivotIdentity
  pivot: WirePivot
}

export interface PivotLifecycleWireResult<T extends WirePivotRemove | WirePivotUpdate> {
  request?: T
  skipped: string[]
}

/** Build the fail-closed wire request for removing a hydrated native pivot. */
export function toWirePivotRemove(spec: Pick<PivotSpec, 'id' | 'nativeIdentity'>): PivotLifecycleWireResult<WirePivotRemove> {
  if (!spec.nativeIdentity?.part) {
    return { skipped: [`pivot ${spec.id}: it has no hydrated native identity`] }
  }
  return {
    request: { operation: 'remove', identity: { ...spec.nativeIdentity } },
    skipped: [],
  }
}

/** Build a one-object native update. Conversion reuses the add serializer,
 * but refuses to fall back to an add when stable native identity is absent. */
export function toWirePivotUpdate(spec: PivotSpec, ctx: PivotWireContext): PivotLifecycleWireResult<WirePivotUpdate> {
  if (!spec.nativeIdentity?.part) {
    return { skipped: [`pivot ${spec.id}: it has no hydrated native identity`] }
  }
  const converted = toWirePivots([spec], ctx)
  if (converted.pivots.length !== 1) return { skipped: converted.skipped }
  return {
    request: {
      operation: 'update',
      identity: { ...spec.nativeIdentity },
      pivot: converted.pivots[0],
    },
    skipped: [],
  }
}

export function toWirePivots(specs: PivotSpec[], ctx: PivotWireContext): PivotWireResult {
  const pivots: WirePivot[] = []
  const skipped: string[] = []
  for (const spec of specs) {
    const label = spec.id
    const sourceSheet = ctx.sheetNameOf(spec.source.sheetId)
    const targetSheet = ctx.sheetNameOf(spec.target.sheetId)
    if (!sourceSheet || !targetSheet) {
      skipped.push(`pivot ${label}: its sheet no longer exists`)
      continue
    }
    const header = ctx.headerRowOf(spec.source)
    const fields = header.map((h, i) => {
      const s = h == null ? '' : String(h).trim()
      return s === '' ? `Column ${i + 1}` : s
    })
    if (fields.length === 0) {
      skipped.push(`pivot ${label}: its source has no header row`)
      continue
    }
    if (spec.values.length === 0) {
      skipped.push(`pivot ${label}: it has no value fields`)
      continue
    }
    const members = new Map<string, readonly PivotFieldMember[]>()
    const assessment = assessPivotRepresentability(spec, {
      fields,
      membersOf: ctx.fieldMembersOf
        ? (field) => {
            const found = ctx.fieldMembersOf!(spec.source, field)
            if (found) members.set(field, found)
            return found
          }
        : undefined,
    })
    if (!assessment.representable) {
      const reasons = assessment.issues.filter(({ severity }) => severity === 'unsupported').map(({ code, message }) => `${code}: ${message}`)
      skipped.push(`pivot ${label}: ${reasons.join('; ')}`)
      continue
    }
    const nativeFilters = new Map<string, Set<string>>()
    for (const [field, included] of Object.entries(spec.filters ?? {})) {
      const allowed = new Set(included)
      nativeFilters.set(field, new Set((members.get(field) ?? []).map(({ value }) => value).filter((value) => !allowed.has(value))))
    }
    for (const filter of spec.memberFilters ?? []) {
      const excluded = filter.mode === 'exclude'
        ? filter.values
        : (members.get(filter.field) ?? []).map(({ value }) => value).filter((value) => !new Set(filter.values).has(value))
      nativeFilters.set(filter.field, new Set(excluded))
    }
    const memberFields = new Set([
      ...nativeFilters.keys(),
      ...(spec.pageFields ?? []).filter(({ selectedItem }) => selectedItem !== undefined).map(({ field }) => field),
    ])
    const pivot: WirePivot = {
      sourceSheetName: sourceSheet,
      sourceRef: rangeA1(spec.source),
      fields,
      targetSheetName: targetSheet,
      targetCellRef: a1(spec.target.startRow, spec.target.startColumn),
      rowFields: spec.rows.length ? spec.rows : undefined,
      colFields: spec.columns.length ? spec.columns : undefined,
      dataFields: spec.values.map((v) => ({ field: v.field, agg: v.agg })),
      name: `InjOffice_${spec.id.replace(/[^A-Za-z0-9]+/g, '_')}`,
    }
    if (spec.pageFields?.length) pivot.pageFields = spec.pageFields.map((page) => ({ ...page }))
    if (nativeFilters.size) pivot.memberFilters = [...nativeFilters].map(([field, values]) => ({ field, excludedItems: [...values] }))
    if (spec.sorts?.length) pivot.sorts = spec.sorts.map((sort) => ({ ...sort }))
    if (memberFields.size) pivot.fieldMembers = [...memberFields].map((field) => ({ field, items: [...(members.get(field) ?? [])] }))
    if (spec.grandTotals === false) pivot.grandTotals = false
    pivots.push(pivot)
  }
  return { pivots, skipped }
}

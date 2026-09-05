import type { PivotFieldMember, PivotSpec } from './types.js'

export type PivotRepresentabilitySeverity = 'warning' | 'unsupported'

export type PivotRepresentabilityCode =
  | 'UNKNOWN_FIELD'
  | 'DUPLICATE_FIELD_CONFIG'
  | 'MULTIPLE_COLUMN_FIELDS'
  | 'PAGE_AXIS_CONFLICT'
  | 'PAGE_FILTER_CONFLICT'
  | 'SORT_AXIS_UNSUPPORTED'
  | 'MEMBER_INVENTORY_REQUIRED'
  | 'UNKNOWN_MEMBER'
  | 'DUPLICATE_MEMBER'
  | 'LEGACY_FILTER_TRANSLATED'
  | 'UNSUPPORTED_MEMBER_TYPE'

export interface PivotRepresentabilityIssue {
  severity: PivotRepresentabilitySeverity
  code: PivotRepresentabilityCode
  message: string
  field?: string
}

export interface PivotRepresentabilityResult {
  representable: boolean
  issues: PivotRepresentabilityIssue[]
}

export interface PivotRepresentabilityContext {
  fields: readonly string[]
  membersOf?: (field: string) => readonly PivotFieldMember[] | null
}

/** Assess the bounded native page/member-filter/label-sort subset. */
export function assessPivotRepresentability(
  spec: Pick<PivotSpec, 'rows' | 'columns' | 'values' | 'filters' | 'memberFilters' | 'pageFields' | 'sorts'>,
  context: PivotRepresentabilityContext,
): PivotRepresentabilityResult {
  const issues: PivotRepresentabilityIssue[] = []
  const fields = new Set(context.fields)
  const configured = [
    ...spec.rows,
    ...spec.columns,
    ...spec.values.map(({ field }) => field),
    ...Object.keys(spec.filters ?? {}),
    ...(spec.memberFilters ?? []).map(({ field }) => field),
    ...(spec.pageFields ?? []).map(({ field }) => field),
    ...(spec.sorts ?? []).map(({ field }) => field),
  ]
  for (const field of new Set(configured)) {
    if (!fields.has(field)) issues.push(unsupported('UNKNOWN_FIELD', `Field ${field} is not present in the source header.`, field))
  }
  if (spec.columns.length > 1) issues.push(unsupported('MULTIPLE_COLUMN_FIELDS', 'Native conversion supports at most one column field.'))
  duplicateIssues(spec.memberFilters ?? [], 'member filter', issues)
  duplicateIssues(spec.pageFields ?? [], 'page field', issues)
  duplicateIssues(spec.sorts ?? [], 'sort', issues)
  const axes = new Set([...spec.rows, ...spec.columns])
  for (const page of spec.pageFields ?? []) {
    if (axes.has(page.field)) issues.push(unsupported('PAGE_AXIS_CONFLICT', `Page field ${page.field} is also a row or column field.`, page.field))
    if (page.selectedItem !== undefined && (spec.filters?.[page.field] || (spec.memberFilters ?? []).some(({ field }) => field === page.field))) {
      issues.push(unsupported('PAGE_FILTER_CONFLICT', `Page field ${page.field} also has a member filter.`, page.field))
    }
  }
  for (const filter of spec.memberFilters ?? []) if (spec.filters?.[filter.field]) {
    issues.push(unsupported('DUPLICATE_FIELD_CONFIG', `Field ${filter.field} has both legacy and explicit member filters.`, filter.field))
  }
  for (const sort of spec.sorts ?? []) {
    if (!axes.has(sort.field)) issues.push(unsupported('SORT_AXIS_UNSUPPORTED', `Sort field ${sort.field} is not a row or column label field.`, sort.field))
  }
  const memberFields = new Set([
    ...Object.keys(spec.filters ?? {}),
    ...(spec.memberFilters ?? []).map(({ field }) => field),
    ...(spec.pageFields ?? []).filter(({ selectedItem }) => selectedItem !== undefined).map(({ field }) => field),
  ])
  for (const field of memberFields) {
    const members = context.membersOf?.(field)
    if (!members) {
      issues.push(unsupported('MEMBER_INVENTORY_REQUIRED', `Field ${field} needs a complete member inventory for native filtering.`, field))
      continue
    }
    const values = new Set<string>()
    for (const member of members) {
      if (!['string', 'number', 'boolean', 'blank'].includes(member.kind)) {
        issues.push(unsupported('UNSUPPORTED_MEMBER_TYPE', `Field ${field} contains unsupported member type ${member.kind}.`, field))
      }
      if (values.has(member.value)) issues.push(unsupported('DUPLICATE_MEMBER', `Field ${field} contains duplicate member ${member.value}.`, field))
      values.add(member.value)
    }
    const requested = [
      ...(spec.filters?.[field] ?? []),
      ...(spec.memberFilters ?? []).filter((filter) => filter.field === field).flatMap(({ values }) => values),
      ...(spec.pageFields ?? []).filter((page) => page.field === field && page.selectedItem !== undefined).map((page) => page.selectedItem!),
    ]
    for (const value of requested) if (!values.has(value)) {
      issues.push(unsupported('UNKNOWN_MEMBER', `Field ${field} does not contain member ${value}.`, field))
    }
  }
  for (const field of Object.keys(spec.filters ?? {})) {
    issues.push({ severity: 'warning', code: 'LEGACY_FILTER_TRANSLATED', message: `Legacy include filter ${field} will be translated to native hidden members.`, field })
  }
  return { representable: !issues.some(({ severity }) => severity === 'unsupported'), issues }
}

function duplicateIssues(values: readonly { field: string }[], label: string, issues: PivotRepresentabilityIssue[]): void {
  const seen = new Set<string>()
  for (const { field } of values) {
    if (seen.has(field)) issues.push(unsupported('DUPLICATE_FIELD_CONFIG', `Field ${field} has more than one ${label}.`, field))
    seen.add(field)
  }
}

function unsupported(code: PivotRepresentabilityCode, message: string, field?: string): PivotRepresentabilityIssue {
  return { severity: 'unsupported', code, message, field }
}

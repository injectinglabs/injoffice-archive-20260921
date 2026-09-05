import type {
  NativeOOXMLCellType,
  NativeWorkbookCapabilityV2,
  NativeWorkbookColumnDimensionV2,
  NativeWorkbookEffectiveStyleV2,
  NativeWorkbookFormulaV2,
  NativeWorkbookNormalStyleV2,
  NativeWorkbookPassthroughPartV2,
  NativeWorkbookRowDimensionV2,
  NativeWorkbookSheetFormatV2,
  NativeWorkbookSourceV2,
  NativeWorkbookUnsupportedV2,
  NativeWorkbookV2,
  NativeWorkbookValueV2,
  NativeSheetState,
} from './nativeContractV2.generated.js'
import { NativeWorkbookV2ValidationError, validateNativeWorkbookV2 } from './nativeValidationV2.js'

export const NATIVE_SHEET_RENDER_MODEL_V2_PROTOCOL = 'injoffice.xlsx.render-model'
export const NATIVE_SHEET_RENDER_MODEL_V2_VERSION = 2 as const
const nativeWorkbookRenderModelBrand: unique symbol = Symbol('injoffice.xlsx.render-model.v2')
const projectedNativeWorkbookModels = new WeakSet<object>()

export interface NativeStyleProvenanceV2 {
  readonly style_id: number
  readonly source_revision: string
  readonly source_package_sha256: string
  readonly projection: NativeWorkbookEffectiveStyleV2['projection']
  readonly raw_projection_sha256: string
}

export interface NativeRenderStyleV2 {
  readonly id: number
  readonly effective: NativeWorkbookEffectiveStyleV2
  readonly provenance: NativeStyleProvenanceV2
}

export type NativeRenderCellContentV2 =
  | { readonly kind: 'blank' }
  | { readonly kind: 'literal'; readonly value: NativeWorkbookValueV2 }
  | { readonly kind: 'formula'; readonly formula: NativeWorkbookFormulaV2 }

export interface NativeRenderCellV2 {
  readonly row: number
  readonly column: number
  readonly ref: string
  /** Absence is distinct from the explicit OOXML type `n`. */
  readonly ooxml_type?: NativeOOXMLCellType
  readonly style_id: number
  readonly content: NativeRenderCellContentV2
  /** Authoritative mutation decision from the source-bound native contract. */
  readonly editable: boolean
}

export interface NativeRenderMergedRangeV2 {
  /** Canonical bounded A1 rectangle from the source worksheet. */
  readonly ref: string
  /** Zero-based, inclusive grid coordinates; hidden dimensions do not renumber them. */
  readonly row: number
  readonly column: number
  readonly end_row: number
  readonly end_column: number
  readonly row_span: number
  readonly column_span: number
  readonly top_left: {
    readonly row: number
    readonly column: number
    readonly ref: string
  }
  /** Only the top-left cell supplies displayed content for this merged rectangle. */
  readonly content_authority: 'top-left'
  /** Covered cells display no content; sparse style-only cell records remain preserved. */
  readonly covered_cell_content: 'blank'
  /** Merge mutation has no native writer authority in v1. */
  readonly editable: false
  readonly source_part: string
  readonly source_revision: string
}

export interface NativeSheetMutationAuthorityV2 {
  readonly editable: boolean
  readonly refusal_code?: string
  readonly source_part: string
  readonly source_revision: string
}

export interface NativeSheetRenderModelV2 {
  readonly id: string
  readonly name: string
  readonly order: number
  readonly state: NativeSheetState
  readonly sheet_format?: NativeWorkbookSheetFormatV2
  readonly rows: ReadonlyArray<NativeWorkbookRowDimensionV2>
  readonly columns: ReadonlyArray<NativeWorkbookColumnDimensionV2>
  readonly cells: ReadonlyArray<NativeRenderCellV2>
  readonly merged_ranges: ReadonlyArray<NativeRenderMergedRangeV2>
  readonly mutation_authority: NativeSheetMutationAuthorityV2
}

export interface NativeWorkbookRenderModelV2 {
  readonly [nativeWorkbookRenderModelBrand]: true
  readonly protocol: typeof NATIVE_SHEET_RENDER_MODEL_V2_PROTOCOL
  readonly version: typeof NATIVE_SHEET_RENDER_MODEL_V2_VERSION
  readonly schema_sha256: string
  readonly document_id: string
  readonly revision: string
  readonly source: NativeWorkbookSourceV2
  readonly normal_style?: NativeWorkbookNormalStyleV2
  readonly date1904?: boolean
  readonly styles: ReadonlyArray<NativeRenderStyleV2>
  readonly sheets: ReadonlyArray<NativeSheetRenderModelV2>
  /** Retained verbatim so policy can reason about source-authoritative content. */
  readonly capabilities: ReadonlyArray<NativeWorkbookCapabilityV2>
  readonly passthrough_parts: ReadonlyArray<NativeWorkbookPassthroughPartV2>
  readonly unsupported: ReadonlyArray<NativeWorkbookUnsupportedV2>
}

/**
 * Builds a sparse, renderer-neutral sheet model. No values are evaluated or
 * coerced: numeric/date lexicals, cached formula values, and absent-vs-explicit
 * cell types survive exactly. Any invalid authority contract is refused before
 * a partial model can escape.
 */
export function projectNativeWorkbookV2(input: unknown): NativeWorkbookRenderModelV2 {
  const validation = validateNativeWorkbookV2(input)
  if (!validation.ok) throw new NativeWorkbookV2ValidationError(validation.issues)
  const workbook = validation.value
  const model = {
    protocol: NATIVE_SHEET_RENDER_MODEL_V2_PROTOCOL,
    version: NATIVE_SHEET_RENDER_MODEL_V2_VERSION,
    schema_sha256: workbook.schema_sha256,
    document_id: workbook.document_id,
    revision: workbook.revision,
    source: copySource(workbook.source),
    ...(workbook.normal_style === undefined ? {} : { normal_style: { ...workbook.normal_style } }),
    ...(workbook.date1904 === undefined ? {} : { date1904: workbook.date1904 }),
    styles: workbook.styles.map((style) => ({
      id: style.id,
      effective: copyStyle(style.effective),
      provenance: {
        style_id: style.id,
        source_revision: workbook.revision,
        source_package_sha256: workbook.source.package_sha256,
        projection: style.effective.projection,
        raw_projection_sha256: style.raw_projection_sha256,
      },
    })),
    sheets: workbook.sheets.map((sheet) => ({
      id: sheet.id,
      name: sheet.name,
      order: sheet.order,
      state: sheet.state,
      ...(sheet.sheet_format === undefined ? {} : { sheet_format: { ...sheet.sheet_format } }),
      rows: sheet.rows.map((row) => ({ ...row })),
      columns: sheet.columns.map((column) => ({ ...column })),
      cells: sheet.cells.map((cell) => ({
        row: cell.row,
        column: cell.column,
        ref: cell.ref,
        ...(cell.ooxml_type === undefined ? {} : { ooxml_type: cell.ooxml_type }),
        style_id: cell.style_id,
        content: cell.formula ? { kind: 'formula', formula: copyFormula(cell.formula) } : cell.value ? { kind: 'literal', value: copyValue(cell.value) } : { kind: 'blank' },
        editable: cell.editable,
      })),
      merged_ranges: sheet.merged_ranges.map((range) => ({
        ref: range.ref,
        row: range.row,
        column: range.column,
        end_row: range.end_row,
        end_column: range.end_column,
        row_span: range.end_row - range.row + 1,
        column_span: range.end_column - range.column + 1,
        top_left: {
          row: range.row,
          column: range.column,
          ref: range.ref.slice(0, range.ref.indexOf(':')),
        },
        content_authority: 'top-left',
        covered_cell_content: 'blank',
        editable: false,
        source_part: sheet.part_name,
        source_revision: workbook.revision,
      })),
      mutation_authority: {
        editable: sheet.editable,
        ...(sheet.refusal_code === undefined ? {} : { refusal_code: sheet.refusal_code }),
        source_part: sheet.part_name,
        source_revision: workbook.revision,
      },
    })),
    capabilities: workbook.capabilities.map((item) => ({ ...item })),
    passthrough_parts: workbook.passthrough_parts.map((item) => ({ ...item })),
    unsupported: workbook.unsupported.map((item) => ({ ...item })),
  } as unknown as NativeWorkbookRenderModelV2
  Object.defineProperty(model, nativeWorkbookRenderModelBrand, { value: true, enumerable: false, configurable: false, writable: false })
  projectedNativeWorkbookModels.add(model)
  return deepFreeze(model)
}

export function isProjectedNativeWorkbookV2(input: unknown): input is NativeWorkbookRenderModelV2 {
  return typeof input === 'object' && input !== null && projectedNativeWorkbookModels.has(input) && (input as { [nativeWorkbookRenderModelBrand]?: unknown })[nativeWorkbookRenderModelBrand] === true && isDeepFrozen(input)
}

function copySource(source: NativeWorkbookSourceV2): NativeWorkbookSourceV2 { return { ...source } }
function copyValue(value: NativeWorkbookValueV2): NativeWorkbookValueV2 {
  return { ...value, ...(value.runs === undefined ? {} : { runs: value.runs.map((run) => ({ ...run })) }) }
}
function copyFormula(formula: NativeWorkbookFormulaV2): NativeWorkbookFormulaV2 { return { ...formula, ...(formula.cached ? { cached: copyValue(formula.cached) } : {}) } }
function copyStyle(style: NativeWorkbookEffectiveStyleV2): NativeWorkbookEffectiveStyleV2 {
  return {
    ...style,
    ...(style.fill === undefined ? {} : { fill: { ...style.fill } }),
    ...(style.border === undefined ? {} : {
      border: {
        ...style.border,
        ...(style.border.left === undefined ? {} : { left: { ...style.border.left } }),
        ...(style.border.right === undefined ? {} : { right: { ...style.border.right } }),
        ...(style.border.top === undefined ? {} : { top: { ...style.border.top } }),
        ...(style.border.bottom === undefined ? {} : { bottom: { ...style.border.bottom } }),
      },
    }),
    unsupported: [...style.unsupported],
  }
}

function deepFreeze<T>(value: T): T {
  if (typeof value !== 'object' || value === null || Object.isFrozen(value)) return value
  for (const child of Object.values(value)) deepFreeze(child)
  return Object.freeze(value)
}

function isDeepFrozen(value: unknown, seen = new Set<object>()): boolean {
  if (typeof value !== 'object' || value === null || seen.has(value)) return true
  if (!Object.isFrozen(value)) return false
  seen.add(value)
  return Object.values(value).every((child) => isDeepFrozen(child, seen))
}

import { describe, expect, it } from 'vitest'
import {
  EXCEL_MAX_COLUMNS,
  EXCEL_MAX_ROWS,
  MAX_WORKBOOK_MUTATIONS,
  WORKBOOK_MUTATION_PROTOCOL,
  WorkbookMutationValidationError,
  decodeWorkbookMutationBatch,
  encodeWorkbookMutationBatch,
  type WorkbookMutationBatchV1,
} from './mutationProtocol.js'

const base = (operations: unknown[]) => ({
  protocol: WORKBOOK_MUTATION_PROTOCOL,
  version: 1,
  batch_id: 'save-0001',
  expected_revision: 'sha256:original-workbook',
  operations,
})

const cell = { row: 2, column: 3 }
const op = (kind: string, extra: Record<string, unknown> = {}, operationId = kind) => ({
  operation_id: operationId,
  kind,
  sheet_id: 'sheet-stable-1',
  ...extra,
})

describe('workbook mutation protocol v1', () => {
  it('accepts the complete supported operation vocabulary in listed order', () => {
    const input = base([
      op('cell.set_value', { cell, value: 42 }, '1'),
      op('cell.clear_value', { cell }, '2'),
      op('cell.set_formula', { cell, formula: '=SUM(A1:A2)' }, '3'),
      op('cell.clear_formula', { cell }, '4'),
      op('style.patch', {
        range: { row: 0, column: 0, end_row: 2, end_column: 3 },
        style: { number_format: '#,##0.00', bold: true, fill_color: '#AABBCC', wrap_text: null },
      }, '5'),
      op('row.set_height', { row: 7, height_points: 20.5 }, '6'),
      op('column.set_width', { column: 4, width: 12.25 }, '7'),
      op('range.merge', { range: { row: 0, column: 0, end_row: 0, end_column: 2 } }, '8'),
      op('range.unmerge', { range: { row: 0, column: 0, end_row: 0, end_column: 2 } }, '9'),
    ])

    const decoded = decodeWorkbookMutationBatch(input)
    expect(decoded.ok).toBe(true)
    if (!decoded.ok) return
    expect(decoded.value.operations.map((operation) => operation.operation_id)).toEqual(['1', '2', '3', '4', '5', '6', '7', '8', '9'])
    expect(JSON.parse(encodeWorkbookMutationBatch(decoded.value))).toEqual(input)
  })

  it('serializes equivalent decoded values deterministically without sorting operations', () => {
    const scrambled = {
      operations: [
        { value: 'Ada', cell: { column: 1, row: 0 }, sheet_id: 's1', kind: 'cell.set_value', operation_id: 'b' },
        { height_points: 22, row: 4, kind: 'row.set_height', operation_id: 'a', sheet_id: 's1' },
      ],
      batch_id: 'save-retry-42',
      expected_revision: 'revision-1',
      version: 1,
      protocol: WORKBOOK_MUTATION_PROTOCOL,
    }
    const decoded = decodeWorkbookMutationBatch(scrambled)
    expect(decoded.ok).toBe(true)
    if (!decoded.ok) return
    expect(encodeWorkbookMutationBatch(decoded.value)).toBe(
      '{"protocol":"injoffice.xlsx.mutations","version":1,"batch_id":"save-retry-42","expected_revision":"revision-1","operations":[{"operation_id":"b","sheet_id":"s1","kind":"cell.set_value","cell":{"row":0,"column":1},"value":"Ada"},{"operation_id":"a","sheet_id":"s1","kind":"row.set_height","row":4,"height_points":22}]}',
    )
  })

  it('requires a bounded canonical batch id distinct from the expected revision', () => {
    const missing = decodeWorkbookMutationBatch({ ...base([op('cell.clear_value', { cell })]), batch_id: undefined })
    expect(missing).toEqual({ ok: false, issues: [{ code: 'REQUIRED', path: '/batch_id', message: 'field is required' }] })

    const invalid = decodeWorkbookMutationBatch({ ...base([op('cell.clear_value', { cell })]), batch_id: 'save id with spaces' })
    expect(invalid.ok).toBe(false)
    if (!invalid.ok) expect(invalid.issues).toEqual([expect.objectContaining({ code: 'INVALID_VALUE', path: '/batch_id' })])

    const decoded = decodeWorkbookMutationBatch({
      ...base([op('cell.clear_value', { cell })]),
      batch_id: 'retry-key-7',
      expected_revision: 'opaque-CAS-revision-99',
    })
    expect(decoded.ok).toBe(true)
    if (decoded.ok) {
      expect(decoded.value.batch_id).toBe('retry-key-7')
      expect(decoded.value.expected_revision).toBe('opaque-CAS-revision-99')
    }
  })

  it('refuses oversized batches before reading entries or accumulating per-operation issues', () => {
    const backing = new Array(MAX_WORKBOOK_MUTATIONS + 1)
    const operations = new Proxy(backing, {
      get(target, property, receiver) {
        if (typeof property === 'string' && /^\d+$/.test(property)) throw new Error('oversized batch entry was inspected')
        return Reflect.get(target, property, receiver)
      },
    })
    expect(decodeWorkbookMutationBatch(base(operations))).toEqual({
      ok: false,
      issues: [{
        code: 'BATCH_TOO_LARGE',
        path: '/operations',
        message: `must contain at most ${MAX_WORKBOOK_MUTATIONS} operations`,
      }],
    })
  })

  it.each([
    'sheet.add',
    'sheet.delete',
    'sheet.rename',
    'sheet.reorder',
    'row.insert',
    'row.delete',
    'row.move',
    'column.insert',
    'column.delete',
    'column.move',
    'range.move',
  ])('refuses recognized structural operation %s', (kind) => {
    const decoded = decodeWorkbookMutationBatch(base([op(kind)]))
    expect(decoded).toEqual({
      ok: false,
      issues: [{
        code: 'UNSUPPORTED_OPERATION',
        path: '/operations/0/kind',
        message: `structural operation ${JSON.stringify(kind)} is not supported by protocol v1`,
        operation_index: 0,
        operation_id: kind,
      }],
    })
  })

  it('distinguishes unknown operations from explicitly unsupported structural edits', () => {
    const decoded = decodeWorkbookMutationBatch(base([op('cell.teleport')]))
    expect(decoded.ok).toBe(false)
    if (decoded.ok) return
    expect(decoded.issues).toContainEqual(expect.objectContaining({ code: 'UNKNOWN_OPERATION', path: '/operations/0/kind' }))
  })

  it('rejects unknown fields at every modeled level including style fields', () => {
    const decoded = decodeWorkbookMutationBatch({
      ...base([op('style.patch', {
        range: { row: 0, column: 0, end_row: 1, end_column: 1, endColumn: 1 },
        style: { bold: true, border: 'thin' },
        ignored: true,
      })]),
      extra: true,
    })
    expect(decoded.ok).toBe(false)
    if (decoded.ok) return
    expect(decoded.issues.map(({ code, path }) => ({ code, path }))).toEqual(expect.arrayContaining([
      { code: 'UNKNOWN_FIELD', path: '/extra' },
      { code: 'UNKNOWN_FIELD', path: '/operations/0/ignored' },
      { code: 'UNKNOWN_FIELD', path: '/operations/0/range/endColumn' },
      { code: 'UNKNOWN_FIELD', path: '/operations/0/style/border' },
    ]))
  })

  it('rejects missing revision, empty batches, duplicate ids, and an empty style delta', () => {
    const missingRevision = { ...base([]), expected_revision: ' ' }
    const empty = decodeWorkbookMutationBatch(missingRevision)
    expect(empty.ok).toBe(false)
    if (!empty.ok) expect(empty.issues.map((entry) => entry.code)).toEqual(expect.arrayContaining(['INVALID_VALUE', 'EMPTY_BATCH']))

    const duplicate = decodeWorkbookMutationBatch(base([
      op('cell.clear_value', { cell }, 'same'),
      op('cell.clear_formula', { cell }, 'same'),
      op('style.patch', { range: { row: 0, column: 0, end_row: 0, end_column: 0 }, style: {} }, 'style'),
    ]))
    expect(duplicate.ok).toBe(false)
    if (!duplicate.ok) expect(duplicate.issues.map((entry) => entry.code)).toEqual(expect.arrayContaining(['DUPLICATE_OPERATION_ID', 'EMPTY_STYLE_DELTA']))
  })

  it('rejects non-finite, fractional, reversed, and out-of-grid coordinates', () => {
    const decoded = decodeWorkbookMutationBatch(base([
      op('cell.set_value', { cell: { row: Number.NaN, column: 0 }, value: 1 }, 'nan'),
      op('cell.set_value', { cell: { row: 1.5, column: 0 }, value: 1 }, 'fraction'),
      op('cell.set_value', { cell: { row: EXCEL_MAX_ROWS, column: EXCEL_MAX_COLUMNS }, value: 1 }, 'outside'),
      op('range.merge', { range: { row: 5, column: 5, end_row: 4, end_column: 8 } }, 'backwards'),
      op('row.set_height', { row: 0, height_points: Number.POSITIVE_INFINITY }, 'infinite'),
    ]))
    expect(decoded.ok).toBe(false)
    if (decoded.ok) return
    expect(decoded.issues.map((entry) => entry.path)).toEqual(expect.arrayContaining([
      '/operations/0/cell/row',
      '/operations/1/cell/row',
      '/operations/2/cell/row',
      '/operations/2/cell/column',
      '/operations/3/range',
      '/operations/4/height_points',
    ]))
  })

  it('rejects malformed values, formulas, colors, alignments, and single-cell merges', () => {
    const decoded = decodeWorkbookMutationBatch(base([
      op('cell.set_value', { cell, value: null }, 'null-value'),
      op('cell.set_formula', { cell, formula: 'SUM(A1:A2)' }, 'formula'),
      op('style.patch', {
        range: { row: 0, column: 0, end_row: 0, end_column: 0 },
        style: { fill_color: '#aabbcc', horizontal_alignment: 'justify' },
      }, 'style'),
      op('range.merge', { range: { row: 0, column: 0, end_row: 0, end_column: 0 } }, 'merge'),
    ]))
    expect(decoded.ok).toBe(false)
    if (decoded.ok) return
    expect(decoded.issues.map((entry) => entry.code)).toEqual(expect.arrayContaining(['INVALID_TYPE', 'INVALID_VALUE', 'INVALID_RANGE']))
  })

  it('throws the structured validation error when encoding an invalid typed value', () => {
    const invalid = base([op('row.insert')]) as unknown as WorkbookMutationBatchV1
    expect(() => encodeWorkbookMutationBatch(invalid)).toThrow(WorkbookMutationValidationError)
    try {
      encodeWorkbookMutationBatch(invalid)
    } catch (error) {
      expect(error).toBeInstanceOf(WorkbookMutationValidationError)
      expect((error as WorkbookMutationValidationError).issues[0]?.code).toBe('UNSUPPORTED_OPERATION')
    }
  })
})

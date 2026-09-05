import { createHash } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import {
  NATIVE_SHEET_DECORATION_LIMITS,
  NativeSheetDecorationError,
  NativeWorkbookValidationError,
  compileNativeSheetDecorationsV1,
  compileNativeSheetGeometryV1,
  createNativeSheetDecorationRecordingSurfaceV1,
  decodeNativeWorkbookV1,
  emitNativeSheetDecorationCommandsV1,
  nativeWorkbookStyleRawProjectionSha256V1,
  projectNativeWorkbookV1,
  replayNativeSheetDecorationCommandsV1,
  validateNativeSheetDecorationPlanV1,
} from './index.js'
import type {
  NativeMaximumDigitWidthAuthorityV1,
  NativeSheetDecorationCommandV1,
  NativeSheetDecorationPlanV1,
  NativeSheetGeometryV1,
  NativeWorkbookRenderModelV1,
  NativeWorkbookUnsupportedV1,
  NativeWorkbookV1,
} from './index.js'

const repositoryRoot = resolve(import.meta.dirname, '../../..')
const fixtureSource = readFileSync(resolve(repositoryRoot, 'go/xlsxpatch/testdata/native-xlsx-v1/valid/lexical-render.json'), 'utf8').trim()
const excelFixtureSource = readFileSync(resolve(repositoryRoot, 'go/xlsxpatch/testdata/native-xlsx-v1/valid/excel-authored-happy-tree.json'), 'utf8').trim()

function fixture(): NativeWorkbookV1 {
  const workbook = JSON.parse(fixtureSource) as NativeWorkbookV1
  const mutable = workbook as unknown as MutableWorkbook
  mutable.normal_style = {
    style_xf_id: 0, font_id: 0, font_name: 'Calibri', font_size_points: 11, font_bold: false, font_italic: false,
    font_record_sha256: `sha256:${'e'.repeat(64)}`,
  }
  mutable.styles[0]!.effective = {
    ...mutable.styles[0]!.effective,
    fill: { origin: 'styles-record', fill_id: 0, record_sha256: `sha256:${'3'.repeat(64)}` },
    border: {
      origin: 'styles-record', border_id: 0, record_sha256: `sha256:${'1'.repeat(64)}`,
      right: { style: 'thin', color: '#102030' },
    },
  }
  mutable.styles[1]!.effective = {
    ...mutable.styles[1]!.effective,
    fill_color: '#DDEEFF',
    fill: { origin: 'styles-record', fill_id: 1, record_sha256: `sha256:${'4'.repeat(64)}`, color: '#DDEEFF' },
    border: {
      origin: 'styles-record', border_id: 1, record_sha256: `sha256:${'2'.repeat(64)}`,
      left: { style: 'thin', color: '#102030' },
      top: { style: 'double', color: '#405060' },
      right: { style: 'dashed', color: '#708090' },
      bottom: { style: 'medium', color: '#A0B0C0' },
    },
  }
  sealStyles(mutable)
  return workbook
}

type MutableWorkbook = {
  normal_style?: NativeWorkbookV1['normal_style']
  styles: Array<{ id: number; effective: MutableEffective; raw_projection_sha256: string }>
  sheets: Array<{
    id: string; part_name: string; editable: boolean; refusal_code?: string
    rows: Array<{ row: number; height_points?: number; hidden: boolean; custom_height: boolean; style_id?: number }>
    columns: Array<{ column: number; end_column: number; width?: number; hidden: boolean; custom_width: boolean; best_fit: boolean; style_id?: number }>
    cells: Array<Record<string, unknown>>
    merged_ranges: Array<Record<string, unknown>>
  }>
  unsupported: NativeWorkbookUnsupportedV1[]
}

type MutableEffective = Record<string, unknown> & {
  number_format?: string; font_color?: string; fill_color?: string
  fill?: { origin: string; fill_id?: number; record_sha256?: string; color?: string }
  border?: {
    origin: string; border_id?: number; record_sha256?: string
    left?: { style: string; color: string }; right?: { style: string; color: string }
    top?: { style: string; color: string }; bottom?: { style: string; color: string }
  }
  projection: 'full' | 'partial'; unsupported: string[]
}

function sealStyles(workbook: MutableWorkbook): void {
  for (const style of workbook.styles) style.raw_projection_sha256 = nativeWorkbookStyleRawProjectionSha256V1(style.effective as NativeWorkbookV1['styles'][number]['effective'])
}

function freezeRecursively(value: unknown): void {
  if (typeof value !== 'object' || value === null || Object.isFrozen(value)) return
  for (const child of Object.values(value)) freezeRecursively(child)
  Object.freeze(value)
}

function metric(workbook: NativeWorkbookV1): NativeMaximumDigitWidthAuthorityV1 {
  const normal = workbook.normal_style!
  return {
    source_revision: workbook.revision,
    source_package_sha256: workbook.source.package_sha256,
    normal_style_xf_id: normal.style_xf_id,
    normal_style_font_id: normal.font_id,
    font_name: normal.font_name,
    font_size_points: normal.font_size_points,
    font_bold: normal.font_bold,
    font_italic: normal.font_italic,
    normal_font_record_sha256: normal.font_record_sha256 as `sha256:${string}`,
    font_sha256: `sha256:${'f'.repeat(64)}`,
    provider_id: 'fixture-opentype-parser',
    provider_revision: '1',
    measurement_dpi: 96,
    maximum_digit_width_pixels: 7,
  }
}

function setup(workbook: NativeWorkbookV1 = fixture()): { workbook: NativeWorkbookV1; model: NativeWorkbookRenderModelV1; geometry: NativeSheetGeometryV1 } {
  const model = projectNativeWorkbookV1(workbook)
  const geometry = compileNativeSheetGeometryV1(model, '7', { row: 0, column: 0, end_row: 0, end_column: 1 }, metric(workbook))
  return { workbook, model, geometry }
}

function issue(code: string, capability: string, scopeID: string, partName: string, extra: Partial<NativeWorkbookUnsupportedV1> = {}): NativeWorkbookUnsupportedV1 {
  const location = [code, capability, scopeID, partName, extra.cell_ref ?? '', extra.range_ref ?? ''].join('\0')
  return {
    id: `unsupported:${createHash('sha256').update(location).digest('hex')}`,
    code, capability, scope_id: scopeID, part_name: partName,
    preservation: 'preserve-exact', message: 'source-authoritative appearance', ...extra,
  }
}

function record(plan: NativeSheetDecorationPlanV1): ReadonlyArray<NativeSheetDecorationCommandV1> {
  const surface = createNativeSheetDecorationRecordingSurfaceV1()
  emitNativeSheetDecorationCommandsV1(plan, surface)
  return surface.finish()
}

function resealPlan(plan: NativeSheetDecorationPlanV1): NativeSheetDecorationPlanV1 {
  const { decoration_sha256: _discarded, ...unsigned } = structuredClone(plan)
  return { ...unsigned, decoration_sha256: `sha256:${createHash('sha256').update(JSON.stringify(unsigned)).digest('hex')}` }
}

function uncheckedCommands(plan: NativeSheetDecorationPlanV1): NativeSheetDecorationCommandV1[] {
  return [
    {
      kind: 'beginDecorations', protocol: plan.protocol, version: plan.version,
      document_id: plan.document_id, sheet_id: plan.sheet_id, source_part: plan.source_part,
      source_revision: plan.source_revision, source_package_sha256: plan.source_package_sha256,
      geometry_sha256: plan.geometry_sha256, coordinate_space: plan.coordinate_space,
      decoration_sha256: plan.decoration_sha256,
    },
    { kind: 'clipRect', rect: plan.bounds },
    ...plan.fills.map((decoration) => ({ kind: 'fillRect' as const, decoration })),
    ...plan.border_segments.map((decoration) => ({ kind: 'borderSegment' as const, decoration })),
    { kind: 'endDecorations' },
  ]
}

describe('native XLSX exact fill and border decorations', () => {
  it('compiles deterministic integer-EMU fills and deduplicated identical joins', () => {
    const { model, geometry } = setup()
    const before = JSON.stringify(model)
    const first = compileNativeSheetDecorationsV1(model, geometry)
    const second = compileNativeSheetDecorationsV1(model, geometry)
    expect(JSON.stringify(first)).toBe(JSON.stringify(second))
    expect(first.decoration_sha256).toMatch(/^sha256:[0-9a-f]{64}$/)
    expect(first.fills).toEqual([{
      cell_ref: 'B1', row: 0, column: 1, style_id: 1, fill_id: 1, fill_record_sha256: `sha256:${'4'.repeat(64)}`,
      rect: { x_emu: 828_675, y_emu: 0, width_emu: 828_675, height_emu: 231_775 }, color: '#DDEEFF',
    }])
    const join = first.border_segments.find((segment) => segment.x1_emu === 828_675 && segment.x2_emu === 828_675)
    expect(join).toMatchObject({ orientation: 'vertical', y1_emu: 0, y2_emu: 231_775, border_style: 'thin', color: '#102030' })
    expect(join?.sources).toEqual([
      { cell_ref: 'A1', edge: 'right', style_id: 0, border_id: 0, border_record_sha256: `sha256:${'1'.repeat(64)}` },
      { cell_ref: 'B1', edge: 'left', style_id: 1, border_id: 1, border_record_sha256: `sha256:${'2'.repeat(64)}` },
    ])
    expect(first.border_segments).toHaveLength(4)
    expect(Object.isFrozen(model)).toBe(true)
    expect(Object.isFrozen(model.styles[1]?.effective.border)).toBe(true)
    expect(Object.isFrozen(first.border_segments[0]?.sources)).toBe(true)
    expect(JSON.stringify(model)).toBe(before)
  })

  it('validates the complete plan before the first push and validates replay before the first callback', () => {
    const { model, geometry } = setup()
    const plan = compileNativeSheetDecorationsV1(model, geometry)
    const commands = record(plan)
    expect(commands.map((command) => command.kind)).toEqual([
      'beginDecorations', 'clipRect', 'fillRect', 'borderSegment', 'borderSegment', 'borderSegment', 'borderSegment', 'endDecorations',
    ])
    expect(commands[0]).toMatchObject({
      kind: 'beginDecorations', protocol: 'injoffice.xlsx.sheet-decoration', version: 1,
      source_part: 'Worksheets/Sheet1.xml', source_revision: model.revision,
    })
    const replayed: string[] = []
    replayNativeSheetDecorationCommandsV1(replayed, commands, { execute(target, command) { target.push(command.kind) } })
    expect(replayed).toEqual(commands.map((command) => command.kind))

    const stalePlan = structuredClone(plan) as unknown as { fills: Array<{ color: string }> }
    stalePlan.fills[0]!.color = '#ABCDEF'
    let pushes = 0
    expect(() => emitNativeSheetDecorationCommandsV1(stalePlan as unknown as NativeSheetDecorationPlanV1, { push() { pushes++ } })).toThrowError(expect.objectContaining({ code: 'decoration.planDigest' }))
    expect(pushes).toBe(0)

    const staleCommands = structuredClone(commands)
    const fill = staleCommands.find((command) => command.kind === 'fillRect') as Extract<NativeSheetDecorationCommandV1, { kind: 'fillRect' }>
    ;(fill.decoration as { color: string }).color = '#ABCDEF'
    let callbacks = 0
    expect(() => replayNativeSheetDecorationCommandsV1({}, staleCommands, { execute() { callbacks++ } })).toThrowError(expect.objectContaining({ code: 'decoration.planDigest' }))
    expect(callbacks).toBe(0)

    const capacity = createNativeSheetDecorationRecordingSurfaceV1(commands.length)
    emitNativeSheetDecorationCommandsV1(plan, capacity, commands.length)
    expect(capacity.finish()).toHaveLength(commands.length)
    const tooSmall = createNativeSheetDecorationRecordingSurfaceV1(commands.length - 1)
    expect(() => emitNativeSheetDecorationCommandsV1(plan, tooSmall, commands.length)).toThrowError(expect.objectContaining({ code: 'decoration.commandBudget' }))
    expect(tooSmall.commands).toEqual([])
  })

  it('rejects malformed, oversized, fractional, escaping, and source-forged plans atomically', () => {
    const { model, geometry } = setup()
    const plan = compileNativeSheetDecorationsV1(model, geometry)
    const cases: unknown[] = []
    for (const [key, value] of [
      ['protocol', 'other'], ['version', 2], ['document_id', ''], ['sheet_id', '0'], ['source_part', '../sheet.xml'],
      ['source_revision', `rev:${'b'.repeat(64)}`], ['coordinate_space', 'page'],
    ] as const) cases.push({ ...structuredClone(plan), [key]: value })
    const fractional = structuredClone(plan) as unknown as { fills: Array<{ rect: { x_emu: number } }> }; fractional.fills[0]!.rect.x_emu = 0.5; cases.push(fractional)
    const escaping = structuredClone(plan) as unknown as { fills: Array<{ rect: { width_emu: number } }> }; escaping.fills[0]!.rect.width_emu = Number.MAX_SAFE_INTEGER; cases.push(escaping)
    const malformedSource = structuredClone(plan) as unknown as { border_segments: Array<{ sources: Array<{ cell_ref: string }> }> }; malformedSource.border_segments[0]!.sources[0]!.cell_ref = 'XFE1'; cases.push(malformedSource)
    const fractionalBorder = structuredClone(plan) as unknown as { border_segments: Array<{ x1_emu: number }> }; fractionalBorder.border_segments[0]!.x1_emu = 0.5; cases.push(fractionalBorder)
    cases.push({ ...structuredClone(plan), source_part: 'xl/worksheets/%ZZ.xml' })
    const oversized = structuredClone(plan) as unknown as { fills: unknown[] }
    oversized.fills = new Array(NATIVE_SHEET_DECORATION_LIMITS.maxFills + 1)
    cases.push(oversized)
    const oversizedBorders = structuredClone(plan) as unknown as { border_segments: unknown[] }
    oversizedBorders.border_segments = new Array(NATIVE_SHEET_DECORATION_LIMITS.maxBorderSegments + 1)
    cases.push(oversizedBorders)
    for (const candidate of cases) {
      let pushes = 0
      expect(() => emitNativeSheetDecorationCommandsV1(candidate as NativeSheetDecorationPlanV1, { push() { pushes++ } })).toThrow(NativeSheetDecorationError)
      expect(pushes).toBe(0)
    }
    expect(() => validateNativeSheetDecorationPlanV1({})).toThrowError(expect.objectContaining({ code: 'decoration.planInvalid' }))

    const commands = structuredClone(record(plan))
    const malformedStreams: unknown[] = [commands.slice(1), [...commands, ...new Array(NATIVE_SHEET_DECORATION_LIMITS.maxCommands).fill({ kind: 'endDecorations' })]]
    for (const stream of malformedStreams) {
      let callbacks = 0
      expect(() => replayNativeSheetDecorationCommandsV1({}, stream as NativeSheetDecorationCommandV1[], { execute() { callbacks++ } })).toThrow(NativeSheetDecorationError)
      expect(callbacks).toBe(0)
    }
  })

  it('rejects mutable handcrafted models and valid-looking style tampering with a stale raw projection digest', () => {
    const workbook = fixture()
    const model = projectNativeWorkbookV1(workbook)
    const geometry = compileNativeSheetGeometryV1(model, '7', { row: 0, column: 0, end_row: 0, end_column: 1 }, metric(workbook))
    const handcrafted = structuredClone(model)
    expect(() => compileNativeSheetGeometryV1(handcrafted, '7', geometry.viewport, metric(workbook))).toThrowError(expect.objectContaining({ code: 'geometry.sourceUnsupported' }))
    expect(() => compileNativeSheetDecorationsV1(handcrafted, geometry)).toThrowError(expect.objectContaining({ code: 'decoration.modelAuthority' }))
    for (const symbol of Object.getOwnPropertySymbols(model)) Object.defineProperty(handcrafted, symbol, { value: true })
    freezeRecursively(handcrafted)
    expect(() => compileNativeSheetDecorationsV1(handcrafted, geometry)).toThrowError(expect.objectContaining({ code: 'decoration.modelAuthority' }))
    const skeletal = { protocol: model.protocol, version: model.version, document_id: model.document_id, revision: model.revision, styles: [], sheets: [] }
    expect(() => compileNativeSheetGeometryV1(skeletal as unknown as NativeWorkbookRenderModelV1, '7', geometry.viewport, metric(workbook))).toThrowError(expect.objectContaining({ code: 'geometry.sourceUnsupported' }))
    expect(() => compileNativeSheetDecorationsV1(skeletal as unknown as NativeWorkbookRenderModelV1, geometry)).toThrowError(expect.objectContaining({ code: 'decoration.modelAuthority' }))
    const oversizedHandcrafted = { ...model, styles: new Array(100_001).fill(model.styles[0]) }
    expect(() => compileNativeSheetGeometryV1(oversizedHandcrafted as NativeWorkbookRenderModelV1, '7', geometry.viewport, metric(workbook))).toThrowError(expect.objectContaining({ code: 'geometry.sourceUnsupported' }))
    expect(() => compileNativeSheetDecorationsV1(oversizedHandcrafted as NativeWorkbookRenderModelV1, geometry)).toThrowError(expect.objectContaining({ code: 'decoration.modelAuthority' }))
    const oversizedNative = structuredClone(workbook) as unknown as MutableWorkbook
    oversizedNative.styles = new Array(100_001).fill(oversizedNative.styles[0])
    expect(() => projectNativeWorkbookV1(oversizedNative)).toThrow(NativeWorkbookValidationError)

    for (const tamper of [
      (candidate: MutableWorkbook) => { candidate.styles[1]!.effective.fill_color = '#ABCDEF'; candidate.styles[1]!.effective.fill!.color = '#ABCDEF' },
      (candidate: MutableWorkbook) => { candidate.styles[1]!.effective.border!.top!.style = 'thin' },
      (candidate: MutableWorkbook) => { candidate.styles[1]!.effective.border!.top!.color = '#ABCDEF' },
    ]) {
      const candidate = structuredClone(workbook) as unknown as MutableWorkbook
      tamper(candidate)
      expect(() => projectNativeWorkbookV1(candidate)).toThrow(NativeWorkbookValidationError)
    }
  })

  it('keeps exact fill and border lanes independent from unrelated unsupported metadata', () => {
    const workbook = fixture() as unknown as MutableWorkbook
    const style = workbook.styles[1]!.effective
    delete style.font_color
    delete style.number_format
    style.projection = 'partial'
    style.unsupported = ['font-color', 'number-format']
    workbook.unsupported.push(issue('STYLE_FONT_COLOR', 'styles', 'style:1', 'XL/Styles.xml'))
    workbook.unsupported.push(issue('STYLE_NUMBER_FORMAT', 'styles', 'style:1', 'XL/Styles.xml'))
    sealStyles(workbook)
    const { model, geometry } = setup(workbook as unknown as NativeWorkbookV1)
    const plan = compileNativeSheetDecorationsV1(model, geometry)
    expect(plan.fills).toHaveLength(1)
    expect(plan.border_segments).toHaveLength(4)
  })

  it('requires an exact bijection between partial style codes and source diagnostics', () => {
    const missing = fixture() as unknown as MutableWorkbook
    missing.styles[1]!.effective.projection = 'partial'
    missing.styles[1]!.effective.unsupported = ['font-color']
    sealStyles(missing)
    expect(() => projectNativeWorkbookV1(missing)).toThrow(NativeWorkbookValidationError)

    const contradictory = fixture() as unknown as MutableWorkbook
    contradictory.unsupported.push(issue('STYLE_BORDER', 'styles', 'style:1', 'XL/Styles.xml'))
    expect(() => projectNativeWorkbookV1(contradictory)).toThrow(NativeWorkbookValidationError)

    const nonCanonicalScope = fixture() as unknown as MutableWorkbook
    delete nonCanonicalScope.styles[1]!.effective.font_color
    nonCanonicalScope.styles[1]!.effective.projection = 'partial'
    nonCanonicalScope.styles[1]!.effective.unsupported = ['font-color']
    sealStyles(nonCanonicalScope)
    const unsupportedIndex = nonCanonicalScope.unsupported.length
    nonCanonicalScope.unsupported.push(issue('STYLE_FONT_COLOR', 'styles', 'style:+1', 'XL/Styles.xml'))
    try {
      projectNativeWorkbookV1(nonCanonicalScope)
      throw new Error('non-canonical style scope was accepted')
    } catch (error) {
      expect(error).toBeInstanceOf(NativeWorkbookValidationError)
      expect((error as NativeWorkbookValidationError).issues).toContainEqual(expect.objectContaining({ path: `/unsupported/${unsupportedIndex}/scope_id` }))
    }
  })

  it('rejects recomputed-digest duplicate, orientation-invalid, and non-canonical border topology atomically', () => {
    const { model, geometry } = setup()
    const plan = compileNativeSheetDecorationsV1(model, geometry)
    const duplicate = structuredClone(plan) as unknown as { border_segments: Array<Record<string, unknown>> }
    const conflicting = structuredClone(duplicate.border_segments[0]!)
    conflicting.color = '#ABCDEF'
    duplicate.border_segments.splice(1, 0, conflicting)
    const wrongEdge = structuredClone(plan) as unknown as { border_segments: Array<{ orientation: string; sources: Array<{ edge: string }> }> }
    const horizontal = wrongEdge.border_segments.find((segment) => segment.orientation === 'horizontal')!
    horizontal.sources[0]!.edge = 'left'
    const reversedJoin = structuredClone(plan) as unknown as { border_segments: Array<{ sources: Array<Record<string, unknown>> }> }
    reversedJoin.border_segments.find((segment) => segment.sources.length === 2)!.sources.reverse()
    const sameCellJoin = structuredClone(plan) as unknown as { border_segments: Array<{ sources: Array<{ cell_ref: string }> }> }
    const sameCellSources = sameCellJoin.border_segments.find((segment) => segment.sources.length === 2)!.sources
    sameCellSources[1]!.cell_ref = sameCellSources[0]!.cell_ref
    const nonAdjacentJoin = structuredClone(plan) as unknown as { border_segments: Array<{ sources: Array<{ cell_ref: string }> }> }
    nonAdjacentJoin.border_segments.find((segment) => segment.sources.length === 2)!.sources[1]!.cell_ref = 'C1'
    const reusedSource = structuredClone(plan) as unknown as { border_segments: Array<{ sources: Array<{ cell_ref: string; edge: string }> }> }
    reusedSource.border_segments[1]!.sources[0] = { ...reusedSource.border_segments[0]!.sources[0]! }
    const outOfOrder = structuredClone(plan) as unknown as { border_segments: Array<Record<string, unknown>> }
    outOfOrder.border_segments.reverse()
    for (const candidate of [
      resealPlan(duplicate as unknown as NativeSheetDecorationPlanV1),
      resealPlan(wrongEdge as unknown as NativeSheetDecorationPlanV1),
      resealPlan(reversedJoin as unknown as NativeSheetDecorationPlanV1),
      resealPlan(sameCellJoin as unknown as NativeSheetDecorationPlanV1),
      resealPlan(nonAdjacentJoin as unknown as NativeSheetDecorationPlanV1),
      resealPlan(reusedSource as unknown as NativeSheetDecorationPlanV1),
      resealPlan(outOfOrder as unknown as NativeSheetDecorationPlanV1),
    ]) {
      let pushes = 0
      expect(() => emitNativeSheetDecorationCommandsV1(candidate, { push() { pushes++ } })).toThrowError(expect.objectContaining({ code: 'decoration.planInvalid' }))
      expect(pushes).toBe(0)
      let callbacks = 0
      expect(() => replayNativeSheetDecorationCommandsV1({}, uncheckedCommands(candidate), { execute() { callbacks++ } })).toThrow(NativeSheetDecorationError)
      expect(callbacks).toBe(0)
    }
  })

  it.each([
    ['row hidden', (workbook: MutableWorkbook) => { workbook.sheets[0]!.rows[0]!.hidden = true }],
    ['row height zero', (workbook: MutableWorkbook) => { workbook.sheets[0]!.rows[0]!.height_points = 0 }],
    ['column hidden', (workbook: MutableWorkbook) => { workbook.sheets[0]!.columns[0]!.hidden = true }],
    ['column width zero', (workbook: MutableWorkbook) => { workbook.sheets[0]!.columns[0]!.width = 0 }],
  ])('fails closed for styled zero-size dimensions: %s', (_name, mutate) => {
    const workbook = fixture() as unknown as MutableWorkbook
    mutate(workbook)
    const native = workbook as unknown as NativeWorkbookV1
    const model = projectNativeWorkbookV1(native)
    const geometry = compileNativeSheetGeometryV1(model, '7', { row: 0, column: 0, end_row: 0, end_column: 1 }, metric(native))
    expect(() => compileNativeSheetDecorationsV1(model, geometry)).toThrowError(expect.objectContaining({ code: 'decoration.hiddenDimension' }))
  })

  it.each([
    ['CONDITIONAL_FORMATTING', 'conditional-formatting', 'sheet:7', 'Worksheets/Sheet1.xml', undefined],
    ['TABLE_REFERENCE', 'tables', 'sheet:7', 'Worksheets/Sheet1.xml', undefined],
    ['DRAWING_REFERENCE', 'drawings', 'sheet:7', 'Worksheets/Sheet1.xml', undefined],
    ['EXTERNAL_RELATIONSHIP', 'external-links', 'workbook', 'XL/_rels/Workbook.xml.rels', undefined],
    ['FORMULA_ATTRIBUTES', 'formula-groups', 'sheet:7', 'Worksheets/Sheet1.xml', 'A1'],
  ])('refuses every source-authoritative appearance class: %s', (code, capability, scopeID, partName, cellRef) => {
    const workbook = fixture() as unknown as MutableWorkbook
    if (cellRef) workbook.sheets[0]!.cells[0] = { row: 0, column: 0, ref: 'A1', style_id: 0, formula: { text: '1+1', type: 'normal' }, editable: false }
    workbook.unsupported.push(issue(code, capability, scopeID, partName, cellRef ? { cell_ref: cellRef } : {}))
    const native = workbook as unknown as NativeWorkbookV1
    const { model, geometry } = setup(native)
    expect(() => compileNativeSheetDecorationsV1(model, geometry)).toThrowError(expect.objectContaining({ code: 'decoration.sourceUnsupported' }))
  })

  it('refuses generic worksheet-root authority even when its inventory is internally valid', () => {
    const workbook = fixture() as unknown as MutableWorkbook
    workbook.sheets[0]!.editable = false
    workbook.sheets[0]!.refusal_code = 'UNSAFE_WORKSHEET_ATTRIBUTES'
    workbook.unsupported.push(issue('WORKSHEET_ATTRIBUTES', 'worksheet-features', 'sheet:7', 'Worksheets/Sheet1.xml'))
    const native = workbook as unknown as NativeWorkbookV1
    const model = projectNativeWorkbookV1(native)
    expect(() => compileNativeSheetGeometryV1(model, '7', { row: 0, column: 0, end_row: 0, end_column: 1 }, metric(native))).toThrowError(expect.objectContaining({ code: 'geometry.sourceUnsupported' }))
  })

  it('runs an unmodified Excel-authored Go extraction through project, geometry, decorations, commands, and deterministic replay', () => {
    const decoded = decodeNativeWorkbookV1(excelFixtureSource)
    expect(decoded.ok).toBe(true)
    if (!decoded.ok) return
    const workbook = decoded.value
    const model = projectNativeWorkbookV1(workbook)
    const geometry = compileNativeSheetGeometryV1(model, '5', { row: 1, column: 11, end_row: 1, end_column: 11 }, metric(workbook))
    const first = compileNativeSheetDecorationsV1(model, geometry)
    const second = compileNativeSheetDecorationsV1(model, geometry)
    expect(JSON.stringify(first)).toBe(JSON.stringify(second))
    expect(first.fills).toHaveLength(1)
    expect(first.fills[0]).toMatchObject({ cell_ref: 'L2', color: '#FFFF00', fill_id: 2 })
    const commands = record(first)
    const replayed: string[] = []
    replayNativeSheetDecorationCommandsV1(replayed, commands, { execute(target, command) { target.push(JSON.stringify(command)) } })
    expect(replayed).toEqual(commands.map((command) => JSON.stringify(command)))
  })
})

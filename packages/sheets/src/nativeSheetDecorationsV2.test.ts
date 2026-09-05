import { createHash } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { createRequire } from 'node:module'
import { describe, expect, it } from 'vitest'
import {
  NATIVE_SHEET_DECORATION_V2_LIMITS,
  NativeSheetDecorationV2Error as NativeSheetDecorationError,
  NativeWorkbookV2ValidationError,
  compileNativeSheetDecorationsV2,
  compileNativeSheetGeometryV2,
  createNativeMaximumDigitWidthAuthorityV2,
  createNativeSheetDecorationRecordingSurfaceV2,
  decodeNativeWorkbookV2,
  emitNativeSheetDecorationCommandsV2,
  nativeWorkbookStyleRawProjectionSha256V2,
  projectNativeWorkbookV2,
  replayNativeSheetDecorationCommandsV2,
  validateNativeSheetDecorationPlanV2,
} from './index.js'
import type {
  NativeMaximumDigitWidthAuthorityV2,
  NativeSheetDecorationCommandV2,
  NativeSheetDecorationPlanV2,
  NativeSheetGeometryV2,
  NativeWorkbookRenderModelV2,
  NativeWorkbookUnsupportedV2,
  NativeWorkbookV2,
} from './index.js'

const repositoryRoot = resolve(import.meta.dirname, '../../..')
const require = createRequire(import.meta.url)
const FONT_BYTES = new Uint8Array(readFileSync(require.resolve('dejavu-fonts-ttf/ttf/DejaVuSans.ttf')))
const fixtureSource = readFileSync(resolve(repositoryRoot, 'go/xlsxpatch/testdata/native-xlsx-v2/valid/lexical-render.json'), 'utf8').trim()
const excelFixtureSource = readFileSync(resolve(repositoryRoot, 'go/xlsxpatch/testdata/native-xlsx-v2/valid/excel-authored-happy-tree.json'), 'utf8').trim()

function fixture(): NativeWorkbookV2 {
  const workbook = JSON.parse(fixtureSource) as NativeWorkbookV2
  const mutable = workbook as unknown as MutableWorkbook
  mutable.normal_style = {
    style_xf_id: 0, font_id: 0, font_name: 'DejaVu Sans', font_size_points: 11, font_bold: false, font_italic: false,
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
  normal_style?: NativeWorkbookV2['normal_style']
  styles: Array<{ id: number; effective: MutableEffective; raw_projection_sha256: string }>
  sheets: Array<{
    id: string; part_name: string; editable: boolean; refusal_code?: string
    rows: Array<{ row: number; height_points?: number; hidden: boolean; custom_height: boolean; style_id?: number }>
    columns: Array<{ column: number; end_column: number; width?: number; hidden: boolean; custom_width: boolean; best_fit: boolean; style_id?: number }>
    cells: Array<Record<string, unknown>>
    merged_ranges: Array<Record<string, unknown>>
  }>
  unsupported: NativeWorkbookUnsupportedV2[]
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
  for (const style of workbook.styles) style.raw_projection_sha256 = nativeWorkbookStyleRawProjectionSha256V2(style.effective as NativeWorkbookV2['styles'][number]['effective'])
}

function freezeRecursively(value: unknown): void {
  if (typeof value !== 'object' || value === null || Object.isFrozen(value)) return
  for (const child of Object.values(value)) freezeRecursively(child)
  Object.freeze(value)
}

function metric(workbook: NativeWorkbookV2): NativeMaximumDigitWidthAuthorityV2 {
  return createNativeMaximumDigitWidthAuthorityV2(projectNativeWorkbookV2(workbook), FONT_BYTES)
}

function setup(workbook: NativeWorkbookV2 = fixture()): { workbook: NativeWorkbookV2; model: NativeWorkbookRenderModelV2; geometry: NativeSheetGeometryV2 } {
  const model = projectNativeWorkbookV2(workbook)
  const geometry = compileNativeSheetGeometryV2(model, '7', { row: 0, column: 0, end_row: 0, end_column: 1 }, metric(workbook))
  return { workbook, model, geometry }
}

function issue(code: string, capability: string, scopeID: string, partName: string, extra: Partial<NativeWorkbookUnsupportedV2> = {}): NativeWorkbookUnsupportedV2 {
  const location = [code, capability, scopeID, partName, extra.cell_ref ?? '', extra.range_ref ?? ''].join('\0')
  return {
    id: `unsupported:${createHash('sha256').update(location).digest('hex')}`,
    code, capability, scope_id: scopeID, part_name: partName,
    preservation: 'preserve-exact', message: 'source-authoritative appearance', ...extra,
  }
}

function record(plan: NativeSheetDecorationPlanV2): ReadonlyArray<NativeSheetDecorationCommandV2> {
  const surface = createNativeSheetDecorationRecordingSurfaceV2()
  emitNativeSheetDecorationCommandsV2(plan, surface)
  return surface.finish()
}

function resealPlan(plan: NativeSheetDecorationPlanV2): NativeSheetDecorationPlanV2 {
  const { decoration_sha256: _discarded, ...unsigned } = structuredClone(plan)
  return { ...unsigned, decoration_sha256: `sha256:${createHash('sha256').update(`injoffice.xlsx.sheet-decoration.v1\0${JSON.stringify(unsigned)}`).digest('hex')}` }
}

function uncheckedCommands(plan: NativeSheetDecorationPlanV2): NativeSheetDecorationCommandV2[] {
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
    const first = compileNativeSheetDecorationsV2(model, geometry)
    const second = compileNativeSheetDecorationsV2(model, geometry)
    expect(JSON.stringify(first)).toBe(JSON.stringify(second))
    expect(first.decoration_sha256).toMatch(/^sha256:[0-9a-f]{64}$/)
    expect(first.fills).toEqual([{
      cell_ref: 'B1', row: 0, column: 1, style_id: 1, fill_id: 1, fill_record_sha256: `sha256:${'4'.repeat(64)}`,
      rect: { x_emu: 1_066_800, y_emu: 0, width_emu: 1_066_800, height_emu: 231_775 }, color: '#DDEEFF',
    }])
    const join = first.border_segments.find((segment) => segment.x1_emu === 1_066_800 && segment.x2_emu === 1_066_800)
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

  it('uses a one-cell source halo to prove matching, outside-only, and conflicting shared edges', () => {
    const workbook = fixture()
    const model = projectNativeWorkbookV2(workbook)
    const geometry = compileNativeSheetGeometryV2(model, '7', { row: 0, column: 0, end_row: 0, end_column: 0 }, metric(workbook))
    const matching = compileNativeSheetDecorationsV2(model, geometry)
    const right = matching.border_segments.find((segment) => segment.orientation === 'vertical' && segment.x1_emu === geometry.bounds.width_emu)
    expect(right?.sources).toEqual([
      { cell_ref: 'A1', edge: 'right', style_id: 0, border_id: 0, border_record_sha256: `sha256:${'1'.repeat(64)}` },
      { cell_ref: 'B1', edge: 'left', style_id: 1, border_id: 1, border_record_sha256: `sha256:${'2'.repeat(64)}` },
    ])

    const outsideOnly = fixture() as unknown as MutableWorkbook
    outsideOnly.sheets[0]!.cells.push({ row: 1, column: 0, ref: 'A2', style_id: 1, editable: true })
    const outsideNative = outsideOnly as unknown as NativeWorkbookV2
    const outsideModel = projectNativeWorkbookV2(outsideNative)
    const outsideGeometry = compileNativeSheetGeometryV2(outsideModel, '7', { row: 0, column: 0, end_row: 0, end_column: 0 }, metric(outsideNative))
    const bottom = compileNativeSheetDecorationsV2(outsideModel, outsideGeometry).border_segments.find((segment) => segment.orientation === 'horizontal' && segment.y1_emu === outsideGeometry.bounds.height_emu)
    expect(bottom).toMatchObject({ border_style: 'double', color: '#405060', sources: [{ cell_ref: 'A2', edge: 'top' }] })

    const conflicting = fixture() as unknown as MutableWorkbook
    conflicting.styles[1]!.effective.border!.left!.color = '#FFFFFF'
    sealStyles(conflicting)
    const conflictNative = conflicting as unknown as NativeWorkbookV2
    const conflictModel = projectNativeWorkbookV2(conflictNative)
    const conflictGeometry = compileNativeSheetGeometryV2(conflictModel, '7', { row: 0, column: 0, end_row: 0, end_column: 0 }, metric(conflictNative))
    expect(() => compileNativeSheetDecorationsV2(conflictModel, conflictGeometry)).toThrowError(expect.objectContaining({ code: 'decoration.borderConflict' }))
  })

  it('validates the complete plan before the first push and validates replay before the first callback', () => {
    const { model, geometry } = setup()
    const plan = compileNativeSheetDecorationsV2(model, geometry)
    const commands = record(plan)
    expect(commands.map((command) => command.kind)).toEqual([
      'beginDecorations', 'clipRect', 'fillRect', 'borderSegment', 'borderSegment', 'borderSegment', 'borderSegment', 'endDecorations',
    ])
    expect(commands[0]).toMatchObject({
      kind: 'beginDecorations', protocol: 'injoffice.xlsx.sheet-decoration', version: 1,
      source_part: 'Worksheets/Sheet1.xml', source_revision: model.revision,
    })
    const replayed: string[] = []
    replayNativeSheetDecorationCommandsV2(replayed, commands, { execute(target, command) { target.push(command.kind) } })
    expect(replayed).toEqual(commands.map((command) => command.kind))

    const stalePlan = structuredClone(plan) as unknown as { fills: Array<{ color: string }> }
    stalePlan.fills[0]!.color = '#ABCDEF'
    let pushes = 0
    expect(() => emitNativeSheetDecorationCommandsV2(stalePlan as unknown as NativeSheetDecorationPlanV2, { push() { pushes++ } })).toThrowError(expect.objectContaining({ code: 'decoration.planDigest' }))
    expect(pushes).toBe(0)

    const staleCommands = structuredClone(commands)
    const fill = staleCommands.find((command) => command.kind === 'fillRect') as Extract<NativeSheetDecorationCommandV2, { kind: 'fillRect' }>
    ;(fill.decoration as { color: string }).color = '#ABCDEF'
    let callbacks = 0
    expect(() => replayNativeSheetDecorationCommandsV2({}, staleCommands, { execute() { callbacks++ } })).toThrowError(expect.objectContaining({ code: 'decoration.planDigest' }))
    expect(callbacks).toBe(0)

    const capacity = createNativeSheetDecorationRecordingSurfaceV2(commands.length)
    emitNativeSheetDecorationCommandsV2(plan, capacity, commands.length)
    expect(capacity.finish()).toHaveLength(commands.length)
    const tooSmall = createNativeSheetDecorationRecordingSurfaceV2(commands.length - 1)
    expect(() => emitNativeSheetDecorationCommandsV2(plan, tooSmall, commands.length)).toThrowError(expect.objectContaining({ code: 'decoration.commandBudget' }))
    expect(tooSmall.commands).toEqual([])
  })

  it('rejects malformed, oversized, fractional, escaping, and source-forged plans atomically', () => {
    const { model, geometry } = setup()
    const plan = compileNativeSheetDecorationsV2(model, geometry)
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
    oversized.fills = new Array(NATIVE_SHEET_DECORATION_V2_LIMITS.maxFills + 1)
    cases.push(oversized)
    const oversizedBorders = structuredClone(plan) as unknown as { border_segments: unknown[] }
    oversizedBorders.border_segments = new Array(NATIVE_SHEET_DECORATION_V2_LIMITS.maxBorderSegments + 1)
    cases.push(oversizedBorders)
    for (const candidate of cases) {
      let pushes = 0
      expect(() => emitNativeSheetDecorationCommandsV2(candidate as NativeSheetDecorationPlanV2, { push() { pushes++ } })).toThrow(NativeSheetDecorationError)
      expect(pushes).toBe(0)
    }
    expect(() => validateNativeSheetDecorationPlanV2({})).toThrowError(expect.objectContaining({ code: 'decoration.planInvalid' }))

    const commands = structuredClone(record(plan))
    const malformedStreams: unknown[] = [commands.slice(1), [...commands, ...new Array(NATIVE_SHEET_DECORATION_V2_LIMITS.maxCommands).fill({ kind: 'endDecorations' })]]
    for (const stream of malformedStreams) {
      let callbacks = 0
      expect(() => replayNativeSheetDecorationCommandsV2({}, stream as NativeSheetDecorationCommandV2[], { execute() { callbacks++ } })).toThrow(NativeSheetDecorationError)
      expect(callbacks).toBe(0)
    }
  })

  it('rejects mutable handcrafted models and valid-looking style tampering with a stale raw projection digest', () => {
    const workbook = fixture()
    const model = projectNativeWorkbookV2(workbook)
    const geometry = compileNativeSheetGeometryV2(model, '7', { row: 0, column: 0, end_row: 0, end_column: 1 }, metric(workbook))
    const handcrafted = structuredClone(model)
    expect(() => compileNativeSheetGeometryV2(handcrafted, '7', geometry.viewport, metric(workbook))).toThrowError(expect.objectContaining({ code: 'geometry.sourceUnsupported' }))
    expect(() => compileNativeSheetDecorationsV2(handcrafted, geometry)).toThrowError(expect.objectContaining({ code: 'decoration.modelAuthority' }))
    for (const symbol of Object.getOwnPropertySymbols(model)) Object.defineProperty(handcrafted, symbol, { value: true })
    freezeRecursively(handcrafted)
    expect(() => compileNativeSheetDecorationsV2(handcrafted, geometry)).toThrowError(expect.objectContaining({ code: 'decoration.modelAuthority' }))
    const skeletal = { protocol: model.protocol, version: model.version, document_id: model.document_id, revision: model.revision, styles: [], sheets: [] }
    expect(() => compileNativeSheetGeometryV2(skeletal as unknown as NativeWorkbookRenderModelV2, '7', geometry.viewport, metric(workbook))).toThrowError(expect.objectContaining({ code: 'geometry.sourceUnsupported' }))
    expect(() => compileNativeSheetDecorationsV2(skeletal as unknown as NativeWorkbookRenderModelV2, geometry)).toThrowError(expect.objectContaining({ code: 'decoration.modelAuthority' }))
    const oversizedHandcrafted = { ...model, styles: new Array(100_001).fill(model.styles[0]) }
    expect(() => compileNativeSheetGeometryV2(oversizedHandcrafted as NativeWorkbookRenderModelV2, '7', geometry.viewport, metric(workbook))).toThrowError(expect.objectContaining({ code: 'geometry.sourceUnsupported' }))
    expect(() => compileNativeSheetDecorationsV2(oversizedHandcrafted as NativeWorkbookRenderModelV2, geometry)).toThrowError(expect.objectContaining({ code: 'decoration.modelAuthority' }))
    const oversizedNative = structuredClone(workbook) as unknown as MutableWorkbook
    oversizedNative.styles = new Array(100_001).fill(oversizedNative.styles[0])
    expect(() => projectNativeWorkbookV2(oversizedNative)).toThrow(NativeWorkbookV2ValidationError)

    for (const tamper of [
      (candidate: MutableWorkbook) => { candidate.styles[1]!.effective.fill_color = '#ABCDEF'; candidate.styles[1]!.effective.fill!.color = '#ABCDEF' },
      (candidate: MutableWorkbook) => { candidate.styles[1]!.effective.border!.top!.style = 'thin' },
      (candidate: MutableWorkbook) => { candidate.styles[1]!.effective.border!.top!.color = '#ABCDEF' },
    ]) {
      const candidate = structuredClone(workbook) as unknown as MutableWorkbook
      tamper(candidate)
      expect(() => projectNativeWorkbookV2(candidate)).toThrow(NativeWorkbookV2ValidationError)
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
    const { model, geometry } = setup(workbook as unknown as NativeWorkbookV2)
    const plan = compileNativeSheetDecorationsV2(model, geometry)
    expect(plan.fills).toHaveLength(1)
    expect(plan.border_segments).toHaveLength(4)
  })

  it('requires an exact bijection between partial style codes and source diagnostics', () => {
    const missing = fixture() as unknown as MutableWorkbook
    missing.styles[1]!.effective.projection = 'partial'
    missing.styles[1]!.effective.unsupported = ['font-color']
    sealStyles(missing)
    expect(() => projectNativeWorkbookV2(missing)).toThrow(NativeWorkbookV2ValidationError)

    const contradictory = fixture() as unknown as MutableWorkbook
    contradictory.unsupported.push(issue('STYLE_BORDER', 'styles', 'style:1', 'XL/Styles.xml'))
    expect(() => projectNativeWorkbookV2(contradictory)).toThrow(NativeWorkbookV2ValidationError)
  })

  it('rejects recomputed-digest duplicate and orientation-invalid border topology atomically', () => {
    const { model, geometry } = setup()
    const plan = compileNativeSheetDecorationsV2(model, geometry)
    const duplicate = structuredClone(plan) as unknown as { border_segments: Array<Record<string, unknown>> }
    const conflicting = structuredClone(duplicate.border_segments[0]!)
    conflicting.color = '#ABCDEF'
    duplicate.border_segments.splice(1, 0, conflicting)
    const wrongEdge = structuredClone(plan) as unknown as { border_segments: Array<{ orientation: string; sources: Array<{ edge: string }> }> }
    const horizontal = wrongEdge.border_segments.find((segment) => segment.orientation === 'horizontal')!
    horizontal.sources[0]!.edge = 'left'
    const repeatedSource = structuredClone(plan) as unknown as { border_segments: Array<{ sources: Array<Record<string, unknown>> }> }
    repeatedSource.border_segments[1]!.sources[0] = structuredClone(repeatedSource.border_segments[0]!.sources[0]!)
    const nonAdjacent = structuredClone(plan) as unknown as { border_segments: Array<{ sources: Array<{ cell_ref: string; edge: string }> }> }
    const joined = nonAdjacent.border_segments.find((segment) => segment.sources.length === 2)!
    joined.sources[1]!.cell_ref = 'C1'
    for (const candidate of [duplicate, wrongEdge, repeatedSource, nonAdjacent].map((value) => resealPlan(value as unknown as NativeSheetDecorationPlanV2))) {
      let pushes = 0
      expect(() => emitNativeSheetDecorationCommandsV2(candidate, { push() { pushes++ } })).toThrowError(expect.objectContaining({ code: 'decoration.planInvalid' }))
      expect(pushes).toBe(0)
      let callbacks = 0
      expect(() => replayNativeSheetDecorationCommandsV2({}, uncheckedCommands(candidate), { execute() { callbacks++ } })).toThrow(NativeSheetDecorationError)
      expect(callbacks).toBe(0)
    }
  })

  it('refuses accessor and Proxy plans or callbacks without invoking them', () => {
    const { model, geometry } = setup()
    const plan = compileNativeSheetDecorationsV2(model, geometry)
    let getterCalls = 0
    const accessor = structuredClone(plan) as NativeSheetDecorationPlanV2 & { trap?: string }
    Object.defineProperty(accessor, 'trap', { enumerable: true, get() { getterCalls++; return 'unsafe' } })
    expect(() => emitNativeSheetDecorationCommandsV2(accessor, { push() {} })).toThrow(NativeSheetDecorationError)
    expect(getterCalls).toBe(0)
    let reads = 0
    const proxied = new Proxy(plan, { get(target, property, receiver) { reads++; return Reflect.get(target, property, receiver) } })
    expect(() => emitNativeSheetDecorationCommandsV2(proxied, { push() {} })).toThrow(NativeSheetDecorationError)
    expect(reads).toBe(0)
    const surface = {}
    Object.defineProperty(surface, 'push', { enumerable: true, get() { getterCalls++; return () => undefined } })
    expect(() => emitNativeSheetDecorationCommandsV2(plan, surface as never)).toThrowError(expect.objectContaining({ code: 'decoration.commandInvalid' }))
    expect(getterCalls).toBe(0)
  })

  it.each([
    ['row hidden', (workbook: MutableWorkbook) => { workbook.sheets[0]!.rows[0]!.hidden = true }],
    ['row height zero', (workbook: MutableWorkbook) => { workbook.sheets[0]!.rows[0]!.height_points = 0 }],
    ['column hidden', (workbook: MutableWorkbook) => { workbook.sheets[0]!.columns[0]!.hidden = true }],
    ['column width zero', (workbook: MutableWorkbook) => { workbook.sheets[0]!.columns[0]!.width = 0 }],
  ])('fails closed for styled zero-size dimensions: %s', (_name, mutate) => {
    const workbook = fixture() as unknown as MutableWorkbook
    mutate(workbook)
    const native = workbook as unknown as NativeWorkbookV2
    const model = projectNativeWorkbookV2(native)
    const geometry = compileNativeSheetGeometryV2(model, '7', { row: 0, column: 0, end_row: 0, end_column: 1 }, metric(native))
    expect(() => compileNativeSheetDecorationsV2(model, geometry)).toThrowError(expect.objectContaining({ code: 'decoration.hiddenDimension' }))
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
    const native = workbook as unknown as NativeWorkbookV2
    const { model, geometry } = setup(native)
    expect(() => compileNativeSheetDecorationsV2(model, geometry)).toThrowError(expect.objectContaining({ code: 'decoration.sourceUnsupported' }))
  })

  it('refuses generic worksheet-root authority even when its inventory is internally valid', () => {
    const workbook = fixture() as unknown as MutableWorkbook
    workbook.sheets[0]!.editable = false
    workbook.sheets[0]!.refusal_code = 'UNSAFE_WORKSHEET_ATTRIBUTES'
    workbook.unsupported.push(issue('WORKSHEET_ATTRIBUTES', 'worksheet-features', 'sheet:7', 'Worksheets/Sheet1.xml'))
    const native = workbook as unknown as NativeWorkbookV2
    const model = projectNativeWorkbookV2(native)
    expect(() => compileNativeSheetGeometryV2(model, '7', { row: 0, column: 0, end_row: 0, end_column: 1 }, metric(native))).toThrowError(expect.objectContaining({ code: 'geometry.sourceUnsupported' }))
  })

  it('projects the unmodified Excel-authored bridge with lockstep v1 Normal identity and refuses a mismatched font metric', () => {
    const decoded = decodeNativeWorkbookV2(excelFixtureSource)
    expect(decoded.ok).toBe(true)
    if (!decoded.ok) return
    const workbook = decoded.value
    const model = projectNativeWorkbookV2(workbook)
    expect(model.normal_style?.font_name).toBe('Calibri')
    expect(() => metric(workbook)).toThrowError(expect.objectContaining({ code: 'geometry.metricAuthority' }))
  })
})

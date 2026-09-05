import type { PrintConfigurationSnapshotV1, PrintIssue, PrintLayoutConfig, PrintMargins, PrintRange, PrintRenderConfig, PrintResult } from './types'

const AREAS = new Set(['CurrentSheet', 'Workbook', 'CurrentSelection', 'AllSelection'])
const PAPERS = new Set(['Letter', 'Tabloid', 'Legal', 'Statement', 'Executive', 'Folio', 'A3', 'A4', 'A5', 'B4', 'B5', 'Custom'])
const DIRECTIONS = new Set(['Portrait', 'Landscape'])
const SCALES = new Set(['Origin', 'FitWidth', 'FitHeight', 'FitPage', 'Custom'])
const FREEZE = new Set(['Row', 'Column'])
const MARGINS = new Set(['Normal', 'Narrow', 'Wide', 'None', 'Custom'])
const ALIGN = new Set(['Start', 'End', 'Middle'])
const HEADER_FOOTER = new Set(['PageSize', 'WorkbookTitle', 'WorksheetTitle', 'Date', 'Time'])

const issue = (issues: PrintIssue[], path: string, code: PrintIssue['code'], message: string) => issues.push({ path, code, message })
const finitePositive = (value: number) => Number.isFinite(value) && value > 0

function validateRange(range: PrintRange, path: string, issues: PrintIssue[]): void {
  const values = [range.startRow, range.startColumn, range.endRow, range.endColumn]
  if (!values.every(Number.isSafeInteger) || range.startRow < 0 || range.startColumn < 0 || range.endRow < range.startRow || range.endColumn < range.startColumn || range.endRow > 1_048_575 || range.endColumn > 16_383) {
    issue(issues, path, 'INVALID_RANGE', 'range must be a bounded zero-based inclusive Excel range')
  }
}

function validateMargins(value: PrintMargins | undefined, path: string, issues: PrintIssue[]): void {
  if (!value) {
    issue(issues, path, 'REQUIRED', 'custom margins are required')
    return
  }
  for (const [name, margin] of Object.entries(value)) {
    if (!Number.isFinite(margin) || margin < 0 || margin > 20) issue(issues, `${path}/${name}`, 'INVALID_NUMBER', 'margin must be within 0..20 inches')
  }
}

function validateRepeatedRange(value: unknown, axis: 'row' | 'column', path: string, issues: PrintIssue[]): void {
  if (value === undefined) return
  if (!value || typeof value !== 'object') {
    issue(issues, path, 'INVALID_RANGE', `repeated ${axis} range must be an object`)
    return
  }
  const start = (value as Record<string, unknown>)[axis === 'row' ? 'startRow' : 'startColumn']
  const end = (value as Record<string, unknown>)[axis === 'row' ? 'endRow' : 'endColumn']
  const max = axis === 'row' ? 1_048_575 : 16_383
  if (!Number.isSafeInteger(start) || !Number.isSafeInteger(end) || (start as number) < 0 || (end as number) < (start as number) || (end as number) > max) issue(issues, path, 'INVALID_RANGE', `repeated ${axis} range must be zero-based, inclusive, and bounded`)
}

function validateWatermark(value: PrintRenderConfig['watermark'], issues: PrintIssue[]): void {
  if (value === undefined) return
  if (!value || typeof value !== 'object' || (value.kind !== 'text' && value.kind !== 'image')) {
    issue(issues, '/watermark', 'INVALID_ENUM', 'watermark must be a text or image watermark')
    return
  }
  if (value.kind === 'text' && (typeof value.text !== 'string' || value.text.length === 0 || value.text.length > 512)) issue(issues, '/watermark/text', 'INVALID_NUMBER', 'watermark text must contain 1-512 characters')
  if (value.kind === 'image' && (typeof value.sourceId !== 'string' || value.sourceId.length === 0 || value.sourceId.length > 512)) issue(issues, '/watermark/sourceId', 'REQUIRED', 'watermark image sourceId must contain 1-512 characters')
  if (value.opacity !== undefined && (!Number.isFinite(value.opacity) || value.opacity < 0 || value.opacity > 1)) issue(issues, '/watermark/opacity', 'INVALID_NUMBER', 'watermark opacity must be within 0..1')
  if (value.kind === 'text') {
    if (value.rotation !== undefined && (!Number.isFinite(value.rotation) || value.rotation < -180 || value.rotation > 180)) issue(issues, '/watermark/rotation', 'INVALID_NUMBER', 'watermark rotation must be within -180..180 degrees')
    if (value.fontSize !== undefined && (!Number.isFinite(value.fontSize) || value.fontSize <= 0 || value.fontSize > 1_000)) issue(issues, '/watermark/fontSize', 'INVALID_NUMBER', 'watermark font size must be within 0..1000 points')
    if (value.color !== undefined && (typeof value.color !== 'string' || !/^#[0-9A-Fa-f]{6}(?:[0-9A-Fa-f]{2})?$/.test(value.color))) issue(issues, '/watermark/color', 'INVALID_NUMBER', 'watermark color must be #RRGGBB or #RRGGBBAA')
  } else if (value.scale !== undefined && (!Number.isFinite(value.scale) || value.scale <= 0 || value.scale > 100)) issue(issues, '/watermark/scale', 'INVALID_NUMBER', 'watermark image scale must be within 0..100')
}

export function validatePrintLayout(input: PrintLayoutConfig): PrintResult<PrintLayoutConfig> {
  const issues: PrintIssue[] = []
  if (!AREAS.has(input.area)) issue(issues, '/area', 'INVALID_ENUM', 'unsupported print area')
  if (!PAPERS.has(input.paperSize)) issue(issues, '/paperSize', 'INVALID_ENUM', 'unsupported paper size')
  if (!DIRECTIONS.has(input.direction)) issue(issues, '/direction', 'INVALID_ENUM', 'unsupported direction')
  if (!SCALES.has(input.scale)) issue(issues, '/scale', 'INVALID_ENUM', 'unsupported scale')
  if (!MARGINS.has(input.margin)) issue(issues, '/margin', 'INVALID_ENUM', 'unsupported margin preset')
  const targets = Array.isArray(input.subUnitIds) ? input.subUnitIds : []
  if (targets.length === 0) issue(issues, '/subUnitIds', 'REQUIRED', 'at least one sheet target is required')
  const seen = new Set<string>()
  targets.forEach((target, index) => {
    const id = typeof target === 'string' ? target : target?.id
    if (typeof id !== 'string' || id.length === 0 || id.length > 128) issue(issues, `/subUnitIds/${index}`, 'REQUIRED', 'sheet id must contain 1-128 characters')
    else if (seen.has(id) && typeof target === 'string') issue(issues, `/subUnitIds/${index}`, 'DUPLICATE', `duplicate whole-sheet target ${id}`)
    else seen.add(id)
    if (typeof target === 'object' && target?.range) validateRange(target.range, `/subUnitIds/${index}/range`, issues)
  })
  if (!Number.isSafeInteger(input.customScale) || input.customScale < 10 || input.customScale > 400) issue(issues, '/customScale', 'INVALID_NUMBER', 'custom scale must be an integer within 10..400')
  if (input.fitToWidthPages !== undefined && (!Number.isSafeInteger(input.fitToWidthPages) || input.fitToWidthPages < 0 || input.fitToWidthPages > 100_000)) issue(issues, '/fitToWidthPages', 'INVALID_NUMBER', 'fit width pages must be an integer within 0..100000')
  if (input.fitToHeightPages !== undefined && (!Number.isSafeInteger(input.fitToHeightPages) || input.fitToHeightPages < 0 || input.fitToHeightPages > 100_000)) issue(issues, '/fitToHeightPages', 'INVALID_NUMBER', 'fit height pages must be an integer within 0..100000')
  if (!Number.isSafeInteger(input.maxRowsEachPage) || input.maxRowsEachPage < 0) issue(issues, '/maxRowsEachPage', 'INVALID_NUMBER', 'row page limit must be a non-negative integer')
  if (!Number.isSafeInteger(input.maxColumnsEachPage) || input.maxColumnsEachPage < 0) issue(issues, '/maxColumnsEachPage', 'INVALID_NUMBER', 'column page limit must be a non-negative integer')
  const freeze = Array.isArray(input.freeze) ? input.freeze : []
  if (!Array.isArray(input.freeze) || new Set(freeze).size !== freeze.length || freeze.some((value) => !FREEZE.has(value))) issue(issues, '/freeze', 'INVALID_ENUM', 'freeze values must be unique Row or Column values')
  if (input.margin === 'Custom') validateMargins(input.customMargins, '/customMargins', issues)
  if (input.paperSize === 'Custom' && (!input.pageSizeCustom || !finitePositive(input.pageSizeCustom.width) || !finitePositive(input.pageSizeCustom.height))) issue(issues, '/pageSizeCustom', 'INVALID_NUMBER', 'custom paper requires positive width and height')
  validateRepeatedRange(input.repeatRows, 'row', '/repeatRows', issues)
  validateRepeatedRange(input.repeatColumns, 'column', '/repeatColumns', issues)
  return issues.length ? { ok: false, issues } : { ok: true, value: structuredClone(input) }
}

export function validatePrintRender(input: PrintRenderConfig): PrintResult<PrintRenderConfig> {
  const issues: PrintIssue[] = []
  if (typeof input.gridlines !== 'boolean') issue(issues, '/gridlines', 'INVALID_ENUM', 'gridlines must be boolean')
  if (typeof input.headings !== 'boolean') issue(issues, '/headings', 'INVALID_ENUM', 'headings must be boolean')
  if (input.isCustomHeaderFooter !== undefined && typeof input.isCustomHeaderFooter !== 'boolean') issue(issues, '/isCustomHeaderFooter', 'INVALID_ENUM', 'custom header/footer flag must be boolean')
  if (!ALIGN.has(input.hAlign)) issue(issues, '/hAlign', 'INVALID_ENUM', 'unsupported horizontal alignment')
  if (!ALIGN.has(input.vAlign)) issue(issues, '/vAlign', 'INVALID_ENUM', 'unsupported vertical alignment')
  const headerFooter = Array.isArray(input.headerFooter) ? input.headerFooter : []
  if (!Array.isArray(input.headerFooter) || new Set(headerFooter).size !== headerFooter.length || headerFooter.some((value) => !HEADER_FOOTER.has(value))) issue(issues, '/headerFooter', 'INVALID_ENUM', 'header/footer tokens must be unique supported values')
  if (!input.headerFooterSetting || typeof input.headerFooterSetting !== 'object') {
    issue(issues, '/headerFooterSetting', 'REQUIRED', 'header/footer settings are required')
  } else {
    for (const name of ['topLeft', 'topCenter', 'topRight', 'bottomLeft', 'bottomCenter', 'bottomRight'] as const) {
      const value = input.headerFooterSetting[name]
      if (typeof value !== 'string' || value.length > 1024) issue(issues, `/headerFooterSetting/${name}`, 'INVALID_NUMBER', 'header/footer text must not exceed 1024 characters')
    }
  }
  validateWatermark(input.watermark, issues)
  return issues.length ? { ok: false, issues } : { ok: true, value: structuredClone(input) }
}

/** Validate untrusted persisted settings without allowing malformed input or a
 * structured-clone failure to escape as an implementation exception. */
export function validatePrintConfigurationSnapshot(input: unknown): PrintResult<PrintConfigurationSnapshotV1> {
  if (!input || typeof input !== 'object' || Array.isArray(input)) return { ok: false, issues: [{ path: '/', code: 'INVALID_TYPE', message: 'print configuration must be an object' }] }
  const value = input as Partial<PrintConfigurationSnapshotV1>
  const issues: PrintIssue[] = []
  if (value.version !== 1) issue(issues, '/version', 'INVALID_ENUM', 'print configuration version must be 1')
  if (!value.layout || typeof value.layout !== 'object' || Array.isArray(value.layout)) issue(issues, '/layout', 'INVALID_TYPE', 'layout must be an object')
  if (!value.render || typeof value.render !== 'object' || Array.isArray(value.render)) issue(issues, '/render', 'INVALID_TYPE', 'render must be an object')
  if (issues.length) return { ok: false, issues }
  try {
    const layout = validatePrintLayout(value.layout as PrintLayoutConfig)
    const render = validatePrintRender(value.render as PrintRenderConfig)
    if (!layout.ok) issues.push(...layout.issues.map((entry) => ({ ...entry, path: `/layout${entry.path}` })))
    if (!render.ok) issues.push(...render.issues.map((entry) => ({ ...entry, path: `/render${entry.path}` })))
    if (!layout.ok || !render.ok) return { ok: false, issues }
    return { ok: true, value: { version: 1, layout: layout.value, render: render.value } }
  } catch {
    return { ok: false, issues: [{ path: '/', code: 'INVALID_TYPE', message: 'print configuration must be structured-cloneable' }] }
  }
}

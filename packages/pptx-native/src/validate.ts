import { PPTX_NATIVE_RESOURCE_LIMITS, PPTX_NATIVE_SCHEMA } from './schema.generated'
import type {
  NativeCompatibility,
  NativeElement,
  NativePptxDeck,
  NativeSourceAnchor,
  NativeTableCell,
} from './types'

export interface NativeValidationIssue { path: string; code: string; message: string }
export type NativeValidationResult =
  | { ok: true; value: NativePptxDeck }
  | { ok: false; issues: NativeValidationIssue[] }

type JsonSchema = Record<string, unknown>

interface SchemaBudget {
  nodes: number
  exhausted: boolean
}

interface SemanticBudget {
  elements: number
  textCodeUnits: number
  tableCells: number
  inlineAssetBase64CodeUnits: number
}

export function validateNativePptx(input: unknown): NativeValidationResult {
  const issues: NativeValidationIssue[] = []
  validateSchema(input, PPTX_NATIVE_SCHEMA as unknown as JsonSchema, '$', issues, 0, { nodes: 0, exhausted: false })
  if (issues.length) return { ok: false, issues }

  const deck = input as NativePptxDeck
  validateSemantics(deck, issues)
  return issues.length ? { ok: false, issues } : { ok: true, value: deck }
}

export function assertNativePptx(input: unknown): asserts input is NativePptxDeck {
  const result = validateNativePptx(input)
  if (!result.ok) {
    const detail = result.issues.slice(0, 8).map((issue) => `${issue.path}: ${issue.message}`).join('; ')
    throw new TypeError(`invalid native PPTX contract: ${detail}`)
  }
}

function validateSchema(value: unknown, schema: JsonSchema, path: string, issues: NativeValidationIssue[], depth: number, budget: SchemaBudget): void {
  if (budget.exhausted) return
  budget.nodes++
  if (budget.nodes > PPTX_NATIVE_RESOURCE_LIMITS.maxNodes) {
    budget.exhausted = true
    add(issues, path, 'native.resourceBudget', `contract exceeds ${PPTX_NATIVE_RESOURCE_LIMITS.maxNodes} JSON nodes`)
    return
  }
  if (depth > PPTX_NATIVE_RESOURCE_LIMITS.maxDepth) {
    add(issues, path, 'native.resourceDepth', `contract nesting exceeds ${PPTX_NATIVE_RESOURCE_LIMITS.maxDepth}`)
    return
  }
  if (typeof schema.$ref === 'string') {
    const prefix = '#/$defs/'
    if (!schema.$ref.startsWith(prefix)) return add(issues, path, 'schema.ref', `unsupported schema reference ${schema.$ref}`)
    const target = (PPTX_NATIVE_SCHEMA.$defs as unknown as Record<string, JsonSchema>)[schema.$ref.slice(prefix.length)]
    if (!target) return add(issues, path, 'schema.ref', `missing schema reference ${schema.$ref}`)
    return validateSchema(value, target, path, issues, depth, budget)
  }

  if (Array.isArray(schema.oneOf)) {
    const discriminated = discriminatedAlternative(value, schema.oneOf as JsonSchema[])
    if (discriminated) return validateSchema(value, discriminated, path, issues, depth, budget)
    let matches = 0
    for (const alternative of schema.oneOf as JsonSchema[]) {
      const candidate: NativeValidationIssue[] = []
      validateSchema(value, alternative, path, candidate, depth, budget)
      if (budget.exhausted) {
        add(issues, path, 'native.resourceBudget', `contract exceeds ${PPTX_NATIVE_RESOURCE_LIMITS.maxNodes} JSON nodes`)
        return
      }
      if (!candidate.length) matches++
    }
    if (matches !== 1) add(issues, path, 'schema.oneOf', 'must match exactly one element kind')
    return
  }

  if ('const' in schema && value !== schema.const) {
    add(issues, path, 'schema.const', `must equal ${JSON.stringify(schema.const)}`)
    return
  }
  if (Array.isArray(schema.enum) && !schema.enum.includes(value)) {
    add(issues, path, 'schema.enum', `must be one of ${schema.enum.join(', ')}`)
    return
  }

  switch (schema.type) {
    case 'object':
      if (!isRecord(value)) return add(issues, path, 'schema.type', 'must be an object')
      validateObject(value, schema, path, issues, depth, budget)
      return
    case 'array':
      if (!Array.isArray(value)) return add(issues, path, 'schema.type', 'must be an array')
      if (typeof schema.minItems === 'number' && value.length < schema.minItems) add(issues, path, 'schema.minItems', `must contain at least ${schema.minItems} item(s)`)
      if (typeof schema.maxItems === 'number' && value.length > schema.maxItems) add(issues, path, 'schema.maxItems', `must contain at most ${schema.maxItems} item(s)`)
      if (isRecord(schema.items)) {
        const count = Math.min(value.length, typeof schema.maxItems === 'number' ? schema.maxItems : value.length)
        for (let index = 0; index < count && !budget.exhausted; index++) validateSchema(value[index], schema.items as JsonSchema, `${path}[${index}]`, issues, depth + 1, budget)
      }
      return
    case 'string':
      if (typeof value !== 'string') return add(issues, path, 'schema.type', 'must be a string')
      if (typeof schema.minLength === 'number' && value.length < schema.minLength) add(issues, path, 'schema.minLength', 'is too short')
      if (typeof schema.maxLength === 'number' && value.length > schema.maxLength) {
        add(issues, path, 'schema.maxLength', 'is too long')
        return
      }
      if (typeof schema.pattern === 'string' && !new RegExp(schema.pattern).test(value)) {
        add(issues, path, 'schema.pattern', 'has an invalid format')
        return
      }
      if (schema['x-native-format'] === 'canonical-opc-part-name' && !isCanonicalOpcPartName(value)) {
        add(issues, path, 'schema.pattern', 'must be a safe canonical package-relative part name')
      }
      return
    case 'integer':
      if (!Number.isSafeInteger(value) || Object.is(value, -0)) return add(issues, path, 'schema.integer', 'must be a safe non-negative-zero integer')
      if (typeof schema.minimum === 'number' && (value as number) < schema.minimum) add(issues, path, 'schema.minimum', `must be at least ${schema.minimum}`)
      if (typeof schema.maximum === 'number' && (value as number) > schema.maximum) add(issues, path, 'schema.maximum', `must be at most ${schema.maximum}`)
      return
    case 'boolean':
      if (typeof value !== 'boolean') add(issues, path, 'schema.type', 'must be a boolean')
      return
  }
}

function validateObject(value: Record<string, unknown>, schema: JsonSchema, path: string, issues: NativeValidationIssue[], depth: number, budget: SchemaBudget): void {
  const properties = (schema.properties ?? {}) as Record<string, JsonSchema>
  for (const name of (schema.required ?? []) as string[]) {
    if (!Object.prototype.hasOwnProperty.call(value, name)) add(issues, `${path}.${name}`, 'schema.required', 'is required')
  }
  let propertyCount = 0
  if (schema.additionalProperties === false) {
    for (const name of Object.keys(value).sort()) {
      if (!Object.prototype.hasOwnProperty.call(value, name)) continue
      propertyCount++
      if (!Object.prototype.hasOwnProperty.call(properties, name)) {
        budget.nodes++
        if (budget.nodes > PPTX_NATIVE_RESOURCE_LIMITS.maxNodes) {
          budget.exhausted = true
          add(issues, path, 'native.resourceBudget', `contract exceeds ${PPTX_NATIVE_RESOURCE_LIMITS.maxNodes} JSON nodes`)
          return
        }
        add(issues, `${path}.${name}`, 'schema.additionalProperty', 'is not allowed')
      }
    }
  }
  if (typeof schema.minProperties === 'number' && propertyCount < schema.minProperties) {
    add(issues, path, 'schema.minProperties', `must contain at least ${schema.minProperties} property`)
  }
  for (const [name, child] of Object.entries(properties)) {
    if (Object.prototype.hasOwnProperty.call(value, name)) validateSchema(value[name], child, `${path}.${name}`, issues, depth + 1, budget)
  }
}

function discriminatedAlternative(value: unknown, alternatives: JsonSchema[]): JsonSchema | undefined {
  if (!isRecord(value) || typeof value.kind !== 'string') return undefined
  const matches = alternatives.filter((alternative) => {
    if (typeof alternative.$ref !== 'string' || !alternative.$ref.startsWith('#/$defs/')) return false
    const target = (PPTX_NATIVE_SCHEMA.$defs as unknown as Record<string, JsonSchema>)[alternative.$ref.slice('#/$defs/'.length)]
    return isRecord(target?.properties) && isRecord(target.properties.kind) && target.properties.kind.const === value.kind
  })
  return matches.length === 1 ? matches[0] : undefined
}

function validateSemantics(deck: NativePptxDeck, issues: NativeValidationIssue[]): void {
  const ids = new Set<string>()
  const assets = new Map(deck.assets.map((asset) => [asset.id, asset]))
  const slides = new Set(deck.slides.map((slide) => slide.id))
  const elements = new Map<string, string>()
  const budget: SemanticBudget = { elements: 0, textCodeUnits: 0, tableCells: 0, inlineAssetBase64CodeUnits: 0 }
  let worst: NativeCompatibility['status'] = 'editable'

  if (deck.origin === 'parsed' && !deck.sourceRevision) add(issues, '$.sourceRevision', 'native.sourceRevision', 'is required for a parsed deck')
  if (deck.origin === 'authored' && deck.sourceRevision) add(issues, '$.sourceRevision', 'native.sourceRevision', 'is not allowed for an authored deck')
  registerId(deck.documentId, '$.documentId', ids, issues)
  validateCompatibility(deck.compatibility, '$.compatibility', issues)

  deck.assets.forEach((asset, index) => {
    const path = `$.assets[${index}]`
    registerId(asset.id, `${path}.id`, ids, issues)
    validateSourceState(deck, asset.provenance, asset.source, asset.passthrough, path, issues)
    if (asset.provenance === 'parsed' && asset.dataBase64 === undefined) {
      const hasReadCapability = asset.source !== undefined && asset.passthrough.some((ref) => ref.ownerPart === asset.source?.partName && ref.fingerprintSha256 === asset.sha256)
      if (!hasReadCapability) add(issues, `${path}.passthrough`, 'native.assetReadCapability', 'parsed source-only assets require a capability bound to the source part and asset digest')
    }
    if (asset.dataBase64 !== undefined) budget.inlineAssetBase64CodeUnits += asset.dataBase64.length
    if (asset.dataBase64 !== undefined && budget.inlineAssetBase64CodeUnits <= PPTX_NATIVE_RESOURCE_LIMITS.maxTotalInlineAssetBase64CodeUnits && decodedBase64Length(asset.dataBase64) !== asset.byteLength) {
      add(issues, `${path}.dataBase64`, 'native.assetLength', 'decoded length does not match byteLength')
    }
  })
  if (budget.inlineAssetBase64CodeUnits > PPTX_NATIVE_RESOURCE_LIMITS.maxTotalInlineAssetBase64CodeUnits) {
    add(issues, '$.assets', 'native.resourceBudget', `inline asset data exceeds ${PPTX_NATIVE_RESOURCE_LIMITS.maxTotalInlineAssetBase64CodeUnits} UTF-16 code units`)
  }

  deck.slides.forEach((slide, slideIndex) => {
    const path = `$.slides[${slideIndex}]`
    registerId(slide.id, `${path}.id`, ids, issues)
    validateSourceState(deck, slide.provenance, slide.source, slide.passthrough, path, issues)
    validateCompatibility(slide.compatibility, `${path}.compatibility`, issues)
    worst = worseStatus(worst, slide.compatibility.status)
    let slideWorst = slide.compatibility.status
    validateTransition(slide.transition, `${path}.transition`, issues)
    slide.elements.forEach((element, elementIndex) => {
      slideWorst = worseStatus(slideWorst, validateElement(deck, element, `${path}.elements[${elementIndex}]`, ids, elements, assets, issues, budget, 1, slide.id, slide.source?.partName))
    })
    if (statusRank(slide.compatibility.status) < statusRank(slideWorst)) {
      add(issues, `${path}.compatibility.status`, 'native.compatibilityAggregate', `must be at least ${slideWorst} because an element has that status`)
    }
    worst = worseStatus(worst, slideWorst)
  })

  if (statusRank(deck.compatibility.status) < statusRank(worst)) {
    add(issues, '$.compatibility.status', 'native.compatibilityAggregate', `must be at least ${worst} because a descendant has that status`)
  }
  if (budget.textCodeUnits > PPTX_NATIVE_RESOURCE_LIMITS.maxTotalTextCodeUnits) add(issues, '$', 'native.resourceBudget', `text exceeds ${PPTX_NATIVE_RESOURCE_LIMITS.maxTotalTextCodeUnits} UTF-16 code units`)
  if (budget.tableCells > PPTX_NATIVE_RESOURCE_LIMITS.maxTableCells) add(issues, '$', 'native.resourceBudget', `tables exceed ${PPTX_NATIVE_RESOURCE_LIMITS.maxTableCells} cells`)
  for (const [path, compatibility] of compatibilityScopes(deck)) {
    for (let diagnosticIndex = 0; diagnosticIndex < compatibility.diagnostics.length; diagnosticIndex++) {
      const diagnostic = compatibility.diagnostics[diagnosticIndex]
      const diagnosticPath = `${path}[${diagnosticIndex}].scope`
      if (diagnostic.scope?.slideId && !slides.has(diagnostic.scope.slideId)) add(issues, `${diagnosticPath}.slideId`, 'native.scopeReference', 'references an unknown slide id')
      if (diagnostic.scope?.elementId && !elements.has(diagnostic.scope.elementId)) add(issues, `${diagnosticPath}.elementId`, 'native.scopeReference', 'references an unknown element id')
      if (diagnostic.scope?.slideId && diagnostic.scope.elementId && elements.has(diagnostic.scope.elementId) && elements.get(diagnostic.scope.elementId) !== diagnostic.scope.slideId) {
        add(issues, `${diagnosticPath}.elementId`, 'native.scopeOwnership', 'element does not belong to the scoped slide')
      }
    }
  }
}

function validateElement(
  deck: NativePptxDeck,
  element: NativeElement,
  path: string,
  ids: Set<string>,
  elements: Map<string, string>,
  assets: Map<string, NativePptxDeck['assets'][number]>,
  issues: NativeValidationIssue[],
  budget: SemanticBudget,
  depth: number,
  slideId: string,
  slidePart: string | undefined,
): NativeCompatibility['status'] {
  budget.elements++
  if (budget.elements > PPTX_NATIVE_RESOURCE_LIMITS.maxTotalElements) {
    add(issues, path, 'native.resourceBudget', `deck exceeds ${PPTX_NATIVE_RESOURCE_LIMITS.maxTotalElements} elements`)
    return 'refused'
  }
  if (depth > PPTX_NATIVE_RESOURCE_LIMITS.maxDepth) {
    add(issues, path, 'native.resourceDepth', `group depth exceeds ${PPTX_NATIVE_RESOURCE_LIMITS.maxDepth}`)
    return 'refused'
  }
  registerId(element.id, `${path}.id`, ids, issues)
  elements.set(element.id, slideId)
  validateSourceState(deck, element.provenance, element.source, element.passthrough, path, issues)
  if (element.provenance === 'parsed' && !slidePart) {
    add(issues, `${path}.provenance`, 'native.sourceOwnership', 'parsed elements require an owning parsed slide with a source anchor')
  } else if (element.provenance === 'parsed' && element.source && element.source.partName !== slidePart) {
    add(issues, `${path}.source.partName`, 'native.sourcePart', 'parsed elements must be anchored to their owning slide part')
  }
  validateCompatibility(element.compatibility, `${path}.compatibility`, issues)
  validateAnimation(element.animation, `${path}.animation`, issues)
  let worst = element.compatibility.status

  if (element.kind === 'text' || element.kind === 'shape') {
    for (const paragraph of element.paragraphs) for (const run of paragraph.runs) budget.textCodeUnits += run.text.length
    if (element.textBody) validateTextBody(element.textBody, element.transform, `${path}.textBody`, issues)
  }

  if (element.kind === 'shape' && element.preset === undefined && element.compatibility.status !== 'refused') {
    add(issues, `${path}.preset`, 'native.shapePreset', 'is required unless the shape is explicitly refused')
  }
  if ((element.kind === 'shape' || element.kind === 'connector') && element.stroke) {
    const metadataCount = Number(element.stroke.cap !== undefined) + Number(element.stroke.join !== undefined) + Number(element.stroke.dash !== undefined)
    if (metadataCount !== 0 && metadataCount !== 3) add(issues, `${path}.stroke`, 'native.strokeMetadata', 'cap, join, and dash must be supplied together')
    if (element.stroke.join === 'miter' && element.stroke.miterLimit === undefined) add(issues, `${path}.stroke.miterLimit`, 'native.stroke', 'is required for a miter join')
    if (element.stroke.join !== 'miter' && element.stroke.miterLimit !== undefined) add(issues, `${path}.stroke.miterLimit`, 'native.stroke', 'is allowed only for a miter join')
  }

  if (element.kind === 'picture') {
    const asset = assets.get(element.assetId)
    if (!asset) add(issues, `${path}.assetId`, 'native.assetReference', 'references an unknown asset id')
    else if (!asset.contentType.startsWith('image/')) add(issues, `${path}.assetId`, 'native.assetType', 'picture assets must have an image content type')
  } else if (element.kind === 'table') {
    let authoritativeCells = 0
    let legacyCells = 0
    for (let rowIndex = 0; rowIndex < element.table.rows.length; rowIndex++) {
      const row = element.table.rows[rowIndex]!
      budget.tableCells += row.length
      for (let columnIndex = 0; columnIndex < row.length; columnIndex++) {
        const cell = row[columnIndex]!
        const cellPath = `${path}.table.rows[${rowIndex}][${columnIndex}]`
        const hasParagraphs = cell.paragraphs !== undefined
        const hasTextBody = cell.textBody !== undefined
        if (hasParagraphs !== hasTextBody) add(issues, cellPath, 'native.tableTextAuthority', 'paragraphs and textBody must be supplied together')
        if (cell.paragraphs && cell.textBody) {
          authoritativeCells++
          for (const paragraph of cell.paragraphs) for (const run of paragraph.runs) budget.textCodeUnits += run.text.length
          if (cell.align !== undefined) add(issues, `${cellPath}.align`, 'native.tableTextAuthority', 'legacy align is not allowed with authoritative cell paragraphs')
          if (!tableCellTextMatches(cell.text, cell.paragraphs)) add(issues, `${cellPath}.text`, 'native.tableTextAuthority', 'must equal the newline-joined authoritative paragraph text')
          const width = element.table.columnWidths[columnIndex]
          const height = element.table.rowHeights[rowIndex]
          if (width !== undefined && height !== undefined) validateTextBody(cell.textBody, { x: 0, y: 0, cx: width, cy: height }, `${cellPath}.textBody`, issues)
        } else {
          legacyCells++
          budget.textCodeUnits += cell.text.length
        }
      }
    }
    if (authoritativeCells !== 0 && legacyCells !== 0) {
      add(issues, `${path}.table.rows`, 'native.tableTextAuthority', 'authoritative and legacy table cells cannot be mixed')
    }
    if (authoritativeCells !== 0) {
      if (element.table.rowHeights.length !== element.table.rows.length) {
        add(issues, `${path}.table.rowHeights`, 'native.tableGeometry', 'authoritative tables require one exact height per row')
      }
      const columnTotal = exactSafeTrackTotal(element.table.columnWidths)
      const rowTotal = exactSafeTrackTotal(element.table.rowHeights)
      if (columnTotal === undefined || columnTotal !== element.transform.cx) {
        add(issues, `${path}.table.columnWidths`, 'native.tableGeometry', 'authoritative column tracks must sum exactly to the table frame width')
      }
      if (rowTotal === undefined || rowTotal !== element.transform.cy) {
        add(issues, `${path}.table.rowHeights`, 'native.tableGeometry', 'authoritative row tracks must sum exactly to the table frame height')
      }
    }
    if (element.table.rowHeights.length !== 0 && element.table.rowHeights.length !== element.table.rows.length) {
      add(issues, `${path}.table.rowHeights`, 'native.tableDimensions', 'must be empty or contain one height per row')
    }
    element.table.rows.forEach((row, rowIndex) => {
      if (row.length !== element.table.columnWidths.length) add(issues, `${path}.table.rows[${rowIndex}]`, 'native.tableDimensions', 'must contain one cell per column')
    })
  } else if (element.kind === 'chart') {
    if (element.compatibility.status !== 'preserveOnly') add(issues, `${path}.compatibility.status`, 'native.opaqueChart', 'opaque charts must be preserveOnly')
    if (element.chart.previewAssetId && !assets.has(element.chart.previewAssetId)) add(issues, `${path}.chart.previewAssetId`, 'native.assetReference', 'references an unknown asset id')
    else if (element.chart.previewAssetId && !assets.get(element.chart.previewAssetId)?.contentType.startsWith('image/')) add(issues, `${path}.chart.previewAssetId`, 'native.assetType', 'chart previews must reference an image asset')
    if (element.chart.opaqueRef.ownerPart !== element.chart.chartPart) add(issues, `${path}.chart.opaqueRef.ownerPart`, 'native.chartReference', 'must equal chartPart')
    if (element.source.relationshipId !== element.chart.relationshipId) add(issues, `${path}.chart.relationshipId`, 'native.chartReference', 'must equal source.relationshipId')
  } else if (element.kind === 'group') {
    if (element.provenance === 'parsed' && element.childTransform === undefined) {
      add(issues, `${path}.childTransform`, 'native.groupTransform', 'parsed groups require an authoritative DrawingML child coordinate transform')
    }
    if (element.childTransform !== undefined) validateExactGroupTransform(element.transform, element.childTransform, `${path}.childTransform`, issues)
    element.children.forEach((child, index) => {
      worst = worseStatus(worst, validateElement(deck, child, `${path}.children[${index}]`, ids, elements, assets, issues, budget, depth + 1, slideId, slidePart))
    })
    if (statusRank(element.compatibility.status) < statusRank(worst)) {
      add(issues, `${path}.compatibility.status`, 'native.compatibilityAggregate', `must be at least ${worst} because a child has that status`)
    }
  }
  return worst
}

function exactSafeTrackTotal(tracks: readonly number[]): number | undefined {
  let total = 0
  for (const track of tracks) {
    total += track
    if (!Number.isSafeInteger(total)) return undefined
  }
  return total
}

function tableCellTextMatches(text: string, paragraphs: NonNullable<NativeTableCell['paragraphs']>): boolean {
  let cursor = 0
  for (let paragraphIndex = 0; paragraphIndex < paragraphs.length; paragraphIndex++) {
    if (paragraphIndex !== 0 && text[cursor++] !== '\n') return false
    for (const run of paragraphs[paragraphIndex]!.runs) {
      if (!text.startsWith(run.text, cursor)) return false
      cursor += run.text.length
    }
  }
  return cursor === text.length
}

function validateTextBody(
  body: Extract<NativeElement, { kind: 'text' | 'shape' }>['textBody'] & {},
  transform: NativeElement['transform'],
  path: string,
  issues: NativeValidationIssue[],
): void {
  const width = transform.cx - body.leftInsetEmu - body.rightInsetEmu
  const height = transform.cy - body.topInsetEmu - body.bottomInsetEmu
  if (!Number.isSafeInteger(width) || width <= 0) add(issues, path, 'native.textBodyBounds', 'horizontal insets must leave a positive safe-integer text-body width')
  if (!Number.isSafeInteger(height) || height <= 0) add(issues, path, 'native.textBodyBounds', 'vertical insets must leave a positive safe-integer text-body height')
}

function validateExactGroupTransform(transform: NativeElement['transform'], childTransform: NativeElement['transform'], path: string, issues: NativeValidationIssue[]): void {
  const ppm = 1_000_000n
  const scale = (extent: number, childExtent: number): bigint | undefined => {
    const numerator = BigInt(extent) * ppm
    const divisor = BigInt(childExtent)
    if (divisor <= 0n || numerator % divisor !== 0n) return undefined
    const value = numerator / divisor
    return value > 0n && value <= BigInt(Number.MAX_SAFE_INTEGER) ? value : undefined
  }
  const scaleX = scale(transform.cx, childTransform.cx)
  const scaleY = scale(transform.cy, childTransform.cy)
  const exactTranslation = (offset: number, childOffset: number, ratio: bigint | undefined): boolean => {
    if (ratio === undefined) return false
    const product = BigInt(childOffset) * ratio
    if (product % ppm !== 0n) return false
    const result = BigInt(offset) - product / ppm
    return result >= BigInt(Number.MIN_SAFE_INTEGER) && result <= BigInt(Number.MAX_SAFE_INTEGER)
  }
  if (scaleX === undefined || scaleY === undefined || !exactTranslation(transform.x, childTransform.x, scaleX) || !exactTranslation(transform.y, childTransform.y, scaleY)) {
    add(issues, path, 'native.groupTransform', 'must compose to exact safe integer-PPM scale and integer-EMU translation')
  }
}

function validateSourceState(deck: NativePptxDeck, provenance: 'authored' | 'parsed', source: NativeSourceAnchor | undefined, passthrough: { token: string }[], path: string, issues: NativeValidationIssue[]): void {
  if (deck.origin === 'authored' && provenance !== 'authored') add(issues, `${path}.provenance`, 'native.provenance', 'objects in an authored deck must be authored')
  if (provenance === 'authored' && (source || passthrough.length)) add(issues, path, 'native.authoredSource', 'authored objects cannot claim source anchors or passthrough tokens')
  if (provenance === 'parsed' && deck.origin !== 'parsed') add(issues, `${path}.provenance`, 'native.provenance', 'parsed objects require a parsed deck')
  if (provenance === 'parsed' && !source) add(issues, `${path}.source`, 'native.parsedSource', 'parsed objects require a source anchor')
}

function validateCompatibility(compatibility: NativeCompatibility, path: string, issues: NativeValidationIssue[]): void {
  const refusals = compatibility.diagnostics.filter((diagnostic) => diagnostic.severity === 'refusal').length
  const warnings = compatibility.diagnostics.filter((diagnostic) => diagnostic.severity === 'warning').length
  if (compatibility.status === 'editable' && refusals) add(issues, path, 'native.compatibility', 'editable content cannot contain refusal diagnostics')
  if (compatibility.status === 'preserveOnly' && !warnings) add(issues, path, 'native.compatibility', 'preserveOnly content requires a warning diagnostic')
  if (compatibility.status === 'preserveOnly' && refusals) add(issues, path, 'native.compatibility', 'preserveOnly content cannot contain refusal diagnostics')
  if (compatibility.status === 'refused' && !refusals) add(issues, path, 'native.compatibility', 'refused content requires a refusal diagnostic')
}

function validateAnimation(animation: NativeElement['animation'], path: string, issues: NativeValidationIssue[]): void {
  if (!animation) return
  if (animation.effect === 'fade' && (animation.direction !== undefined || animation.distancePpm !== undefined)) add(issues, path, 'native.animation', 'fade cannot declare direction or distancePpm')
}

function validateTransition(transition: NativePptxDeck['slides'][number]['transition'], path: string, issues: NativeValidationIssue[]): void {
  if (transition?.type === 'fade' && transition.direction !== undefined) add(issues, path, 'native.transition', 'fade cannot declare a direction')
}

function* compatibilityScopes(deck: NativePptxDeck): Generator<[string, NativeCompatibility]> {
  yield ['$.compatibility.diagnostics', deck.compatibility]
  for (let slideIndex = 0; slideIndex < deck.slides.length; slideIndex++) {
    const slide = deck.slides[slideIndex]
    yield [`$.slides[${slideIndex}].compatibility.diagnostics`, slide.compatibility]
    yield* elementCompatibilityScopes(slide.elements, `$.slides[${slideIndex}].elements`)
  }
}

function* elementCompatibilityScopes(elements: NativeElement[], path: string): Generator<[string, NativeCompatibility]> {
  for (let index = 0; index < elements.length; index++) {
    const element = elements[index]
    yield [`${path}[${index}].compatibility.diagnostics`, element.compatibility]
    if (element.kind === 'group') yield* elementCompatibilityScopes(element.children, `${path}[${index}].children`)
  }
}

function decodedBase64Length(value: string): number {
  if (!/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(value)) return -1
  return value.length === 0 ? 0 : (value.length / 4) * 3 - (value.endsWith('==') ? 2 : value.endsWith('=') ? 1 : 0)
}

function isCanonicalOpcPartName(value: string): boolean {
  if (!value || value.startsWith('/')) return false
  for (const component of value.split('/')) {
    if (!component) return false
    for (let index = 0; index < component.length; index++) {
      if (component[index] !== '%') continue
      if (!/^[0-9A-F]{2}$/.test(component.slice(index + 1, index + 3))) return false
      index += 2
    }
    let decoded: string
    try {
      decoded = decodeURIComponent(component)
    } catch {
      return false
    }
    if (decoded === '.' || decoded === '..' || decoded.endsWith('.') || decoded.includes('/') || decoded.includes('\\')) return false
    for (const character of decoded) {
      const codePoint = character.codePointAt(0) ?? 0
      if (codePoint < 0x20 || codePoint === 0x7f) return false
    }
  }
  return true
}

function registerId(id: string, path: string, seen: Set<string>, issues: NativeValidationIssue[]): void {
  if (seen.has(id)) add(issues, path, 'native.duplicateId', `duplicate durable id ${id}`)
  seen.add(id)
}

function worseStatus(left: NativeCompatibility['status'], right: NativeCompatibility['status']): NativeCompatibility['status'] {
  return statusRank(left) >= statusRank(right) ? left : right
}

function statusRank(status: NativeCompatibility['status']): number {
  return status === 'editable' ? 0 : status === 'preserveOnly' ? 1 : 2
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function add(issues: NativeValidationIssue[], path: string, code: string, message: string): void {
  if (issues.length < 100) issues.push({ path, code, message })
}

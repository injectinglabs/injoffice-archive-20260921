import { sha256 } from '@noble/hashes/sha2.js'
import { bytesToHex } from '@noble/hashes/utils.js'
import {
  PPTX_NATIVE_RESOURCE_LIMITS,
  validateNativePptx,
  type NativeAnimation,
  type NativeCompatibility,
  type NativeConnectorElement,
  type NativeElement,
  type NativeParagraph,
  type NativePptxDeck,
  type NativeShapeElement,
  type NativeSlide,
  type NativeStroke,
  type NativeTextBodyLayout,
  type NativeTextElement,
  type NativeTransition,
} from '@injoffice/pptx-native'
import {
  BUILTIN_THEMES,
  compileDeckToWire,
  type DeckSpec,
  type DeckTheme,
  type DiagramSpec,
  type OrgChartNode,
  type ShapeAnimationSpec,
  type ShapePositionOverride,
  type SlideSpec,
  type SlideTransitionSpec,
  type WireDeck,
  type WireParagraph,
  type WireShape,
  type WireSlide,
  type WireTextRun,
} from '@injoffice/slides/authoring'

export interface AuthoredDeckCompileIssue {
  readonly path: string
  readonly code: string
  readonly message: string
}

export type AuthoredDeckCompileResult =
  | { readonly ok: true; readonly deck: NativePptxDeck }
  | { readonly ok: false; readonly issues: readonly AuthoredDeckCompileIssue[] }

const DEFAULT_SLIDE_CX = 12_192_000
const DEFAULT_SLIDE_CY = 6_858_000
const DEFAULT_TEXT_BODY: NativeTextBodyLayout = Object.freeze({
  leftInsetEmu: 91_440,
  rightInsetEmu: 91_440,
  topInsetEmu: 45_720,
  bottomInsetEmu: 45_720,
  wrap: 'square',
  verticalAnchor: 'top',
  autoFit: 'none',
  horizontalOverflow: 'overflow',
  verticalOverflow: 'overflow',
})
const MAX_RENDER_COORDINATE_EMU = 281_474_976_710_655
const MAX_ISSUES = 128
const MAX_INPUT_DEPTH = 64
const MAX_INPUT_NODES = 1_000_000
const XML_CONTROL = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/u
const COLOR = /^#?[0-9A-Fa-f]{6}$/
const BUILTIN_THEME_IDS = new Set(BUILTIN_THEMES.map((theme) => theme.id))
const SLIDE_KINDS = new Set(['title', 'section', 'bullets', 'two-col', 'quote', 'closing'])
const SHAPE_KINDS = new Set(['textBox', 'line', 'rect', 'roundRect', 'ellipse', 'triangle', 'diamond', 'rightArrow', 'pentagon', 'hexagon', 'star5'])
const PLACEHOLDERS = new Set(['', 'title', 'ctrTitle', 'subTitle', 'body'])
const ALIGNS = new Set(['', 'l', 'ctr', 'r'])
const DIRECTIONS = new Set(['left', 'right', 'top', 'bottom'])

interface ValidationContext {
  readonly issues: AuthoredDeckCompileIssue[]
  readonly seen: WeakSet<object>
  nodes: number
  totalTextCodeUnits: number
}

interface IdentityContext {
  readonly source: 'deckSpec' | 'wireDeck'
  readonly documentKey?: string
  readonly slideKeys?: readonly string[]
  readonly elementScopes?: readonly string[]
}

const editable = (): NativeCompatibility => ({ status: 'editable', diagnostics: [] })

function addIssue(context: ValidationContext, path: string, code: string, message: string): void {
  if (context.issues.length < MAX_ISSUES) context.issues.push({ path, code, message })
}

function enterContainer(value: object, path: string, context: ValidationContext, depth: number): boolean {
  if (depth > MAX_INPUT_DEPTH) {
    addIssue(context, path, 'authored.resourceDepth', `input nesting exceeds ${MAX_INPUT_DEPTH}`)
    return false
  }
  context.nodes++
  if (context.nodes > MAX_INPUT_NODES) {
    addIssue(context, path, 'authored.resourceBudget', `input exceeds ${MAX_INPUT_NODES} containers`)
    return false
  }
  if (context.seen.has(value)) {
    addIssue(context, path, 'authored.nonJson', 'cycles and shared object references are not allowed')
    return false
  }
  context.seen.add(value)
  return true
}

function record(value: unknown, path: string, context: ValidationContext, depth: number): Record<string, unknown> | undefined {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    addIssue(context, path, 'authored.type', 'must be a plain object')
    return undefined
  }
  const prototype = Object.getPrototypeOf(value)
  if (prototype !== Object.prototype && prototype !== null) {
    addIssue(context, path, 'authored.nonJson', 'must be a plain JSON object')
    return undefined
  }
  if (!enterContainer(value, path, context, depth)) return undefined
  return value as Record<string, unknown>
}

function array(value: unknown, path: string, context: ValidationContext, depth: number, maxItems: number): unknown[] | undefined {
  if (!Array.isArray(value)) {
    addIssue(context, path, 'authored.type', 'must be an array')
    return undefined
  }
  if (!enterContainer(value, path, context, depth)) return undefined
  if (value.length > maxItems) addIssue(context, path, 'authored.resourceBudget', `must contain at most ${maxItems} items`)
  const remainingNodes = Math.max(0, MAX_INPUT_NODES - context.nodes)
  if (value.length > remainingNodes) addIssue(context, path, 'authored.resourceBudget', `input exceeds ${MAX_INPUT_NODES} bounded JSON nodes`)
  const inspected = Math.min(value.length, maxItems, remainingNodes)
  context.nodes += inspected
  return inspected === value.length ? value : value.slice(0, inspected)
}

function own(recordValue: Record<string, unknown>, key: string): boolean {
  return Object.prototype.hasOwnProperty.call(recordValue, key)
}

function onlyKeys(value: Record<string, unknown>, allowed: ReadonlySet<string>, path: string, context: ValidationContext): void {
  for (const key of Object.keys(value).sort()) {
    if (!allowed.has(key)) addIssue(context, `${path}.${key}`, 'authored.unsupportedField', 'is not supported by the native authored compiler')
  }
}

function requiredString(value: unknown, path: string, context: ValidationContext, maxLength = 1_048_576, allowEmpty = true): string | undefined {
  if (typeof value !== 'string') {
    addIssue(context, path, 'authored.type', 'must be a string')
    return undefined
  }
  if ((!allowEmpty && value.length === 0) || value.length > maxLength) {
    addIssue(context, path, 'authored.stringLength', `must contain ${allowEmpty ? 'at most' : 'between 1 and'} ${maxLength} UTF-16 code units`)
    return undefined
  }
  if (XML_CONTROL.test(value) || hasUnpairedSurrogate(value)) {
    addIssue(context, path, 'authored.xmlText', 'contains a character that cannot be represented exactly in Office Open XML')
    return undefined
  }
  context.totalTextCodeUnits += value.length
  if (context.totalTextCodeUnits > PPTX_NATIVE_RESOURCE_LIMITS.maxTotalTextCodeUnits) {
    addIssue(context, path, 'authored.resourceBudget', `deck text exceeds ${PPTX_NATIVE_RESOURCE_LIMITS.maxTotalTextCodeUnits} UTF-16 code units`)
  }
  return value
}

function optionalString(recordValue: Record<string, unknown>, key: string, path: string, context: ValidationContext, maxLength = 1_048_576, allowEmpty = true): void {
  if (own(recordValue, key)) requiredString(recordValue[key], `${path}.${key}`, context, maxLength, allowEmpty)
}

function booleanField(recordValue: Record<string, unknown>, key: string, path: string, context: ValidationContext): void {
  if (own(recordValue, key) && typeof recordValue[key] !== 'boolean') addIssue(context, `${path}.${key}`, 'authored.type', 'must be a boolean')
}

function finiteNumber(value: unknown, path: string, context: ValidationContext): number | undefined {
  if (typeof value !== 'number' || !Number.isFinite(value) || Object.is(value, -0)) {
    addIssue(context, path, 'authored.number', 'must be a finite non-negative-zero number')
    return undefined
  }
  return value
}

function safeInteger(value: unknown, path: string, context: ValidationContext, minimum: number, maximum = MAX_RENDER_COORDINATE_EMU): number | undefined {
  const numberValue = finiteNumber(value, path, context)
  if (numberValue === undefined) return undefined
  if (!Number.isSafeInteger(numberValue) || numberValue < minimum || numberValue > maximum) {
    addIssue(context, path, 'authored.integerEmu', `must be a safe integer between ${minimum} and ${maximum}`)
    return undefined
  }
  return numberValue
}

function enumValue(value: unknown, allowed: ReadonlySet<string>, path: string, context: ValidationContext): string | undefined {
  if (typeof value !== 'string' || !allowed.has(value)) {
    addIssue(context, path, 'authored.enum', `must be one of ${[...allowed].join(', ')}`)
    return undefined
  }
  return value
}

function colorValue(value: unknown, path: string, context: ValidationContext): string | undefined {
  if (typeof value !== 'string' || !COLOR.test(value)) {
    addIssue(context, path, 'authored.color', 'must be a six-digit sRGB color, with an optional leading #')
    return undefined
  }
  return value.replace(/^#/, '').toUpperCase()
}

function shapePaintValue(value: unknown, path: string, context: ValidationContext): string | undefined {
  if (value === '') return ''
  return colorValue(value, path, context)
}

function validateStringArray(value: unknown, path: string, context: ValidationContext, depth: number, maxItems: number = PPTX_NATIVE_RESOURCE_LIMITS.maxParagraphsPerElement): void {
  const values = array(value, path, context, depth, maxItems)
  values?.forEach((item, index) => requiredString(item, `${path}[${index}]`, context))
}

function validateDeckTheme(value: unknown, path: string, context: ValidationContext, depth: number): void {
  if (typeof value === 'string') {
    if (!BUILTIN_THEME_IDS.has(value)) addIssue(context, path, 'authored.theme', 'unknown built-in theme would silently fall back')
    return
  }
  const theme = record(value, path, context, depth)
  if (!theme) return
  const keys = new Set(['id', 'background', 'surface', 'ink', 'muted', 'accent', 'displayFont', 'bodyFont', 'monoFont'])
  onlyKeys(theme, keys, path, context)
  for (const key of keys) {
    if (!own(theme, key)) addIssue(context, `${path}.${key}`, 'authored.required', 'is required for a self-contained custom theme')
  }
  requiredString(theme.id, `${path}.id`, context, 256, false)
  for (const key of ['background', 'surface', 'ink', 'muted', 'accent']) colorValue(theme[key], `${path}.${key}`, context)
  for (const key of ['displayFont', 'bodyFont', 'monoFont']) requiredString(theme[key], `${path}.${key}`, context, 1024, false)
}

function validatePositionOverride(value: unknown, path: string, context: ValidationContext, depth: number): void {
  const override = record(value, path, context, depth)
  if (!override) return
  onlyKeys(override, new Set(['x', 'y', 'cx', 'cy']), path, context)
  if (Object.keys(override).length === 0) addIssue(context, path, 'authored.emptyOverride', 'must change at least one geometry field')
  if (own(override, 'x')) safeInteger(override.x, `${path}.x`, context, -MAX_RENDER_COORDINATE_EMU)
  if (own(override, 'y')) safeInteger(override.y, `${path}.y`, context, -MAX_RENDER_COORDINATE_EMU)
  if (own(override, 'cx')) safeInteger(override.cx, `${path}.cx`, context, 1)
  if (own(override, 'cy')) safeInteger(override.cy, `${path}.cy`, context, 1)
}

function validateAnimationSpec(value: unknown, path: string, context: ValidationContext, depth: number): void {
  const animation = record(value, path, context, depth)
  if (!animation) return
  onlyKeys(animation, new Set(['enter', 'direction', 'delayMs', 'durationMs', 'distance']), path, context)
  const effect = enumValue(animation.enter, new Set(['fade', 'flyIn']), `${path}.enter`, context)
  if (own(animation, 'direction')) enumValue(animation.direction, DIRECTIONS, `${path}.direction`, context)
  if (own(animation, 'delayMs')) safeInteger(animation.delayMs, `${path}.delayMs`, context, 0, 86_400_000)
  if (own(animation, 'durationMs')) safeInteger(animation.durationMs, `${path}.durationMs`, context, 0, 86_400_000)
  if (own(animation, 'distance')) {
    const distance = finiteNumber(animation.distance, `${path}.distance`, context)
    if (distance !== undefined && (distance < 0 || distance > 1 || !Number.isSafeInteger(distance * 1_000_000))) {
      addIssue(context, `${path}.distance`, 'authored.animationDistance', 'must be an exact integer part-per-million fraction between 0 and 1')
    }
  }
  if (effect === 'fade' && (own(animation, 'direction') || own(animation, 'distance'))) {
    addIssue(context, path, 'authored.animation', 'fade cannot carry ignored direction or distance semantics')
  }
}

function validateTransitionSpec(value: unknown, path: string, context: ValidationContext, depth: number): void {
  const transition = record(value, path, context, depth)
  if (!transition) return
  onlyKeys(transition, new Set(['kind', 'direction']), path, context)
  const kind = enumValue(transition.kind, new Set(['fade', 'push', 'wipe']), `${path}.kind`, context)
  if (own(transition, 'direction')) enumValue(transition.direction, DIRECTIONS, `${path}.direction`, context)
  if (kind === 'fade' && own(transition, 'direction')) addIssue(context, path, 'authored.transition', 'fade cannot carry an ignored direction')
}

function validateOrgNode(value: unknown, path: string, context: ValidationContext, depth: number, diagramBudget: { nodes: number }): void {
  diagramBudget.nodes++
  if (diagramBudget.nodes > 5_000) {
    addIssue(context, path, 'authored.diagramBudget', 'an org chart may contain at most 5000 nodes so its boxes and edges fit one native slide budget')
    return
  }
  const node = record(value, path, context, depth)
  if (!node) return
  onlyKeys(node, new Set(['label', 'children']), path, context)
  requiredString(node.label, `${path}.label`, context)
  if (own(node, 'children')) {
    const children = array(node.children, `${path}.children`, context, depth + 1, 10_000)
    children?.forEach((child, index) => validateOrgNode(child, `${path}.children[${index}]`, context, depth + 2, diagramBudget))
  }
}

function validateDiagram(value: unknown, path: string, context: ValidationContext, depth: number): void {
  const diagram = record(value, path, context, depth)
  if (!diagram) return
  const type = enumValue(diagram.type, new Set(['process', 'orgChart']), `${path}.type`, context)
  if (type === 'process') {
    onlyKeys(diagram, new Set(['type', 'steps']), path, context)
    validateStringArray(diagram.steps, `${path}.steps`, context, depth + 1, 5_000)
  } else if (type === 'orgChart') {
    onlyKeys(diagram, new Set(['type', 'root']), path, context)
    validateOrgNode(diagram.root, `${path}.root`, context, depth + 1, { nodes: 0 })
  }
}

const COMMON_SLIDE_KEYS = ['id', 'kind', 'shapeOverrides', 'shapeAnimations', 'transition', 'diagram', 'notes'] as const
const SLIDE_CONTENT_KEYS: Record<string, readonly string[]> = {
  title: ['eyebrow', 'title', 'subtitle', 'body'],
  section: ['eyebrow', 'title', 'subtitle', 'body'],
  bullets: ['eyebrow', 'title', 'body', 'bullets'],
  'two-col': ['eyebrow', 'title', 'bullets', 'bulletsRight', 'colTitles'],
  quote: ['quote', 'subtitle'],
  closing: ['title', 'subtitle'],
}

function validateSlideSpec(value: unknown, path: string, context: ValidationContext, depth: number): void {
  const slide = record(value, path, context, depth)
  if (!slide) return
  const kind = enumValue(slide.kind, SLIDE_KINDS, `${path}.kind`, context)
  const allowed = new Set([...COMMON_SLIDE_KEYS, ...(kind ? SLIDE_CONTENT_KEYS[kind] ?? [] : [])])
  onlyKeys(slide, allowed, path, context)
  requiredString(slide.id, `${path}.id`, context, 1024, false)
  if (own(slide, 'notes')) addIssue(context, `${path}.notes`, 'authored.unsupportedNotes', 'speaker notes are not represented by pptx-native/v1')
  for (const key of ['eyebrow', 'title', 'subtitle', 'body', 'quote']) optionalString(slide, key, path, context)
  for (const key of ['bullets', 'bulletsRight']) if (own(slide, key)) validateStringArray(slide[key], `${path}.${key}`, context, depth + 1)
  if (own(slide, 'colTitles')) {
    const titles = array(slide.colTitles, `${path}.colTitles`, context, depth + 1, 2)
    if (titles && titles.length !== 2) addIssue(context, `${path}.colTitles`, 'authored.tuple', 'must contain exactly two headings')
    titles?.forEach((title, index) => requiredString(title, `${path}.colTitles[${index}]`, context))
  }
  for (const [field, validator] of [
    ['shapeOverrides', validatePositionOverride],
    ['shapeAnimations', validateAnimationSpec],
  ] as const) {
    if (!own(slide, field)) continue
    const entries = record(slide[field], `${path}.${field}`, context, depth + 1)
    if (!entries) continue
    for (const [key, entry] of Object.entries(entries).sort(([left], [right]) => left < right ? -1 : left > right ? 1 : 0)) {
      requiredString(key, `${path}.${field}{key}`, context, 256, false)
      validator(entry, `${path}.${field}.${key}`, context, depth + 2)
    }
  }
  if (own(slide, 'transition')) validateTransitionSpec(slide.transition, `${path}.transition`, context, depth + 1)
  if (own(slide, 'diagram')) validateDiagram(slide.diagram, `${path}.diagram`, context, depth + 1)
}

function validateDeckSpec(value: unknown, context: ValidationContext): DeckSpec | undefined {
  const deck = record(value, '$', context, 0)
  if (!deck) return undefined
  onlyKeys(deck, new Set(['id', 'title', 'theme', 'slides']), '$', context)
  requiredString(deck.id, '$.id', context, 1024, false)
  requiredString(deck.title, '$.title', context)
  if (own(deck, 'theme')) validateDeckTheme(deck.theme, '$.theme', context, 1)
  const slides = array(deck.slides, '$.slides', context, 1, PPTX_NATIVE_RESOURCE_LIMITS.maxSlides)
  slides?.forEach((slide, index) => validateSlideSpec(slide, `$.slides[${index}]`, context, 2))
  if (context.issues.length) return undefined
  return cloneJson(value) as DeckSpec
}

function validateWireRun(value: unknown, path: string, context: ValidationContext, depth: number): void {
  const run = record(value, path, context, depth)
  if (!run) return
  onlyKeys(run, new Set(['text', 'bold', 'italic', 'sizePt', 'color', 'font']), path, context)
  const text = requiredString(run.text, `${path}.text`, context)
  booleanField(run, 'bold', path, context)
  booleanField(run, 'italic', path, context)
  if (own(run, 'sizePt')) {
    const size = finiteNumber(run.sizePt, `${path}.sizePt`, context)
    if (size !== undefined && (!Number.isSafeInteger(size * 100) || size * 100 < 1 || size * 100 > 400_000)) {
      addIssue(context, `${path}.sizePt`, 'authored.fontSize', 'must convert exactly to an integer 1/100 point between 1 and 400000')
    }
  }
  if (own(run, 'color')) colorValue(run.color, `${path}.color`, context)
  if (own(run, 'font')) requiredString(run.font, `${path}.font`, context, 256, false)
  if (text && (!own(run, 'sizePt') || !own(run, 'color') || !own(run, 'font'))) {
    addIssue(context, path, 'authored.inheritedTextStyle', 'non-empty runs require explicit sizePt, color, and font for native layout authority')
  }
}

function validateWireParagraph(value: unknown, path: string, context: ValidationContext, depth: number): void {
  const paragraph = record(value, path, context, depth)
  if (!paragraph) return
  onlyKeys(paragraph, new Set(['runs', 'align', 'level', 'bullet']), path, context)
  const runs = array(paragraph.runs, `${path}.runs`, context, depth + 1, PPTX_NATIVE_RESOURCE_LIMITS.maxRunsPerParagraph)
  runs?.forEach((run, index) => validateWireRun(run, `${path}.runs[${index}]`, context, depth + 2))
  if (own(paragraph, 'align')) enumValue(paragraph.align, ALIGNS, `${path}.align`, context)
  if (own(paragraph, 'level')) safeInteger(paragraph.level, `${path}.level`, context, 0, 8)
  booleanField(paragraph, 'bullet', path, context)
}

function validateWireAnimation(shape: Record<string, unknown>, path: string, context: ValidationContext): void {
  if (!own(shape, 'enter')) {
    for (const key of ['enterDirection', 'enterDelayMs', 'enterDurationMs', 'enterDistance']) {
      if (own(shape, key)) addIssue(context, `${path}.${key}`, 'authored.orphanAnimation', 'requires enter')
    }
    return
  }
  const effect = enumValue(shape.enter, new Set(['fade', 'flyIn']), `${path}.enter`, context)
  if (own(shape, 'enterDirection')) enumValue(shape.enterDirection, DIRECTIONS, `${path}.enterDirection`, context)
  if (own(shape, 'enterDelayMs')) safeInteger(shape.enterDelayMs, `${path}.enterDelayMs`, context, 0, 86_400_000)
  if (own(shape, 'enterDurationMs')) safeInteger(shape.enterDurationMs, `${path}.enterDurationMs`, context, 0, 86_400_000)
  if (own(shape, 'enterDistance')) {
    const distance = finiteNumber(shape.enterDistance, `${path}.enterDistance`, context)
    if (distance !== undefined && (distance < 0 || distance > 1 || !Number.isSafeInteger(distance * 1_000_000))) {
      addIssue(context, `${path}.enterDistance`, 'authored.animationDistance', 'must be an exact integer part-per-million fraction between 0 and 1')
    }
  }
  if (effect === 'fade' && (own(shape, 'enterDirection') || own(shape, 'enterDistance'))) {
    addIssue(context, path, 'authored.animation', 'fade cannot carry ignored direction or distance semantics')
  }
}

const WIRE_SHAPE_KEYS = new Set([
  'kind', 'placeholder', 'name', 'x', 'y', 'cx', 'cy', 'fill', 'stroke', 'strokeWidthPt', 'paragraphs',
  'headArrow', 'tailArrow', 'flipH', 'key', 'enter', 'enterDirection', 'enterDelayMs', 'enterDurationMs', 'enterDistance',
])

function validateWireShape(value: unknown, path: string, context: ValidationContext, depth: number): void {
  const shape = record(value, path, context, depth)
  if (!shape) return
  onlyKeys(shape, WIRE_SHAPE_KEYS, path, context)
  const kind = enumValue(shape.kind, SHAPE_KINDS, `${path}.kind`, context)
  safeInteger(shape.x, `${path}.x`, context, -MAX_RENDER_COORDINATE_EMU)
  safeInteger(shape.y, `${path}.y`, context, -MAX_RENDER_COORDINATE_EMU)
  safeInteger(shape.cx, `${path}.cx`, context, 1)
  safeInteger(shape.cy, `${path}.cy`, context, 1)
  if (own(shape, 'placeholder')) enumValue(shape.placeholder, PLACEHOLDERS, `${path}.placeholder`, context)
  optionalString(shape, 'name', path, context, 1024)
  optionalString(shape, 'key', path, context, 256, false)
  if (own(shape, 'fill')) shapePaintValue(shape.fill, `${path}.fill`, context)
  if (own(shape, 'stroke')) shapePaintValue(shape.stroke, `${path}.stroke`, context)
  if (own(shape, 'strokeWidthPt')) {
    const width = finiteNumber(shape.strokeWidthPt, `${path}.strokeWidthPt`, context)
    const effective = width !== undefined && width <= 0 ? 12_700 : (width ?? 0) * 12_700
    if (width !== undefined && (width < 0 || !Number.isSafeInteger(effective) || effective < 0 || effective > PPTX_NATIVE_RESOURCE_LIMITS.maxLineWidthEmu)) {
      addIssue(context, `${path}.strokeWidthPt`, 'authored.strokeWidth', `must convert exactly to integer EMU at or below ${PPTX_NATIVE_RESOURCE_LIMITS.maxLineWidthEmu}`)
    }
    if (typeof shape.stroke !== 'string' || shape.stroke === '') addIssue(context, `${path}.strokeWidthPt`, 'authored.ignoredStrokeWidth', 'would be ignored because stroke is absent')
  }
  for (const key of ['headArrow', 'tailArrow', 'flipH']) booleanField(shape, key, path, context)
  if (own(shape, 'paragraphs')) {
    const paragraphs = array(shape.paragraphs, `${path}.paragraphs`, context, depth + 1, PPTX_NATIVE_RESOURCE_LIMITS.maxParagraphsPerElement)
    paragraphs?.forEach((paragraph, index) => validateWireParagraph(paragraph, `${path}.paragraphs[${index}]`, context, depth + 2))
  }
  validateWireAnimation(shape, path, context)

  if (kind === 'line') {
    for (const key of ['placeholder', 'fill', 'paragraphs']) if (own(shape, key)) addIssue(context, `${path}.${key}`, 'authored.connectorField', 'is not meaningful for a line connector')
    if ((shape.headArrow === true || shape.tailArrow === true) && (typeof shape.stroke !== 'string' || shape.stroke === '')) addIssue(context, path, 'authored.invisibleArrow', 'arrowheads without an explicit stroke are discarded by the wire writer')
  } else if (kind) {
    for (const key of ['headArrow', 'tailArrow', 'flipH']) if (own(shape, key)) addIssue(context, `${path}.${key}`, 'authored.shapeField', 'is only meaningful for a line connector')
    if (kind === 'textBox' && (Boolean(shape.fill) || Boolean(shape.stroke) || own(shape, 'strokeWidthPt'))) {
      addIssue(context, path, 'authored.textBoxPaint', 'pptx-native/v1 text elements cannot carry the wire text-box fill or outline exactly')
    }
  }
}

function validateWireSlide(value: unknown, path: string, context: ValidationContext, depth: number): void {
  const slide = record(value, path, context, depth)
  if (!slide) return
  onlyKeys(slide, new Set(['shapes', 'background', 'transition', 'transitionDirection']), path, context)
  const shapes = array(slide.shapes, `${path}.shapes`, context, depth + 1, PPTX_NATIVE_RESOURCE_LIMITS.maxElementsPerContainer)
  shapes?.forEach((shape, index) => validateWireShape(shape, `${path}.shapes[${index}]`, context, depth + 2))
  if (!own(slide, 'background')) addIssue(context, `${path}.background`, 'authored.inheritedBackground', 'an explicit solid background is required for native rendering authority')
  else colorValue(slide.background, `${path}.background`, context)
  if (own(slide, 'transition')) enumValue(slide.transition, new Set(['fade', 'push', 'wipe']), `${path}.transition`, context)
  if (own(slide, 'transitionDirection')) enumValue(slide.transitionDirection, DIRECTIONS, `${path}.transitionDirection`, context)
  if (!own(slide, 'transition') && own(slide, 'transitionDirection')) addIssue(context, `${path}.transitionDirection`, 'authored.orphanTransition', 'requires transition')
  if (slide.transition === 'fade' && own(slide, 'transitionDirection')) addIssue(context, path, 'authored.transition', 'fade cannot carry an ignored direction')
}

function validateWireDeck(value: unknown, context: ValidationContext): WireDeck | undefined {
  const deck = record(value, '$', context, 0)
  if (!deck) return undefined
  onlyKeys(deck, new Set(['slides', 'cx', 'cy']), '$', context)
  if (own(deck, 'cx')) safeInteger(deck.cx, '$.cx', context, 1)
  if (own(deck, 'cy')) safeInteger(deck.cy, '$.cy', context, 1)
  const slides = array(deck.slides, '$.slides', context, 1, PPTX_NATIVE_RESOURCE_LIMITS.maxSlides)
  slides?.forEach((slide, index) => validateWireSlide(slide, `$.slides[${index}]`, context, 2))
  if (context.issues.length) return undefined
  return cloneJson(value) as WireDeck
}

function freshContext(): ValidationContext {
  return { issues: [], seen: new WeakSet(), nodes: 0, totalTextCodeUnits: 0 }
}

function hasUnpairedSurrogate(value: string): boolean {
  for (let index = 0; index < value.length; index++) {
    const code = value.charCodeAt(index)
    if (code >= 0xD800 && code <= 0xDBFF) {
      const next = value.charCodeAt(index + 1)
      if (!(next >= 0xDC00 && next <= 0xDFFF)) return true
      index++
    } else if (code >= 0xDC00 && code <= 0xDFFF) return true
  }
  return false
}

function cloneJson<T>(value: T): T {
  if (Array.isArray(value)) return value.map((child) => cloneJson(child)) as T
  if (typeof value !== 'object' || value === null) return value
  const cloned: Record<string, unknown> = Object.create(null) as Record<string, unknown>
  for (const key of Object.keys(value as Record<string, unknown>)) cloned[key] = cloneJson((value as Record<string, unknown>)[key])
  return cloned as T
}

function canonicalValue(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(canonicalValue)
  if (typeof value !== 'object' || value === null) return value
  return Object.fromEntries(Object.entries(value as Record<string, unknown>)
    .sort(([left], [right]) => left < right ? -1 : left > right ? 1 : 0)
    .map(([key, child]) => [key, canonicalValue(child)]))
}

function seedText(domain: string, value: unknown, ordinal = 0): string {
  return `${domain}\u0000${JSON.stringify(canonicalValue(value))}\u0000${ordinal}`
}

class IdFactory {
  readonly #ids = new Map<string, string>()

  create(prefix: 'deck' | 'slide' | 'element', domain: string, value: unknown, ordinal = 0): string {
    const seed = seedText(domain, value, ordinal)
    const id = `authored.${prefix}.${bytesToHex(sha256(new TextEncoder().encode(seed)))}`
    const previous = this.#ids.get(id)
    if (previous !== undefined && previous !== seed) throw new Error(`SHA-256 identifier collision for ${prefix}`)
    if (previous !== undefined) throw new Error(`duplicate identifier seed for ${prefix}`)
    this.#ids.set(id, seed)
    return id
  }
}

function occurrence(occurrences: Map<string, number>, value: unknown): number {
  const key = JSON.stringify(canonicalValue(value))
  const current = occurrences.get(key) ?? 0
  occurrences.set(key, current + 1)
  return current
}

function wireSlideIdentity(slide: WireSlide): unknown {
  const shapes = [...slide.shapes]
  shapes.sort((left, right) => {
    const leftKey = JSON.stringify(canonicalValue(left))
    const rightKey = JSON.stringify(canonicalValue(right))
    return leftKey < rightKey ? -1 : leftKey > rightKey ? 1 : 0
  })
  return {
    background: slide.background,
    transition: slide.transition,
    transitionDirection: slide.transitionDirection,
    shapes,
  }
}

function normalizeColor(value: string): string {
  return value.replace(/^#/, '').toUpperCase()
}

function direction(value: 'left' | 'right' | 'top' | 'bottom'): 'left' | 'right' | 'up' | 'down' {
  return value === 'top' ? 'up' : value === 'bottom' ? 'down' : value
}

function nativeAnimation(shape: WireShape): NativeAnimation | undefined {
  if (!shape.enter) return undefined
  if (shape.enter === 'fade') {
    return { effect: 'fade', delayMs: shape.enterDelayMs ?? 0, durationMs: shape.enterDurationMs && shape.enterDurationMs > 0 ? shape.enterDurationMs : 500 }
  }
  const distance = shape.enterDistance && shape.enterDistance > 0 ? shape.enterDistance : 0.25
  return {
    effect: 'flyIn',
    direction: direction(shape.enterDirection ?? 'bottom'),
    delayMs: shape.enterDelayMs ?? 0,
    durationMs: shape.enterDurationMs && shape.enterDurationMs > 0 ? shape.enterDurationMs : 500,
    distancePpm: Math.round(distance * 1_000_000),
  }
}

function nativeTransition(slide: WireSlide): NativeTransition | undefined {
  if (!slide.transition) return undefined
  if (slide.transition === 'fade') return { type: 'fade' }
  return { type: slide.transition, direction: direction(slide.transitionDirection ?? 'left') }
}

function nativeStroke(shape: WireShape): NativeStroke | undefined {
  if (!shape.stroke) return undefined
  const widthPt = shape.strokeWidthPt && shape.strokeWidthPt > 0 ? shape.strokeWidthPt : 1
  return { color: normalizeColor(shape.stroke), widthEmu: Math.round(widthPt * 12_700) }
}

function nativeParagraphs(paragraphs: readonly WireParagraph[] | undefined): NativeParagraph[] {
  return (paragraphs ?? []).map((paragraph) => ({
    runs: paragraph.runs.map(nativeRun),
    align: paragraph.align === 'ctr' ? 'center' : paragraph.align === 'r' ? 'right' : 'left',
    level: paragraph.level ?? 0,
    bullet: paragraph.bullet ?? false,
  }))
}

function nativeRun(run: WireTextRun): NativeParagraph['runs'][number] {
  return {
    text: run.text,
    ...(run.bold ? { bold: true } : {}),
    ...(run.italic ? { italic: true } : {}),
    ...(run.sizePt === undefined ? {} : { fontSizeHundredthPt: Math.round(run.sizePt * 100) }),
    ...(run.color === undefined ? {} : { color: normalizeColor(run.color) }),
    ...(run.font === undefined ? {} : { fontFamily: run.font }),
  }
}

function nativeElement(shape: WireShape, id: string): NativeElement {
  const common = {
    id,
    provenance: 'authored' as const,
    name: shape.name || (shape.kind === 'line' ? 'line' : shape.kind),
    transform: { x: shape.x, y: shape.y, cx: shape.cx, cy: shape.cy },
    ...(nativeAnimation(shape) ? { animation: nativeAnimation(shape) } : {}),
    passthrough: [],
    compatibility: editable(),
  }
  if (shape.kind === 'line') {
    const connector: NativeConnectorElement = {
      ...common,
      kind: 'connector',
      ...(nativeStroke(shape) ? { stroke: nativeStroke(shape) } : {}),
      ...(shape.headArrow ? { headArrow: true } : {}),
      ...(shape.tailArrow ? { tailArrow: true } : {}),
      ...(shape.flipH ? { flipH: true } : {}),
    }
    return connector
  }
  const paragraphs = nativeParagraphs(shape.paragraphs)
  if (paragraphs.length === 0 && shape.placeholder) paragraphs.push({ runs: [], align: 'left', level: 0, bullet: false })
  const placeholder = shape.placeholder || 'body'
  const hasTextBody = paragraphs.length > 0
  if (shape.kind === 'textBox') {
    const text: NativeTextElement = {
      ...common,
      kind: 'text',
      placeholder,
      paragraphs,
      ...(hasTextBody ? { textBody: { ...DEFAULT_TEXT_BODY } } : {}),
    }
    return text
  }
  const preset: NativeShapeElement = {
    ...common,
    kind: 'shape',
    preset: shape.kind,
    placeholder,
    ...(shape.fill ? { fill: normalizeColor(shape.fill) } : {}),
    ...(nativeStroke(shape) ? { stroke: nativeStroke(shape) } : {}),
    paragraphs,
    ...(hasTextBody ? { textBody: { ...DEFAULT_TEXT_BODY } } : {}),
  }
  return preset
}

function stableShapeSeed(shape: WireShape, slideIdentity: string): unknown {
  if (shape.key !== undefined) return { slideIdentity, elementKey: shape.key, kind: shape.kind }
  return { slideIdentity, shape }
}

function compileValidatedWire(wire: WireDeck, identity: IdentityContext): AuthoredDeckCompileResult {
  const issues: AuthoredDeckCompileIssue[] = []
  const factory = new IdFactory()
  try {
    const documentSeed = identity.documentKey === undefined ? wire : { source: identity.source, documentKey: identity.documentKey }
    const documentId = factory.create('deck', `${identity.source}.document`, documentSeed)
    const slideOccurrences = new Map<string, number>()
    const slides: NativeSlide[] = wire.slides.map((wireSlide, slideIndex) => {
      const slideKey = identity.slideKeys?.[slideIndex]
      const slideSeed = slideKey === undefined ? wireSlideIdentity(wireSlide) : { documentKey: identity.documentKey, slideKey }
      const slideOrdinal = slideKey === undefined ? occurrence(slideOccurrences, slideSeed) : 0
      const slideId = factory.create('slide', `${identity.source}.slide`, slideSeed, slideOrdinal)
      const elementOccurrences = new Map<string, number>()
      const seenKeys = new Set<string>()
      const elements: NativeElement[] = wireSlide.shapes.map((shape, elementIndex) => {
        if (shape.key !== undefined) {
          if (seenKeys.has(shape.key)) throw new Error(`slide ${slideIndex} has duplicate shape key ${shape.key}`)
          seenKeys.add(shape.key)
        }
        const seed = stableShapeSeed(shape, identity.elementScopes?.[slideIndex] ?? slideKey ?? slideId)
        const ordinal = shape.key === undefined ? occurrence(elementOccurrences, seed) : 0
        const id = factory.create('element', `${identity.source}.element`, seed, ordinal)
        return nativeElement(shape, id)
      })
      return {
        id: slideId,
        provenance: 'authored',
        background: normalizeColor(wireSlide.background!),
        ...(nativeTransition(wireSlide) ? { transition: nativeTransition(wireSlide) } : {}),
        elements,
        passthrough: [],
        compatibility: editable(),
      }
    })
    const deck: NativePptxDeck = {
      contractVersion: 'pptx-native/v1',
      documentId,
      origin: 'authored',
      size: { cx: wire.cx ?? DEFAULT_SLIDE_CX, cy: wire.cy ?? DEFAULT_SLIDE_CY },
      assets: [],
      slides,
      compatibility: editable(),
    }
    const validation = validateNativePptx(deck)
    if (!validation.ok) {
      return { ok: false, issues: validation.issues.map((issue) => ({ path: issue.path, code: `authored.output.${issue.code}`, message: issue.message })) }
    }
    return { ok: true, deck: deepFreeze(deck) }
  } catch (error) {
    issues.push({ path: '$', code: 'authored.identifier', message: error instanceof Error ? error.message : String(error) })
    return { ok: false, issues }
  }
}

function deepFreeze<T>(value: T, seen = new Set<object>()): T {
  if (typeof value !== 'object' || value === null || seen.has(value)) return value
  seen.add(value)
  for (const child of Object.values(value as Record<string, unknown>)) deepFreeze(child, seen)
  return Object.freeze(value)
}

/** Atomically compile strict WireDeck JSON to an immutable authored native v1 deck. */
export function compileWireDeckToNativeV1(input: unknown): AuthoredDeckCompileResult {
  const context = freshContext()
  const wire = validateWireDeck(input, context)
  if (!wire) return { ok: false, issues: context.issues }
  return compileValidatedWire(wire, { source: 'wireDeck' })
}

/** Atomically compile strict DeckSpec JSON through the pure authoring boundary. */
export function compileDeckSpecToNativeV1(input: unknown): AuthoredDeckCompileResult {
  const context = freshContext()
  const spec = validateDeckSpec(input, context)
  if (!spec) return { ok: false, issues: context.issues }
  let wire: WireDeck
  try {
    wire = compileDeckToWire(spec)
  } catch (error) {
    return { ok: false, issues: [{ path: '$', code: 'authored.deckSpecCompile', message: error instanceof Error ? error.message : String(error) }] }
  }
  for (let slideIndex = 0; slideIndex < spec.slides.length; slideIndex++) {
    const slide = spec.slides[slideIndex]!
    const keys = new Set(wire.slides[slideIndex]!.shapes.map((shape) => shape.key).filter((key): key is string => key !== undefined))
    for (const field of ['shapeOverrides', 'shapeAnimations'] as const) {
      for (const key of Object.keys(slide[field] ?? {}).sort()) {
        if (!keys.has(key)) context.issues.push({ path: `$.slides[${slideIndex}].${field}.${key}`, code: 'authored.staleShapeKey', message: 'does not match a compiled shape and would be silently ignored' })
      }
    }
  }
  if (context.issues.length) return { ok: false, issues: context.issues }
  const wireContext = freshContext()
  const validatedWire = validateWireDeck(wire, wireContext)
  if (!validatedWire) {
    return { ok: false, issues: wireContext.issues.map((issue) => ({ ...issue, path: `$.compiledWire${issue.path.slice(1)}` })) }
  }
  return compileValidatedWire(validatedWire, {
    source: 'deckSpec',
    documentKey: spec.id,
    slideKeys: spec.slides.map((slide) => slide.id),
    elementScopes: spec.slides.map((slide) => `${slide.id}:${slide.kind}`),
  })
}

// Compile-time guards: the strict decoders deliberately operate on unknown,
// while the values passed to the legacy pure layout compiler retain the public
// authoring types. These assignments also catch drift in the imported boundary.
const _deckSpecBoundary: DeckSpec | undefined = undefined
const _wireDeckBoundary: WireDeck | undefined = undefined
const _themeBoundary: DeckTheme | undefined = undefined
const _slideBoundary: SlideSpec | undefined = undefined
const _diagramBoundary: DiagramSpec | undefined = undefined
const _orgBoundary: OrgChartNode | undefined = undefined
const _animationBoundary: ShapeAnimationSpec | undefined = undefined
const _overrideBoundary: ShapePositionOverride | undefined = undefined
const _transitionBoundary: SlideTransitionSpec | undefined = undefined
void [_deckSpecBoundary, _wireDeckBoundary, _themeBoundary, _slideBoundary, _diagramBoundary, _orgBoundary, _animationBoundary, _overrideBoundary, _transitionBoundary]

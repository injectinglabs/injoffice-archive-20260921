import type { NativeElement, NativePptxDeck, NativeShapeElement } from '@injoffice/pptx-native'
import type {
  PptxNativeExactAutoShapeV1,
  PptxNativeExactParagraphV1,
  PptxNativeMutationRequestV1,
} from '@injoffice/pptx-wasm'

const COLOR = /^[0-9A-F]{6}$/

export type EditablePptxTextTarget = {
  operationKind: 'text.replace'
  slideId: string
  slideIndex: number
  elementId: string
  elementName: string
  sourcePartName: string
  sourceObjectId: string
  expectedFingerprintSha256: string
  paragraphs: ReadonlyArray<PptxNativeExactParagraphV1>
}

export type EditablePptxShapeTarget = {
  operationKind: 'autoshape.update'
  slideId: string
  slideIndex: number
  elementId: string
  elementName: string
  sourcePartName: string
  sourceObjectId: string
  expectedFingerprintSha256: string
  autoShape: PptxNativeExactAutoShapeV1
}

export type EditablePptxTarget = EditablePptxTextTarget | EditablePptxShapeTarget

export type PptxMutationEvidence = {
  operation: 'text.replace' | 'autoshape.update'
  operation_id: string
  target_id: string
  source_part: string
  source_object: string
  expected_fingerprint_sha256: string
  expected_revision: string
  requested: Record<string, string | number>
}

export type PptxRoundTripProof = {
  target: string
  before: string
  after: string
  previousRevision: string
  revision: string
  artifactIdentity: 'stable' | 'browser memory' | 'request-local'
  preservedElements: number
}

const EXACT_SHAPE_PRESETS = ['rect', 'ellipse', 'triangle', 'diamond'] as const
export type ExactPptxShapePreset = typeof EXACT_SHAPE_PRESETS[number]
export const exactPptxShapePresets: ReadonlyArray<ExactPptxShapePreset> = EXACT_SHAPE_PRESETS

export function editablePptxTextTargets(deck: NativePptxDeck): EditablePptxTextTarget[] {
  const targets: EditablePptxTextTarget[] = []
  deck.slides.forEach((slide, slideIndex) => {
    visitElements(slide.elements, (element) => {
      if ((element.kind !== 'text' && element.kind !== 'shape') || !element.source || element.compatibility.status !== 'editable') return
      const paragraphs = exactParagraphs(element.paragraphs)
      if (!paragraphs || paragraphs[0]?.runs[0] === undefined) return
      targets.push({
        operationKind: 'text.replace',
        slideId: slide.id,
        slideIndex,
        elementId: element.id,
        elementName: element.name || `${element.kind} ${element.source.objectId}`,
        sourcePartName: element.source.partName,
        sourceObjectId: element.source.objectId,
        expectedFingerprintSha256: element.source.fingerprintSha256,
        paragraphs,
      })
    })
  })
  return targets
}

export function editablePptxShapeTargets(deck: NativePptxDeck): EditablePptxShapeTarget[] {
  const targets: EditablePptxShapeTarget[] = []
  deck.slides.forEach((slide, slideIndex) => {
    visitElements(slide.elements, (element) => {
      if (element.kind !== 'shape' || !element.source || element.compatibility.status !== 'editable') return
      const autoShape = exactAutoShape(element)
      if (!autoShape) return
      targets.push({
        operationKind: 'autoshape.update',
        slideId: slide.id,
        slideIndex,
        elementId: element.id,
        elementName: element.name || `shape ${element.source.objectId}`,
        sourcePartName: element.source.partName,
        sourceObjectId: element.source.objectId,
        expectedFingerprintSha256: element.source.fingerprintSha256,
        autoShape,
      })
    })
  })
  return targets
}

export function editablePptxTargets(deck: NativePptxDeck): EditablePptxTarget[] {
  return [...editablePptxTextTargets(deck), ...editablePptxShapeTargets(deck)]
    .sort((left, right) => left.slideIndex - right.slideIndex
      || (left.operationKind === right.operationKind ? left.elementName.localeCompare(right.elementName) : left.operationKind === 'text.replace' ? -1 : 1))
}

export function pptxTargetKey(target: Pick<EditablePptxTarget, 'sourcePartName' | 'sourceObjectId'> & { operationKind?: EditablePptxTarget['operationKind'] }): string {
  return `${target.operationKind ?? 'text.replace'}\u0000${target.sourcePartName}\u0000${target.sourceObjectId}`
}

export function firstPptxRunText(target: EditablePptxTextTarget | undefined): string {
  return target?.paragraphs[0]?.runs[0]?.text ?? ''
}

export function findPptxTextTarget(
  deck: NativePptxDeck,
  anchor: Pick<EditablePptxTextTarget, 'sourcePartName' | 'sourceObjectId'>,
): EditablePptxTextTarget | undefined {
  const key = pptxTargetKey(anchor)
  return editablePptxTextTargets(deck).find((target) => pptxTargetKey(target) === key)
}

export function findPptxTarget(deck: NativePptxDeck, anchor: EditablePptxTarget): EditablePptxTarget | undefined {
  const key = pptxTargetKey(anchor)
  return editablePptxTargets(deck).find((target) => pptxTargetKey(target) === key)
}

export function buildPptxTextMutation(
  deck: NativePptxDeck,
  target: EditablePptxTextTarget,
  text: string,
  operationId: string,
): PptxNativeMutationRequestV1 {
  if (!deck.sourceRevision) throw new Error('PPTX text mutation requires an exact parsed source revision.')
  if (text === firstPptxRunText(target)) throw new Error('PPTX text mutation must make a semantic change.')
  const paragraphs = target.paragraphs.map((paragraph, paragraphIndex) => ({
    align: paragraph.align,
    level: paragraph.level,
    bullet: false as const,
    runs: paragraph.runs.map((run, runIndex) => ({
      text: paragraphIndex === 0 && runIndex === 0 ? text : run.text,
      bold: run.bold,
      italic: run.italic,
      fontSizeHundredthPt: run.fontSizeHundredthPt,
      color: run.color,
      fontFamily: run.fontFamily,
    })),
  }))
  return {
    expectedSourceRevision: deck.sourceRevision,
    operations: [{
      operationId,
      kind: 'text.replace',
      elementId: target.elementId,
      expectedFingerprintSha256: target.expectedFingerprintSha256,
      paragraphs,
    }],
  }
}

export function buildPptxShapeMutation(
  deck: NativePptxDeck,
  target: EditablePptxShapeTarget,
  preset: ExactPptxShapePreset,
  fill: string | undefined,
  operationId: string,
): PptxNativeMutationRequestV1 {
  if (!deck.sourceRevision) throw new Error('PPTX AutoShape mutation requires an exact parsed source revision.')
  if (fill !== undefined && !COLOR.test(fill)) throw new Error('PPTX AutoShape fill must be six uppercase hexadecimal digits or omitted.')
  if (preset === target.autoShape.preset && fill === target.autoShape.fill) throw new Error('PPTX AutoShape mutation must make a semantic change.')
  return {
    expectedSourceRevision: deck.sourceRevision,
    operations: [{
      operationId,
      kind: 'autoshape.update',
      elementId: target.elementId,
      expectedFingerprintSha256: target.expectedFingerprintSha256,
      autoShape: {
        transform: target.autoShape.transform,
        preset,
        ...(fill === undefined ? {} : { fill }),
        ...(target.autoShape.stroke === undefined ? {} : { stroke: target.autoShape.stroke }),
      },
    }],
  }
}

export function pptxTargetValue(target: EditablePptxTarget | undefined): string {
  if (!target) return ''
  return target.operationKind === 'text.replace'
    ? firstPptxRunText(target)
    : `${target.autoShape.preset} · ${target.autoShape.fill ?? 'no fill'}`
}

export function buildPptxMutationEvidence(target: EditablePptxTarget, mutation: PptxNativeMutationRequestV1): PptxMutationEvidence {
  const operation = mutation.operations[0]
  if (!operation) throw new Error('PPTX mutation evidence requires one operation.')
  let requested: Record<string, string | number>
  if (operation.kind === 'text.replace') {
    requested = {
      replacement_utf16_units: operation.paragraphs[0]?.runs[0]?.text.length ?? 0,
      replacement_preview: boundedPreview(operation.paragraphs[0]?.runs[0]?.text ?? ''),
    }
  } else requested = { preset: operation.autoShape.preset, fill: operation.autoShape.fill ?? 'none' }
  return {
    operation: operation.kind,
    operation_id: operation.operationId,
    target_id: operation.elementId,
    source_part: target.sourcePartName,
    source_object: target.sourceObjectId,
    expected_fingerprint_sha256: operation.expectedFingerprintSha256,
    expected_revision: mutation.expectedSourceRevision,
    requested,
  }
}

export function untouchedPptxElementInventory(deck: NativePptxDeck, edited: EditablePptxTarget): string[] {
  const inventory: string[] = []
  for (const slide of deck.slides) {
    const visit = (element: NativeElement): void => {
      if (sameSource(element, edited)) return
      if (element.kind === 'group' && containsSource(element.children, edited)) {
        element.children.forEach(visit)
        return
      }
      inventory.push(JSON.stringify({ slideId: slide.id, elementId: element.id, kind: element.kind, source: element.source ?? null }))
      if (element.kind === 'group') element.children.forEach(visit)
    }
    slide.elements.forEach(visit)
  }
  return inventory.sort()
}

export function verifyPptxRoundTrip(
  before: NativePptxDeck,
  after: NativePptxDeck,
  target: EditablePptxTarget,
  requested: { text?: string; preset?: ExactPptxShapePreset; fill?: string },
  responseRevision?: string,
  responsePackageSHA256?: string,
  beforeArtifactId = '',
  responseArtifactId?: string,
  browserLocal = false,
): PptxRoundTripProof {
  const nextTarget = findPptxTarget(after, target)
  if (!nextTarget || nextTarget.operationKind !== target.operationKind) throw new Error(`Readback target is missing for ${target.elementName}.`)
  if (after.documentId !== before.documentId) throw new Error('Readback changed the native presentation identity.')
  if (!before.sourceRevision || !after.sourceRevision || after.sourceRevision === before.sourceRevision) throw new Error('Readback source revision did not advance.')
  if (responseRevision && responseRevision !== after.sourceRevision) throw new Error('Readback revision does not match the saved package.')
  if (responsePackageSHA256 && responsePackageSHA256 !== packageSHAFromRevision(after.sourceRevision)) throw new Error('Readback digest does not match the saved package.')
  if (beforeArtifactId && responseArtifactId !== beforeArtifactId) throw new Error('Mutation response changed the server artifact identity.')

  let afterValue: string
  if (target.operationKind === 'text.replace') {
    if (nextTarget.operationKind !== 'text.replace' || firstPptxRunText(nextTarget) !== requested.text) throw new Error(`Text readback mismatch for ${target.elementName}.`)
    afterValue = firstPptxRunText(nextTarget)
  } else {
    if (nextTarget.operationKind !== 'autoshape.update' || nextTarget.autoShape.preset !== requested.preset || nextTarget.autoShape.fill !== requested.fill) throw new Error(`AutoShape readback mismatch for ${target.elementName}.`)
    afterValue = pptxTargetValue(nextTarget)
  }

  const untouchedBefore = untouchedPptxElementInventory(before, target)
  const untouchedAfter = untouchedPptxElementInventory(after, target)
  if (JSON.stringify(untouchedAfter) !== JSON.stringify(untouchedBefore)) throw new Error('Readback changed an untouched element source anchor.')

  return {
    target: `Slide ${target.slideIndex + 1} · ${target.elementName} · ${target.operationKind}`,
    before: pptxTargetValue(target),
    after: afterValue,
    previousRevision: before.sourceRevision,
    revision: after.sourceRevision,
    artifactIdentity: browserLocal ? 'browser memory' : beforeArtifactId ? 'stable' : 'request-local',
    preservedElements: untouchedAfter.length,
  }
}

export function pptxDownloadName(sourceName: string): string {
  const base = sourceName.replace(/\.pptx$/i, '').replace(/[^A-Za-z0-9._-]+/g, '-').replace(/^-+|-+$/g, '')
  return `${base || 'presentation'}-injoffice.pptx`
}

function exactParagraphs(value: unknown): ReadonlyArray<PptxNativeExactParagraphV1> | undefined {
  if (!Array.isArray(value) || value.length === 0) return undefined
  const paragraphs: PptxNativeExactParagraphV1[] = []
  for (const item of value) {
    if (!isRecord(item) || !Array.isArray(item.runs) || (item.align !== 'left' && item.align !== 'center' && item.align !== 'right')) return undefined
    if (!Number.isSafeInteger(item.level) || (item.level as number) < 0 || (item.level as number) > 8 || item.bullet !== false) return undefined
    const runs: PptxNativeExactParagraphV1['runs'][number][] = []
    for (const run of item.runs) {
      if (!isRecord(run) || typeof run.text !== 'string' || /[\t\r\n]/.test(run.text) || typeof run.bold !== 'boolean' || typeof run.italic !== 'boolean') return undefined
      if (!Number.isSafeInteger(run.fontSizeHundredthPt) || (run.fontSizeHundredthPt as number) < 1 || (run.fontSizeHundredthPt as number) > 400_000) return undefined
      if (typeof run.color !== 'string' || !COLOR.test(run.color) || typeof run.fontFamily !== 'string' || run.fontFamily.length === 0 || run.fontFamily.length > 256) return undefined
      runs.push({
        text: run.text,
        bold: run.bold,
        italic: run.italic,
        fontSizeHundredthPt: run.fontSizeHundredthPt as number,
        color: run.color,
        fontFamily: run.fontFamily,
      })
    }
    paragraphs.push({ runs, align: item.align, level: item.level as number, bullet: false })
  }
  return paragraphs
}

function exactAutoShape(element: NativeShapeElement): PptxNativeExactAutoShapeV1 | undefined {
  if (!element.preset || !EXACT_SHAPE_PRESETS.includes(element.preset as ExactPptxShapePreset)) return undefined
  if (element.fill !== undefined && !COLOR.test(element.fill)) return undefined
  const stroke = element.stroke
  if (stroke && (!COLOR.test(stroke.color) || !Number.isSafeInteger(stroke.widthEmu) || stroke.cap === undefined || stroke.join === undefined || stroke.dash !== 'solid')) return undefined
  if (stroke?.join === 'miter' && !Number.isSafeInteger(stroke.miterLimit)) return undefined
  if (stroke?.join !== 'miter' && stroke?.miterLimit !== undefined) return undefined
  return {
    transform: element.transform,
    preset: element.preset as ExactPptxShapePreset,
    ...(element.fill === undefined ? {} : { fill: element.fill }),
    ...(stroke === undefined ? {} : {
      stroke: {
        color: stroke.color,
        widthEmu: stroke.widthEmu,
        cap: stroke.cap!,
        join: stroke.join!,
        dash: 'solid',
        ...(stroke.miterLimit === undefined ? {} : { miterLimit: stroke.miterLimit }),
      },
    }),
  }
}

function packageSHAFromRevision(revision: string): string {
  if (!/^rev-[0-9a-f]{64}$/.test(revision)) throw new Error('The native PPTX source revision is malformed.')
  return `sha256:${revision.slice(4)}`
}

function sameSource(element: NativeElement, target: EditablePptxTarget): boolean {
  return element.source?.partName === target.sourcePartName && element.source?.objectId === target.sourceObjectId
}

function containsSource(elements: ReadonlyArray<NativeElement>, target: EditablePptxTarget): boolean {
  return elements.some((element) => sameSource(element, target) || (element.kind === 'group' && containsSource(element.children, target)))
}

function boundedPreview(text: string): string {
  return text.length > 120 ? `${text.slice(0, 120)}…` : text
}

function visitElements(elements: ReadonlyArray<NativeElement>, visit: (element: NativeElement) => void): void {
  for (const element of elements) {
    visit(element)
    if (element.kind === 'group') visitElements(element.children, visit)
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
}

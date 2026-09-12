import { PPTX_RENDER_LIMITS, RenderCompileError, type RenderNode, type RenderParagraphNode, type RenderPathCommand, type RenderRect, type RenderStroke, type RenderTextBodyNode, type RenderTextRunNode, type RenderTransform, type SlideRenderTree } from './types.js'

import type { NativePictureCrop,NativeArrowEnd } from '@injoffice/pptx-native'

export type PaintCommand =
  | { readonly kind: 'beginSlide'; readonly size: { readonly cx: number; readonly cy: number }; readonly background: string }
  | { readonly kind: 'endSlide' }
  | { readonly kind: 'save' }
  | { readonly kind: 'restore' }
  | { readonly kind: 'transform'; readonly transform: RenderTransform }
  | { readonly kind: 'clipRect'; readonly rect: RenderRect }
  | { readonly kind: 'clipRoundRect'; readonly rect: RenderRect; readonly radiusEmu: number }
  | { readonly kind: 'path'; readonly sourceElementId: string; readonly path: readonly RenderPathCommand[]; readonly fill?: string; readonly stroke?: RenderStroke; readonly headArrow?: boolean; readonly tailArrow?: boolean; readonly headEnd?:Readonly<NativeArrowEnd>;readonly tailEnd?:Readonly<NativeArrowEnd> }
  | { readonly kind: 'image'; readonly sourceElementId: string; readonly role: 'picture' | 'chartPreview'; readonly assetId: string; readonly rect: RenderRect; readonly crop?: Readonly<NativePictureCrop> }
  | { readonly kind: 'glyphRun'; readonly sourceElementId: string; readonly run: RenderTextRunNode }
  | { readonly kind: 'placeholder'; readonly sourceElementId: string; readonly rect: RenderRect; readonly reason: string; readonly label: string }

export interface PaintSurface {
  push(command: PaintCommand): void
}

export interface RecordingPaintSurface extends PaintSurface {
  readonly commands: readonly PaintCommand[]
  finish(): readonly PaintCommand[]
}

export function createRecordingPaintSurface(maxCommands: number = PPTX_RENDER_LIMITS.maxPaintCommands): RecordingPaintSurface {
  if (!Number.isSafeInteger(maxCommands) || maxCommands < 1 || maxCommands > PPTX_RENDER_LIMITS.maxPaintCommands) {
    throw new RenderCompileError('render.invalidLimit', '$.maxCommands', `maxCommands must be from 1 through ${PPTX_RENDER_LIMITS.maxPaintCommands}`)
  }
  const commands: PaintCommand[] = []
  let finished: readonly PaintCommand[] | undefined
  const freezeCommand = (command: PaintCommand): PaintCommand => {
    const freeze = (value: unknown): void => {
      if (typeof value !== 'object' || value === null || Object.isFrozen(value)) return
      for (const child of Object.values(value)) freeze(child)
      Object.freeze(value)
    }
    freeze(command)
    return command
  }
  return {
    get commands() { return finished ?? Object.freeze([...commands]) },
    push(command) {
      if (finished) throw new Error('recording paint surface is already finished')
      if (commands.length >= maxCommands) throw new RenderCompileError('render.paintBudget', '$.paint', `paint commands exceed ${maxCommands}`)
      commands.push(freezeCommand(command))
    },
    finish() {
      if (!finished) finished = Object.freeze([...commands])
      return finished
    },
  }
}

function paintParagraphs(paragraphs: readonly RenderParagraphNode[], surface: PaintSurface): void {
  for (const paragraph of paragraphs) {
    for (const run of paragraph.marker ? [paragraph.marker,...paragraph.runs] : paragraph.runs) {
      if (run.status === 'refused') {
        surface.push({
          kind: 'placeholder', sourceElementId: run.sourceElementId,
          rect: { x: run.x, y: paragraph.y, cx: Math.max(1, Math.round(run.fontSizeMilliPoints * 127 / 10)), cy: Math.max(1, run.lineHeightEmu) },
          reason: 'textRefusal', label: 'Text shaping refused',
        })
      } else {
        surface.push({ kind: 'glyphRun', sourceElementId: run.sourceElementId, run })
      }
    }
  }
}

function paintTextBody(textBody: RenderTextBodyNode, surface: PaintSurface): void {
  if (textBody.status === 'refused') {
    surface.push({
      kind: 'placeholder', sourceElementId: textBody.sourceElementId, rect: textBody.bounds,
      reason: 'textRefusal', label: textBody.refusalLabel ?? 'Text layout refused',
    })
    return
  }
  if (textBody.transform) {
    surface.push({kind:'save'})
    surface.push({kind:'transform',transform:textBody.transform})
  }
  paintParagraphs(textBody.paragraphs, surface)
  if (textBody.transform) surface.push({kind:'restore'})
}

function paintNode(node: RenderNode, surface: PaintSurface): void {
  surface.push({ kind: 'save' })
  surface.push({ kind: 'transform', transform: node.transform })
  if (node.clip) surface.push(node.clip.kind === 'roundRect' ? { kind: 'clipRoundRect', rect: node.clip.rect, radiusEmu: node.clip.radiusEmu } : { kind: 'clipRect', rect: node.clip.rect })
  switch (node.kind) {
    case 'shape':
      surface.push({ kind: 'path', sourceElementId: node.sourceElementId, path: node.path, fill: node.fill?.color, stroke: node.stroke })
      if (node.textBody) paintTextBody(node.textBody, surface)
      break
    case 'text':
      paintTextBody(node.textBody, surface)
      break
    case 'connector':
      surface.push({ kind: 'path', sourceElementId: node.sourceElementId, path: node.path, stroke: node.stroke, headArrow: node.headArrow, tailArrow: node.tailArrow,...(node.headEnd?{headEnd:node.headEnd}:{}),...(node.tailEnd?{tailEnd:node.tailEnd}:{}) })
      break
    case 'image':
      surface.push({ kind: 'image', sourceElementId: node.sourceElementId, role: node.role, assetId: node.assetId, rect: node.bounds, ...(node.crop ? { crop: node.crop } : {}) })
      break
    case 'table':
      for (const cell of node.cells) {
        if (!cell.fill && !cell.border) continue
        surface.push({ kind: 'save' })
        surface.push({ kind: 'transform', transform: { aPpm: 1_000_000, bPpm: 0, cPpm: 0, dPpm: 1_000_000, txEmu: cell.bounds.x, tyEmu: cell.bounds.y } })
        surface.push({
          kind: 'path', sourceElementId: node.sourceElementId,
          path: [{ kind: 'rect', rect: { x: 0, y: 0, cx: cell.bounds.cx, cy: cell.bounds.cy } }],
          fill: cell.fill?.color, stroke: cell.border,
        })
        surface.push({ kind: 'restore' })
      }
      // Paint every cell background before any cell text. Native table text may
      // overflow its cell, and a later sibling fill must not cover that text.
      for (const cell of node.cells) {
        surface.push({ kind: 'save' })
        surface.push({ kind: 'transform', transform: { aPpm: 1_000_000, bPpm: 0, cPpm: 0, dPpm: 1_000_000, txEmu: cell.bounds.x, tyEmu: cell.bounds.y } })
        if (cell.textBody) {
          paintTextBody(cell.textBody, surface)
        } else if (cell.paragraph) {
          surface.push({ kind: 'clipRect', rect: { x: 0, y: 0, cx: cell.bounds.cx, cy: cell.bounds.cy } })
          paintParagraphs([cell.paragraph], surface)
        }
        surface.push({ kind: 'restore' })
      }
      break
    case 'group':
      for (const child of node.children) paintNode(child, surface)
      break
    case 'placeholder':
      surface.push({ kind: 'placeholder', sourceElementId: node.sourceElementId, rect: node.bounds, reason: node.reason, label: node.label })
      break
  }
  surface.push({ kind: 'restore' })
}

/** Traverse in stable z-order into any renderer-neutral command sink. */
export function paintSlideRenderTree(tree: SlideRenderTree, surface: PaintSurface, maxCommands: number = PPTX_RENDER_LIMITS.maxPaintCommands): void {
  if (!Number.isSafeInteger(maxCommands) || maxCommands < 1 || maxCommands > PPTX_RENDER_LIMITS.maxPaintCommands) {
    throw new RenderCompileError('render.invalidLimit', '$.maxCommands', `maxCommands must be from 1 through ${PPTX_RENDER_LIMITS.maxPaintCommands}`)
  }
  // Stage the complete bounded command list before touching a caller surface.
  // A native paint-budget refusal is therefore atomic; arbitrary host-surface
  // exceptions during the subsequent replay remain owned by that host.
  const staging = createRecordingPaintSurface(maxCommands)
  staging.push({ kind: 'beginSlide', size: tree.size, background: tree.background.color })
  staging.push({ kind: 'save' })
  staging.push({ kind: 'clipRect', rect: tree.clip.rect })
  for (const node of tree.nodes) paintNode(node, staging)
  staging.push({ kind: 'restore' })
  staging.push({ kind: 'endSlide' })
  for (const command of staging.finish()) surface.push(command)
}

/**
 * Minimal host boundary for a Canvas2D-backed application. Context is generic on
 * purpose: this package does not name a browser global or allocate a drawing surface.
 */
export interface Canvas2DCommandAdapter<HostContext> {
  execute(context: HostContext, command: PaintCommand): void
}

export function replayPaintCommandsToCanvas2D<HostContext>(
  context: HostContext,
  commands: readonly PaintCommand[],
  adapter: Canvas2DCommandAdapter<HostContext>,
): void {
  for (const command of commands) adapter.execute(context, command)
}

export function paintSlideRenderTreeToCanvas2D<HostContext>(
  tree: SlideRenderTree,
  context: HostContext,
  adapter: Canvas2DCommandAdapter<HostContext>,
): void {
  const recording = createRecordingPaintSurface()
  paintSlideRenderTree(tree, recording)
  replayPaintCommandsToCanvas2D(context, recording.finish(), adapter)
}

import { useMemo, useState, type ReactNode } from 'react'
import {
  createRecordingPaintSurface,
  paintSlideRenderTree,
  presetPath,
  type PaintCommand,
  type RenderPathCommand,
  type SlideRenderTree,
} from '@injoffice/pptx-render'
import type { NativeShapePreset } from '@injoffice/pptx-native'
import { DsCallout, DsChip, DsField, DsSelect } from '../design-system/primitives'
import '../design-system/live-tools.css'

const PRESETS: readonly NativeShapePreset[] = ['rect', 'roundRect', 'ellipse', 'triangle', 'diamond', 'rightArrow', 'pentagon', 'hexagon', 'star5']
const SHAPE_CX = 4_200_000
const SHAPE_CY = 2_500_000

function makeTree(preset: NativeShapePreset, includeConnector: boolean): SlideRenderTree {
  const nodes: SlideRenderTree['nodes'][number][] = [{
    kind: 'shape',
    sourceElementId: 'demo-shape',
    sourceKind: 'shape',
    zIndex: 0,
    transform: { aPpm: 1_000_000, bPpm: 0, cPpm: 0, dPpm: 1_000_000, txEmu: 1_100_000, tyEmu: 1_050_000 },
    bounds: { x: 0, y: 0, cx: SHAPE_CX, cy: SHAPE_CY },
    compatibility: 'editable',
    preset,
    path: presetPath(preset, SHAPE_CX, SHAPE_CY),
    fill: { color: '0F6F8F' },
    stroke: { color: '1A6B52', widthEmu: 22_860, cap: 'flat', join: 'round', dash: 'solid' },
  }]

  if (includeConnector) {
    nodes.push({
      kind: 'connector',
      sourceElementId: 'demo-connector',
      sourceKind: 'connector',
      zIndex: 1,
      transform: { aPpm: 1_000_000, bPpm: 0, cPpm: 0, dPpm: 1_000_000, txEmu: 6_000_000, tyEmu: 2_000_000 },
      bounds: { x: 0, y: 0, cx: 2_400_000, cy: 1_200_000 },
      compatibility: 'editable',
      path: [{ kind: 'moveTo', x: 0, y: 0 }, { kind: 'lineTo', x: 2_400_000, y: 1_200_000 }],
      stroke: { color: 'C23A2B', widthEmu: 36_000, cap: 'round', join: 'round', dash: 'solid' },
      headArrow: false,
      tailArrow: true,
    })
  }

  return {
    version: 'pptx-render-tree/v2',
    documentId: 'render-command-demo',
    slideId: 'slide-one',
    slideIndex: 0,
    size: { cx: 10_000_000, cy: 5_625_000 },
    background: { color: 'F5F6F6' },
    clip: { kind: 'rect', rect: { x: 0, y: 0, cx: 10_000_000, cy: 5_625_000 } },
    nodes,
    assets: [],
    diagnostics: [],
  }
}

function recordCommands(tree: SlideRenderTree): readonly PaintCommand[] {
  const surface = createRecordingPaintSurface()
  paintSlideRenderTree(tree, surface)
  return surface.finish()
}

function geometryPreview(commands: readonly RenderPathCommand[]): ReactNode {
  const first = commands[0]
  if (!first) return null
  if (first.kind === 'rect') return <rect x="0" y="0" width={first.rect.cx} height={first.rect.cy} rx="0" />
  if (first.kind === 'roundRect') return <rect x="0" y="0" width={first.rect.cx} height={first.rect.cy} rx={first.radiusEmu} />
  if (first.kind === 'ellipse') return <ellipse cx={first.rect.cx / 2} cy={first.rect.cy / 2} rx={first.rect.cx / 2} ry={first.rect.cy / 2} />

  const data = commands.map((command) => {
    if (command.kind === 'moveTo') return `M ${command.x} ${command.y}`
    if (command.kind === 'lineTo') return `L ${command.x} ${command.y}`
    return command.kind === 'close' ? 'Z' : ''
  }).join(' ')
  return <path d={data} />
}

export default function PptxRenderPage() {
  const [preset, setPreset] = useState<NativeShapePreset>('roundRect')
  const [includeConnector, setIncludeConnector] = useState(true)
  const tree = useMemo(() => makeTree(preset, includeConnector), [includeConnector, preset])
  const commands = useMemo(() => recordCommands(tree), [tree])
  const shape = tree.nodes[0]

  return (
    <section className="capability-page capability-page--pptx-render ds" data-demo-surface="pptx-render" aria-labelledby="pptx-render-title">
      <header className="capability-intro ds-surf-head">
        <div>
          <p className="capability-package ds-surf-pkg">@injoffice/pptx-render</p>
          <h2 id="pptx-render-title">Turn a RenderTree into deterministic paint commands</h2>
          <p>The library emits renderer-neutral geometry and an ordered command stream. The small SVG is a host preview; it is not a browser-layout substitute for native Office text.</p>
        </div>
        <span className="capability-runtime" role="status">Browser-safe command layer</span>
      </header>

      <div className="capability-workspace ds-split ds-split--wide">
        <div className="capability-controls ds-split-main">
          <DsField label="Native shape preset">
            <DsSelect value={preset} onChange={(event) => setPreset(event.target.value as NativeShapePreset)}>
              {PRESETS.map((value) => <option key={value} value={value}>{value}</option>)}
            </DsSelect>
          </DsField>
          <label className="ds-check capability-check">
            <input type="checkbox" checked={includeConnector} onChange={(event) => setIncludeConnector(event.target.checked)} />
            Include a connector in the command stream
          </label>
          <div className="capability-slide-preview ds-shape-stage" aria-label={`Host preview of the ${preset} geometry`}>
            <svg viewBox={`0 0 ${SHAPE_CX} ${SHAPE_CY}`} role="img" aria-label={`${preset} native preset path`}>
              <g fill="var(--ds-select-soft)" stroke="var(--ds-select)" strokeWidth="36000">{shape.kind === 'shape' ? geometryPreview(shape.path) : null}</g>
            </svg>
          </div>
          <DsCallout tone="note" title="Host preview is not native text layout">
            Native text compilation additionally requires a manifest, exact font bytes, a resolver, and a trusted shaper. This page intentionally does not fake those host inputs.
          </DsCallout>
        </div>

        <div className="capability-result ds-split-side" aria-live="polite">
          <div className="capability-result-heading">
            <DsChip tone="green"><span className="capability-status capability-status--success">Recording complete</span></DsChip>
            <strong>{commands.length} commands</strong>
          </div>
          <ol className="capability-command-list">
            {commands.map((command, index) => (
              <li className="ds-command" key={`${index}-${command.kind}`}>
                <code>{command.kind}</code>
                {'sourceElementId' in command ? <span className="ds-muted">{command.sourceElementId}</span> : null}
              </li>
            ))}
          </ol>
        </div>
      </div>
    </section>
  )
}

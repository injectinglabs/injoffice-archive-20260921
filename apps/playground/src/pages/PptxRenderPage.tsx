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
    fill: { color: '625BF6' },
    stroke: { color: '312E81', widthEmu: 22_860, cap: 'flat', join: 'round', dash: 'solid' },
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
      stroke: { color: 'FF7867', widthEmu: 36_000, cap: 'round', join: 'round', dash: 'solid' },
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
    background: { color: 'F7FAFF' },
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
    <section className="capability-page capability-page--pptx-render" data-demo-surface="pptx-render" aria-labelledby="pptx-render-title">
      <header className="capability-intro">
        <div>
          <p className="capability-package">@injoffice/pptx-render</p>
          <h2 id="pptx-render-title">Turn a RenderTree into deterministic paint commands</h2>
          <p>The library emits renderer-neutral geometry and an ordered command stream. The small SVG is a host preview; it is not a browser-layout substitute for native Office text.</p>
        </div>
        <span className="capability-runtime" role="status">Browser-safe command layer</span>
      </header>

      <div className="capability-workspace">
        <div className="capability-controls">
          <label className="capability-field">
            Native shape preset
            <select value={preset} onChange={(event) => setPreset(event.target.value as NativeShapePreset)}>
              {PRESETS.map((value) => <option key={value} value={value}>{value}</option>)}
            </select>
          </label>
          <label className="capability-check">
            <input type="checkbox" checked={includeConnector} onChange={(event) => setIncludeConnector(event.target.checked)} />
            Include a connector in the command stream
          </label>
          <div className="capability-slide-preview" aria-label={`Host preview of the ${preset} geometry`}>
            <svg viewBox={`0 0 ${SHAPE_CX} ${SHAPE_CY}`} role="img" aria-label={`${preset} native preset path`}>
              <g fill="#625BF6" stroke="#312E81" strokeWidth="36000">{shape.kind === 'shape' ? geometryPreview(shape.path) : null}</g>
            </svg>
          </div>
          <p className="capability-note">Native text compilation additionally requires a manifest, exact font bytes, a resolver, and a trusted shaper. This page intentionally does not fake those host inputs.</p>
        </div>

        <div className="capability-result" aria-live="polite">
          <div className="capability-result-heading">
            <span className="capability-status capability-status--success">Recording complete</span>
            <strong>{commands.length} commands</strong>
          </div>
          <ol className="capability-command-list">
            {commands.map((command, index) => (
              <li key={`${index}-${command.kind}`}>
                <code>{command.kind}</code>
                {'sourceElementId' in command ? <span>{command.sourceElementId}</span> : null}
              </li>
            ))}
          </ol>
        </div>
      </div>
    </section>
  )
}

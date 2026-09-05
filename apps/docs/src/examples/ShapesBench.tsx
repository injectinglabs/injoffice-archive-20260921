import { useState } from 'react'
import { ShapePreview, SHAPE_KINDS, shapeLabel, type ShapeKind } from '@injoffice/shapes'
import { LiveBench } from '../components/LiveBench'

const PICKS = ['rect', 'roundRect', 'ellipse', 'triangle', 'diamond', 'plus', 'star5', 'hexagon'] as const

export function ShapesBench() {
  const [kind, setKind] = useState<ShapeKind>('roundRect')
  return (
    <LiveBench title="Live example" hint="@injoffice/shapes · ShapePreview">
      <div className="bench-controls">
        <label className="field">
          Preset
          <select value={kind} onChange={(event) => setKind(event.target.value as ShapeKind)}>
            {PICKS.map((item) => <option key={item} value={item}>{shapeLabel(item)}</option>)}
          </select>
        </label>
        <span className="badge">{SHAPE_KINDS.length} kinds in the package</span>
      </div>
      <ShapePreview kind={kind} width={280} height={140} fill="#f8e4d8" stroke="#d24a1a" strokeWidth={2} />
    </LiveBench>
  )
}

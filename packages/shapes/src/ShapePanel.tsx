import { useEffect, useReducer, useState } from 'react'
import type { ShapeManager } from './manager'
import { hasText } from './types'
import type { ShapeSpec } from './types'

// ShapePanel — the shape style side panel. Same binding pattern as
// ChartPanel/PivotPanel: onChange forces a re-render, every control writes
// through updateSpec so the live float updates immediately.

const COLOR_FIELDS: Array<{ key: keyof ShapeSpec; label: string }> = [
  { key: 'fill', label: 'Fill' },
  { key: 'stroke', label: 'Stroke' },
]

export function ShapePanel({ manager }: { manager: ShapeManager }) {
  const [, force] = useReducer((x: number) => x + 1, 0)
  const [selectedId, setSelectedId] = useState<string | null>(null)

  useEffect(() => manager.onChange(force), [manager])

  const shapes = manager.list()
  const active: ShapeSpec | undefined = (selectedId ? manager.getSpec(selectedId) : undefined) ?? shapes[shapes.length - 1]

  if (shapes.length === 0) {
    return <div className="ioc-panel ioc-panel--empty">No shapes yet — insert one from the toolbar.</div>
  }
  if (!active) return null

  const patch = (p: Partial<Omit<ShapeSpec, 'id' | 'kind'>>) => manager.updateSpec(active.id, p)

  return (
    <div className="ioc-panel" aria-label="Shape style">
      {shapes.length > 1 && (
        <label className="ioc-field">
          <span>Shape</span>
          <select value={active.id} onChange={(e) => setSelectedId(e.target.value)}>
            {shapes.map((s, i) => (
              <option key={s.id} value={s.id}>
                {i + 1}. {s.kind}
                {s.text ? ` — ${s.text.slice(0, 24)}` : ''}
              </option>
            ))}
          </select>
        </label>
      )}
      {COLOR_FIELDS.map(({ key, label }) => (
        <label className="ioc-field" key={key}>
          <span>{label}</span>
          <input
            type="color"
            value={(active[key] as string) || '#000000'}
            onChange={(e) => patch({ [key]: e.target.value } as Partial<ShapeSpec>)}
          />
        </label>
      ))}
      <label className="ioc-field">
        <span>Stroke width</span>
        <input
          type="number"
          min={0}
          max={12}
          step={0.5}
          value={active.strokeWidth ?? 1.5}
          onChange={(e) => patch({ strokeWidth: Number(e.target.value) })}
        />
      </label>
      {hasText(active.kind) && (
        <>
          <label className="ioc-field">
            <span>Text color</span>
            <input
              type="color"
              value={active.textColor || '#1d2427'}
              onChange={(e) => patch({ textColor: e.target.value })}
            />
          </label>
          <label className="ioc-field">
            <span>Font size</span>
            <input
              type="number"
              min={8}
              max={72}
              value={active.fontSize ?? 14}
              onChange={(e) => patch({ fontSize: Number(e.target.value) })}
            />
          </label>
        </>
      )}
      <button type="button" className="ioc-remove" onClick={() => manager.remove(active.id)}>
        Delete shape
      </button>
    </div>
  )
}

/** Toolbar's shape-kind picker data — grouped by category (a flat 82-item
 *  <select> is unusable). Re-exported from types.ts's SHAPE_CATEGORIES so
 *  the host only needs one import for the picker AND the panel. */
export { SHAPE_CATEGORIES } from './types'

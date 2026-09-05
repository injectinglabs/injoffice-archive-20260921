import { useEffect, useMemo, useReducer, useState } from 'react'
import { PivotCommandController } from './commands'
import type { PivotManager } from './manager'
import { AGG_KINDS } from './types'
import type { AggKind, PivotSpec, PivotValueField } from './types'

// PivotPanel — the pivot EDITING panel. Every mutation is routed through a
// PivotCommandController so hosts can connect the panel to the same snapshot
// undo path as public Univer commands.
//
// Slicers live here as value chips per filter field: click to toggle values
// in/out of the pivot. Same ioc-* styling hooks as the chart panel.

function toggle<T>(list: T[], item: T): T[] {
  return list.includes(item) ? list.filter((x) => x !== item) : [...list, item]
}

export type PivotPanelProps =
  | { controller: PivotCommandController; manager?: never }
  | { /** @deprecated Pass the controller returned by registerUniverPivotCommands to enable host undo. */ manager: PivotManager; controller?: never }

export function createPivotPanelCommandBindings(controller: PivotCommandController, id: string) {
  return {
    remove: () => controller.remove(id),
    update: (patch: Partial<Omit<PivotSpec, 'id' | 'nativeIdentity'>>) => controller.update(id, patch),
    setRowFields: (rows: string[]) => controller.setRowFields(id, rows),
    setColumnFields: (columns: string[]) => controller.setColumnFields(id, columns),
    setValueFields: (values: PivotValueField[]) => controller.setValueFields(id, values),
    setFilters: (filters?: Record<string, string[]>) => controller.setFilters(id, filters),
  }
}

export function PivotPanel({ controller: suppliedController, manager: legacyManager }: PivotPanelProps) {
  const controller = useMemo(
    () => {
      if (suppliedController) return suppliedController
      if (legacyManager) return new PivotCommandController(legacyManager)
      throw new TypeError('PivotPanel requires a controller or manager')
    },
    [suppliedController, legacyManager],
  )
  const manager = controller.manager
  const [, force] = useReducer((x: number) => x + 1, 0)
  const [selectedId, setSelectedId] = useState<string | null>(null)
  const [slicerField, setSlicerField] = useState<string>('')

  useEffect(() => manager.onChange(force), [manager])

  const pivots = manager.list()
  const active: PivotSpec | undefined =
    (selectedId ? manager.getSpec(selectedId) : undefined) ?? pivots[pivots.length - 1]

  if (pivots.length === 0) {
    return <div className="ioc-panel ioc-panel--empty">No pivots yet — select a data block (headers included) and insert one.</div>
  }
  if (!active) return null

  const fields = manager.sourceFieldsFor(active.id)
  const commands = createPivotPanelCommandBindings(controller, active.id)

  const slicerValues = slicerField ? manager.fieldValuesFor(active.id, slicerField) : []
  const activeFilter = slicerField ? active.filters?.[slicerField] : undefined

  const setFilter = (field: string, values: string[] | undefined) => {
    const next = { ...(active.filters ?? {}) }
    if (values === undefined) delete next[field]
    else next[field] = values
    commands.setFilters(Object.keys(next).length ? next : undefined)
  }

  return (
    <div className="ioc-panel">
      <div className="ioc-row">
        <select
          className="ioc-chart-select"
          value={active.id}
          onChange={(e) => setSelectedId(e.target.value)}
          aria-label="Pivot"
        >
          {pivots.map((p, i) => (
            <option key={p.id} value={p.id}>{`Pivot ${i + 1}`}</option>
          ))}
        </select>
        <button type="button" className="ioc-remove" onClick={commands.remove}>
          Remove
        </button>
      </div>

      <div className="ioc-series">
        <div className="ioc-series-head">Rows</div>
        {fields.map((f) => (
          <label key={f} className="ioc-check">
            <input
              type="checkbox"
              checked={active.rows.includes(f)}
              onChange={() => commands.setRowFields(toggle(active.rows, f))}
            />
            {f}
          </label>
        ))}
      </div>

      <label className="ioc-field">
        <span>Column field</span>
        <select
          value={active.columns[0] ?? ''}
          onChange={(e) => commands.setColumnFields(e.target.value ? [e.target.value] : [])}
        >
          <option value="">none</option>
          {fields.map((f) => (
            <option key={f} value={f}>{f}</option>
          ))}
        </select>
      </label>

      <div className="ioc-series">
        <div className="ioc-series-head">Values</div>
        {active.values.map((v, i) => (
          <div key={`${v.field}-${i}`} className="ioc-value-row">
            <select
              aria-label="Value field"
              value={v.field}
              onChange={(e) => {
                const values = active.values.slice()
                values[i] = { ...v, field: e.target.value }
                commands.setValueFields(values)
              }}
            >
              {fields.map((f) => (
                <option key={f} value={f}>{f}</option>
              ))}
            </select>
            <select
              aria-label="Aggregation"
              value={v.agg}
              onChange={(e) => {
                const values = active.values.slice()
                values[i] = { ...v, agg: e.target.value as AggKind }
                commands.setValueFields(values)
              }}
            >
              {AGG_KINDS.map((a) => (
                <option key={a} value={a}>{a}</option>
              ))}
            </select>
            <button
              type="button"
              aria-label="Remove value"
              disabled={active.values.length <= 1}
              onClick={() => commands.setValueFields(active.values.filter((_, j) => j !== i))}
            >
              ×
            </button>
          </div>
        ))}
        <button
          type="button"
          className="ioc-add"
          onClick={() => commands.setValueFields([...active.values, { field: fields[0], agg: 'sum' }])}
        >
          + value
        </button>
      </div>

      <div className="ioc-series">
        <div className="ioc-series-head">Slicer</div>
        <select aria-label="Slicer field" value={slicerField} onChange={(e) => setSlicerField(e.target.value)}>
          <option value="">choose field…</option>
          {fields.map((f) => (
            <option key={f} value={f}>{f}</option>
          ))}
        </select>
        {slicerField && (
          <div className="ioc-chips">
            {slicerValues.map((v) => {
              const selected = activeFilter === undefined || activeFilter.includes(v)
              return (
                <button
                  key={v}
                  type="button"
                  className={`ioc-chip${selected ? ' ioc-chip--on' : ''}`}
                  aria-pressed={selected}
                  onClick={() => {
                    // No filter yet = everything selected; first toggle
                    // materializes the full list minus the clicked value.
                    const current = activeFilter ?? slicerValues
                    const next = toggle(current, v)
                    setFilter(slicerField, next.length === slicerValues.length ? undefined : next)
                  }}
                >
                  {v || '(blank)'}
                </button>
              )
            })}
            {activeFilter !== undefined && (
              <button type="button" className="ioc-chip" onClick={() => setFilter(slicerField, undefined)}>
                clear
              </button>
            )}
          </div>
        )}
      </div>

      <label className="ioc-field">
        <span>Grand totals</span>
        <input
          type="checkbox"
          checked={active.grandTotals !== false}
          onChange={(e) => commands.update({ grandTotals: e.target.checked ? undefined : false })}
        />
      </label>
    </div>
  )
}

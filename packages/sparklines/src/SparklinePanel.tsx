import { useEffect, useMemo, useReducer, useState, type SVGProps } from 'react'
import { SparklineCommandController } from './commands'
import type { CreateSparklineInput } from './manager'
import { SPARKLINE_TYPES } from './types'
import type {
  EmptyCellBehavior,
  SparklineGeometry,
  SparklineSpec,
  SparklineType,
  SparklineValueReader,
  SparklineViewport,
} from './types'
import { validateSparkline } from './validation'

const EMPTY_CELLS: EmptyCellBehavior[] = ['gap', 'zero', 'connect']
const DEFAULT_VIEWPORT: SparklineViewport = { width: 120, height: 24, padding: 2 }

export interface SparklineInsertDraft {
  type: SparklineType
  sheetId: string
  startRow: string
  startColumn: string
  endRow: string
  endColumn: string
  targetRow: string
  targetColumn: string
  emptyCells: EmptyCellBehavior
  showMarkers: boolean
  showHigh: boolean
  showLow: boolean
  showFirst: boolean
  showLast: boolean
  showNegative: boolean
}

export function defaultSparklineInsertDraft(sheetId = 'sheet-1'): SparklineInsertDraft {
  return {
    type: 'line',
    sheetId,
    startRow: '0',
    startColumn: '0',
    endRow: '0',
    endColumn: '5',
    targetRow: '0',
    targetColumn: '6',
    emptyCells: 'gap',
    showMarkers: false,
    showHigh: false,
    showLow: false,
    showFirst: false,
    showLast: false,
    showNegative: false,
  }
}

function integerField(value: string, label: string): number {
  const parsed = Number(value)
  if (!Number.isInteger(parsed) || parsed < 0) throw new TypeError(`${label} must be a non-negative integer`)
  return parsed
}

/** Convert a labeled insert form into a validated create payload. */
export function sparklineInputFromDraft(draft: SparklineInsertDraft): CreateSparklineInput {
  const sheetId = draft.sheetId.trim()
  if (!sheetId) throw new TypeError('sheetId must not be empty')
  const input: CreateSparklineInput = {
    type: draft.type,
    source: {
      sheetId,
      startRow: integerField(draft.startRow, 'startRow'),
      startColumn: integerField(draft.startColumn, 'startColumn'),
      endRow: integerField(draft.endRow, 'endRow'),
      endColumn: integerField(draft.endColumn, 'endColumn'),
    },
    target: {
      sheetId,
      row: integerField(draft.targetRow, 'targetRow'),
      column: integerField(draft.targetColumn, 'targetColumn'),
    },
    options: {
      emptyCells: draft.emptyCells,
      ...(draft.showMarkers ? { showMarkers: true } : {}),
      ...(draft.showHigh ? { showHigh: true } : {}),
      ...(draft.showLow ? { showLow: true } : {}),
      ...(draft.showFirst ? { showFirst: true } : {}),
      ...(draft.showLast ? { showLast: true } : {}),
      ...(draft.showNegative ? { showNegative: true } : {}),
    },
  }
  const issues = validateSparkline({ ...input, id: 'draft' })
  if (issues.length) throw new TypeError(issues.map((issue) => `${issue.path} ${issue.message}`).join('; '))
  return input
}

export function createSparklinePanelCommandBindings(controller: SparklineCommandController) {
  return {
    create: (input: CreateSparklineInput) => controller.create(input),
    update: (id: string, patch: Partial<Omit<SparklineSpec, 'id' | 'groupId'>>) => controller.update(id, patch),
    remove: (id: string) => controller.remove(id),
    group: (memberIds: readonly string[]) => controller.group(memberIds),
    ungroup: (memberIds: readonly string[]) => controller.ungroup(memberIds),
  }
}

export function SparklineSvg({
  geometry,
  title,
  ...rest
}: { geometry: SparklineGeometry; title?: string } & SVGProps<SVGSVGElement>) {
  return (
    <svg
      xmlns="http://www.w3.org/2000/svg"
      viewBox={`0 0 ${geometry.width} ${geometry.height}`}
      role="img"
      aria-label={title ?? `${geometry.type} sparkline`}
      {...rest}
    >
      <line
        x1={0}
        y1={geometry.baselineY}
        x2={geometry.width}
        y2={geometry.baselineY}
        stroke="currentColor"
        strokeWidth={0.5}
      />
      {geometry.paths.map((path, index) => (
        <polyline
          key={`path-${index}`}
          points={path.points.map((point) => `${point.x},${point.y}`).join(' ')}
          fill="none"
          stroke={path.color}
          strokeWidth={path.width}
        />
      ))}
      {geometry.bars.map((bar, index) => (
        <rect
          key={`bar-${index}`}
          x={bar.x}
          y={bar.y}
          width={bar.width}
          height={bar.height}
          fill={bar.color}
        />
      ))}
      {geometry.markers.map((marker, index) => (
        <circle key={`marker-${index}`} cx={marker.x} cy={marker.y} r={1.5} fill={marker.color} />
      ))}
    </svg>
  )
}

export interface SparklinePanelProps {
  controller: SparklineCommandController
  defaultSheetId?: string
  readValues?: SparklineValueReader
  previewViewport?: SparklineViewport
}

export function SparklinePanel({
  controller,
  defaultSheetId = 'sheet-1',
  readValues,
  previewViewport = DEFAULT_VIEWPORT,
}: SparklinePanelProps) {
  const manager = controller.manager
  const commands = useMemo(() => createSparklinePanelCommandBindings(controller), [controller])
  const [, force] = useReducer((count: number) => count + 1, 0)
  const [selectedId, setSelectedId] = useState<string | null>(null)
  const [groupIds, setGroupIds] = useState<string[]>([])
  const [draft, setDraft] = useState(() => defaultSparklineInsertDraft(defaultSheetId))
  const [error, setError] = useState<string | null>(null)

  useEffect(() => manager.onChange(force), [manager])

  const sparklines = manager.list()
  const groups = manager.listGroups()
  const active = (selectedId ? manager.get(selectedId) : undefined) ?? sparklines[sparklines.length - 1]
  const preview = (() => {
    if (!active || !readValues) return null
    try { return manager.render(active.id, readValues, previewViewport) } catch { return null }
  })()

  const patchDraft = (patch: Partial<SparklineInsertDraft>) => setDraft((current) => ({ ...current, ...patch }))
  const toggleGroupMember = (id: string) => {
    setGroupIds((current) => current.includes(id) ? current.filter((member) => member !== id) : [...current, id])
  }
  const run = (work: () => void) => {
    setError(null)
    try { work() } catch (caught) { setError(caught instanceof Error ? caught.message : String(caught)) }
  }

  const sameType = groupIds.length >= 2 && groupIds.every((id) => manager.get(id)?.type === manager.get(groupIds[0]!)?.type)

  return (
    <div className="ioc-panel ioc-sparkline-panel">
      <section className="ioc-series" aria-labelledby="ioc-sparkline-insert-title">
        <div className="ioc-series-head" id="ioc-sparkline-insert-title">Insert sparkline</div>
        <label className="ioc-field">
          <span>Type</span>
          <select aria-label="Sparkline type" value={draft.type} onChange={(event) => patchDraft({ type: event.target.value as SparklineType })}>
            {SPARKLINE_TYPES.map((type) => <option key={type} value={type}>{type}</option>)}
          </select>
        </label>
        <label className="ioc-field">
          <span>Sheet</span>
          <input aria-label="Sparkline sheet" value={draft.sheetId} onChange={(event) => patchDraft({ sheetId: event.target.value })} />
        </label>
        <div className="ioc-row" role="group" aria-label="Source range">
          <label className="ioc-field"><span>Start row</span><input inputMode="numeric" aria-label="Source start row" value={draft.startRow} onChange={(event) => patchDraft({ startRow: event.target.value })} /></label>
          <label className="ioc-field"><span>Start column</span><input inputMode="numeric" aria-label="Source start column" value={draft.startColumn} onChange={(event) => patchDraft({ startColumn: event.target.value })} /></label>
        </div>
        <div className="ioc-row" role="group" aria-label="Source end">
          <label className="ioc-field"><span>End row</span><input inputMode="numeric" aria-label="Source end row" value={draft.endRow} onChange={(event) => patchDraft({ endRow: event.target.value })} /></label>
          <label className="ioc-field"><span>End column</span><input inputMode="numeric" aria-label="Source end column" value={draft.endColumn} onChange={(event) => patchDraft({ endColumn: event.target.value })} /></label>
        </div>
        <div className="ioc-row" role="group" aria-label="Target cell">
          <label className="ioc-field"><span>Target row</span><input inputMode="numeric" aria-label="Target row" value={draft.targetRow} onChange={(event) => patchDraft({ targetRow: event.target.value })} /></label>
          <label className="ioc-field"><span>Target column</span><input inputMode="numeric" aria-label="Target column" value={draft.targetColumn} onChange={(event) => patchDraft({ targetColumn: event.target.value })} /></label>
        </div>
        <label className="ioc-field">
          <span>Empty cells</span>
          <select aria-label="Empty cell behavior" value={draft.emptyCells} onChange={(event) => patchDraft({ emptyCells: event.target.value as EmptyCellBehavior })}>
            {EMPTY_CELLS.map((value) => <option key={value} value={value}>{value}</option>)}
          </select>
        </label>
        <div className="ioc-series" role="group" aria-label="Markers">
          {([
            ['showMarkers', 'Markers'],
            ['showHigh', 'High'],
            ['showLow', 'Low'],
            ['showFirst', 'First'],
            ['showLast', 'Last'],
            ['showNegative', 'Negative'],
          ] as const).map(([key, label]) => (
            <label key={key} className="ioc-check">
              <input type="checkbox" checked={draft[key]} onChange={(event) => patchDraft({ [key]: event.target.checked })} />
              {label}
            </label>
          ))}
        </div>
        <button
          type="button"
          className="ioc-primary"
          onClick={() => run(() => {
            const created = commands.create(sparklineInputFromDraft(draft))
            setSelectedId(created.id)
          })}
        >
          Insert sparkline
        </button>
      </section>

      {error ? <div role="alert">{error}</div> : null}

      {sparklines.length === 0 ? (
        <div className="ioc-panel ioc-panel--empty">No sparklines yet — insert one from a one-row or one-column range.</div>
      ) : active ? (
        <>
          <div className="ioc-row">
            <select className="ioc-chart-select" aria-label="Sparkline" value={active.id} onChange={(event) => setSelectedId(event.target.value)}>
              {sparklines.map((spec, index) => (
                <option key={spec.id} value={spec.id}>{spec.id || `Sparkline ${index + 1}`} ({spec.type})</option>
              ))}
            </select>
            <button type="button" className="ioc-remove" onClick={() => run(() => { commands.remove(active.id); setGroupIds((ids) => ids.filter((id) => id !== active.id)) })}>
              Remove
            </button>
          </div>

          {preview ? <SparklineSvg geometry={preview} title={`${active.type} sparkline preview`} /> : null}

          <label className="ioc-field">
            <span>Type</span>
            <select aria-label="Selected sparkline type" value={active.type} onChange={(event) => run(() => commands.update(active.id, { type: event.target.value as SparklineType }))}>
              {SPARKLINE_TYPES.map((type) => <option key={type} value={type}>{type}</option>)}
            </select>
          </label>
          <label className="ioc-field">
            <span>Empty cells</span>
            <select
              aria-label="Selected empty cell behavior"
              value={active.options?.emptyCells ?? 'gap'}
              onChange={(event) => run(() => commands.update(active.id, { options: { ...active.options, emptyCells: event.target.value as EmptyCellBehavior } }))}
            >
              {EMPTY_CELLS.map((value) => <option key={value} value={value}>{value}</option>)}
            </select>
          </label>
          <div className="ioc-series" role="group" aria-label="Selected markers">
            {([
              ['showMarkers', 'Markers'],
              ['showHigh', 'High'],
              ['showLow', 'Low'],
              ['showFirst', 'First'],
              ['showLast', 'Last'],
              ['showNegative', 'Negative'],
            ] as const).map(([key, label]) => (
              <label key={key} className="ioc-check">
                <input
                  type="checkbox"
                  checked={Boolean(active.options?.[key])}
                  onChange={(event) => run(() => commands.update(active.id, { options: { ...active.options, [key]: event.target.checked || undefined } }))}
                />
                {label}
              </label>
            ))}
          </div>

          <section className="ioc-series" aria-labelledby="ioc-sparkline-group-title">
            <div className="ioc-series-head" id="ioc-sparkline-group-title">Group workflow</div>
            {sparklines.map((spec) => (
              <label key={spec.id} className="ioc-check">
                <input type="checkbox" checked={groupIds.includes(spec.id)} onChange={() => toggleGroupMember(spec.id)} />
                {spec.id}{spec.groupId ? ` · ${spec.groupId}` : ''}
              </label>
            ))}
            <div className="ioc-row">
              <button type="button" disabled={!sameType} onClick={() => run(() => { commands.group(groupIds); setGroupIds([]) })}>
                Group selected
              </button>
              <button type="button" disabled={groupIds.length === 0} onClick={() => run(() => { commands.ungroup(groupIds); setGroupIds([]) })}>
                Ungroup selected
              </button>
            </div>
            <p className="ioc-note">{groups.length} group{groups.length === 1 ? '' : 's'} · select two sparklines of the same type to group.</p>
          </section>
        </>
      ) : null}
    </div>
  )
}

import type { OutlineCommandController } from './commands'
import type { OutlineAxis, OutlineGroup } from './types'

export interface OutlineGutterControl {
  id: string
  axis: OutlineAxis
  sheetId: string
  start: number
  end: number
  depth: number
  collapsed: boolean
}

export interface OutlineGutterModel {
  axis: OutlineAxis
  sheetId: string
  maxLevel: number
  controls: OutlineGutterControl[]
}

export function outlineGroupDepth(groups: readonly OutlineGroup[], group: OutlineGroup): number {
  return groups.filter((candidate) =>
    candidate.sheetId === group.sheetId
    && candidate.axis === group.axis
    && candidate.id !== group.id
    && candidate.start <= group.start
    && candidate.end >= group.end,
  ).length + 1
}

/** Nested outline margin for one sheet axis, ordered outer-to-inner. */
export function layoutOutlineGutter(
  groups: readonly OutlineGroup[],
  sheetId: string,
  axis: OutlineAxis,
): OutlineGutterModel {
  const scoped = groups.filter((group) => group.sheetId === sheetId && group.axis === axis)
  const controls = scoped
    .map((group) => ({
      id: group.id,
      axis: group.axis,
      sheetId: group.sheetId,
      start: group.start,
      end: group.end,
      depth: outlineGroupDepth(scoped, group),
      collapsed: group.collapsed,
    }))
    .sort((left, right) => left.start - right.start || right.end - left.end || left.id.localeCompare(right.id))
  return {
    axis,
    sheetId,
    maxLevel: controls.reduce((max, control) => Math.max(max, control.depth), 0),
    controls,
  }
}

/** Excel-style level buttons: level 1 collapses every group, higher levels
 * reveal nested bands whose depth is strictly less than the requested level. */
export function collapseToOutlineLevel(
  model: OutlineGutterModel,
  level: number,
): Array<{ id: string; collapsed: boolean }> {
  if (!Number.isInteger(level) || level < 1) throw new TypeError('outline level must be a positive integer')
  return model.controls.map((control) => ({ id: control.id, collapsed: control.depth >= level }))
}

export function applyOutlineLevel(
  controller: OutlineCommandController,
  sheetId: string,
  axis: OutlineAxis,
  level: number,
): number {
  const model = layoutOutlineGutter(controller.manager.list(sheetId, axis), sheetId, axis)
  let changed = 0
  for (const next of collapseToOutlineLevel(model, level)) {
    const current = controller.manager.get(next.id)
    if (!current || current.collapsed === next.collapsed) continue
    controller.setCollapsed(next.id, next.collapsed)
    changed += 1
  }
  return changed
}

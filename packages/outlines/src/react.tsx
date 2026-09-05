import { useEffect, useReducer, useState, type KeyboardEvent } from 'react'
import { OutlineCommandController } from './commands'
import { applyOutlineLevel, layoutOutlineGutter, type OutlineGutterControl } from './gutter'
import type { OutlineAxis } from './types'

export { applyOutlineLevel, collapseToOutlineLevel, layoutOutlineGutter, outlineGroupDepth } from './gutter'
export type { OutlineGutterControl, OutlineGutterModel } from './gutter'

export function createOutlineGutterCommandBindings(controller: OutlineCommandController) {
  return {
    expand: (id: string) => controller.setCollapsed(id, false),
    collapse: (id: string) => controller.setCollapsed(id, true),
    toggle: (id: string) => {
      const group = controller.manager.get(id)
      return controller.setCollapsed(id, group ? !group.collapsed : false)
    },
    showLevel: (sheetId: string, axis: OutlineAxis, level: number) => applyOutlineLevel(controller, sheetId, axis, level),
  }
}

function controlLabel(control: OutlineGutterControl): string {
  const dimension = control.axis === 'row' ? 'rows' : 'columns'
  const action = control.collapsed ? 'Expand' : 'Collapse'
  return `${action} ${dimension} ${control.start + 1}–${control.end + 1}`
}

export interface OutlineGutterProps {
  controller: OutlineCommandController
  sheetId: string
  axis: OutlineAxis
}

export function OutlineGutter({ controller, sheetId, axis }: OutlineGutterProps) {
  const [, force] = useReducer((count: number) => count + 1, 0)
  const [activeId, setActiveId] = useState<string | null>(null)
  useEffect(() => controller.manager.onChange(force), [controller])

  const model = layoutOutlineGutter(controller.manager.list(sheetId, axis), sheetId, axis)
  const commands = createOutlineGutterCommandBindings(controller)
  const levels = Array.from({ length: model.maxLevel }, (_, index) => index + 1)
  const dimension = axis === 'row' ? 'row' : 'column'

  const focusControl = (id: string | undefined) => {
    if (!id) return
    setActiveId(id)
  }

  const onListKeyDown = (event: KeyboardEvent<HTMLElement>) => {
    if (model.controls.length === 0) return
    const currentIndex = Math.max(0, model.controls.findIndex((control) => control.id === activeId))
    if (event.key === 'ArrowDown' || event.key === 'ArrowRight') {
      event.preventDefault()
      focusControl(model.controls[(currentIndex + 1) % model.controls.length]?.id)
    } else if (event.key === 'ArrowUp' || event.key === 'ArrowLeft') {
      event.preventDefault()
      focusControl(model.controls[(currentIndex - 1 + model.controls.length) % model.controls.length]?.id)
    } else if (event.key === 'Home') {
      event.preventDefault()
      focusControl(model.controls[0]?.id)
    } else if (event.key === 'End') {
      event.preventDefault()
      focusControl(model.controls.at(-1)?.id)
    } else if ((event.key === 'Enter' || event.key === ' ') && activeId) {
      event.preventDefault()
      commands.toggle(activeId)
    }
  }

  if (model.controls.length === 0) {
    return <div className="ioc-panel ioc-panel--empty">No {dimension} groups on this sheet.</div>
  }

  return (
    <div className="ioc-outline-gutter" role="region" aria-label={`${dimension} outline margin`}>
      <div className="ioc-outline-levels" role="toolbar" aria-label={`${dimension} outline levels`}>
        {levels.map((level) => (
          <button
            key={level}
            type="button"
            aria-label={`Show outline level ${level}`}
            onClick={() => commands.showLevel(sheetId, axis, level)}
          >
            {level}
          </button>
        ))}
      </div>
      <ol
        className="ioc-outline-controls"
        role="list"
        aria-label={`${dimension} outline groups`}
        onKeyDown={onListKeyDown}
      >
        {model.controls.map((control) => (
          <li key={control.id} style={{ paddingLeft: (control.depth - 1) * 12 }}>
            <button
              type="button"
              aria-expanded={!control.collapsed}
              aria-current={activeId === control.id ? 'true' : undefined}
              aria-label={controlLabel(control)}
              onClick={() => {
                setActiveId(control.id)
                commands.toggle(control.id)
              }}
            >
              {control.collapsed ? '+' : '−'}
            </button>
          </li>
        ))}
      </ol>
    </div>
  )
}

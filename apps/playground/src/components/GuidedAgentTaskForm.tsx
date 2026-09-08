import { useEffect, useId, useRef, useState } from 'react'
import type { JsonObject } from '@injoffice/agent-tools'
import type { AgentTool } from '../route'
import { buildGuidedTaskRequest, getGuidedTaskOptions, initialGuidedTaskValues, type GuidedTaskValues } from '../guidedAgentTasks'
import './GuidedAgentTaskForm.css'

type Props = { tool: AgentTool; context: JsonObject; defaultRequest: string; disabled: boolean; onRequestChange: (request: string) => void }

/** Mount once per inspected sample. Unrelated parent renders never reset an edit. */
export default function GuidedAgentTaskForm({ tool, context, defaultRequest, disabled, onRequestChange }: Props) {
  const id = useId()
  const [values, setValues] = useState(() => initialGuidedTaskValues(tool, context, defaultRequest))
  const lastEmitted = useRef<string | undefined>(undefined)
  const options = getGuidedTaskOptions(tool, context)
  const result = buildGuidedTaskRequest(tool, context, values)
  const target = options.targets.find((item) => item.id === values.target)
  const update = (key: keyof GuidedTaskValues, value: string) => setValues((previous) => ({ ...previous, [key]: value }))
  useEffect(() => {
    if (lastEmitted.current === result.request) return
    lastEmitted.current = result.request
    onRequestChange(result.request)
  }, [result.request, onRequestChange])

  return <fieldset className="agent-guided-task" data-agent-guided-task={tool} disabled={disabled} aria-describedby={`${id}-help${result.error ? ` ${id}-error` : ''}`}>
    <legend>Customize this task</legend>
    {tool === 'pdf' ? <div className="agent-guided-task__fields">
      <div><label htmlFor={`${id}-page`}>Page</label><select id={`${id}-page`} data-agent-task-page value={values.page} onChange={(event) => update('page', event.target.value)}>
        {!options.pages.length && <option value="">No editable pages</option>}
        {options.pages.map((page, index) => <option key={`${page}-${index}`} value={page}>Page {page}</option>)}
      </select></div>
      <div><label htmlFor={`${id}-degrees`}>Rotate by</label><select id={`${id}-degrees`} data-agent-task-degrees value={values.degrees} onChange={(event) => update('degrees', event.target.value)}>
        {!options.degrees.length && <option value="">No available rotations</option>}
        {options.degrees.map((degrees, index) => <option key={`${degrees}-${index}`} value={degrees}>{Math.abs(degrees)}° {degrees < 0 ? 'counterclockwise' : 'clockwise'}</option>)}
      </select></div>
    </div> : <>
      <label htmlFor={`${id}-target`}>{tool === 'sheets' ? 'Workstream' : 'Text to replace'}</label>
      <select id={`${id}-target`} data-agent-task-target value={values.target} onChange={(event) => update('target', event.target.value)}>
        {!options.targets.some((item) => !item.unavailable) && <option value="">No unique editable targets</option>}
        {options.targets.map((item, index) => <option key={`${item.id}-${index}`} value={item.id} disabled={!!item.unavailable}>{item.label}{item.unavailable ? ' (unavailable)' : ''}</option>)}
      </select>
      {target && <p className="agent-guided-task__current" data-agent-task-current><span>{tool === 'sheets' ? 'Current status' : 'Current text'}:</span> {target.currentValue}</p>}
      <label htmlFor={`${id}-value`}>{tool === 'sheets' ? 'New status' : 'Replacement text'}</label>
      {tool === 'sheets' ? <select id={`${id}-value`} data-agent-task-value value={values.value} onChange={(event) => update('value', event.target.value)}>
        {!options.values.length && <option value="">No available statuses</option>}
        {options.values.map((value, index) => <option key={`${value}-${index}`} value={value}>{value}</option>)}
      </select> : <input id={`${id}-value`} data-agent-task-value type="text" maxLength={1000} value={values.value} onChange={(event) => update('value', event.target.value)} aria-invalid={!!result.error} aria-describedby={`${id}-help${result.error ? ` ${id}-error` : ''}`} />}
    </>}
    <p id={`${id}-help`} className="agent-guided-task__help">{tool === 'pdf' ? 'Pages come from the inspected sample. Rotation is relative to the current orientation.' : tool === 'sheets' ? 'Workstreams and statuses come from the inspected sample.' : 'Choose unique inspected text. This sample supports one single-line replacement without double quotes.'} Nothing is written until you approve.</p>
    {result.error && <p id={`${id}-error`} className="agent-guided-task__error" role="alert" data-agent-task-error>{result.error}</p>}
  </fieldset>
}

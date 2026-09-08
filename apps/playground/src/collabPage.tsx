import { useId, useState } from 'react'
import { DsChip } from './design-system/primitives'
import './design-system/live-tools.css'
import { CollabDemo as ServerCollabDemo } from './collabDemo'
import { CollabSimulator } from './collabSimulator'
import type { CollabCompositionProps } from './collabComposition'

type Mode = 'simulation' | 'server'

export function CollabDemo({ fixedFormat, initialHash }: CollabCompositionProps = {}) {
  const [mode, setMode] = useState<Mode>('simulation')
  const id = useId()
  const chooseMode = (next: Mode) => {
    if (next === mode) return
    if (!window.confirm('Switch collaboration mode? The current room will close and unsaved demo edits will be lost.')) return
    setMode(next)
  }

  return (
    <div className="collab-page ds">
      <nav className="collab-mode-switcher ds-workstrip" aria-label="Collaboration demo mode">
        <div>
          <strong>Demo mode</strong>
          <span className="ds-muted">{mode === 'simulation' ? 'No server or upload required' : 'Exercises the HTTP and SSE adapter'}</span>
        </div>
        <div className="ds-segment" role="tablist" aria-label="Choose collaboration runtime" onKeyDown={event => {
          if (!['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) return
          event.preventDefault()
          const tabs = Array.from(event.currentTarget.querySelectorAll<HTMLButtonElement>('[role="tab"]'))
          const index = tabs.indexOf(document.activeElement as HTMLButtonElement)
          tabs[event.key === 'Home' ? 0 : event.key === 'End' ? tabs.length - 1 : (index + (event.key === 'ArrowLeft' ? -1 : 1) + tabs.length) % tabs.length]?.focus()
        }}>
          <button id={`${id}-simulation`} aria-controls={`${id}-panel`} tabIndex={mode === 'simulation' ? 0 : -1} type="button" role="tab" aria-selected={mode === 'simulation'} onClick={() => chooseMode('simulation')}>Two-editor simulation</button>
          <button id={`${id}-server`} aria-controls={`${id}-panel`} tabIndex={mode === 'server' ? 0 : -1} type="button" role="tab" aria-selected={mode === 'server'} onClick={() => chooseMode('server')}>HTTP + SSE integration</button>
        </div>
        {mode === 'simulation' ? <DsChip tone="green">No server or upload required</DsChip> : <DsChip tone="blue">HTTP + SSE integration</DsChip>}
      </nav>
      <div id={`${id}-panel`} role="tabpanel" aria-labelledby={`${id}-${mode}`} tabIndex={0}>
        {mode === 'simulation' ? <CollabSimulator fixedFormat={fixedFormat} initialHash={initialHash} /> : <ServerCollabDemo fixedFormat={fixedFormat} initialHash={initialHash} />}
      </div>
    </div>
  )
}

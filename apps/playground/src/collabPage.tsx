import { useState } from 'react'
import { DsChip } from './design-system/primitives'
import './design-system/live-tools.css'
import { CollabDemo as ServerCollabDemo } from './collabDemo'
import { CollabSimulator } from './collabSimulator'

type Mode = 'simulation' | 'server'

export function CollabDemo() {
  const [mode, setMode] = useState<Mode>('simulation')

  return (
    <div className="collab-page ds">
      <nav className="collab-mode-switcher ds-workstrip" aria-label="Collaboration demo mode">
        <div>
          <strong>Demo mode</strong>
          <span className="ds-muted">{mode === 'simulation' ? 'No server or upload required' : 'Exercises the HTTP and SSE adapter'}</span>
        </div>
        <div className="ds-segment" role="tablist" aria-label="Choose collaboration runtime">
          <button type="button" role="tab" aria-selected={mode === 'simulation'} onClick={() => setMode('simulation')}>Two-editor simulation</button>
          <button type="button" role="tab" aria-selected={mode === 'server'} onClick={() => setMode('server')}>HTTP + SSE integration</button>
        </div>
        {mode === 'simulation' ? <DsChip tone="green">No server or upload required</DsChip> : <DsChip tone="blue">HTTP + SSE integration</DsChip>}
      </nav>
      <div role="tabpanel">
        {mode === 'simulation' ? <CollabSimulator /> : <ServerCollabDemo />}
      </div>
    </div>
  )
}

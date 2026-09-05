import { useState } from 'react'
import { CollabDemo as ServerCollabDemo } from './collabDemo'
import { CollabSimulator } from './collabSimulator'

type Mode = 'simulation' | 'server'

export function CollabDemo() {
  const [mode, setMode] = useState<Mode>('simulation')

  return (
    <div className="collab-page">
      <nav className="collab-mode-switcher" aria-label="Collaboration demo mode">
        <div>
          <strong>Demo mode</strong>
          <span>{mode === 'simulation' ? 'No server or upload required' : 'Exercises the HTTP and SSE adapter'}</span>
        </div>
        <div role="tablist" aria-label="Choose collaboration runtime">
          <button type="button" role="tab" aria-selected={mode === 'simulation'} onClick={() => setMode('simulation')}>Two-editor simulation</button>
          <button type="button" role="tab" aria-selected={mode === 'server'} onClick={() => setMode('server')}>HTTP + SSE integration</button>
        </div>
      </nav>
      <div role="tabpanel">
        {mode === 'simulation' ? <CollabSimulator /> : <ServerCollabDemo />}
      </div>
    </div>
  )
}

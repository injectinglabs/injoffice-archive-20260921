import { useState } from 'react'
import type { DemoRecipeSource } from '../demoRecipes'

// Keep source text out of the initial bundle; load only the selected example.
const sourceFiles = import.meta.glob<string>(['../pages/*.tsx', '../collabSimulator.tsx'], { query: '?raw', import: 'default' })

export default function DemoSource({ source }: { source: DemoRecipeSource }) {
  const [code, setCode] = useState('')
  const [status, setStatus] = useState<'idle' | 'loading' | 'ready' | 'error'>('idle')
  const [copied, setCopied] = useState(false)
  const loadSource = async () => {
    if (status === 'ready' || status === 'loading') return
    const load = sourceFiles[source.path.replace('apps/playground/src/', '../')]
    setStatus('loading')
    try {
      if (!load) throw new Error('Source is not bundled')
      setCode(await load())
      setStatus('ready')
    } catch {
      setStatus('error')
    }
  }

  return (
    <details className="demo-source" onToggle={(event) => { if (event.currentTarget.open) void loadSource() }}>
      <summary>View this demo’s source</summary>
      <div className="demo-source__toolbar">
        <a href={source.href} target="_blank" rel="noreferrer">{source.path}</a>
        {status === 'ready' ? <button type="button" onClick={async () => {
          try { await navigator.clipboard.writeText(code); setCopied(true) } catch { setCopied(false) }
        }}>{copied ? 'Copied' : 'Copy source'}</button> : null}
      </div>
      {status === 'loading' ? <p role="status">Loading source…</p> : null}
      {status === 'error' ? <p role="status">Source could not load. <button type="button" onClick={() => { void loadSource() }}>Retry</button> or open the repository link above.</p> : null}
      {status === 'ready' ? <pre tabIndex={0} aria-label="Demo source code"><code>{code}</code></pre> : null}
    </details>
  )
}

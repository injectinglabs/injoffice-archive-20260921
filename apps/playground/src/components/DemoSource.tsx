import { useState } from 'react'
import type { DemoRecipeSource } from '../demoRecipes'

// Keep source text out of the initial bundle; load only the selected example.
const sourceFiles = import.meta.glob<string>(['../pages/*.tsx', '../collabSimulator.tsx'], { query: '?raw', import: 'default' })
const agentIntegration = `import { createAgentSession } from '@injoffice/agent-tools'

// Your host supplies the artifact, format adapter, actor and operations.
const session = await createAgentSession({ artifact, adapter, actor,
  confirmDestructive: async ({ confirmation }) => confirmation === 'approved',
})
const change = await session.plan(operations, {
  expectedRevision: session.identity.revision,
  expectedFingerprint: session.identity.fingerprint,
})
const diff = await change.diff()
const validation = await change.validate()
// Show the diff and wait for approval of this exact change.
if (validation.valid && approvedByUser) {
  const receipt = await change.commit({ idempotencyKey, confirmation: 'approved' })
  showResult(receipt.verification)
}`

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
    <>
    {source.path.endsWith('/AgentPage.tsx') && <details className="demo-source">
      <summary>Minimal host integration</summary>
      <p>The host supplies file I/O through an adapter and owns approval. This sketch has no model dependency; the full demo source below shows how the sample adapters are used.</p>
      <pre className="ds-code" tabIndex={0} aria-label="Agent integration sketch"><code>{agentIntegration}</code></pre>
    </details>}
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
      {status === 'ready' ? <pre className="ds-code" tabIndex={0} aria-label="Demo source code"><code>{code}</code></pre> : null}
    </details>
    </>
  )
}

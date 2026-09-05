import { useEffect, type ReactNode } from 'react'
import { PresenceStack, type PresenceSource } from '../../../../packages/collab/src/index.js'
import { COLLAB_COPY, parseCollabQuery, writeCollabQuery } from '../collabScope'

export type ConnectionState = 'idle' | 'joining' | 'live' | 'error'

export const XLSX_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
export const SAMPLE_XLSX = `${import.meta.env.BASE_URL}collab-demo.xlsx`
export const COLLAB_API_BASE = (import.meta.env.VITE_INJOFFICE_API_BASE ?? '').replace(/\/$/, '')

export function defaultCollabName(): string {
  const stored = sessionStorage.getItem('injoffice-collab-name')
  if (stored) return stored
  return `Tab ${Math.floor(Math.random() * 90 + 10)}`
}

export async function mintCollabArtifact(file: Blob): Promise<string> {
  let response: Response
  try {
    response = await fetch(`${COLLAB_API_BASE}/v1/xlsx/extract`, {
      method: 'POST',
      headers: { 'Content-Type': XLSX_TYPE },
      body: file,
    })
  } catch {
    throw new Error('Extract failed. Start the local sidecar with `go run ./cmd/injoffice-server` from go/injoffice-server.')
  }
  const detail = await response.text()
  if (!response.ok) throw new Error(`Extract failed (${response.status}): ${detail}`)
  const id = response.headers.get('X-InjOffice-Artifact-Id')?.trim()
  if (!id) throw new Error('Extract did not return an opaque artifact ID.')
  return id
}

export async function fetchSampleWorkbook(): Promise<Blob> {
  const response = await fetch(SAMPLE_XLSX)
  if (!response.ok) throw new Error(`Sample workbook missing (${response.status}).`)
  return response.blob()
}

function connectionLabel(connection: ConnectionState): string {
  return connection === 'live' ? 'Connected' : connection === 'joining' ? 'Joining' : connection === 'error' ? 'Not connected' : 'Local only'
}

export function CollabRoomChrome({
  title,
  connection,
  name,
  onNameChange,
  artifact,
  onArtifactChange,
  busy,
  joinDisabled,
  onJoin,
  actions,
  presence,
  status,
  error,
  proof,
  editorTitle,
  editorHint,
  editorMeta,
  metrics,
  tryIt,
  protocolExtra,
  children,
}: {
  title: string
  connection: ConnectionState
  name: string
  onNameChange: (name: string) => void
  artifact: string
  onArtifactChange: (artifact: string) => void
  busy: boolean
  joinDisabled?: boolean
  onJoin: () => void
  actions?: ReactNode
  presence?: PresenceSource | null
  status: string
  error: string
  proof: string
  editorTitle: string
  editorHint: string
  editorMeta: string
  metrics: { peers: number; applied: number; pending: number }
  tryIt: ReactNode
  protocolExtra?: ReactNode
  children: ReactNode
}) {
  useEffect(() => {
    sessionStorage.setItem('injoffice-collab-name', name)
  }, [name])

  useEffect(() => {
    const { format } = parseCollabQuery(window.location.href)
    writeCollabQuery({ artifact, format })
  }, [artifact])

  return (
    <>
      <div className="native-toolbar collab-toolbar workbench-toolbar" role="group" aria-label="Collaboration actions">
        <strong>{title}</strong>
        <span className={`collab-connection collab-connection--${connection} workbench-badge`} role="status" aria-label={`Connection: ${connection}`}>
          <i aria-hidden="true" />
          {connectionLabel(connection)}
        </span>
        <label>
          Name
          <input value={name} onChange={(event) => onNameChange(event.target.value)} disabled={connection === 'live'} />
        </label>
        <label>
          Artifact
          <input
            value={artifact}
            onChange={(event) => onArtifactChange(event.target.value.trim())}
            placeholder="art_…"
            spellCheck={false}
            disabled={busy}
          />
        </label>
        <button type="button" className="workbench-button workbench-button--primary" disabled={busy || !artifact || joinDisabled} onClick={onJoin}>
          Join
        </button>
        {actions}
        <button type="button" className="workbench-button" disabled={!artifact} onClick={() => window.open(window.location.href, '_blank')}>
          Open second tab
        </button>
        {presence ? <PresenceStack manager={presence} /> : null}
      </div>

      <p className="native-status workbench-status" role="status" aria-live="polite" aria-atomic="true" data-state={error ? 'error' : connection}>{status}</p>
      {error ? <p className="native-error workbench-callout workbench-callout--error" role="alert">{error}</p> : null}

      <section className="collab-boundaries workbench-boundary" aria-labelledby="collab-boundaries-title">
        <div>
          <h2 id="collab-boundaries-title">What this proves</h2>
          <p>{proof}</p>
        </div>
        <div className="collab-security workbench-callout workbench-callout--warning" role="note">
          <strong>Local development only</strong>
          <p>{COLLAB_COPY.security}</p>
        </div>
        <p className="collab-completion">{COLLAB_COPY.completion}</p>
      </section>

      <div className="collab-body collab-body--univer">
        <div className="collab-univer-wrap">
          <div className="collab-univer-heading">
            <div>
              <h2>{editorTitle}</h2>
              <p>{editorHint}</p>
            </div>
            <span>{editorMeta}</span>
          </div>
          {children}
        </div>
        <aside className="native-side collab-protocol-side workbench-inspector">
          <h3>Protocol state</h3>
          <dl className="collab-metrics">
            <div><dt>Peers</dt><dd>{metrics.peers}</dd></div>
            <div><dt>Applied sequence</dt><dd>{metrics.applied}</dd></div>
            <div><dt>Pending operations</dt><dd>{metrics.pending}</dd></div>
          </dl>
          {protocolExtra}
          <h4>Room artifact</h4>
          <p className="native-muted">{artifact || 'No artifact yet.'}</p>
          <h4>Try it</h4>
          {tryIt}
        </aside>
      </div>
    </>
  )
}

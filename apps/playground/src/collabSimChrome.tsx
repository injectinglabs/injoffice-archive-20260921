import { type ReactNode, type Ref } from 'react'
import { PresenceStack, type PresenceSource } from '../../../packages/collab/src/index.js'
import { DsAvatar, DsChip } from './design-system/primitives'
import './design-system/live-tools.css'
import { COLLAB_FORMATS, type CollabFormat } from './collabScope'

export type SimConnection = 'starting' | 'live' | 'error'
export type SimMetrics = { peers: number; applied: number; pending: number }

export type SimEditorProfile = {
  id: string
  name: string
  color: string
  label: string
}

export const SIM_EDITORS: SimEditorProfile[] = [
  { id: 'mira', name: 'Mira', color: '#0f6f8f', label: 'Product editor' },
  { id: 'noah', name: 'Noah', color: '#1a6b52', label: 'Review editor' },
]

export const SIM_FORMAT_LABEL: Record<CollabFormat, string> = {
  sheets: 'Sheets',
  slides: 'Slides',
  docs: 'Docs',
  pdf: 'PDF',
}

export const SIM_FORMAT_HINT: Record<CollabFormat, string> = {
  sheets: 'Two Univer workbooks',
  slides: 'Two DeckSpec canvases',
  docs: 'Two ProseMirror documents',
  pdf: 'Two PDF annotators',
}

export const SIM_ROOMS: Record<CollabFormat, string> = {
  sheets: 'quarterly-plan',
  slides: 'shared-deck',
  docs: 'shared-document',
  pdf: 'shared-pdf',
}

export function SimEditorFrame({
  profile,
  status,
  presence,
  metrics,
  canvasClassName,
  extraFooter,
  stageRef,
  children,
}: {
  profile: SimEditorProfile
  status: SimConnection
  presence: PresenceSource | null
  metrics: SimMetrics
  canvasClassName?: string
  extraFooter?: ReactNode
  stageRef?: Ref<HTMLDivElement>
  children: ReactNode
}) {
  return (
    <section className="collab-sim-editor ds-editor-card" aria-labelledby={`sim-editor-${profile.id}`}>
      <header className="collab-sim-editor__header">
        <div className="collab-sim-identity ds-row">
          <DsAvatar initials={profile.name.slice(0, 1)} tone={profile.id === 'noah' ? 2 : 3} />
          <div>
            <h3 id={`sim-editor-${profile.id}`}>{profile.name}</h3>
            <p>{profile.label}</p>
          </div>
        </div>
        <div className="collab-sim-editor__state">
          {presence ? <PresenceStack manager={presence} /> : null}
          <DsChip tone={status === 'live' ? 'green' : status === 'error' ? 'refuse' : 'plain'}>
            <span className={`collab-connection collab-connection--${status === 'live' ? 'live' : status === 'error' ? 'error' : 'joining'}`} role="status">
              <i aria-hidden="true" />
              {status === 'live' ? 'Connected' : status === 'error' ? 'Unavailable' : 'Connecting'}
            </span>
          </DsChip>
        </div>
      </header>
      <div ref={stageRef} className={`collab-sim-editor__stage${canvasClassName ? ` ${canvasClassName}` : ''}`}>
        {children}
      </div>
      <footer className="collab-sim-editor__footer">
        <span>{metrics.peers} collaborator{metrics.peers === 1 ? '' : 's'} visible</span>
        <span>Applied {metrics.applied}</span>
        <span>{metrics.pending ? `${metrics.pending} pending` : 'Synced'}</span>
        {extraFooter}
      </footer>
    </section>
  )
}

export function SimFormatTabs({
  format,
  onFormat,
  hint,
}: {
  format: CollabFormat
  onFormat: (next: CollabFormat) => void
  hint?: string
}) {
  return (
    <div className="view-switcher collab-format-bar ds-workstrip" role="group" aria-label="Collaboration format">
      <div className="collab-format-bar__label">
        <strong>Format</strong>
        <span>One room, four editors</span>
      </div>
      <div className="tool-segment ds-segment">
        {COLLAB_FORMATS.map((item) => (
          <button
            key={item}
            type="button"
            aria-pressed={format === item}
            onClick={() => onFormat(item)}
          >
            {SIM_FORMAT_LABEL[item]}
          </button>
        ))}
      </div>
      {hint ? <span className="ds-muted">{hint}</span> : null}
    </div>
  )
}

import type { CSSProperties, ComponentType } from 'react'
import type { PeerInfo, SheetSelection } from './types'

export interface SheetPresenceLabelProps {
  peer: PeerInfo
  selection: SheetSelection
}

type UniverPopupProps = {
  popup?: {
    extraProps?: Partial<SheetPresenceLabelProps> & {
      labelComponent?: SheetPresenceLabelComponent
    }
  }
}

/**
 * Default cell-anchored collaborator flag. Hosts can replace this component
 * through PresenceOptions.peerLabelComponent; class names and CSS custom
 * properties are also stable styling hooks.
 */
export function SheetPresenceLabel({ peer, selection }: Partial<SheetPresenceLabelProps>) {
  if (!peer || !selection) return null
  const editing = selection.mode === 'editing'
  const style = {
    '--ioc-peer-color': peer.color,
    alignItems: 'center',
    background: 'var(--ioc-peer-color)',
    borderRadius: '4px 4px 4px 0',
    boxShadow: '0 2px 6px rgba(22, 31, 43, .22)',
    color: '#fff',
    display: 'inline-flex',
    fontFamily: 'system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif',
    fontSize: '11px',
    fontWeight: 650,
    gap: '5px',
    lineHeight: 1,
    maxWidth: '160px',
    minHeight: '22px',
    overflow: 'hidden',
    padding: '0 7px',
    pointerEvents: 'none',
    whiteSpace: 'nowrap',
  } as CSSProperties

  return (
    <span
      className={`ioc-sheet-presence-label${editing ? ' ioc-sheet-presence-label--editing' : ''}`}
      style={style}
      role="status"
      aria-label={`${peer.name} is ${editing ? 'editing' : 'selecting this cell'}`}
    >
      <span className="ioc-sheet-presence-label__name">{peer.name}</span>
      {editing ? (
        <span className="ioc-sheet-presence-label__activity" aria-hidden="true">
          <i /> editing
        </span>
      ) : null}
      {editing && selection.draft !== undefined ? (
        <span
          className="ioc-sheet-presence-label__draft"
          style={{ maxWidth: '90px', overflow: 'hidden', textOverflow: 'ellipsis' }}
        >
          {selection.draft || '…'}
        </span>
      ) : null}
    </span>
  )
}

export type SheetPresenceLabelComponent = ComponentType<SheetPresenceLabelProps>

/** @internal Adapts Univer's `{ popup }` render prop to our public component API. */
export function SheetPresencePopup({ popup }: UniverPopupProps) {
  const details = popup?.extraProps
  const peer = details?.peer
  const selection = details?.selection
  if (!peer || !selection) return null
  const Label = details.labelComponent ?? SheetPresenceLabel
  return <Label peer={peer} selection={selection} />
}

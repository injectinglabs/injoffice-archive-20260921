import { useEffect, useReducer } from 'react'
import { initials, selectionA1 } from './selection'
import type { FileChange, PeerInfo } from './types'

/** Anything with peers — the sheet PresenceManager, the DocPresenceManager,
 *  or a test double. Structural on purpose so hosts never need casts. */
export interface PresenceSource {
  onChange(cb: () => void): () => void
  peers(): PeerInfo<unknown>[]
}

// PresenceStack — the avatar row for a file: who else has it open, coloured
// to match their selection on the grid. Binds to a PresenceManager via
// onChange. Structure-only (ioc-presence-* hooks); the host styles it.

const MAX_VISIBLE = 5

export function PresenceStack({ manager, max = MAX_VISIBLE }: { manager: PresenceSource; max?: number }) {
  const [, bump] = useReducer((n: number) => n + 1, 0)
  useEffect(() => manager.onChange(bump), [manager])
  const peers = manager.peers()
  if (peers.length === 0) return null
  const shown = peers.slice(0, max)
  const extra = peers.length - shown.length
  return (
    <div className="ioc-presence" role="group" aria-label={`${peers.length} other ${peers.length === 1 ? 'person has' : 'people have'} this file open`}>
      {shown.map((p) => (
        <PeerAvatar key={p.client_id} peer={p} />
      ))}
      {extra > 0 && (
        <span className="ioc-presence-avatar ioc-presence-avatar--more" title={peers.slice(max).map((p) => p.name).join(', ')}>
          +{extra}
        </span>
      )}
    </div>
  )
}

function PeerAvatar({ peer }: { peer: PeerInfo<unknown> }) {
  const selection = peer.selection
  const where = selection && typeof selection === 'object' && 'sheet' in selection && 'ranges' in selection
    ? selectionA1(selection as Parameters<typeof selectionA1>[0])
    : null
  return (
    <span
      className="ioc-presence-avatar"
      style={{ background: peer.color }}
      title={where ? `${peer.name} · ${where}` : peer.name}
      aria-label={where ? `${peer.name}, at ${where}` : peer.name}
    >
      {initials(peer.name)}
    </span>
  )
}

/** Who/what a FileChange came from, for the banner. */
export function describeFileChange(change: FileChange, nameOf?: (clientId: string | undefined) => string | null): string {
  const who = change.author === 'agent' ? 'The agent' : change.user_name || nameOf?.(change.origin) || 'Someone else'
  switch (change.action) {
    case 'delivered':
      return `${who} delivered a new version of this file.`
    case 'print-setup':
      return `${who} changed the print setup.`
    default:
      return `${who} saved a new version of this file.`
  }
}

// FileChangedBanner — shown by a host that has LOCAL unsaved edits when the
// file changed underneath it; a host with no local edits just reloads. The
// choice is the user's: reload (discarding their edits) or keep editing —
// their next Save overwrites, but every version is in History.
export function FileChangedBanner({
  message,
  onReload,
  onDismiss,
}: {
  message: string
  onReload: () => void
  onDismiss: () => void
}) {
  return (
    <div className="ioc-file-changed" role="status">
      <span className="ioc-file-changed-msg">{message} You have unsaved edits.</span>
      <button type="button" className="ioc-file-changed-btn ioc-file-changed-btn--primary" onClick={onReload}>
        Reload (discard mine)
      </button>
      <button type="button" className="ioc-file-changed-btn" onClick={onDismiss}>
        Keep editing
      </button>
    </div>
  )
}

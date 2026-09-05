import type { SheetSelection } from './types'

// Pure helpers for the selection wire shape: validation of what a remote
// peer sent (never trust it — it's an opaque blob to the gateway), equality
// (skip redundant publishes), and A1 rendering for labels/tooltips.

export const MAX_RANGES = 64
export const MAX_DRAFT_LENGTH = 512
const MAX_INDEX = 1_048_576 // Excel's row limit; columns are far smaller

function isIndex(n: unknown): n is number {
  return typeof n === 'number' && Number.isInteger(n) && n >= 0 && n <= MAX_INDEX
}

/** Parse an untrusted selection blob; null when malformed. Normalizes each range so start ≤ end. */
export function parseSelection(raw: unknown): SheetSelection | null {
  if (!raw || typeof raw !== 'object') return null
  const o = raw as Record<string, unknown>
  if (typeof o.sheet !== 'string' || !o.sheet || !Array.isArray(o.ranges)) return null
  if (o.ranges.length > MAX_RANGES) return null
  const ranges: number[][] = []
  for (const r of o.ranges) {
    if (!Array.isArray(r) || r.length !== 4 || !r.every(isIndex)) return null
    const [r0, c0, r1, c1] = r as number[]
    ranges.push([Math.min(r0, r1), Math.min(c0, c1), Math.max(r0, r1), Math.max(c0, c1)])
  }
  const out: SheetSelection = { sheet: o.sheet, ranges }
  if (Array.isArray(o.active) && o.active.length === 2 && o.active.every(isIndex)) {
    out.active = [o.active[0] as number, o.active[1] as number]
  }
  if (o.mode === 'selecting' || o.mode === 'editing') out.mode = o.mode
  if (out.mode === 'editing' && typeof o.draft === 'string' && o.draft.length <= MAX_DRAFT_LENGTH) out.draft = o.draft
  return out
}

export function sameSelection(a: SheetSelection | null | undefined, b: SheetSelection | null | undefined): boolean {
  if (!a || !b) return a === b || (!a && !b)
  if (a.sheet !== b.sheet || a.ranges.length !== b.ranges.length || a.mode !== b.mode || a.draft !== b.draft) return false
  for (let i = 0; i < a.ranges.length; i++) {
    const x = a.ranges[i]
    const y = b.ranges[i]
    for (let k = 0; k < 4; k++) if (x[k] !== y[k]) return false
  }
  const ax = a.active
  const bx = b.active
  if (!ax !== !bx) return false
  return !ax || !bx || (ax[0] === bx[0] && ax[1] === bx[1])
}

export function colName(c: number): string {
  let s = ''
  let n = c + 1
  while (n > 0) {
    const rem = (n - 1) % 26
    s = String.fromCharCode(65 + rem) + s
    n = Math.floor((n - 1) / 26)
  }
  return s
}

export function rangeA1(r: number[]): string {
  const [r0, c0, r1, c1] = r
  const a = `${colName(c0)}${r0 + 1}`
  if (r0 === r1 && c0 === c1) return a
  return `${a}:${colName(c1)}${r1 + 1}`
}

/** "B2:D5" or "A1, C3:C9" — what a tooltip shows next to a peer's name. */
export function selectionA1(sel: SheetSelection | null | undefined): string {
  if (!sel || sel.ranges.length === 0) return ''
  return sel.ranges.slice(0, 3).map(rangeA1).join(', ') + (sel.ranges.length > 3 ? ', …' : '')
}

/** Two-letter initials for an avatar, from a display name. */
export function initials(name: string): string {
  const parts = name.trim().split(/\s+/).filter(Boolean)
  if (parts.length === 0) return '?'
  if (parts.length === 1) return parts[0].slice(0, 2).toUpperCase()
  return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase()
}

/** #rrggbb + alpha → rgba(); passes through anything that isn't 6-digit hex. */
export function withAlpha(hex: string, alpha: number): string {
  const m = /^#([0-9a-f]{6})$/i.exec(hex.trim())
  if (!m) return hex
  const n = parseInt(m[1], 16)
  return `rgba(${(n >> 16) & 255}, ${(n >> 8) & 255}, ${n & 255}, ${alpha})`
}

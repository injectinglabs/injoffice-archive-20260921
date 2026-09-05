// Bidirectional registry between ShapeManager and the ShapeFloat components
// Univer mounts (same reason as @injoffice/charts' registry: float-dom props
// carry only serializable data, so live wiring meets here). Unlike charts,
// shapes flow BOTH ways: the manager pushes spec changes down, and the
// component pushes inline text edits back up.

import type { ShapeSpec } from './types'

type Listener = () => void

const specs = new Map<string, ShapeSpec>()
const listeners = new Map<string, Set<Listener>>()
let textEditHandler: ((id: string, text: string) => void) | null = null

export function publishShapeSpec(spec: ShapeSpec): void {
  specs.set(spec.id, spec)
  listeners.get(spec.id)?.forEach((l) => l())
}

export function getShapeSpec(id: string): ShapeSpec | undefined {
  return specs.get(id)
}

export function dropShape(id: string): void {
  specs.delete(id)
  listeners.delete(id)
}

export function subscribeShape(id: string, listener: Listener): () => void {
  let set = listeners.get(id)
  if (!set) {
    set = new Set()
    listeners.set(id, set)
  }
  set.add(listener)
  return () => {
    set.delete(listener)
  }
}

/** The manager registers itself as the single upstream text-edit sink. */
export function setShapeTextEditHandler(fn: ((id: string, text: string) => void) | null): void {
  textEditHandler = fn
}

export function pushShapeTextEdit(id: string, text: string): void {
  textEditHandler?.(id, text)
}

import { createContext, useContext } from 'react'
import { parseCollabQuery, type CollabFormat } from './collabScope'

export type CollabCompositionProps = { fixedFormat?: CollabFormat; initialHash?: string }
export const CollabCompositionContext = createContext<CollabCompositionProps>({})
export const useCollabComposition = () => useContext(CollabCompositionContext)

/** Embedded rooms may consume their explicit deep link, never another workspace's URL. */
export function initialCollabArtifact(scope: CollabCompositionProps, currentHref: string): string {
  if (scope.fixedFormat !== undefined && !scope.initialHash) return ''
  const href = scope.initialHash ?? currentHref
  const parsed = parseCollabQuery(href)
  // Canonical workspace URLs name their format in the path, not ?format=.
  // Keep the same precedence as workspace routing: a direct tool path wins.
  const path = new URL(href, 'http://collab.invalid').hash.replace(/^#\/?/, '').split(/[/?#]/)[0]?.toLowerCase()
  const format = path === 'sheets' || path === 'docs' || path === 'slides' || path === 'pdf' ? path : parsed.format
  if (scope.fixedFormat !== undefined && format !== scope.fixedFormat) return ''
  return parsed.artifact
}

/** Share an embedded room without changing the currently active tool or global query. */
export function collabRoomHref(currentHref: string, artifact: string, fixedFormat?: CollabFormat): string {
  if (fixedFormat === undefined) return currentHref
  const url = new URL(currentHref)
  url.searchParams.delete('artifact')
  url.searchParams.delete('format')
  url.hash = `/collab?${new URLSearchParams({ format: fixedFormat, artifact })}`
  return url.href
}

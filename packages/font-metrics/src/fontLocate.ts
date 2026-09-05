import {
  getFontIndex,
  indexedFaces,
  norm,
  readFaceBytes,
  refCoversCodepoints,
  styleScore,
  styleTokens,
} from './sfnt.js'

export { isTruetype } from './sfnt.js'

/** Report whether the system inventory contains an exact normalized family name. */
export function isFamilyInstalled(family: string): boolean {
  const key = norm(family)
  return key.length > 0 && (getFontIndex().byFamily.get(key)?.length ?? 0) > 0
}

/**
 * Locate a system face by exact PostScript name, then by family plus the style
 * words carried by the requested PostScript name. Collection faces are copied
 * into a standalone sfnt before being returned.
 */
export function findSystemFont(psName: string, family: string): Buffer | null {
  const index = getFontIndex()
  const exact = index.byPs.get(norm(psName))
  if (exact) return readFaceBytes(exact)

  const candidates = index.byFamily.get(norm(family))
  if (!candidates?.length) return null
  const wanted = styleTokens(psName)
  const ordered = candidates
    .map((face, order) => ({ face, order, score: styleScore(face, wanted) }))
    .sort((left, right) => right.score - left.score || left.order - right.order)
  for (const candidate of ordered) {
    const bytes = readFaceBytes(candidate.face)
    if (bytes) return bytes
  }
  return null
}

function requestedCodepoints(text: string): number[] {
  const ignored = new Set([0x0009, 0x000a, 0x000d, 0x200c, 0x200d, 0xfeff])
  const points = new Set<number>()
  for (const scalar of text) {
    const cp = scalar.codePointAt(0)!
    if (ignored.has(cp) || (cp >= 0xfe00 && cp <= 0xfe0f) || (cp >= 0xe0100 && cp <= 0xe01ef)) continue
    points.add(cp)
  }
  return [...points]
}

/** Find the first deterministic system face whose Unicode cmap maps all text scalars. */
export function findFontCovering(text: string): Buffer | null {
  const codepoints = requestedCodepoints(text)
  if (codepoints.length === 0) return null
  for (const face of indexedFaces(getFontIndex())) {
    if (!refCoversCodepoints(face, codepoints)) continue
    const bytes = readFaceBytes(face)
    if (bytes) return bytes
  }
  return null
}

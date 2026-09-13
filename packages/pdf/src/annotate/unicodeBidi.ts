import { getBidiCharTypeName, getEmbeddingLevels, getReorderedIndices } from './unicodeBidiVendor/index.js'

export type UnicodeBaseDirection = 'auto' | 'ltr' | 'rtl'
export interface BidiScalar {
  value: string
  start: number
  end: number
  type: string
  level: number
}
export interface UnicodeBidiResult {
  scalars: readonly BidiScalar[]
  /** Scalar indices in display order; cluster shaping still applies rule L3. */
  order: readonly number[]
  paragraphLevel: number
}

/** Unicode17 levels over scalars, with exact original UTF-16 source spans. */
export function resolveUnicodeBidi(value: string, direction: UnicodeBaseDirection = 'auto'): UnicodeBidiResult {
  if (!['auto', 'ltr', 'rtl'].includes(direction)) throw new Error('invalid embedded appearance base direction')
  if (value.length > 4096) throw new Error('embedded appearance text exceeds 4096 UTF-16 units')
  const characters = [...value]
  let offset = 0
  const scalars = characters.map(character => {
    const code = character.codePointAt(0)!
    if (code >= 0xd800 && code <= 0xdfff) throw new Error('embedded appearance text contains an unpaired surrogate')
    const start = offset
    offset += character.length
    return { value: character, start, end: offset, type: getBidiCharTypeName(character), level: 0 }
  })
  const resolved = getEmbeddingLevels(characters, direction)
  if (resolved.paragraphs.length > 1 || scalars.some(scalar => scalar.type === 'B' || scalar.type === 'S')) throw new Error('embedded appearance requires a single line')
  scalars.forEach((scalar, index) => { scalar.level = resolved.levels[index]! })
  return { scalars, order: getReorderedIndices(characters, resolved), paragraphLevel: resolved.paragraphs[0]?.level ?? (direction === 'rtl' ? 1 : 0) }
}

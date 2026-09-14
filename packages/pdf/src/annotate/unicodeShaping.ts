import { unicodeScriptItems } from './unicodeScript.js'
import { resolveUnicodeBidi, type UnicodeBaseDirection } from './unicodeBidi.js'
/** Source clusters and glyph positions are distinct: neither is a Unicode scalar. */
export interface UnicodeGlyph {
  id: number
  cluster: number
  end: number
  x: number
  y: number
  advance: number
}
export interface UnicodeRun {
  value: string
  unitsPerEm: number
  width: number
  glyphs: readonly UnicodeGlyph[]
  requiresActualText?: boolean
}
export interface UnicodeShapingOptions { direction?: UnicodeBaseDirection; language?: string }

/** Browser-safe HarfBuzz adapter. No system fonts, normalization or Node loaders. */
export async function prepareUnicodeShaper(bytes: Uint8Array, unitsPerEm: number, numGlyphs: number, options: UnicodeShapingOptions = {}) {
  const language = options.language ?? 'und'
  const direction = options.direction
  if (typeof language !== 'string' || language.length > 63 || !/^[A-Za-z0-9]{1,8}(?:-[A-Za-z0-9]{1,8})*$/.test(language)) throw new Error('invalid embedded appearance language tag')
  if (direction !== undefined && !['auto', 'ltr', 'rtl'].includes(direction)) throw new Error('invalid embedded appearance base direction')
  const hb = await import('harfbuzzjs')
  const blob = new hb.Blob(bytes.slice().buffer)
  const face = new hb.Face(blob, 0)
  const font = new hb.Font(face)
  font.setScale(unitsPerEm, unitsPerEm)
  return (value: string): UnicodeRun => {
    if (value.length > 4096) throw new Error('embedded appearance text exceeds 4096 UTF-16 units')
    const bidi = resolveUnicodeBidi(value, direction)
    // UAX9 resolves levels first. Each script/direction item is then shaped
    // with the original full source as context and global UTF-16 clusters.
    // C0/C1 controls are not single-line appearance content. Bidi format,
    // joiner and variation controls retain exact source semantics through HB.
    if (/[\p{Control}\p{Surrogate}]/u.test(value)) throw new Error('embedded appearance requires single-line Unicode text')
    const items = unicodeScriptItems(value, bidi)
    const visualRanks = new Map(bidi.order.map((scalarIndex, visualIndex) => [scalarIndex, visualIndex]))
    const ranked = items.map(item => ({ item, rank: Math.min(...item.scalarIndices.map(index => visualRanks.get(index)!)) }))
    ranked.sort((a, b) => a.rank - b.rank)
    const shaped = ranked.flatMap(({ item }) => {
      const buffer = new hb.Buffer()
      // Whole source plus item bounds retains global UTF-16 source clusters.
      buffer.addText(value, item.start, item.end - item.start)
      buffer.setDirection(item.level % 2 ? hb.Direction.RTL : hb.Direction.LTR)
      buffer.setScript(item.script)
      buffer.setLanguage(language)
      buffer.setFlags(hb.BufferFlag.REMOVE_DEFAULT_IGNORABLES)
      buffer.setClusterLevel(0)
      hb.shape(font, buffer)
      const infos = buffer.getGlyphInfos()
      const positions = buffer.getGlyphPositions()
      if (infos.length !== positions.length) throw new Error('invalid shaping position count')
      return infos.map((info, i) => ({ ...info, ...positions[i]! }))
    })
    if (shaped.length > 16384) throw new Error('embedded appearance glyph limit exceeded')
    const boundaries = [...new Set(shaped.map(g => g.cluster))].sort((a, b) => a - b)
    // Removed formatting controls remain in the exact source span of the next
    // visible cluster (leading controls) or prior cluster (interior/trailing).
    const firstCluster = boundaries[0]
    const ends = new Map(boundaries.map((start, i) => [start, boundaries[i + 1] ?? value.length]))
    let x = 0
    let y = 0
    const glyphs = shaped.map(g => {
      if (!Number.isInteger(g.codepoint) || g.codepoint <= 0 || g.codepoint >= numGlyphs) throw new Error('embedded appearance font is missing a requested glyph')
      if (!Number.isInteger(g.cluster) || g.cluster < 0 || g.cluster >= value.length
        || (g.cluster > 0 && /[\uDC00-\uDFFF]/.test(value[g.cluster]!))
        || ![g.xAdvance, g.yAdvance, g.xOffset, g.yOffset].every(Number.isFinite)
        || [g.xAdvance, g.yAdvance, g.xOffset, g.yOffset].some(n => Math.abs(n) > unitsPerEm * 100)
        || g.yAdvance !== 0) throw new Error('invalid horizontal shaping positions')
      const glyph = { id: g.codepoint, cluster: g.cluster === firstCluster ? 0 : g.cluster, end: ends.get(g.cluster)!, x: x + g.xOffset, y: y + g.yOffset, advance: g.xAdvance }
      x += g.xAdvance
      y += g.yAdvance
      return glyph
    })
    if (!Number.isFinite(x) || x < 0) throw new Error('invalid shaped run width')
    // Within a cluster, keep the advancing/base glyph native whenever possible;
    // remaining same-fill glyphs retain their independent shaped positions.
    const groups = new Map<number, UnicodeGlyph[]>()
    for (const glyph of glyphs) { const group = groups.get(glyph.cluster) ?? []; group.push(glyph); groups.set(glyph.cluster, group) }
    const orderedGlyphs = [...groups.values()].flatMap(group => {
      const primary = group.findIndex(glyph => glyph.advance > 0)
      return primary > 0 ? [group[primary]!, ...group.filter((_, i) => i !== primary)] : group
    })
    return { value, unitsPerEm, width: x, glyphs: orderedGlyphs, requiresActualText: bidi.scalars.some(scalar => scalar.level % 2 !== 0 || /^\p{Bidi_Control}$/u.test(scalar.value) || scalar.value === '\u200C' || scalar.value === '\u200D') }
  }
}

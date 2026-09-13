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
}

/** Browser-safe HarfBuzz adapter. No system fonts, normalization or Node loaders. */
export async function prepareUnicodeShaper(bytes: Uint8Array, unitsPerEm: number, numGlyphs: number) {
  const hb = await import('harfbuzzjs')
  const blob = new hb.Blob(bytes.slice().buffer)
  const face = new hb.Face(blob, 0)
  const font = new hb.Font(face)
  font.setScale(unitsPerEm, unitsPerEm)
  return (value: string): UnicodeRun => {
    if (value.length > 4096) throw new Error('embedded appearance text exceeds 4096 UTF-16 units')
    // Resolve supported LTR scripts before shaping: guessing one script for a
    // mixed paragraph would select the wrong later-script GSUB/GPOS plan.
    const scripts: readonly [RegExp, string][] = [
      [/^\p{Script=Latin}$/u, 'Latn'], [/^\p{Script=Greek}$/u, 'Grek'], [/^\p{Script=Cyrillic}$/u, 'Cyrl'],
      [/^\p{Script=Han}$/u, 'Hani'], [/^\p{Script=Hiragana}$/u, 'Hira'], [/^\p{Script=Katakana}$/u, 'Kana'], [/^\p{Script=Hangul}$/u, 'Hang'],
    ]
    const characters = [...value].map(character => {
      const script = scripts.find(([pattern]) => pattern.test(character))?.[1]
      const neutral = (/^\p{Script=Common}$/u.test(character) && /^[0-9\p{Punctuation}\p{Symbol} ]$/u.test(character))
        || (/^\p{Script=Inherited}$/u.test(character) && /^\p{Mark}$/u.test(character))
      if ((!script && !neutral) || /[\p{Control}\p{Format}\p{Surrogate}]/u.test(character)) throw new Error('embedded appearance requires supported LTR text; bidi itemization is not yet available')
      return { character, script }
    })
    let activeScript = characters.find(c => c.script)?.script ?? 'Zyyy'
    const items: { start: number; end: number; script: string }[] = []
    let offset = 0
    for (const character of characters) {
      if (offset > 0 && /^\p{Mark}$/u.test(character.character) && character.script && character.script !== activeScript) throw new Error('embedded appearance cross-script mark requires grapheme itemization')
      activeScript = character.script ?? activeScript
      const previous = items.at(-1)
      if (previous?.script === activeScript) previous.end += character.character.length
      else items.push({ start: offset, end: offset + character.character.length, script: activeScript })
      offset += character.character.length
    }
    const shaped = items.flatMap(item => {
      const buffer = new hb.Buffer()
      // Whole source plus item bounds retains global UTF-16 source clusters.
      buffer.addText(value, item.start, item.end - item.start)
      buffer.setDirection(hb.Direction.LTR)
      buffer.setScript(item.script)
      buffer.setLanguage('und')
      buffer.setClusterLevel(0)
      hb.shape(font, buffer)
      const infos = buffer.getGlyphInfos()
      const positions = buffer.getGlyphPositions()
      if (infos.length !== positions.length) throw new Error('invalid shaping position count')
      return infos.map((info, i) => ({ ...info, ...positions[i]! }))
    })
    if (shaped.length > 16384) throw new Error('embedded appearance glyph limit exceeded')
    const boundaries = [...new Set(shaped.map(g => g.cluster))].sort((a, b) => a - b)
    if (value && boundaries[0] !== 0) throw new Error('invalid initial shaping cluster')
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
      const glyph = { id: g.codepoint, cluster: g.cluster, end: ends.get(g.cluster)!, x: x + g.xOffset, y: y + g.yOffset, advance: g.xAdvance }
      x += g.xAdvance
      y += g.yAdvance
      return glyph
    })
    if (!Number.isFinite(x) || x < 0) throw new Error('invalid shaped run width')
    return { value, unitsPerEm, width: x, glyphs }
  }
}

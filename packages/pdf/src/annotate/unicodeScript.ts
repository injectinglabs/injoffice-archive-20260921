import { SCRIPT, SCRIPT_EXTENSIONS, BRACKET_PAIRS } from './unicodeScriptData.generated.js'
import { unicodeGraphemes, unicodeRange } from './unicodeGrapheme.js'
import type { UnicodeBidiResult } from './unicodeBidi.js'

const implicit = (script: string) => ['Zyyy', 'Zinh', 'Zzzz'].includes(script)
export function unicodeScript(character: string): { primary: string; extensions: readonly string[] } {
  const code = character.codePointAt(0)!
  const primary = unicodeRange(SCRIPT, code)?.[2] ?? 'Zzzz'
  return { primary, extensions: unicodeRange(SCRIPT_EXTENSIONS, code)?.[2] ?? [primary] }
}
export interface UnicodeScriptItem { start: number; end: number; script: string; level: number; scalarIndices: number[] }

/** Graphemes select script even for scholarly cross-script mark usage.
 * Bidi level transitions still split shaping direction inside an EGC. */
export function unicodeScriptItems(value: string, bidi: UnicodeBidiResult): UnicodeScriptItem[] {
  const clusters = unicodeGraphemes(value).map(cluster => {
    const characters = bidi.scalars.slice(cluster.scalarStart, cluster.scalarEnd)
    let candidates: string[] | undefined, base: string | undefined
    for (const character of characters) {
      const properties = unicodeScript(character.value)
      if (!base && !implicit(properties.primary)) base = properties.primary
      const explicit = properties.extensions.filter(script => !implicit(script))
      if (!explicit.length) continue
      candidates = candidates === undefined ? explicit : candidates.filter(script => explicit.includes(script))
    }
    const level = characters.find(character => !['BN', 'LRE', 'RLE', 'LRO', 'RLO', 'PDF', 'LRI', 'RLI', 'FSI', 'PDI'].includes(character.type))?.level ?? characters[0]!.level
    return { ...cluster, level, base, candidates }
  })
  const following: (string | undefined)[] = []
  let nextBase: string | undefined
  for (let index = clusters.length - 1; index >= 0; index--) {
    following[index] = nextBase
    nextBase = clusters[index]!.base ?? nextBase
  }
  const items: UnicodeScriptItem[] = []
  const brackets: { closing: string; script: string; level: number }[] = []
  let active = 'Zyyy'
  for (const [index, cluster] of clusters.entries()) {
    const next = following[index]
    const candidates = cluster.candidates
    let script = candidates?.length
      ? (candidates.includes(active) ? active : next && candidates.includes(next) ? next : cluster.base && candidates.includes(cluster.base) ? cluster.base : candidates[0]!)
      : cluster.base ?? (!implicit(active) ? active : next ?? 'Zyyy')
    const character = value.slice(cluster.start, cluster.end)
    const paired = BRACKET_PAIRS[character]
    if (paired) brackets.push({ closing: paired, script, level: cluster.level })
    else {
      let opening = brackets.length - 1
      while (opening >= 0 && (brackets[opening]!.closing !== character || brackets[opening]!.level !== cluster.level)) opening--
      if (opening >= 0) { script = brackets[opening]!.script; brackets.length = opening }
    }
    active = script
    // A Prepend scalar may have a different UAX9 level from its substantive
    // base (e.g. U+0600 before Arabic letters). Do not replace the base's RTL
    // level with the prefix's LTR level merely to retain one grapheme item.
    for (let index = cluster.scalarStart; index < cluster.scalarEnd; index++) {
      const scalar = bidi.scalars[index]!, previous = items.at(-1)
      if (previous?.script === script && previous.level === scalar.level) { previous.end = scalar.end; previous.scalarIndices.push(index) }
      else items.push({ start: scalar.start, end: scalar.end, script, level: scalar.level, scalarIndices: [index] })
    }
  }
  return items
}

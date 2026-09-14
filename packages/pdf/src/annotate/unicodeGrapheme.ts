import { CONJUNCT, GRAPHEME, PICTOGRAPHIC } from './unicodeScriptData.generated.js'

/** Binary search over nonoverlapping Unicode scalar ranges. */
export function unicodeRange<T extends readonly [number, number, ...unknown[]]>(rows: readonly T[], code: number): T | undefined {
  let low = 0, high = rows.length - 1
  while (low <= high) {
    const middle = (low + high) >>> 1, row = rows[middle]!
    if (code < row[0]) high = middle - 1
    else if (code > row[1]) low = middle + 1
    else return row
  }
  return undefined
}
export interface UnicodeGrapheme { start: number; end: number; scalarStart: number; scalarEnd: number }

/** Unicode17 extended grapheme clusters, UAX29 GB3–GB999, over original scalars. */
export function unicodeGraphemes(value: string): UnicodeGrapheme[] {
  const scalars = [...value], starts: number[] = []
  let offset = 0
  const properties = scalars.map(character => {
    starts.push(offset); offset += character.length
    const code = character.codePointAt(0)!
    return { type: unicodeRange(GRAPHEME, code)?.[2] ?? 'Other', conjunct: unicodeRange(CONJUNCT, code)?.[2], pictographic: !!unicodeRange(PICTOGRAPHIC, code) }
  })
  starts.push(offset)
  const control = (type: string) => type === 'Control' || type === 'CR' || type === 'LF'
  const breaks = (index: number): boolean => {
    const a = properties[index - 1]!, b = properties[index]!
    if (a.type === 'CR' && b.type === 'LF') return false
    if (control(a.type) || control(b.type)) return true
    if (a.type === 'L' && ['L', 'V', 'LV', 'LVT'].includes(b.type)) return false
    if (['LV', 'V'].includes(a.type) && ['V', 'T'].includes(b.type)) return false
    if (['LVT', 'T'].includes(a.type) && b.type === 'T') return false
    if (b.type === 'Extend' || b.type === 'ZWJ' || b.type === 'SpacingMark' || a.type === 'Prepend') return false
    if (b.conjunct === 'Consonant') {
      let previous = index - 1, linker = false
      while (previous >= 0 && ['Extend', 'Linker'].includes(properties[previous]!.conjunct ?? '')) {
        linker ||= properties[previous]!.conjunct === 'Linker'; previous--
      }
      if (linker && previous >= 0 && properties[previous]!.conjunct === 'Consonant') return false
    }
    if (b.pictographic && a.type === 'ZWJ') {
      let previous = index - 2
      while (previous >= 0 && properties[previous]!.type === 'Extend') previous--
      if (previous >= 0 && properties[previous]!.pictographic) return false
    }
    if (a.type === 'Regional_Indicator' && b.type === 'Regional_Indicator') {
      let count = 0
      for (let previous = index - 1; previous >= 0 && properties[previous]!.type === 'Regional_Indicator'; previous--) count++
      if (count % 2 === 1) return false
    }
    return true
  }
  const result: UnicodeGrapheme[] = []
  let start = 0
  for (let index = 1; index <= scalars.length; index++) {
    if (index === scalars.length || breaks(index)) {
      result.push({ start: starts[start]!, end: starts[index]!, scalarStart: start, scalarEnd: index }); start = index
    }
  }
  return result
}

import { readFileSync } from 'node:fs'
import { gunzipSync } from 'node:zlib'
import { execFileSync } from 'node:child_process'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { resolveUnicodeBidi } from './unicodeBidi.js'
import { getBidiCharTypeName, getEmbeddingLevels, getReorderedIndices } from './unicodeBidiVendor/index.js'

const data = (name: string) => gunzipSync(readFileSync(resolve(import.meta.dirname, `../../testdata/unicode17/${name}.txt.gz`))).toString('utf8')
function check(characters: string[], direction: 'ltr' | 'rtl' | undefined, levels: string[], order: number[], paragraph?: number) {
  const result = getEmbeddingLevels(characters, direction)
  if (paragraph !== undefined && result.paragraphs[0]?.level !== paragraph) return false
  if (!levels.every((value, index) => value === 'x' || Number(value) === result.levels[index])) return false
  return JSON.stringify(getReorderedIndices(characters, result).filter(index => levels[index] !== 'x')) === JSON.stringify(order)
}

describe('Unicode17 bidi data and scalar boundary', () => {
  it('resets trailing whitespace using local indices for a nonzero line start', () => {
    const scalars = [...'xאב ג']
    const resolved = getEmbeddingLevels(scalars, 'ltr')
    expect(getReorderedIndices(scalars, resolved, 1, 3)).toEqual([0, 2, 1, 3, 4])
  })
  it('reproduces generated tables from hash-pinned, licensed official inputs offline', () => {
    expect(execFileSync(process.execPath, [resolve(import.meta.dirname, '../../../../scripts/generate-pdf-unicode-bidi.mjs'), '--check'], { encoding: 'utf8' })).toContain('tables match')
  })
  it('passes all official BidiCharacterTest cases through rule L2', () => {
    let count = 0
    for (const line of data('BidiCharacterTest').split('\n')) {
      const text = line.split('#')[0]!.trim()
      if (!text) continue
      const [codes, direction, paragraph, levels, ordering] = text.split(';') as [string, string, string, string, string]
      const characters = codes.trim().split(/\s+/).map(code => String.fromCodePoint(parseInt(code, 16)))
      const good = check(characters, direction === '0' ? 'ltr' : direction === '1' ? 'rtl' : undefined, levels.trim().split(/\s+/), ordering.trim() ? ordering.trim().split(/\s+/).map(Number) : [], Number(paragraph))
      if (!good) throw new Error(`Unicode17 BidiCharacterTest failure: ${text}`)
      count++
    }
    expect(count).toBe(91707)
  })
  it('passes all official BidiTest class/paragraph permutations through rule L2', () => {
    const representatives: Record<string, string> = { L: 'A', R: 'א', AL: 'ا', EN: '1', ES: '+', ET: '$', AN: '١', CS: ',', NSM: '\u0300', BN: '\u00AD', B: '\u2029', S: '\t', WS: ' ', ON: '!', LRE: '\u202A', RLE: '\u202B', LRO: '\u202D', RLO: '\u202E', PDF: '\u202C', LRI: '\u2066', RLI: '\u2067', FSI: '\u2068', PDI: '\u2069' }
    for (const [type, scalar] of Object.entries(representatives)) expect(getBidiCharTypeName(scalar)).toBe(type)
    let levels: string[] = [], order: number[] = [], count = 0
    for (const line of data('BidiTest').split('\n')) {
      const text = line.split('#')[0]!.trim()
      if (!text) continue
      if (text.startsWith('@Levels:')) { levels = text.slice(8).trim().split(/\s+/); continue }
      if (text.startsWith('@Reorder:')) { order = text.slice(9).trim() ? text.slice(9).trim().split(/\s+/).map(Number) : []; continue }
      if (text.startsWith('@')) continue
      const [types, bitset] = text.split(';') as [string, string]
      const characters = types.trim().split(/\s+/).map(type => representatives[type]!)
      for (const [bit, direction] of [[1, undefined], [2, 'ltr'], [4, 'rtl']] as const) if (parseInt(bitset.trim(), 16) & bit) {
        if (!check(characters, direction, levels, order)) throw new Error(`Unicode17 BidiTest failure: ${text} (${direction ?? 'auto'})`)
        count++
      }
    }
    expect(count).toBe(770241)
  }, 20000)
  it('retains supplementary scalars and current Arabic/Garay/Adlam classes', () => {
    const value = 'A\u{1E900}\u{10D40}\u{10EC2}\u0870'
    const result = resolveUnicodeBidi(value)
    expect(result.scalars.map(s => [s.start, s.end, s.type])).toEqual([[0, 1, 'L'], [1, 3, 'R'], [3, 5, 'AN'], [5, 7, 'AL'], [7, 8, 'AL']])
    expect(resolveUnicodeBidi('\u{1E900}').scalars[0]!.level).toBe(1)
    expect(resolveUnicodeBidi('\u{10D40}').scalars[0]!.level).toBe(2)
    expect(resolveUnicodeBidi('\u{10EC2}').scalars[0]!.level).toBe(1)
  })
  it('resolves isolates and paired brackets while retaining original source spans', () => {
    const value = 'abc \u2067אבג (123)\u2069 xyz'
    const result = resolveUnicodeBidi(value)
    expect(result.scalars.map(s => value.slice(s.start, s.end)).join('')).toBe(value)
    expect(result.scalars.find(s => s.value === 'א')!.level).toBe(1)
    expect(result.scalars.find(s => s.value === '1')!.level).toBe(2)
    expect(new Set(result.order).size).toBe(result.scalars.length)
    expect(resolveUnicodeBidi('123', 'rtl').paragraphLevel).toBe(1)
  })
  it.each(['\ud800', 'x\nY', 'x\tY'])('refuses malformed/single-line-incompatible source %j', value => {
    expect(() => resolveUnicodeBidi(value)).toThrow()
  })
})

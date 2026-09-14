import { readFileSync } from 'node:fs'
import { gunzipSync } from 'node:zlib'
import { resolve } from 'node:path'
import { execFileSync } from 'node:child_process'
import { expect, it } from 'vitest'
import { unicodeGraphemes } from './unicodeGrapheme.js'

it('passes every official Unicode17 extended GraphemeBreakTest case', () => {
  const data = gunzipSync(readFileSync(resolve(import.meta.dirname, '../../testdata/unicode17/GraphemeBreakTest.txt.gz'))).toString('utf8')
  let count = 0
  for (const line of data.split('\n')) {
    const text = line.split('#')[0]!.trim()
    if (!text) continue
    let value = '', expected: number[] = []
    for (const token of text.split(/\s+/)) {
      if (token === '÷') expected.push(value.length)
      else if (token !== '×') value += String.fromCodePoint(parseInt(token, 16))
    }
    const actual = [0, ...unicodeGraphemes(value).map(cluster => cluster.end)]
    if (JSON.stringify(actual) !== JSON.stringify(expected)) throw new Error(`Unicode17 grapheme failure: ${text}`)
    count++
  }
  expect(count).toBe(766)
})
it('regenerates script and grapheme tables from pinned official inputs offline', () => {
  expect(execFileSync(process.execPath, [resolve(import.meta.dirname, '../../../../scripts/generate-pdf-unicode-script.mjs'), '--check'], { encoding: 'utf8' })).toContain('tables match')
})
it('retains supplementary, Indic conjunct, and mixed-script mark source spans', () => {
  for (const value of ['👩‍👩‍👧', 'क्ष', 'a\u0483', 'e\u0301']) expect(unicodeGraphemes(value)).toEqual([{ start: 0, end: value.length, scalarStart: 0, scalarEnd: [...value].length }])
  expect(unicodeGraphemes('')).toEqual([])
})

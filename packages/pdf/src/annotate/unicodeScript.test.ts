import { expect, it } from 'vitest'
import { resolveUnicodeBidi } from './unicodeBidi.js'
import { unicodeScript, unicodeScriptItems } from './unicodeScript.js'
const items = (value: string) => unicodeScriptItems(value, resolveUnicodeBidi(value))
it('uses current script extensions for shared Indic punctuation and Japanese marks', () => {
  expect(unicodeScript('।').extensions).toContain('Deva')
  expect(items('क।क').map(item => item.script)).toEqual(['Deva'])
  expect(items('あー').map(item => item.script)).toEqual(['Hira'])
  expect(items('アー').map(item => item.script)).toEqual(['Kana'])
})
it('never cuts a grapheme around a cross-script mark or Indic conjunct', () => {
  for (const [value, script] of [['a\u0483', 'Latn'], ['x\u05B0', 'Latn'], ['क्षि', 'Deva']]) {
    expect(items(value!)).toMatchObject([{ start: 0, end: value!.length, script }])
  }
})
it('keeps source spans and resolved direction for mixed scripts', () => {
  const value = 'café Ω Ж नमस्ते مرحبا'
  const result = items(value)
  expect(result.map(item => value.slice(item.start, item.end)).join('')).toBe(value)
  expect(result.map(item => item.script)).toEqual(['Latn', 'Grek', 'Cyrl', 'Deva', 'Arab'])
  expect(result.at(-1)!.level % 2).toBe(1)
})
it('restores the opening punctuation script across nested script changes', () => {
  for (const value of ['abc (क) def', 'abc [क (abc) क] def']) {
    const result = items(value)
    const at = (index: number) => result.find(item => index >= item.start && index < item.end)!.script
    expect(at(value.indexOf(value.includes('[') ? '[' : '('))).toBe('Latn')
    expect(at(value.lastIndexOf(value.includes(']') ? ']' : ')'))).toBe('Latn')
  }
})
it('keeps bracket contexts separate at isolate embedding levels', () => {
  const value = 'abc (\u2067ܫܠܡܐ (123)\u2069) xyz', result = items(value)
  expect(result.map(item => value.slice(item.start, item.end)).join('')).toBe(value)
  expect(result.find(item => item.start <= value.lastIndexOf(')') && item.end > value.lastIndexOf(')'))!.script).toBe('Latn')
})
it('preserves resolved direction for a Prepend scalar inside an Arabic grapheme', () => {
  expect(items('\u0600بب')).toEqual([
    { start: 0, end: 1, script: 'Arab', level: 2, scalarIndices: [0] },
    { start: 1, end: 3, script: 'Arab', level: 1, scalarIndices: [1, 2] },
  ])
})

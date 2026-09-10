import { test } from 'node:test'
import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { validateRenderingCorpus } from './qualify-rendering-corpus.mjs'
const fixture = () => JSON.parse(readFileSync(new URL('../qualification/rendering/manifest.json', import.meta.url)))
test('generated corpus covers every format without claiming Office parity', () => {
  const spec = validateRenderingCorpus(fixture())
  assert.deepEqual([...new Set(spec.cases.flatMap(entry => entry.formats))].sort(), ['docx', 'pdf', 'pptx', 'xlsx'])
  assert.equal(spec.externalOfficeReferenceCount, 0)
})
for (const [name, mutate] of [
  ['missing coverage', value => value.cases.pop()],
  ['duplicate case', value => value.cases[1] = value.cases[0]],
  ['arbitrary command', value => value.cases[0].id = '../other-script'],
  ['false independent claim', value => value.externalOfficeReferenceCount = 1],
  ['unclassified reference', value => value.cases[0].referenceKind = 'Office parity'],
  ['missing scope', value => value.cases[0].scope = ''],
  ['removed format', value => value.cases[4].formats = ['docx', 'pptx']],
  ['inflated oracle claim', value => value.cases[2].referenceKind = 'analytical-oracle'],
]) test(`refuses ${name}`, () => { const value = fixture(); mutate(value); assert.throws(() => validateRenderingCorpus(value)) })

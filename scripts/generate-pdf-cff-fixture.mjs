import assert from 'node:assert/strict'
import { createHash } from 'node:crypto'
import { readFileSync, writeFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { resolve } from 'node:path'
import subsetFont from 'subset-font'
import fontkit from '@pdf-lib/fontkit'

const require = createRequire(import.meta.url)
const subsetRequire = createRequire(require.resolve('subset-font'))
const sourcePath = process.argv[2]
if (!sourcePath || sourcePath.startsWith('--')) throw new Error('Usage: node scripts/generate-pdf-cff-fixture.mjs /path/to/pinned/NotoSansJP-Regular.otf [--check]')
const sha256 = bytes => createHash('sha256').update(bytes).digest('hex')
const source = readFileSync(sourcePath)
const expectedSource = 'dff723ba59d57d136764a04b9b2d03205544f7cd785a711442d6d2d085ac5073'
assert.equal(sha256(source), expectedSource, 'pinned source font hash')
assert.equal(require('subset-font/package.json').version, '2.7.0')
assert.equal(subsetRequire('harfbuzzjs/package.json').version, '0.10.3')
const text = 'Aé é Ω 日本語かなカナ'
const options = { targetFormat: 'truetype', glyphNames: true }
const bytes = await subsetFont(source, text, options)
const parsed = fontkit.create(bytes), cff = parsed['CFF ']
assert.equal(cff.isCIDFont, true)
assert.ok(cff.topDict.FDArray.length >= 2, 'retain multiple source font dictionaries')
assert.ok(new Set(cff.topDict.FDSelect.ranges.map(range => range.fd)).size >= 2, 'exercise different source dictionaries')
const directory = resolve(import.meta.dirname, '../packages/pdf/testdata/fonts')
const manifest = {
  repository: 'https://github.com/notofonts/noto-cjk', commit: '523d033d6cb47f4a80c58a35753646f5c3608a78',
  path: 'Sans/SubsetOTF/JP/NotoSansJP-Regular.otf', sourceSha256: expectedSource,
  license: 'NotoCJK-OFL.txt', licenseSha256: sha256(readFileSync(resolve(directory, 'NotoCJK-OFL.txt'))),
  generator: 'node scripts/generate-pdf-cff-fixture.mjs /path/to/pinned/NotoSansJP-Regular.otf',
  subsetFont: '2.7.0', harfbuzzjs: '0.10.3', subsetWasmSha256: sha256(readFileSync(subsetRequire.resolve('harfbuzzjs/hb-subset.wasm'))),
  text, options, glyphs: parsed.numGlyphs, fontDictionaries: cff.topDict.FDArray.length, sha256: sha256(bytes),
}
for (const [name, contents] of [['NotoSansJP-CID-subset.otf', bytes], ['NotoSansJP-CID-subset.otf.json', Buffer.from(JSON.stringify(manifest, null, 2) + '\n')]]) {
  const path = resolve(directory, name)
  if (process.argv.includes('--check')) assert.deepEqual(readFileSync(path), contents)
  else writeFileSync(path, contents)
}
console.log(`PDF CFF fixture: ${bytes.length} bytes, ${parsed.numGlyphs} glyphs, ${cff.topDict.FDArray.length} source font dictionaries`)

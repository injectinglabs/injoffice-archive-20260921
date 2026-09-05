import { createHash } from 'node:crypto'
import { readFile, writeFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = resolve(fileURLToPath(new URL('.', import.meta.url)), '..')
const schemaSource = await readFile(resolve(root, 'schemas/xlsx-native-v2.schema.json'), 'utf8')
const schemaSHA = `sha256:${createHash('sha256').update(schemaSource).digest('hex')}`
const capabilities = [
  { name: 'native-ooxml-parse', level: 'read-only' },
  { name: 'native-geometry', level: 'exact' },
  { name: 'native-decorations', level: 'exact' },
  { name: 'native-grid-commands', level: 'exact' },
  { name: 'unsupported-content', level: 'preserve-exact' },
]
const paths = [
  'go/xlsxpatch/testdata/native-xlsx-v2/valid/lexical-render.json',
  'go/xlsxpatch/testdata/native-xlsx-v2/valid/excel-authored-happy-tree.json',
]

for (const relative of paths) {
  const path = resolve(root, relative)
  const source = JSON.parse(await readFile(path, 'utf8'))
  for (const style of source.styles) style.raw_projection_sha256 = effectiveStyleDigest(style.effective)
  const { protocol, version: _version, schema_sha256: _old, capabilities: _capabilities, ...rest } = source
  await writeFile(path, `${JSON.stringify({ protocol, version: 2, schema_sha256: schemaSHA, ...rest, capabilities })}\n`)
}

function effectiveStyleDigest(style) {
  const order = ['number_format', 'font_name', 'font_size_points', 'bold', 'italic', 'font_color', 'fill_color', 'fill', 'border', 'horizontal_alignment', 'vertical_alignment', 'wrap_text', 'shrink_to_fit', 'text_rotation', 'projection', 'unsupported']
  const canonical = Object.fromEntries(order.filter((name) => style[name] !== undefined).map((name) => [name, style[name]]))
  const json = JSON.stringify(canonical).replaceAll('\u2028', '\\u2028').replaceAll('\u2029', '\\u2029')
  return `sha256:${createHash('sha256').update(json).digest('hex')}`
}

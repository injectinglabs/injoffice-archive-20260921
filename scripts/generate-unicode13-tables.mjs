import { createHash } from 'node:crypto'
import { readFileSync, writeFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = resolve(fileURLToPath(new URL('..', import.meta.url)))
const projectionPath = resolve(root, 'scripts/unicode13/ucd13-projection.txt')
const outputPath = resolve(root, 'packages/font-metrics/src/unicode13Tables.generated.ts')
const scriptPath = fileURLToPath(import.meta.url)
const vendoredSourcePath = resolve(root, 'scripts/unicode13/ucd')
const SOURCE_FILES = Object.freeze({
  Scripts: { file: 'Scripts.txt', sha256: '9a5ed1ec9b5f0d7147e9371ad792ab39203611af7637cff2aa4a5c663b172cde' },
  PropList: { file: 'PropList.txt', sha256: '485b5a3ed25dbf1f94dfa5a9b69d8b4550ffd0c33045ccc55ccfd7c80b2a40cf' },
  DerivedCoreProperties: { file: 'DerivedCoreProperties.txt', sha256: 'a5d45f59b39deaab3c72ce8c1a2e212a5e086dff11b1f9d5bb0e352642e82248' },
  UnicodeData: { file: 'UnicodeData.txt', sha256: 'bdbffbbfc8ad4d3a6d01b5891510458f3d36f7170422af4ea2bed3211a73e8bb' },
  DerivedAge: { file: 'DerivedAge.txt', sha256: 'e779a443d3aa2a3166a15becaa2b737c922480e32c0453d5956093633555078f' },
})
const SCRIPT_TAGS = Object.freeze({
  Latin: 'Latn', Cyrillic: 'Cyrl', Greek: 'Grek', Arabic: 'Arab', Hebrew: 'Hebr',
  Devanagari: 'Deva', Han: 'Hani', Hiragana: 'Kana', Katakana: 'Kana', Hangul: 'Hang',
  Inherited: 'Zinh', Common: 'Zyyy',
})
const PUNCTUATION = new Set(['Ps', 'Pi', 'Pe', 'Pf', 'Po'])
const CONTROL = new Set(['Cc', 'Cf', 'Zl', 'Zp'])

function sha256(value) {
  return createHash('sha256').update(value).digest('hex')
}

function sourceLines(value) {
  return value.split(/\r?\n/).map((line) => line.replace(/#.*/, '').replace(/^[\u0009\u0020]+|[\u0009\u0020]+$/g, '')).filter(Boolean)
}

function parseRange(value) {
  const [start, end = start] = value.split('..').map((part) => Number.parseInt(part, 16))
  if (!Number.isSafeInteger(start) || !Number.isSafeInteger(end) || start < 0 || end < start || end > 0x10ffff) throw new Error(`invalid UCD range ${value}`)
  return [start, end]
}

function mergeRanges(ranges) {
  const sorted = ranges.slice().sort((left, right) => left[0] - right[0] || left[1] - right[1])
  const merged = []
  for (const [start, end] of sorted) {
    const previous = merged.at(-1)
    if (previous && start <= previous[1] + 1) previous[1] = Math.max(previous[1], end)
    else merged.push([start, end])
  }
  return merged
}

function rangeText(ranges) {
  return mergeRanges(ranges).map(([start, end]) => `${start.toString(16).toUpperCase()}${start === end ? '' : `..${end.toString(16).toUpperCase()}`}`).join(',')
}

function readVerified(directory, key) {
  const source = SOURCE_FILES[key]
  const bytes = readFileSync(resolve(directory, source.file))
  const actual = sha256(bytes)
  if (actual !== source.sha256) throw new Error(`${source.file} SHA-256 ${actual} does not match pinned ${source.sha256}`)
  return bytes.toString('utf8')
}

function buildProjection(directory) {
  const scripts = new Map(Object.values(SCRIPT_TAGS).map((tag) => [tag, []]))
  for (const line of sourceLines(readVerified(directory, 'Scripts'))) {
    const [range, property] = line.split(';').map((part) => part.replace(/^[\u0009\u0020]+|[\u0009\u0020]+$/g, ''))
    const tag = SCRIPT_TAGS[property]
    if (tag) scripts.get(tag).push(parseRange(range))
  }
  const whiteSpace = []
  for (const line of sourceLines(readVerified(directory, 'PropList'))) {
    const [range, property] = line.split(';').map((part) => part.replace(/^[\u0009\u0020]+|[\u0009\u0020]+$/g, ''))
    if (property === 'White_Space') whiteSpace.push(parseRange(range))
  }
  const defaultIgnorable = []
  for (const line of sourceLines(readVerified(directory, 'DerivedCoreProperties'))) {
    const [range, property] = line.split(';').map((part) => part.replace(/^[\u0009\u0020]+|[\u0009\u0020]+$/g, ''))
    if (property === 'Default_Ignorable_Code_Point') defaultIgnorable.push(parseRange(range))
  }
  const assigned = []
  for (const line of sourceLines(readVerified(directory, 'DerivedAge'))) {
    const [range] = line.split(';').map((part) => part.replace(/^[\u0009\u0020]+|[\u0009\u0020]+$/g, ''))
    assigned.push(parseRange(range))
  }
  const punctuation = new Map([...PUNCTUATION].map((category) => [category, []]))
  const control = []
  let rangeFirst = null
  for (const line of sourceLines(readVerified(directory, 'UnicodeData'))) {
    const fields = line.split(';')
    if (fields.length !== 15) throw new Error('UnicodeData.txt record does not have 15 fields')
    const codePoint = Number.parseInt(fields[0], 16)
    const name = fields[1]
    const category = fields[2]
    if (name.endsWith(', First>')) {
      if (rangeFirst) throw new Error('nested UnicodeData First range')
      rangeFirst = { codePoint, category, stem: name.slice(0, -8) }
      continue
    }
    let range = [codePoint, codePoint]
    if (name.endsWith(', Last>')) {
      if (!rangeFirst || rangeFirst.category !== category || rangeFirst.stem !== name.slice(0, -7)) throw new Error('UnicodeData Last range does not match First')
      range = [rangeFirst.codePoint, codePoint]
      rangeFirst = null
    } else if (rangeFirst) throw new Error('UnicodeData First range is missing adjacent Last record')
    if (PUNCTUATION.has(category)) punctuation.get(category).push(range)
    if (CONTROL.has(category)) control.push(range)
  }
  if (rangeFirst) throw new Error('unterminated UnicodeData First range')
  const lines = [
    '# InjOffice Unicode 13 deterministic projection v1',
    ...Object.entries(SOURCE_FILES).map(([name, source]) => `source.${name}=${source.file};sha256:${source.sha256}`),
    ...[...scripts].sort(([left], [right]) => left < right ? -1 : left > right ? 1 : 0).map(([tag, ranges]) => `script.${tag}=${rangeText(ranges)}`),
    `property.White_Space=${rangeText(whiteSpace)}`,
    `property.Default_Ignorable_Code_Point=${rangeText(defaultIgnorable)}`,
    `property.Assigned=${rangeText(assigned)}`,
    ...[...punctuation].map(([category, ranges]) => `category.${category}=${rangeText(ranges)}`),
    `category.Control=${rangeText(control)}`,
    '',
  ]
  return lines.join('\n')
}

function parseProjection(value) {
  const entries = new Map()
  for (const line of value.split('\n')) {
    if (line.length === 0 || line.startsWith('#') || line.startsWith('source.')) continue
    const separator = line.indexOf('=')
    if (separator < 1 || entries.has(line.slice(0, separator))) throw new Error('malformed or duplicate Unicode projection entry')
    entries.set(line.slice(0, separator), line.slice(separator + 1))
  }
  return entries
}

function generateRuntime(projection) {
  const entries = parseProjection(projection)
  const scripts = [...entries].filter(([key]) => key.startsWith('script.')).map(([key, ranges]) => [key.slice(7), ranges])
  const encoding = JSON.stringify({
    scripts,
    whiteSpace: entries.get('property.White_Space'),
    defaultIgnorable: entries.get('property.Default_Ignorable_Code_Point'),
    assigned: entries.get('property.Assigned'),
    punctuation: Object.fromEntries([...PUNCTUATION].map((category) => [category, entries.get(`category.${category}`)])),
    control: entries.get('category.Control'),
  })
  const generatorDigest = sha256(readFileSync(scriptPath))
  return `/** Generated by scripts/generate-unicode13-tables.mjs. DO NOT EDIT. */\n` +
    `export const UNICODE_13_VERSION = '13.0.0' as const\n` +
    `export const UNICODE_13_UCD_SOURCE_SHA256 = ${JSON.stringify(Object.fromEntries(Object.entries(SOURCE_FILES).map(([key, value]) => [value.file, `sha256:${value.sha256}`])), null, 2)} as const\n` +
    `export const UNICODE_13_PROJECTION_SHA256 = 'sha256:${sha256(projection)}' as const\n` +
    `export const UNICODE_13_GENERATOR_SHA256 = 'sha256:${generatorDigest}' as const\n` +
    `export const UNICODE_13_TABLES_ENCODING_SHA256 = 'sha256:${sha256(encoding)}' as const\n` +
    `export const UNICODE_13_TABLES_ENCODING = ${JSON.stringify(encoding)} as const\n`
}

function writeOrCheck(path, expected, check) {
  if (check) {
    const actual = readFileSync(path, 'utf8')
    if (actual !== expected) throw new Error(`${path} is stale; run node scripts/generate-unicode13-tables.mjs --from-ucd <Unicode-13-UCD-directory>`)
  } else writeFileSync(path, expected)
}

const args = process.argv.slice(2)
const check = args.includes('--check')
const sourceIndex = args.indexOf('--from-ucd')
const directory = sourceIndex >= 0 ? args[sourceIndex + 1] : vendoredSourcePath
if (!directory) throw new Error('--from-ucd requires the directory containing pinned Unicode 13 source files')
const regeneratedProjection = buildProjection(resolve(directory))
writeOrCheck(projectionPath, regeneratedProjection, check)
const projection = readFileSync(projectionPath, 'utf8')
for (const source of Object.values(SOURCE_FILES)) if (!projection.includes(`source.${Object.keys(SOURCE_FILES).find((key) => SOURCE_FILES[key] === source)}=${source.file};sha256:${source.sha256}`)) throw new Error('Unicode projection source provenance is missing or stale')
writeOrCheck(outputPath, generateRuntime(projection), check)

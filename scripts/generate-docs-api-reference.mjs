import { existsSync, mkdirSync, readdirSync, readFileSync, writeFileSync } from 'node:fs'
import { basename, dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = resolve(fileURLToPath(new URL('.', import.meta.url)), '..')
const outPath = resolve(root, 'docs/api-reference.json')
const check = process.argv.includes('--check')

function packageDirs() {
  const base = resolve(root, 'packages')
  return readdirSync(base, { withFileTypes: true })
    .filter((entry) => entry.isDirectory() && existsSync(join(base, entry.name, 'package.json')))
    .map((entry) => join(base, entry.name))
    .sort()
}

function resolveDts(fromFile, specifier) {
  const base = resolve(dirname(fromFile), specifier)
  const candidates = [
    base,
    base.replace(/\.js$/, '.d.ts'),
    `${base}.d.ts`,
    join(base, 'index.d.ts'),
  ]
  return candidates.find((candidate) => existsSync(candidate) && candidate.endsWith('.d.ts'))
}

function captureSignature(text, start, kind) {
  if (kind === 'function') {
    let depth = 0
    for (let index = start; index < text.length; index += 1) {
      const char = text[index]
      if (char === '(') depth += 1
      else if (char === ')') {
        depth -= 1
        if (depth === 0) {
          const rest = text.slice(index + 1)
          const match = rest.match(/^\s*(:\s*[^;{]+)?/)
          return collapse(text.slice(start, index + 1 + (match?.[0].length ?? 0)))
        }
      }
    }
  }
  const line = text.slice(start, text.indexOf('\n', start) === -1 ? text.length : text.indexOf('\n', start))
  return collapse(line.replace(/[{;]\s*$/, ''))
}

function collapse(value) {
  return value.replace(/\s+/g, ' ').trim()
}

function parseNamedExports(list, typeOnlyGroup) {
  return list.split(',').map((part) => part.trim()).filter(Boolean).map((part) => {
    const typeOnly = typeOnlyGroup || /^type\s+/.test(part)
    const body = part.replace(/^type\s+/, '').replace(/^default\s+as\s+/, '')
    const [exported, alias] = body.split(/\s+as\s+/).map((item) => item.trim())
    return { exported, name: alias || exported, typeOnly }
  }).filter((item) => item.name && item.name !== 'default')
}

function collect(file, seen = new Set()) {
  const abs = resolve(file)
  if (seen.has(abs) || !existsSync(abs)) return []
  seen.add(abs)
  const text = readFileSync(abs, 'utf8')
  const symbols = []
  const byName = new Map()

  const add = (symbol) => {
    if (!symbol?.name || byName.has(symbol.name)) return
    byName.set(symbol.name, symbol)
    symbols.push(symbol)
  }

  const starRe = /export\s+\*\s+from\s+['"](\.[^'"]+)['"]/g
  const fromRe = /export\s+(type\s+)?\{([^}]+)\}\s+from\s+['"](\.[^'"]+)['"]/g
  const localRe = /export\s+(?:declare\s+)?(?:async\s+)?(function|class|const|enum|type|interface)\s+([A-Za-z_$][\w$]*)/g

  let match
  while ((match = starRe.exec(text))) {
    const target = resolveDts(abs, match[1])
    if (target) collect(target, seen).forEach(add)
  }
  while ((match = fromRe.exec(text))) {
    const names = parseNamedExports(match[2], Boolean(match[1]))
    const target = resolveDts(abs, match[3])
    const nested = target ? collect(target, seen) : []
    const nestedByName = new Map(nested.map((symbol) => [symbol.name, symbol]))
    for (const item of names) {
      const found = nestedByName.get(item.exported) ?? nestedByName.get(item.name)
      add(found
        ? { ...found, name: item.name, kind: item.typeOnly ? 'type' : found.kind }
        : { name: item.name, kind: item.typeOnly ? 'type' : 'export', signature: item.name })
    }
  }
  while ((match = localRe.exec(text))) {
    add({
      name: match[2],
      kind: match[1] === 'const' || match[1] === 'enum' ? match[1] : match[1],
      signature: captureSignature(text, match.index, match[1]),
    })
  }
  return symbols
}

function dtsForExport(directory, target) {
  if (typeof target !== 'string') return null
  if (target.endsWith('.d.ts')) return resolve(directory, target)
  if (target.endsWith('.js')) return resolve(directory, target.replace(/\.js$/, '.d.ts'))
  return null
}

function goModules() {
  const goRoot = resolve(root, 'go')
  return readdirSync(goRoot, { withFileTypes: true })
    .filter((entry) => entry.isDirectory() && existsSync(join(goRoot, entry.name, 'go.mod')))
    .map((entry) => {
      const directory = join(goRoot, entry.name)
      const mod = readFileSync(join(directory, 'go.mod'), 'utf8')
      const module = /^module\s+(\S+)/m.exec(mod)?.[1] ?? `go/${entry.name}`
      const readmePath = join(directory, 'README.md')
      const summary = existsSync(readmePath)
        ? (readFileSync(readmePath, 'utf8').split('\n').find((line) => line.trim() && !line.startsWith('#')) ?? '').trim()
        : ''
      return { path: `go/${entry.name}`, module, summary }
    })
    .sort((left, right) => left.path.localeCompare(right.path))
}

function build() {
  const packages = []
  for (const directory of packageDirs()) {
    const manifest = JSON.parse(readFileSync(join(directory, 'package.json'), 'utf8'))
    const exports = []
    for (const [subpath, target] of Object.entries(manifest.exports ?? {})) {
      const conditions = typeof target === 'string' ? { types: target } : target
      const dts = dtsForExport(directory, conditions.types ?? conditions.import ?? conditions.default)
      if (!dts) continue
      if (!existsSync(dts)) {
        throw new Error(`${manifest.name} export ${subpath} is missing ${dts}. Run npm run build first.`)
      }
      const symbols = collect(dts)
        .sort((left, right) => left.name.localeCompare(right.name))
        .map((symbol) => ({ name: symbol.name, kind: symbol.kind, signature: symbol.signature }))
      exports.push({ subpath, file: basename(dts), symbols })
    }
    packages.push({
      name: manifest.name,
      description: manifest.description,
      exports,
    })
  }
  return { schemaVersion: 1, packages, go: goModules() }
}

const document = build()
const serialized = `${JSON.stringify(document, null, 2)}\n`
if (check) {
  if (!existsSync(outPath)) {
    console.error(`missing ${outPath}; run node scripts/generate-docs-api-reference.mjs`)
    process.exit(1)
  }
  const current = readFileSync(outPath, 'utf8')
  if (current !== serialized) {
    console.error('docs/api-reference.json is stale; run node scripts/generate-docs-api-reference.mjs')
    process.exit(1)
  }
  process.exit(0)
}

mkdirSync(dirname(outPath), { recursive: true })
writeFileSync(outPath, serialized)
console.log(`wrote ${outPath} (${document.packages.length} packages, ${document.go.length} Go modules)`)

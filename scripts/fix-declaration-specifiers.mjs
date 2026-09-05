import { existsSync, readdirSync, readFileSync, statSync, unlinkSync, writeFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'

const declarationExtensions = new Map([
  ['.d.ts', '.js'],
  ['.d.mts', '.mjs'],
  ['.d.cts', '.cjs'],
])

const declarationFiles = (directory) => readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
  const path = resolve(directory, entry.name)
  if (entry.isDirectory()) return declarationFiles(path)
  return /\.d\.(?:c|m)?ts$/.test(entry.name) ? [path] : []
})

const hasExplicitExtension = (specifier) => /\.(?:[cm]?[jt]sx?|json|css|wasm|node)$/.test(specifier)

const runtimeSpecifier = (file, specifier) => {
  if (!specifier.startsWith('.') || hasExplicitExtension(specifier)) return undefined
  const base = resolve(dirname(file), specifier)
  for (const [declarationExtension, runtimeExtension] of declarationExtensions) {
    if (existsSync(`${base}${declarationExtension}`)) return `${specifier}${runtimeExtension}`
    if (existsSync(resolve(base, `index${declarationExtension}`))) return `${specifier}/index${runtimeExtension}`
  }
  return undefined
}

const moduleSpecifiers = (source) => {
  const results = []
  const pattern = /(\bfrom\s*|\bimport\s*\(\s*|\brequire\s*\(\s*|\bimport\s*)(['"])(\.{1,2}\/[^'"\r\n]+)\2/g
  for (const match of source.matchAll(pattern)) {
    const specifier = match[3]
    const start = match.index + match[1].length + 1
    results.push({ start, end: start + specifier.length, specifier })
  }
  return results
}

export const rewriteDeclarationFile = (file) => {
  const source = readFileSync(file, 'utf8')
  const replacements = moduleSpecifiers(source).flatMap(({ start, end, specifier }) => {
    const replacement = runtimeSpecifier(file, specifier)
    return replacement ? [{ start, end, replacement }] : []
  }).sort((left, right) => right.start - left.start)
  if (replacements.length === 0) return 0
  let updated = source
  for (const replacement of replacements) {
    updated = `${updated.slice(0, replacement.start)}${replacement.replacement}${updated.slice(replacement.end)}`
  }
  updated = updated.replace(/\r?\n\/\/# sourceMappingURL=[^\r\n]+\r?\n?$/, '\n')
  writeFileSync(file, updated)
  const declarationMap = `${file}.map`
  if (existsSync(declarationMap)) unlinkSync(declarationMap)
  return replacements.length
}

export const rewriteDeclarationTree = (directory) => {
  if (!existsSync(directory) || !statSync(directory).isDirectory()) return { files: 0, specifiers: 0 }
  let files = 0
  let specifiers = 0
  for (const file of declarationFiles(directory)) {
    const rewritten = rewriteDeclarationFile(file)
    if (rewritten > 0) files += 1
    specifiers += rewritten
  }
  return { files, specifiers }
}

export const extensionlessDeclarationSpecifiers = (directory) => {
  if (!existsSync(directory) || !statSync(directory).isDirectory()) return []
  return declarationFiles(directory).flatMap((file) => {
    const source = readFileSync(file, 'utf8')
    return moduleSpecifiers(source)
      .filter(({ specifier }) => !hasExplicitExtension(specifier))
      .map(({ specifier }) => ({ file, specifier }))
  })
}

const run = () => {
  const root = resolve(import.meta.dirname, '..')
  const packageRoot = resolve(root, 'packages')
  let files = 0
  let specifiers = 0
  for (const entry of readdirSync(packageRoot, { withFileTypes: true })) {
    if (!entry.isDirectory()) continue
    const result = rewriteDeclarationTree(resolve(packageRoot, entry.name, 'dist'))
    files += result.files
    specifiers += result.specifiers
  }
  console.log(`Rewrote ${specifiers} relative declaration specifiers across ${files} files.`)
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) run()

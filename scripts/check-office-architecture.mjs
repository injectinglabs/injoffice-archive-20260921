import { existsSync, readFileSync, readdirSync } from 'node:fs'
import { dirname, relative, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const failures = []
const productionExtensions = /\.(?:go|[cm]?[jt]sx?)$/
const ignored = /(?:^|\/)(?:node_modules|dist|coverage)(?:\/|$)|(?:^|\/)[^/]+(?:_test\.go|\.test\.[^/]+)$/
const forbiddenDependency = /^(?:electron|mammoth)$/

function moduleSpecifiers(source) {
  const specifiers = []
  const pattern = /\b(?:import|export)\s+(?:[^'";]*?\s+from\s+)?['"]([^'"]+)['"]|\b(?:import|require)\s*\(\s*['"]([^'"]+)['"]\s*\)/g
  for (const match of source.matchAll(pattern)) specifiers.push(match[1] ?? match[2])
  return specifiers
}

function walk(directory, output = []) {
  if (!existsSync(directory)) return output
  for (const entry of readdirSync(directory, { withFileTypes: true })) {
    const path = resolve(directory, entry.name)
    if (entry.isDirectory()) walk(path, output)
    else if (productionExtensions.test(entry.name)) output.push(path)
  }
  return output
}

const roots = [resolve(root, 'packages'), resolve(root, 'apps'), resolve(root, 'go')]
const files = roots.flatMap((directory) => walk(directory)).filter((path) => {
  const name = relative(root, path).replaceAll('\\', '/')
  return !ignored.test(name) && (name.includes('/src/') || name.startsWith('go/'))
})

for (const path of files) {
  const name = relative(root, path).replaceAll('\\', '/')
  const source = readFileSync(path, 'utf8')
  const imports = moduleSpecifiers(source)
  for (const specifier of imports) {
    if (specifier === 'electron' || specifier.startsWith('electron/')) failures.push(`${name}: production source imports Electron via ${specifier}`)
    if (specifier === 'mammoth' || specifier.startsWith('mammoth/')) failures.push(`${name}: production source imports Mammoth via ${specifier}`)
  }

  const pptxAuthority = name.startsWith('packages/pptx-render/src/') || name.startsWith('packages/pptx-native/src/') || /^go\/pptxpatch\/native[^/]*\.go$/.test(name)
  if (pptxAuthority) {
    for (const token of [/@(?:types\/)?react(?:-dom)?(?:\/|['"])/, /\bkonva\b/i, /\bDOMParser\b/, /\b(?:document|window)\s*\./, /\b(?:innerHTML|outerHTML|HTMLElement)\b/, /\bDeckView\b/]) {
      if (token.test(source)) failures.push(`${name}: native PPTX/render authority contains forbidden DOM or view token ${token}`)
    }
    if (imports.some((specifier) => specifier === '@injoffice/slides' || specifier.startsWith('@injoffice/slides/'))) failures.push(`${name}: native PPTX/render authority imports the legacy slides package`)
  }

  const docxAuthority = /^go\/docxpatch\/native[^/]*\.go$/.test(name) || /^packages\/docs\/src\/native[^/]*\.[jt]sx?$/.test(name)
  if (docxAuthority) {
    for (const token of [/\bDOMParser\b/, /\b(?:innerHTML|outerHTML)\b/, /\bdocument\s*\.\s*(?:createElement|querySelector)/]) {
      if (token.test(source)) failures.push(`${name}: native DOCX authority contains forbidden HTML reconstruction token ${token}`)
    }
  }

  if (!['packages/slides/src/DeckView.tsx', 'packages/slides/src/index.ts'].includes(name) && /(?:import|export)[^\n;]*\bDeckView\b/.test(source)) {
    failures.push(`${name}: imports or re-exports the legacy DOM slide renderer as production authority`)
  }
}

const manifestPaths = [
  resolve(root, 'package.json'),
  ...readdirSync(resolve(root, 'packages'), { withFileTypes: true }).filter((entry) => entry.isDirectory()).map((entry) => resolve(root, 'packages', entry.name, 'package.json')),
  ...readdirSync(resolve(root, 'apps'), { withFileTypes: true }).filter((entry) => entry.isDirectory()).map((entry) => resolve(root, 'apps', entry.name, 'package.json')),
]
for (const manifestPath of manifestPaths) {
  if (!existsSync(manifestPath)) continue
  const manifest = JSON.parse(readFileSync(manifestPath, 'utf8'))
  for (const group of ['dependencies', 'devDependencies', 'peerDependencies', 'optionalDependencies']) {
    for (const [dependency, requested] of Object.entries(manifest[group] ?? {})) {
      if (forbiddenDependency.test(dependency) || typeof requested === 'string' && /(?:^|npm:)(?:electron|mammoth)(?:\/|@|$)/.test(requested)) failures.push(`${relative(root, manifestPath)}: forbidden ${group} entry ${dependency}@${requested}`)
    }
  }
}

const lockSource = readFileSync(resolve(root, 'package-lock.json'), 'utf8')
if (/node_modules\/(?:electron|mammoth)(?:\/|")|npm:(?:electron|mammoth)\//.test(lockSource)) failures.push('package-lock.json: forbidden Electron/Mammoth package or npm alias')

if (failures.length > 0) {
  console.error(failures.map((failure) => `- ${failure}`).join('\n'))
  process.exit(1)
}
console.log(`Validated ${files.length} production files: no Electron dependency, Mammoth native authority, or DOM slide authority regression.`)

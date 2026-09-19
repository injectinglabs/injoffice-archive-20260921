import { existsSync, readFileSync, readdirSync } from 'node:fs'
import { dirname, relative, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const failures = []
const productionExtensions = /\.(?:go|[cm]?[jt]sx?)$/
const ignored = /(?:^|\/)(?:node_modules|dist|coverage)(?:\/|$)|(?:^|\/)[^/]+(?:_test\.go|\.test\.[^/]+)$/
const dependencyGroups = ['dependencies', 'devDependencies', 'peerDependencies', 'optionalDependencies']
// Electron is the desktop host runtime only. packages/*, go/, and production
// /src/ (renderer) must not depend on it; Mammoth stays forbidden everywhere.
const desktopElectronGroups = new Set(['dependencies', 'devDependencies'])
const desktopHostManifest = 'apps/desktop/package.json'
const desktopHostLockPaths = new Set(['apps/desktop', 'node_modules/@injoffice/desktop'])
const allowedElectronInstalls = new Set(['node_modules/electron', 'apps/desktop/node_modules/electron'])
const forbiddenAlias = /(?:^|npm:)(electron|mammoth)(?:\/|@|$)/

function posixName(path) {
  return relative(root, path).replaceAll('\\', '/')
}

function posixLockPath(lockPath) {
  return lockPath.replaceAll('\\', '/')
}

function forbiddenTarget(name, requested) {
  if (name === 'mammoth') return 'mammoth'
  if (name === 'electron') return 'electron'
  if (typeof requested === 'string') {
    const match = requested.match(forbiddenAlias)
    if (match) return match[1]
  }
  return null
}

function isDesktopHostManifest(relativePath) {
  return relativePath === desktopHostManifest
}

function isDesktopHostLockEntry(lockPath, entry = {}) {
  const path = posixLockPath(lockPath)
  return desktopHostLockPaths.has(path) || entry.name === '@injoffice/desktop'
}

function installedPackageName(lockPath, entry = {}) {
  if (typeof entry.name === 'string' && entry.name.length > 0) return entry.name
  const path = posixLockPath(lockPath)
  const marker = 'node_modules/'
  const offset = path.lastIndexOf(marker)
  if (offset === -1) return ''
  const suffix = path.slice(offset + marker.length)
  const parts = suffix.split('/')
  return parts[0]?.startsWith('@') ? parts.slice(0, 2).join('/') : (parts[0] ?? '')
}

function electronInstallAllowed(lockPath) {
  return allowedElectronInstalls.has(posixLockPath(lockPath))
}

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
  const name = posixName(path)
  return !ignored.test(name) && (name.includes('/src/') || name.startsWith('go/'))
})

for (const path of files) {
  const name = posixName(path)
  const source = readFileSync(path, 'utf8')
  const imports = moduleSpecifiers(source)
  for (const specifier of imports) {
    if (specifier === 'electron' || specifier.startsWith('electron/')) failures.push(`${name}: production source imports Electron via ${specifier}`)
    if (specifier === 'mammoth' || specifier.startsWith('mammoth/')) failures.push(`${name}: production source imports Mammoth via ${specifier}`)
  }

  const pptxAuthority = name.startsWith('packages/pptx-render/src/') || name.startsWith('packages/pptx-native/src/') || /^go\/pptxpatch\/native[^/]*\.go$/.test(name)
  if (pptxAuthority) {
    // Prose is not code. A comment ending a sentence on "window." or naming a
    // "document" is not a DOM reference, and matching it fails the gate on a
    // file that never touches the DOM. Strip line and block comments first so
    // the tokens below are judged against the code they are meant to police.
    const code = source.replace(/\/\*[\s\S]*?\*\//g, ' ').replace(/(^|[^:])\/\/[^\n]*/g, '$1')
    for (const token of [/@(?:types\/)?react(?:-dom)?(?:\/|['"])/, /\bkonva\b/i, /\bDOMParser\b/, /\b(?:document|window)\s*\./, /\b(?:innerHTML|outerHTML|HTMLElement)\b/, /\bDeckView\b/]) {
      if (token.test(code)) failures.push(`${name}: native PPTX/render authority contains forbidden DOM or view token ${token}`)
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
  const name = posixName(manifestPath)
  const manifest = JSON.parse(readFileSync(manifestPath, 'utf8'))
  for (const group of dependencyGroups) {
    for (const [dependency, requested] of Object.entries(manifest[group] ?? {})) {
      const target = forbiddenTarget(dependency, requested)
      if (target === 'mammoth') failures.push(`${name}: forbidden ${group} entry ${dependency}@${requested}`)
      else if (target === 'electron' && !(isDesktopHostManifest(name) && desktopElectronGroups.has(group))) {
        failures.push(`${name}: forbidden ${group} entry ${dependency}@${requested}`)
      }
    }
  }
}

const lock = JSON.parse(readFileSync(resolve(root, 'package-lock.json'), 'utf8'))
let desktopPullsElectron = false
const electronInstalls = []
for (const [lockPath, entry] of Object.entries(lock.packages ?? {})) {
  const path = posixLockPath(lockPath)
  const installed = installedPackageName(path, entry)
  if (installed === 'mammoth') failures.push(`package-lock.json: forbidden Mammoth package at ${path || 'root'}`)
  if (installed === 'electron') {
    electronInstalls.push(path || 'root')
    if (!electronInstallAllowed(path)) failures.push(`package-lock.json: Electron at ${path || 'root'} is not owned by @injoffice/desktop`)
  }
  for (const group of dependencyGroups) {
    for (const [dependency, requested] of Object.entries(entry[group] ?? {})) {
      const target = forbiddenTarget(dependency, requested)
      if (target === 'mammoth') failures.push(`package-lock.json: ${path || 'root'} forbidden ${group} entry ${dependency}@${requested}`)
      else if (target === 'electron') {
        if (isDesktopHostLockEntry(path, entry)) desktopPullsElectron = true
        else failures.push(`package-lock.json: ${path || 'root'} forbidden ${group} entry ${dependency}@${requested}`)
      }
    }
  }
}
if (electronInstalls.length > 0 && !desktopPullsElectron) {
  failures.push('package-lock.json: Electron is present but not pulled by @injoffice/desktop')
}

if (failures.length > 0) {
  console.error(failures.map((failure) => `- ${failure}`).join('\n'))
  process.exit(1)
}
console.log(`Validated ${files.length} production files: Electron only as desktop host, no Mammoth native authority, or DOM slide authority regression.`)

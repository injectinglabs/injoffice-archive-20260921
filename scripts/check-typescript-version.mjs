import { existsSync, readFileSync, readdirSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { version as installedVersion } from 'typescript'

const EXPECTED_VERSION = '7.0.2'
const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const failures = []

const manifestPaths = ['package.json', ...['apps', 'packages']
  .flatMap((base) => readdirSync(resolve(root, base), { withFileTypes: true })
    .filter((entry) => entry.isDirectory() && existsSync(resolve(root, base, entry.name, 'package.json')))
    .map((entry) => `${base}/${entry.name}/package.json`))]
  .sort()

for (const path of manifestPaths) {
  const manifest = JSON.parse(readFileSync(resolve(root, path), 'utf8'))
  const declarations = Object.fromEntries(['dependencies', 'devDependencies', 'peerDependencies', 'optionalDependencies']
    .flatMap((kind) => Object.entries(manifest[kind] ?? {}).map(([name, range]) => [`${kind}.${name}`, range])))
  for (const [declaration, range] of Object.entries(declarations)) {
    if (declaration.endsWith('.@typescript/native')) failures.push(`${path}: remove the @typescript/native compatibility alias`)
    if (typeof range === 'string' && range.includes('@typescript/typescript6')) failures.push(`${path}: remove the TypeScript 6 compatibility alias from ${declaration}`)
  }
  if (manifest.devDependencies?.typescript !== EXPECTED_VERSION) {
    failures.push(`${path}: devDependencies.typescript must be pinned to ${EXPECTED_VERSION}`)
  }
}

const lock = JSON.parse(readFileSync(resolve(root, 'package-lock.json'), 'utf8'))
const lockedTypeScriptPaths = Object.keys(lock.packages ?? {}).filter((path) => path.endsWith('node_modules/typescript'))
if (lockedTypeScriptPaths.length !== 1 || lockedTypeScriptPaths[0] !== 'node_modules/typescript') {
  failures.push(`package-lock.json: expected one hoisted TypeScript installation, found ${lockedTypeScriptPaths.join(', ') || 'none'}`)
}
if (lock.packages?.['node_modules/typescript']?.version !== EXPECTED_VERSION) {
  failures.push(`package-lock.json: node_modules/typescript must resolve to ${EXPECTED_VERSION}`)
}
if (installedVersion !== EXPECTED_VERSION) failures.push(`installed TypeScript is ${installedVersion}; expected ${EXPECTED_VERSION}`)

if (failures.length > 0) {
  console.error(failures.map((failure) => `- ${failure}`).join('\n'))
  process.exit(1)
}

console.log(`TypeScript ${EXPECTED_VERSION} is pinned across ${manifestPaths.length - 1} workspaces and root tooling.`)

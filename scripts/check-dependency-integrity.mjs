import { existsSync, readFileSync, readdirSync } from 'node:fs'
import { dirname, relative, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import {
  DEPENDENCY_KINDS,
  installedNameFromLockPath,
  integrityHash,
  packageNameFromLockPath,
  parseGoMod,
  purlName,
  shortestTrace,
} from './dependency-integrity-lib.mjs'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const lock = JSON.parse(readFileSync(resolve(root, 'package-lock.json'), 'utf8'))
const rootManifest = JSON.parse(readFileSync(resolve(root, 'package.json'), 'utf8'))
const evidence = JSON.parse(readFileSync(resolve(root, 'docs/provenance/npm-license-evidence.json'), 'utf8'))
const failures = []

const workspaceManifestPaths = ['apps', 'packages']
  .flatMap((base) => readdirSync(resolve(root, base), { withFileTypes: true })
    .filter((entry) => entry.isDirectory() && existsSync(resolve(root, base, entry.name, 'package.json')))
    .map((entry) => `${base}/${entry.name}/package.json`))
  .sort()

const workspaces = workspaceManifestPaths.map((path) => ({
  path,
  directory: dirname(path),
  manifest: JSON.parse(readFileSync(resolve(root, path), 'utf8')),
}))

const evidenceByIdentity = new Map()
for (const review of evidence.reviews) {
  for (const artifact of review.artifacts) {
    const identity = `${artifact.package}@${review.version}`
    if (evidenceByIdentity.has(identity)) failures.push(`duplicate reviewed-license evidence for ${identity}`)
    evidenceByIdentity.set(identity, { ...review, integrity: artifact.integrity })
  }
}

const components = []
const topLevelByInstalledName = new Map()
for (const [lockPath, entry] of Object.entries(lock.packages ?? {})) {
  if (!lockPath.startsWith('node_modules/') || entry.link || !entry.version) continue
  const installedName = installedNameFromLockPath(lockPath)
  const packageName = packageNameFromLockPath(lockPath, entry)
  const identity = `${installedName}@${entry.version}`
  const review = evidenceByIdentity.get(identity)
  if (review && review.integrity !== entry.integrity) {
    failures.push(`${identity}: reviewed license evidence is bound to a different integrity digest`)
  }
  const license = entry.license ?? review?.license ?? null
  const licenseSource = entry.license ? 'package-lock' : review ? 'reviewed-evidence' : 'unasserted'
  const component = { lockPath, entry, installedName, packageName, identity, license, licenseSource }
  components.push(component)
  if (lockPath === `node_modules/${installedName}`) topLevelByInstalledName.set(installedName, component)
}
components.sort((a, b) => a.lockPath.localeCompare(b.lockPath))

const direct = []
for (const workspace of workspaces) {
  const lockedWorkspace = lock.packages?.[workspace.directory]
  if (!lockedWorkspace) failures.push(`${workspace.path}: workspace is absent from package-lock.json`)
  for (const kind of DEPENDENCY_KINDS) {
    const declared = workspace.manifest[kind] ?? {}
    const locked = lockedWorkspace?.[kind] ?? {}
    for (const [name, range] of Object.entries(declared)) {
      if (locked[name] !== range) failures.push(`${workspace.path}: ${kind}.${name} is stale or absent in package-lock.json`)
      if (name.startsWith('@injoffice/')) continue
      const resolved = topLevelByInstalledName.get(name)
      if (!resolved) failures.push(`${workspace.path}: external ${kind}.${name} does not resolve in package-lock.json`)
      direct.push({ workspace: workspace.path, kind, name, range, resolved })
    }
    for (const name of Object.keys(locked)) {
      if (!(name in declared)) failures.push(`${workspace.path}: package-lock.json has stale ${kind}.${name}`)
    }
  }
}

const graph = new Map()
for (const component of components) {
  if (component.lockPath !== `node_modules/${component.installedName}`) continue
  const edges = new Set()
  for (const kind of ['dependencies', 'optionalDependencies', 'peerDependencies']) {
    for (const name of Object.keys(component.entry[kind] ?? {})) edges.add(name)
  }
  graph.set(component.installedName, edges)
}

const rootNames = [...new Set(direct
  .filter((entry) => entry.kind !== 'devDependencies')
  .map((entry) => entry.name))].sort()
const reachable = new Set()
const queue = [...rootNames]
while (queue.length > 0) {
  const current = queue.shift()
  if (reachable.has(current)) continue
  reachable.add(current)
  for (const next of graph.get(current) ?? []) queue.push(next)
}

const unknownComponents = components.filter((component) => !component.license)
const releaseBlockers = unknownComponents.filter((component) =>
  component.lockPath === `node_modules/${component.installedName}` && reachable.has(component.installedName))

for (const blocker of releaseBlockers) {
  const trace = shortestTrace(graph, rootNames, blocker.installedName)
  failures.push(`${blocker.identity}: no license assertion; production trace ${trace?.join(' -> ') ?? '(unresolved)'}`)
}

const proComponents = components.filter((component) => component.installedName.startsWith('@univerjs-pro/'))
for (const component of proComponents) {
  if (component.license) failures.push(`${component.identity}: Univer Pro dependency must be reviewed even when it declares a license`)
}

const goModules = []
for (const directory of readdirSync(resolve(root, 'go'), { withFileTypes: true }).filter((entry) => entry.isDirectory())) {
  const path = `go/${directory.name}/go.mod`
  if (!existsSync(resolve(root, path))) continue
  const source = readFileSync(resolve(root, path), 'utf8')
  const moduleName = /^module\s+(\S+)/m.exec(source)?.[1]
  const parsed = parseGoMod(source)
  const external = []
  for (const requirement of parsed.requires) {
    if (requirement.module.startsWith('github.com/injectinglabs/injoffice/go/')) {
      const replacement = parsed.replacements.get(requirement.module)
      if (!replacement || !replacement.target.startsWith('../')) failures.push(`${path}: internal module ${requirement.module} must use a local replace directive`)
    } else {
      external.push(requirement)
    }
  }
  goModules.push({ path, module: moduleName, requires: parsed.requires, external })
}
goModules.sort((a, b) => a.path.localeCompare(b.path))

function toCycloneDx() {
  return {
    bomFormat: 'CycloneDX',
    specVersion: '1.5',
    version: 1,
    metadata: {
      component: {
        type: 'application',
        name: rootManifest.name,
        version: rootManifest.version,
        licenses: [{ license: { id: rootManifest.license } }],
      },
      properties: [
        { name: 'injoffice:generator', value: 'scripts/check-dependency-integrity.mjs' },
        { name: 'injoffice:source', value: 'package-lock.json; offline and deterministic' },
        { name: 'injoffice:go-third-party-direct-count', value: String(goModules.reduce((sum, module) => sum + module.external.length, 0)) },
        { name: 'injoffice:unasserted-license-count', value: String(unknownComponents.length) },
        { name: 'injoffice:release-blocker-count', value: String(releaseBlockers.length) },
      ],
    },
    components: components.map((component) => {
      const hash = integrityHash(component.entry.integrity)
      const split = component.packageName.startsWith('@') ? component.packageName.split('/') : [null, component.packageName]
      return {
        'bom-ref': `npm:${component.lockPath}@${component.entry.version}`,
        type: 'library',
        ...(split[0] ? { group: split[0] } : {}),
        name: split[1],
        version: component.entry.version,
        purl: `pkg:npm/${purlName(component.packageName)}@${encodeURIComponent(component.entry.version)}`,
        scope: component.entry.dev || component.entry.optional ? 'optional' : 'required',
        ...(hash ? { hashes: [hash] } : {}),
        ...(component.license ? { licenses: [{ expression: component.license }] } : {}),
        ...(component.entry.resolved ? { externalReferences: [{ type: 'distribution', url: component.entry.resolved }] } : {}),
        properties: [
          { name: 'injoffice:lockfile-path', value: component.lockPath },
          { name: 'injoffice:installed-name', value: component.installedName },
          { name: 'injoffice:license-source', value: component.licenseSource },
          ...(component.license ? [] : [{ name: 'injoffice:license-status', value: 'NOASSERTION' }]),
        ],
      }
    }),
  }
}

if (process.argv.includes('--sbom')) {
  await new Promise((resolveWrite) => process.stdout.write(`${JSON.stringify(toCycloneDx(), null, 2)}\n`, resolveWrite))
  process.exit(0)
}

const directRuntime = direct.filter((entry) => entry.kind !== 'devDependencies')
console.log(`Dependency inventory: ${workspaces.length} workspaces, ${direct.length} direct declarations (${directRuntime.length} runtime/peer/optional), ${components.length} resolved npm components.`)
console.log(`Go inventory: ${goModules.length} modules, ${goModules.reduce((sum, module) => sum + module.external.length, 0)} direct third-party modules.`)
console.log(`License inventory: ${unknownComponents.length} unasserted lockfile components; ${releaseBlockers.length} are production-reachable release blockers.`)
if (proComponents.length > 0) {
  console.log(`Univer Pro inventory: ${proComponents.length} resolved packages.`)
}

if (failures.length > 0) {
  console.error(failures.map((failure) => `- ${failure}`).join('\n'))
  process.exit(1)
}

console.log('Dependency integrity and license assertions are release-ready.')

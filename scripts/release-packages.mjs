import { spawnSync } from 'node:child_process'
import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { basename, join, resolve } from 'node:path'
import { tmpdir } from 'node:os'
import { publishedIntegrity } from './release-registry.mjs'

const root = resolve(import.meta.dirname, '..')
const packageRoot = resolve(root, 'packages')
const tag = process.argv.find((value) => /^v\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?$/.test(value))
const dryRun = process.argv.includes('--dry-run')
const optionValue = (option) => {
  const index = process.argv.indexOf(option)
  return index === -1 ? undefined : process.argv[index + 1]
}
const packDestinationArg = optionValue('--pack-destination')
const stageFromArg = optionValue('--stage-from')
const bootstrapFromArg = optionValue('--bootstrap-from')
const selectedModes = [dryRun, packDestinationArg !== undefined, stageFromArg !== undefined, bootstrapFromArg !== undefined].filter(Boolean).length

if (!tag || selectedModes !== 1 || packDestinationArg === '' || stageFromArg === '' || bootstrapFromArg === '') {
  console.error('Usage: npm run release:packages -- vX.Y.Z (--dry-run | --pack-destination DIR | --stage-from DIR | --bootstrap-from DIR)')
  process.exit(1)
}
if (bootstrapFromArg !== undefined && process.env.CI) {
  console.error('Refusing to bootstrap packages from CI. The first publish must be interactive and protected by 2FA.')
  process.exit(1)
}
if (!existsSync(resolve(root, 'LICENSE'))) {
  console.error('Refusing to publish without a root LICENSE file.')
  process.exit(1)
}

const version = tag.slice(1)
const manifests = new Map()
for (const entry of readdirSync(packageRoot, { withFileTypes: true })) {
  if (!entry.isDirectory()) continue
  const directory = resolve(packageRoot, entry.name)
  const path = resolve(directory, 'package.json')
  if (!existsSync(path)) continue
  const manifest = JSON.parse(readFileSync(path, 'utf8'))
  if (manifest.private === true) continue
  if (manifest.version !== version) {
    console.error(`${manifest.name}: version ${manifest.version} does not match ${tag}`)
    process.exit(1)
  }
  if (!manifest.license || manifest.license === 'UNLICENSED') {
    console.error(`${manifest.name}: a publishable SPDX license is required`)
    process.exit(1)
  }
  manifests.set(manifest.name, { directory, manifest })
}

const ordered = []
const visiting = new Set()
const visited = new Set()
const visit = (name) => {
  if (visited.has(name)) return
  if (visiting.has(name)) throw new Error(`workspace dependency cycle at ${name}`)
  visiting.add(name)
  const item = manifests.get(name)
  for (const dependency of Object.keys(item.manifest.dependencies ?? {})) {
    if (manifests.has(dependency)) visit(dependency)
  }
  visiting.delete(name)
  visited.add(name)
  ordered.push(item)
}
for (const name of manifests.keys()) visit(name)

const npmTag = version.includes('-') ? 'next' : 'latest'
const npmCache = mkdtempSync(join(tmpdir(), 'injoffice-release-cache-'))
const npm = process.platform === 'win32' ? 'npm.cmd' : 'npm'
const npmEnv = { ...process.env, npm_config_cache: npmCache }

const run = (args, options = {}) => {
  const result = spawnSync(npm, args, {
    cwd: root,
    encoding: 'utf8',
    stdio: 'inherit',
    env: npmEnv,
    ...options,
  })
  if (result.status !== 0) {
    const error = new Error(`npm ${args[0]} failed`)
    error.exitCode = result.status ?? 1
    throw error
  }
  return result
}

const sha512 = (path) => createHash('sha512').update(readFileSync(path)).digest('hex')
const sriFromHex = (hex) => `sha512-${Buffer.from(hex, 'hex').toString('base64')}`

try {
  if (packDestinationArg !== undefined) {
    const destination = resolve(root, packDestinationArg)
    mkdirSync(destination, { recursive: true })
    const releaseManifest = { schemaVersion: 1, tag, npmTag, packages: [] }
    for (const { directory, manifest } of ordered) {
      console.log(`Packing ${manifest.name}@${version}`)
      const result = run(['pack', directory, '--pack-destination', destination], { stdio: ['ignore', 'pipe', 'inherit'] })
      const filename = basename(result.stdout.trim().split(/\r?\n/).at(-1) ?? '')
      const tarball = resolve(destination, filename)
      if (!filename.endsWith('.tgz') || !existsSync(tarball)) {
        throw new Error(`${manifest.name}: npm pack did not produce an expected tarball`)
      }
      releaseManifest.packages.push({ name: manifest.name, version, filename, sha512: sha512(tarball) })
    }
    writeFileSync(resolve(destination, 'release-manifest.json'), `${JSON.stringify(releaseManifest, null, 2)}\n`)
    console.log(`Prepared ${releaseManifest.packages.length} release tarballs in ${destination}`)
  } else if (stageFromArg !== undefined || bootstrapFromArg !== undefined) {
    const bootstrap = bootstrapFromArg !== undefined
    const source = resolve(root, bootstrapFromArg ?? stageFromArg)
    const manifestPath = resolve(source, 'release-manifest.json')
    if (!existsSync(manifestPath)) {
      throw new Error(`Missing release manifest: ${manifestPath}`)
    }
    const releaseManifest = JSON.parse(readFileSync(manifestPath, 'utf8'))
    if (releaseManifest.schemaVersion !== 1 || releaseManifest.tag !== tag || releaseManifest.npmTag !== npmTag || !Array.isArray(releaseManifest.packages)) {
      throw new Error('Release manifest does not match the requested release.')
    }
    if (releaseManifest.packages.length !== ordered.length) {
      throw new Error('Release manifest package count does not match the publishable workspace set.')
    }
    const verified = []
    for (let index = 0; index < ordered.length; index += 1) {
      const expected = ordered[index].manifest
      const packed = releaseManifest.packages[index]
      if (!packed || packed.name !== expected.name || packed.version !== version || basename(packed.filename) !== packed.filename || !/^[a-f0-9]{128}$/.test(packed.sha512 ?? '')) {
        throw new Error(`Release manifest entry ${index} is invalid or out of order.`)
      }
      const tarball = resolve(source, packed.filename)
      if (!existsSync(tarball) || sha512(tarball) !== packed.sha512) {
        throw new Error(`${packed.name}: release tarball is missing or failed its SHA-512 check.`)
      }
      verified.push({ ...packed, tarball })
    }

    const registryState = []
    for (const packed of verified) {
      const published = await publishedIntegrity(packed.name, version)
      if (published !== null && published !== sriFromHex(packed.sha512)) {
        throw new Error(`${packed.name}@${version}: published registry integrity does not match the validated tarball.`)
      }
      registryState.push({ packed, published })
    }
    const alreadyPublished = registryState.filter((entry) => entry.published !== null)

    if (!bootstrap && alreadyPublished.length === registryState.length) {
      console.log(`All ${registryState.length} packages are already published with matching SHA-512 integrity.`)
    } else {
      if (!bootstrap && alreadyPublished.length > 0) {
        throw new Error(`Refusing a partial staged release: ${alreadyPublished.length} of ${registryState.length} package versions are already public.`)
      }
      for (const { packed, published } of registryState) {
        if (published !== null) {
          console.log(`Already bootstrapped ${packed.name}@${version}`)
          continue
        }
        console.log(`${bootstrap ? 'Bootstrapping' : 'Staging'} ${packed.name}@${version}`)
        run(bootstrap
          ? ['publish', packed.tarball, '--access', 'public', '--tag', npmTag, '--fetch-retries=0']
          : ['stage', 'publish', packed.tarball, '--access', 'public', '--tag', npmTag, '--fetch-retries=0'])
      }
    }
  } else {
    for (const { directory, manifest } of ordered) {
      const args = ['publish', directory, '--access', 'public', '--tag', npmTag]
      args.push('--dry-run')
      console.log(`Checking ${manifest.name}@${version}`)
      run(args)
    }
  }
} catch (error) {
  console.error(error instanceof Error ? error.message : error)
  process.exitCode = Number.isInteger(error?.exitCode) ? error.exitCode : 1
} finally {
  rmSync(npmCache, { recursive: true, force: true })
}

import { spawnSync } from 'node:child_process'
import { existsSync, mkdtempSync, readdirSync, readFileSync, rmSync, statSync } from 'node:fs'
import { basename, join, resolve } from 'node:path'
import { tmpdir } from 'node:os'
import { register } from 'node:module'
import { pathToFileURL } from 'node:url'
import { extensionlessDeclarationSpecifiers } from './fix-declaration-specifiers.mjs'

const root = resolve(import.meta.dirname, '..')
const packageRoot = resolve(root, 'packages')
register('./workspace-loader.mjs', import.meta.url, { data: { root } })
const packageDirs = readdirSync(packageRoot, { withFileTypes: true })
  .filter((entry) => entry.isDirectory() && existsSync(resolve(packageRoot, entry.name, 'package.json')))
  .map((entry) => resolve(packageRoot, entry.name))
  .sort()

const failures = []
const npmCache = mkdtempSync(join(tmpdir(), 'injoffice-npm-cache-'))
const mebibyte = 1024 * 1024
const maxPackedBytes = 3 * mebibyte
const maxUnpackedBytes = 10 * mebibyte
const maxTotalPackedBytes = 15 * mebibyte
let totalPackedBytes = 0
let totalUnpackedBytes = 0
const requireValue = (condition, message) => {
  if (!condition) failures.push(message)
}

for (const directory of packageDirs) {
  const manifest = JSON.parse(readFileSync(resolve(directory, 'package.json'), 'utf8'))
  const label = manifest.name ?? directory
  requireValue(/^@injoffice\/[a-z0-9-]+$/.test(manifest.name), `${label}: package name must use the @injoffice scope`)
  requireValue(/^\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?$/.test(manifest.version), `${label}: version must be valid SemVer`)
  requireValue(typeof manifest.description === 'string' && manifest.description.trim().length >= 20, `${label}: description is missing or too short`)
  requireValue(manifest.author === 'Injecting Inc.', `${label}: author must be Injecting Inc.`)
  requireValue(manifest.private !== true, `${label}: publishable package must not be private`)
  requireValue(manifest.license === 'Apache-2.0', `${label}: license must be Apache-2.0`)
  requireValue(manifest.type === 'module', `${label}: type must be module`)
  requireValue(manifest.engines?.node === '>=22', `${label}: engines.node must match the supported Node.js baseline (>=22)`)
  requireValue(manifest.main === './dist/index.js', `${label}: main must point at dist/index.js`)
  requireValue(manifest.types === './dist/index.d.ts', `${label}: types must point at dist/index.d.ts`)
  requireValue(manifest.exports?.['.']?.import === './dist/index.js', `${label}: missing ESM exports entry`)
  requireValue(manifest.exports?.['.']?.types === './dist/index.d.ts', `${label}: missing declaration exports entry`)
  requireValue(Array.isArray(manifest.files) && manifest.files.includes('dist') && manifest.files.includes('README.md'), `${label}: files must include dist and README.md`)
  requireValue(manifest.publishConfig?.access === 'public', `${label}: publishConfig.access must be public`)
  requireValue(manifest.repository?.url === 'git+https://github.com/injectinglabs/injoffice.git', `${label}: repository URL is missing`)
  requireValue(manifest.repository?.directory === `packages/${basename(directory)}`, `${label}: repository.directory does not identify the package`)
  requireValue(manifest.homepage === 'https://github.com/injectinglabs/injoffice#readme', `${label}: homepage is missing`)
  requireValue(manifest.bugs?.url === 'https://github.com/injectinglabs/injoffice/issues', `${label}: issue URL is missing`)
  for (const lifecycle of ['preinstall', 'install', 'postinstall']) {
    requireValue(manifest.scripts?.[lifecycle] === undefined, `${label}: published packages must not run ${lifecycle} scripts on consumer install`)
  }

  for (const [subpath, target] of Object.entries(manifest.exports ?? {})) {
    const conditions = typeof target === 'string' ? { default: target } : target
    for (const [condition, path] of Object.entries(conditions ?? {})) {
      requireValue(typeof path === 'string' && path.startsWith('./'), `${label}: export ${subpath} ${condition} target must be package-relative`)
      if (typeof path === 'string' && path.startsWith('./')) requireValue(existsSync(resolve(directory, path)), `${label}: export ${subpath} ${condition} target ${path} is missing`)
    }
  }

  const entry = resolve(directory, 'dist/index.js')
  const declarations = resolve(directory, 'dist/index.d.ts')
  requireValue(existsSync(entry), `${label}: build did not emit dist/index.js`)
  requireValue(existsSync(declarations), `${label}: build did not emit dist/index.d.ts`)
  for (const problem of extensionlessDeclarationSpecifiers(resolve(directory, 'dist'))) {
    requireValue(false, `${label}: ${problem.file.slice(directory.length + 1)} uses extensionless relative declaration import ${JSON.stringify(problem.specifier)}`)
  }
  if (existsSync(entry)) {
    try {
      await import(pathToFileURL(entry))
    } catch (error) {
      failures.push(`${label}: built entry cannot be imported: ${error instanceof Error ? error.message : String(error)}`)
    }
  }
  for (const entryName of Object.keys(manifest.injoffice?.entries ?? { index: 'src/index.ts' })) {
    const js = resolve(directory, `dist/${entryName}.js`)
    const dts = resolve(directory, `dist/${entryName}.d.ts`)
    requireValue(existsSync(js), `${label}: build did not emit dist/${entryName}.js`)
    requireValue(existsSync(dts), `${label}: build did not emit dist/${entryName}.d.ts`)
    if (entryName === 'browser' && existsSync(js)) {
      requireValue(!/from ["']node:/.test(readFileSync(js, 'utf8')), `${label}: browser entry imports a Node builtin`)
    }
  }
  for (const asset of manifest.injoffice?.assets ?? []) {
    requireValue(typeof asset === 'string' && /^[A-Za-z0-9._-]+$/.test(asset), `${label}: invalid packaged asset name ${JSON.stringify(asset)}`)
    const path = resolve(directory, 'dist', asset)
    requireValue(existsSync(path), `${label}: build did not emit dist/${asset}`)
    if (existsSync(path)) requireValue(statSync(path).size > 0, `${label}: packaged asset dist/${asset} is empty`)
  }
  if ((manifest.injoffice?.assets ?? []).some((asset) => asset.endsWith('.wasm')) && existsSync(entry)) {
    requireValue(!readFileSync(entry, 'utf8').includes('data:application/wasm;base64'), `${label}: JavaScript entry inlines a declared WASM asset`)
  }

  const packed = spawnSync(process.platform === 'win32' ? 'npm.cmd' : 'npm', ['pack', '--dry-run', '--json', directory], {
    cwd: root,
    encoding: 'utf8',
    env: { ...process.env, npm_config_cache: npmCache },
  })
  if (packed.status !== 0) {
    failures.push(`${label}: npm pack failed: ${packed.stderr.trim()}`)
    continue
  }
  try {
    const report = JSON.parse(packed.stdout)[0]
    requireValue(Number.isFinite(report.size), `${label}: npm pack report is missing compressed size`)
    requireValue(Number.isFinite(report.unpackedSize), `${label}: npm pack report is missing unpacked size`)
    if (Number.isFinite(report.size)) {
      totalPackedBytes += report.size
      requireValue(report.size <= maxPackedBytes, `${label}: compressed package size ${report.size} exceeds ${maxPackedBytes} bytes`)
    }
    if (Number.isFinite(report.unpackedSize)) {
      totalUnpackedBytes += report.unpackedSize
      requireValue(report.unpackedSize <= maxUnpackedBytes, `${label}: unpacked package size ${report.unpackedSize} exceeds ${maxUnpackedBytes} bytes`)
    }
    const files = new Set(report.files.map((file) => file.path))
    for (const required of ['package.json', 'README.md', 'dist/index.js', 'dist/index.d.ts']) {
      requireValue(files.has(required), `${label}: npm package is missing ${required}`)
    }
    requireValue(files.has('dist/LICENSE'), `${label}: npm package is missing the project license`)
    for (const legalFile of manifest.injoffice?.legalFiles ?? []) {
      requireValue(files.has(`dist/${legalFile.split('/').at(-1)}`), `${label}: npm package is missing legal file ${legalFile}`)
    }
    for (const entryName of Object.keys(manifest.injoffice?.entries ?? { index: 'src/index.ts' })) {
      requireValue(files.has(`dist/${entryName}.js`), `${label}: npm package is missing dist/${entryName}.js`)
      requireValue(files.has(`dist/${entryName}.d.ts`), `${label}: npm package is missing dist/${entryName}.d.ts`)
    }
    for (const asset of manifest.injoffice?.assets ?? []) {
      requireValue(files.has(`dist/${asset}`), `${label}: npm package is missing dist/${asset}`)
    }
    requireValue(![...files].some((file) => file.startsWith('src/')), `${label}: npm package leaks TypeScript source`)
    requireValue(![...files].some((file) => file.includes('.test.')), `${label}: npm package leaks test files`)
  } catch (error) {
    failures.push(`${label}: could not parse npm pack report: ${error instanceof Error ? error.message : String(error)}`)
  }
}

rmSync(npmCache, { recursive: true, force: true })
requireValue(totalPackedBytes <= maxTotalPackedBytes, `release set: compressed package size ${totalPackedBytes} exceeds ${maxTotalPackedBytes} bytes`)

if (failures.length > 0) {
  console.error(failures.map((failure) => `- ${failure}`).join('\n'))
  process.exit(1)
}

console.log(`Validated ${packageDirs.length} publishable packages (${(totalPackedBytes / mebibyte).toFixed(2)} MiB compressed, ${(totalUnpackedBytes / mebibyte).toFixed(2)} MiB unpacked).`)

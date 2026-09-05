import { spawnSync } from 'node:child_process'
import {
  copyFileSync,
  existsSync,
  lstatSync,
  mkdtempSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  realpathSync,
  rmSync,
  statSync,
  writeFileSync,
} from 'node:fs'
import { createServer } from 'node:http'
import { tmpdir } from 'node:os'
import { extname, resolve, sep } from 'node:path'

import { launchChromeForCDP, terminateProcess } from './chrome-cdp-startup.mjs'

const root = resolve(import.meta.dirname, '..')
const packageRoot = resolve(root, 'packages')
const base = '/injoffice-wasm-smoke/'
const temporary = mkdtempSync(resolve(tmpdir(), 'injoffice-wasm-tarball-browser-'))
const profiles = []
const dist = resolve(temporary, 'dist')
const browserRequests = []
const staticRequests = []
const pageErrors = []
const loadingFailures = []
let chrome
let cdp
let staticServer
let chromeDiagnostics = () => '(Chrome was not launched)'

const browserPackageNames = [
  '@injoffice/xlsx-wasm',
  '@injoffice/pptx-wasm',
  '@injoffice/docx-wasm',
]
const workspacePackages = discoverWorkspacePackages()
const localPackages = localDependencyClosure(browserPackageNames)

const fixtures = [
  ['apps/playground/public/native-corpus/pass-excel-defaults.xlsx', 'pass-excel-defaults.xlsx'],
  ['go/officecompat/corpus/generated/packages/docx-transitional-common.docx', 'docx-transitional-common.docx'],
  ['go/officecompat/corpus/generated/packages/pptx-transitional-common.pptx', 'pptx-transitional-common.pptx'],
]

const mime = new Map([
  ['.css', 'text/css; charset=utf-8'],
  ['.html', 'text/html; charset=utf-8'],
  ['.js', 'text/javascript; charset=utf-8'],
  ['.wasm', 'application/wasm'],
  ['.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'],
  ['.docx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'],
  ['.pptx', 'application/vnd.openxmlformats-officedocument.presentationml.presentation'],
])

try {
  console.log('Packing the local WASM package dependency closure...')
  const packedPackages = localPackages.map(packLocalPackage)
  writeFileSync(resolve(temporary, 'package.json'), JSON.stringify({
    name: 'injoffice-wasm-tarball-browser-smoke',
    private: true,
    type: 'module',
  }, null, 2))
  console.log('Installing the package tarballs into a clean consumer...')
  run('npm', [
    'install',
    '--ignore-scripts',
    '--no-package-lock',
    ...(process.env.INJOFFICE_SMOKE_CACHE ? ['--cache', process.env.INJOFFICE_SMOKE_CACHE] : []),
    ...packedPackages.map(({ tarball }) => tarball),
  ], temporary)
  validateInstalledTarballs(packedPackages)

  const publicFixtures = resolve(temporary, 'public', 'fixtures')
  mkdirSync(publicFixtures, { recursive: true })
  for (const [source, fixture] of fixtures) {
    copyFileSync(
      resolve(root, source),
      resolve(publicFixtures, fixture),
    )
  }
  copyFileSync(resolve(root, 'scripts/wasm-tarball-browser-consumer.mjs'), resolve(temporary, 'browser.mjs'))
  writeFileSync(resolve(temporary, 'index.html'), '<script type="module" src="/browser.mjs"></script>\n')
  console.log(`Building the installed consumer with Vite at ${base}...`)
  run(process.execPath, [
    resolve(root, 'node_modules/vite/bin/vite.js'),
    'build',
    '--base', base,
    '--outDir', dist,
  ], temporary)

  staticServer = createServer((request, response) => {
    const url = new URL(request.url ?? '/', 'http://127.0.0.1')
    staticRequests.push(url.pathname)
    if (!url.pathname.startsWith(base)) {
      response.writeHead(404).end('not found')
      return
    }
    const relative = url.pathname === base ? 'index.html' : decodeURIComponent(url.pathname.slice(base.length))
    const file = resolve(dist, relative)
    if (file !== dist && !file.startsWith(`${dist}${sep}`)) {
      response.writeHead(400).end('invalid path')
      return
    }
    try {
      if (!statSync(file).isFile()) throw new Error('not a file')
      response.writeHead(200, { 'Content-Type': mime.get(extname(file)) ?? 'application/octet-stream' })
      response.end(readFileSync(file))
    } catch {
      response.writeHead(404).end('not found')
    }
  })
  await listen(staticServer)
  const address = staticServer.address()
  if (address == null || typeof address === 'string') throw new Error('static server did not expose a TCP port')
  const origin = `http://127.0.0.1:${address.port}`

  console.log('Running the installed packages in headless Chrome...')
  const executable = findChrome()
  const startup = await launchChromeForCDP({
    executable,
    createProfile() {
      const profile = mkdtempSync(resolve(tmpdir(), 'injoffice-wasm-tarball-profile-'))
      profiles.push(profile)
      return profile
    },
    onAttemptFailure(failure, { attempt, attempts }) {
      console.warn(`Chrome startup attempt ${attempt}/${attempts} failed: ${failure.message}`)
    },
  })
  chrome = startup.child
  chromeDiagnostics = startup.diagnostics
  const target = startup.target
  cdp = await connectCDP(target.webSocketDebuggerUrl)
  cdp.on('Network.requestWillBeSent', ({ request }) => browserRequests.push(request.url))
  cdp.on('Runtime.exceptionThrown', ({ exceptionDetails }) => pageErrors.push(
    exceptionDetails.exception?.description ?? exceptionDetails.text,
  ))
  cdp.on('Network.loadingFailed', (event) => loadingFailures.push(`${event.errorText}: ${event.blockedReason ?? ''}`))
  await Promise.all([cdp.send('Network.enable'), cdp.send('Page.enable'), cdp.send('Runtime.enable')])
  await cdp.send('Page.navigate', { url: `${origin}${base}` })

  await pollExpression(
    cdp,
    `['passed', 'failed'].includes(globalThis.__injofficeWasmTarballSmoke?.state)`,
    180_000,
  )
  const result = await evaluate(cdp, 'globalThis.__injofficeWasmTarballSmoke')
  if (result?.state !== 'passed') throw new Error(`browser consumer failed: ${result?.error ?? JSON.stringify(result)}`)
  const formats = result.results?.map(({ format }) => format).join(',')
  if (formats !== 'xlsx,pptx,docx') throw new Error(`browser consumer returned unexpected formats: ${formats}`)
  const remoteRequests = browserRequests.filter((value) => {
    try {
      const url = new URL(value)
      return (url.protocol === 'http:' || url.protocol === 'https:') && url.origin !== origin
    } catch {
      return false
    }
  })
  if (remoteRequests.length > 0) throw new Error(`browser consumer made remote requests: ${remoteRequests.join(', ')}`)
  if (browserRequests.some((value) => {
    try { return new URL(value).pathname.startsWith('/v1/') } catch { return false }
  })) throw new Error('browser consumer made an InjOffice server request')
  if (pageErrors.length > 0) throw new Error(`browser page errors: ${pageErrors.join('; ')}`)
  if (loadingFailures.length > 0) throw new Error(`browser loading failures: ${loadingFailures.join('; ')}`)
  console.log('Installed-tarball browser smoke passed: XLSX, DOCX, and PPTX extract/apply/re-extract with zero server requests.')
} catch (error) {
  const detail = [
    error instanceof Error ? error.message : String(error),
    `Static requests: ${staticRequests.join(', ')}`,
    `Browser requests: ${browserRequests.join(', ')}`,
    `Page errors: ${pageErrors.join(', ')}`,
    `Loading failures: ${loadingFailures.join(', ')}`,
    `Chrome startup diagnostics: ${chromeDiagnostics()}`,
  ].join('\n')
  throw new Error(detail, { cause: error })
} finally {
  cdp?.close()
  if (chrome) await terminateProcess(chrome)
  if (staticServer) await closeServer(staticServer)
  for (const profile of profiles) removeTemporary(profile)
  removeTemporary(temporary)
}

function removeTemporary(directory) {
  try {
    rmSync(directory, { recursive: true, force: true, maxRetries: 10, retryDelay: 100 })
  } catch (error) {
    console.warn(`Could not remove temporary smoke directory ${directory}:`, error)
  }
}

function packLocalPackage(localPackage) {
  const { directory, manifest } = localPackage
  if (!existsSync(resolve(directory, 'dist'))) throw new Error(`${manifest.name}/dist is missing; run npm run build first`)
  validatePublishedLocalImports(localPackage)
  const report = JSON.parse(run('npm', ['pack', '--json', '--pack-destination', temporary, directory], root))[0]
  return { ...localPackage, tarball: resolve(temporary, report.filename) }
}

function discoverWorkspacePackages() {
  const packages = new Map()
  for (const entry of readdirSync(packageRoot, { withFileTypes: true })) {
    if (!entry.isDirectory()) continue
    const directory = resolve(packageRoot, entry.name)
    const manifestPath = resolve(directory, 'package.json')
    if (!existsSync(manifestPath)) continue
    const manifest = JSON.parse(readFileSync(manifestPath, 'utf8'))
    if (typeof manifest.name === 'string') packages.set(manifest.name, { directory, manifest })
  }
  return packages
}

function localDependencyClosure(roots) {
  const result = []
  const seen = new Set()
  const pending = [...roots]
  while (pending.length > 0) {
    const name = pending.shift()
    if (seen.has(name)) continue
    const localPackage = workspacePackages.get(name)
    if (!localPackage) throw new Error(`browser smoke root or dependency ${name} is not a workspace package`)
    seen.add(name)
    result.push(localPackage)
    for (const [dependency, version] of Object.entries(localPackage.manifest.dependencies ?? {})) {
      if (!dependency.startsWith('@injoffice/')) continue
      const target = workspacePackages.get(dependency)
      if (!target) throw new Error(`${name} declares missing local dependency ${dependency}`)
      if (version !== target.manifest.version) {
        throw new Error(`${name} declares ${dependency}@${version}; exact local version is ${target.manifest.version}`)
      }
      pending.push(dependency)
    }
  }
  return result
}

function validatePublishedLocalImports(localPackage) {
  const declared = new Set(Object.keys(localPackage.manifest.dependencies ?? {}))
  for (const file of listFiles(resolve(localPackage.directory, 'dist')).filter((path) => path.endsWith('.js'))) {
    const source = readFileSync(file, 'utf8')
    for (const specifier of moduleSpecifiers(source)) {
      const importedPackage = localImportPackageName(specifier)
      if (importedPackage && !declared.has(importedPackage)) {
        throw new Error(`${localPackage.manifest.name} publishes an import of ${specifier} without declaring ${importedPackage} in dependencies`)
      }
    }
  }
}

function moduleSpecifiers(source) {
  const result = []
  const patterns = [
    /\bfrom\s*['"]([^'"]+)['"]/g,
    /\bimport\s*['"]([^'"]+)['"]/g,
    /\bimport\s*\(\s*['"]([^'"]+)['"]\s*\)/g,
  ]
  for (const pattern of patterns) {
    for (const match of source.matchAll(pattern)) result.push(match[1])
  }
  return result
}

function localImportPackageName(specifier) {
  if (!specifier.startsWith('@injoffice/')) return undefined
  return specifier.split('/').slice(0, 2).join('/')
}

function validateInstalledTarballs(packedPackages) {
  const consumer = JSON.parse(readFileSync(resolve(temporary, 'package.json'), 'utf8'))
  for (const { manifest, tarball } of packedPackages) {
    const requested = consumer.dependencies?.[manifest.name]
    if (typeof requested !== 'string' || !requested.startsWith('file:')) {
      throw new Error(`clean consumer did not record ${manifest.name} as a file tarball dependency`)
    }
    if (realpathSync(resolve(temporary, requested.slice('file:'.length))) !== realpathSync(tarball)) {
      throw new Error(`clean consumer resolved ${manifest.name} from a different tarball`)
    }
    const installedDirectory = resolve(temporary, 'node_modules', ...manifest.name.split('/'))
    if (lstatSync(installedDirectory).isSymbolicLink()) throw new Error(`${manifest.name} was installed as a symlink instead of a tarball`)
    const installed = JSON.parse(readFileSync(resolve(installedDirectory, 'package.json'), 'utf8'))
    if (installed.name !== manifest.name || installed.version !== manifest.version) {
      throw new Error(`installed ${manifest.name} identity does not match its packed workspace manifest`)
    }
  }
}

function listFiles(directory) {
  return readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const path = resolve(directory, entry.name)
    return entry.isDirectory() ? listFiles(path) : [path]
  })
}

function run(command, args, cwd) {
  const result = spawnSync(command, args, { cwd, env: process.env, encoding: 'utf8' })
  if (result.status !== 0) throw new Error([result.stdout, result.stderr].filter(Boolean).join('\n'))
  return result.stdout
}

function findChrome() {
  const candidates = [
    process.env.CHROME_PATH,
    process.env.CHROME_BIN,
    '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
    '/Applications/Chromium.app/Contents/MacOS/Chromium',
    'google-chrome',
    'google-chrome-stable',
    'chromium',
    'chromium-browser',
  ].filter(Boolean)
  for (const candidate of candidates) {
    if (candidate.includes(sep) && !existsSync(candidate)) continue
    if (spawnSync(candidate, ['--version'], { stdio: 'ignore' }).status === 0) return candidate
  }
  throw new Error('Chrome/Chromium is required for the installed-tarball browser smoke; set CHROME_PATH explicitly')
}

function listen(server) {
  return new Promise((resolvePromise, reject) => {
    server.once('error', reject)
    server.listen(0, '127.0.0.1', resolvePromise)
  })
}

function closeServer(server) {
  return new Promise((resolvePromise) => server.close(resolvePromise))
}

async function connectCDP(url) {
  if (typeof WebSocket !== 'function') throw new Error('Node.js 22 or newer is required (global WebSocket is unavailable)')
  const socket = new WebSocket(url)
  await new Promise((resolvePromise, reject) => {
    socket.addEventListener('open', resolvePromise, { once: true })
    socket.addEventListener('error', reject, { once: true })
  })
  let nextId = 0
  const pending = new Map()
  const listeners = new Map()
  const rejectPending = (error) => {
    for (const request of pending.values()) request.reject(error)
    pending.clear()
  }
  socket.addEventListener('error', () => rejectPending(new Error('Chrome DevTools WebSocket failed')))
  socket.addEventListener('close', (event) => rejectPending(
    new Error(`Chrome DevTools WebSocket closed unexpectedly (code ${event.code})`),
  ))
  socket.addEventListener('message', (event) => {
    const message = JSON.parse(event.data)
    if (message.id !== undefined) {
      const request = pending.get(message.id)
      if (!request) return
      pending.delete(message.id)
      if (message.error) request.reject(new Error(message.error.message))
      else request.resolve(message.result)
      return
    }
    for (const listener of listeners.get(message.method) ?? []) listener(message.params)
  })
  return {
    send(method, params = {}) {
      const id = ++nextId
      return new Promise((resolvePromise, reject) => {
        if (socket.readyState !== WebSocket.OPEN) {
          reject(new Error(`Chrome DevTools WebSocket is not open for ${method}`))
          return
        }
        pending.set(id, { resolve: resolvePromise, reject })
        socket.send(JSON.stringify({ id, method, params }))
      })
    },
    on(method, listener) {
      const current = listeners.get(method) ?? []
      current.push(listener)
      listeners.set(method, current)
    },
    close() { socket.close() },
  }
}

async function evaluate(client, expression) {
  const response = await client.send('Runtime.evaluate', { expression, awaitPromise: true, returnByValue: true })
  if (response.exceptionDetails) throw new Error(response.exceptionDetails.exception?.description ?? response.exceptionDetails.text)
  return response.result.value
}

async function pollExpression(client, expression, timeout) {
  const deadline = Date.now() + timeout
  while (Date.now() < deadline) {
    if (await evaluate(client, expression)) return
    await delay(50)
  }
  throw new Error(`timed out waiting for browser consumer: ${JSON.stringify(await evaluate(client, 'globalThis.__injofficeWasmTarballSmoke'))}`)
}

function delay(milliseconds) {
  return new Promise((resolvePromise) => setTimeout(resolvePromise, milliseconds))
}

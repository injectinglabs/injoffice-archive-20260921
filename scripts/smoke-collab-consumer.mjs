import assert from 'node:assert/strict'
import { spawnSync } from 'node:child_process'
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'

const root = resolve(import.meta.dirname, '..')
const temporary = mkdtempSync(join(tmpdir(), 'injoffice-collab-consumer-'))
const npm = process.platform === 'win32' ? 'npm.cmd' : 'npm'
const env = { ...process.env, npm_config_cache: process.env.INJOFFICE_SMOKE_CACHE || join(temporary, 'npm-cache') }
const lock = JSON.parse(readFileSync(resolve(root, 'package-lock.json'), 'utf8'))
const version = name => lock.packages[`node_modules/${name}`].version
function run(command, args, cwd = temporary) {
  const result = spawnSync(command, args, { cwd, env, encoding: 'utf8', maxBuffer: 8 * 1024 * 1024 })
  assert.equal(result.status, 0, [result.error?.message, result.stdout, result.stderr].filter(Boolean).join('\n'))
  return result.stdout
}

try {
  const packed = JSON.parse(run(npm, ['pack', '--json', '--pack-destination', temporary, resolve(root, 'packages/collab')], root))[0]
  writeFileSync(join(temporary, 'package.json'), JSON.stringify({
    name: 'injoffice-collab-consumer', private: true, type: 'module',
    dependencies: { '@injoffice/collab': `file:./${packed.filename}`, react: version('react') },
    devDependencies: { '@types/react': version('@types/react'), typescript: version('typescript') },
  }, null, 2))
  run(npm, ['install', '--ignore-scripts', '--no-audit', '--no-fund'])
  const installed = JSON.parse(readFileSync(join(temporary, 'package-lock.json'), 'utf8'))
  const forbidden = Object.keys(installed.packages).filter(path => /(?:^|\/)node_modules\/(?:@univerjs\/|nanoid(?:\/|$))/.test(path))
  assert.deepEqual(forbidden, [], 'Collaboration-only consumers must not install Univer or NanoID')
  writeFileSync(join(temporary, 'smoke.mjs'), `
import assert from 'node:assert/strict'
import { PresenceManager, PresenceState, DocPresenceManager, PdfPresenceManager, DurableOutboundJournal, COLLAB_EVENTS } from '@injoffice/collab'
for (const value of [PresenceManager, PresenceState, DocPresenceManager, PdfPresenceManager, DurableOutboundJournal]) assert.equal(typeof value, 'function')
assert.ok(COLLAB_EVENTS.size > 0)
const state = new PresenceState()
assert.deepEqual(state.peers(), [])
`)
  run(process.execPath, ['smoke.mjs'])
  writeFileSync(join(temporary, 'smoke.mts'), `
import { PresenceManager, type CollabTransport, type SheetPresenceHost, type SheetPresenceCoordinates } from '@injoffice/collab'
export function attach(host: SheetPresenceHost, transport: CollabTransport): PresenceManager {
  return new PresenceManager(host, transport, { path: '/synthetic/sheet.xlsx', name: 'Consumer' })
}
export const range: SheetPresenceCoordinates = { startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 }
`)
  writeFileSync(join(temporary, 'tsconfig.json'), JSON.stringify({
    compilerOptions: { target: 'ES2022', module: 'NodeNext', moduleResolution: 'NodeNext',
      lib: ['ES2022', 'DOM'], strict: true, skipLibCheck: false, noEmit: true, types: ['react'] },
    files: ['smoke.mts'],
  }, null, 2))
  run(process.execPath, [join(temporary, 'node_modules/typescript/bin/tsc'), '-p', 'tsconfig.json'])
  const audit = JSON.parse(run(npm, ['audit', '--omit=dev', '--json']))
  assert.equal(audit.metadata.vulnerabilities.total, 0, 'The isolated production consumer must have no known advisories')
  console.log('Packed collaboration runtime and strict declarations pass without Univer, NanoID, or consumer overrides.')
  console.log('Isolated production dependency audit: zero known vulnerabilities.')
} finally {
  rmSync(temporary, { recursive: true, force: true })
}

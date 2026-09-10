import { spawn, spawnSync } from 'node:child_process'
import { createHash } from 'node:crypto'
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = resolve(import.meta.dirname, '..')
const commands = Object.freeze({
  'pdf-geometry': ['scripts/qualify-pdf-pixel-oracle.mjs'],
  'pdf-browser': ['scripts/smoke-demo-ux-browser.mjs', '--built'],
  'docx-native': ['scripts/smoke-docx-native-preview-browser.mjs'],
  'pptx-crop': ['scripts/smoke-pptx-crop-browser.mjs', '--styles'],
  'office-real-files': ['scripts/smoke-playground-xlsx-wasm-browser.mjs'],
})
const formats = { 'pdf-geometry': ['pdf'], 'pdf-browser': ['pdf'], 'docx-native': ['docx'], 'pptx-crop': ['pptx'], 'office-real-files': ['xlsx', 'docx', 'pptx'] }

export function validateRenderingCorpus(value) {
  if (!value || value.version !== 1 || value.license !== 'Apache-2.0' || value.externalOfficeReferenceCount !== 0 || !Array.isArray(value.cases)) throw new Error('Generated corpus requires version, license and explicit zero external Office references')
  const ids = new Set()
  for (const entry of value.cases) {
    if (!entry || !Object.hasOwn(commands, entry.id) || ids.has(entry.id) || !['analytical-oracle', 'contract-invariants'].includes(entry.referenceKind) || typeof entry.scope !== 'string' || !entry.scope.trim() || !Array.isArray(entry.formats) || !entry.formats.length || entry.formats.some(format => !['pdf', 'docx', 'pptx', 'xlsx'].includes(format))) throw new Error('Invalid or duplicate rendering case')
    if (JSON.stringify(entry.formats) !== JSON.stringify(formats[entry.id]) || entry.referenceKind !== (['pdf-geometry', 'pptx-crop'].includes(entry.id) ? 'analytical-oracle' : 'contract-invariants')) throw new Error('Case metadata must match the implemented format and oracle contract')
    ids.add(entry.id)
  }
  if (ids.size !== Object.keys(commands).length) throw new Error('Rendering corpus must retain every required format and case')
  return value
}

async function main() {
  const manifestBytes = readFileSync(resolve(root, 'qualification/rendering/manifest.json'))
  const manifest = validateRenderingCorpus(JSON.parse(manifestBytes))
  const args = process.argv.slice(2)
  if (args.length && !(args.length === 2 && args[0] === '--case' && Object.hasOwn(commands, args[1]))) throw new Error('Usage: node scripts/qualify-rendering-corpus.mjs [--case CASE_ID]')
  const selected = args.length ? manifest.cases.filter(entry => entry.id === args[1]) : manifest.cases
  if (selected.some(entry => ['pdf-browser', 'office-real-files'].includes(entry.id))) {
    let html = ''
    try { html = readFileSync(resolve(root, 'apps/playground/dist/index.html'), 'utf8') } catch { /* clear prerequisite error below */ }
    if (!html.includes('/injoffice-smoke/assets/')) throw new Error('First build workspace packages, then: npm run build:renderer -w apps/playground -- --base=/injoffice-smoke/')
  }
  const output = resolve(process.env.SHOWCASE_OUTPUT || resolve(root, 'artifacts/rendering-corpus'))
  mkdirSync(output, { recursive: true })
  writeFileSync(resolve(output, 'DEJAVU-FONTS-LICENSE.txt'), readFileSync(resolve(root, 'node_modules/dejavu-fonts-ttf/LICENSE')))
  const commit = spawnSync('git', ['rev-parse', 'HEAD'], { cwd: root, encoding: 'utf8', timeout: 10000 })
  const status = spawnSync('git', ['status', '--porcelain'], { cwd: root, encoding: 'utf8', timeout: 10000 })
  const report = { version: 1, passed: true, completeGeneratedSuite: false, externalOfficeReferenceCount: 0, officeFidelityQualified: false, manifestSha256: hash(manifestBytes), dependencyLockSha256: hash(readFileSync(resolve(root, 'package-lock.json'))), candidateCommit: commit.status === 0 ? commit.stdout.trim() : 'unknown', worktreeDirty: status.status === 0 ? status.stdout.trim().length > 0 : null, cases: [] }
  for (const entry of selected) {
    const directory = resolve(output, entry.id); mkdirSync(directory, { recursive: true })
    let log = '', failure
    try {
      log = await run(process.execPath, commands[entry.id], { SHOWCASE_OUTPUT: directory })
      if (entry.id === 'pdf-geometry') log += await run('go', ['run', './cmd/visualcheck', '-manifest', resolve(directory, 'manifest.json')], {}, resolve(root, 'go/officecompat'))
    } catch (error) { failure = error.message; report.passed = false; log += failure }
    writeFileSync(resolve(directory, 'runner.log'), log)
    report.cases.push({ ...entry, passed: !failure, evidenceDirectory: entry.id, runnerSha256: hash(readFileSync(resolve(root, commands[entry.id][0]))), ...(failure ? { error: failure } : {}) })
    report.completeGeneratedSuite = report.cases.length === manifest.cases.length
    writeFileSync(resolve(output, 'report.json'), JSON.stringify(report, null, 2) + '\n')
    console.log(`${entry.id}: ${failure ? 'FAIL' : 'PASS'} (${entry.referenceKind})`)
  }
  console.log(`Rendering evidence: ${resolve(output, 'report.json')}; external Office references: 0. This suite does not assert Office visual parity.`)
  if (!report.passed) process.exitCode = 1
}
function hash(bytes) { return createHash('sha256').update(bytes).digest('hex') }
function run(command, args, env = {}, cwd = root) {
  return new Promise((done, reject) => {
    const child = spawn(command, args, { cwd, env: { ...process.env, ...env }, detached: process.platform !== 'win32', stdio: ['ignore', 'pipe', 'pipe'] })
    let output = '', exceeded = false
    const stop = () => { try { if (process.platform !== 'win32' && child.pid) process.kill(-child.pid, 'SIGKILL'); else child.kill('SIGKILL') } catch (error) { if (error.code !== 'ESRCH') throw error } }
    const timer = setTimeout(() => { stop(); reject(new Error(`${command} exceeded 5 minute budget`)) }, 300000)
    const collect = chunk => { output += chunk; if (output.length > 8 * 1024 * 1024) { exceeded = true; stop(); output = output.slice(-65536) } }
    child.stdout.on('data', collect); child.stderr.on('data', collect)
    child.on('error', error => { clearTimeout(timer); reject(error) })
    child.on('exit', code => { clearTimeout(timer); if (code === 0 && !exceeded) done(output); else reject(new Error(`${command} ${args.join(' ')} exited ${code}${exceeded ? ' (output budget exceeded)' : ''}\n${output}`)) })
  })
}
if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) await main()

import assert from 'node:assert/strict'
import { spawnSync } from 'node:child_process'
import { copyFileSync, lstatSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { gzipSync } from 'node:zlib'
import { dirname, join, relative, resolve, sep } from 'node:path'

const root = resolve(import.meta.dirname, '..')
const binary = process.env.GITLEAKS_BIN || 'gitleaks'
const version = spawnSync(binary, ['version'], { encoding: 'utf8' })
assert.equal(version.status, 0, 'Install Gitleaks 8.30.1 or set GITLEAKS_BIN to its executable')
assert.equal(version.stdout.trim(), '8.30.1', 'Use the reviewed Gitleaks 8.30.1 rule set')
const temporary = mkdtempSync(join(tmpdir(), 'injoffice-secret-check-'))
const config = resolve(root, '.gitleaks.toml')
const scan = directory => spawnSync(binary, [
  'dir', directory, '--config', config, '--redact=100', '--no-banner', '--no-color',
  '--ignore-gitleaks-allow', '--gitleaks-ignore-path', join(temporary, 'no-ignore-file'),
  '--max-archive-depth', '2', '--max-decode-depth', '2', '--timeout', '120',
], { encoding: 'utf8', maxBuffer: 4 * 1024 * 1024 })

try {
  // Prove the fixture exception is bounded and inline suppression is disabled.
  const fontKey = Array.from({ length: 16 }, (_, index) => index.toString(16).repeat(2)).join('').toUpperCase()
  const selfTest = join(temporary, 'self-test')
  const fixture = join(selfTest, 'packages/docs/src/nativePagePaintCompilerV1.test.ts')
  mkdirSync(dirname(fixture), { recursive: true })
  writeFileSync(fixture, `const key = '${fontKey}'\n`)
  assert.equal(scan(selfTest).status, 0, 'The known synthetic font key must be allowed')
  writeFileSync(fixture, `const key = '${fontKey.slice(0, -1)}0'\n`)
  assert.equal(scan(selfTest).status, 1, 'A different key in the same file must be detected')
  writeFileSync(fixture, '')
  writeFileSync(join(selfTest, 'unrelated.ts'), `const key = '${fontKey}'\n`)
  assert.equal(scan(selfTest).status, 1, 'The fixture key must not be allowed in unrelated files')
  const syntheticToken = 'ghp_' + 'a1b2c3d4e5f6g7h8i9j0k1l2m3n4o5p6q7r8'
  writeFileSync(join(selfTest, 'unrelated.ts'), `const token = '${syntheticToken}' // gitleaks:allow\n`)
  const blocked = scan(selfTest)
  assert.equal(blocked.status, 1, 'Synthetic GitHub token must be detected despite inline allow comments')
  assert.ok(!`${blocked.stdout}${blocked.stderr}`.includes(syntheticToken), 'Scanner output must not expose a secret')
  writeFileSync(join(selfTest, 'unrelated.ts'), '')
  writeFileSync(join(selfTest, 'evidence.json.gz'), gzipSync(JSON.stringify({
    data: Buffer.from(`const token = '${syntheticToken}'`).toString('base64'),
  })))
  assert.equal(scan(selfTest).status, 1, 'A token in base64-encoded gzip evidence must be detected')

  // Scan current working bytes of indexed files, including staged additions.
  // Build outputs and ignored local credentials never enter this temporary copy.
  const listed = spawnSync('git', ['ls-files', '-z', '--cached'], { cwd: root, encoding: 'utf8' })
  assert.equal(listed.status, 0, 'Cannot inventory tracked files')
  const paths = [...new Set(listed.stdout.split('\0').filter(Boolean))]
  const source = join(temporary, 'source')
  mkdirSync(source)
  for (const path of paths) {
    const input = resolve(root, path)
    assert.ok(relative(root, input).split(sep).join('/') === path && !path.startsWith('../'), `Unsafe tracked path: ${path}`)
    assert.ok(lstatSync(input).isFile(), `Secret scan requires a regular tracked file: ${path}`)
    const output = join(source, path)
    mkdirSync(dirname(output), { recursive: true })
    copyFileSync(input, output)
  }
  const result = scan(source)
  process.stdout.write(result.stdout ?? '')
  process.stderr.write(result.stderr ?? '')
  assert.equal(result.status, 0, 'Tracked source contains a secret finding or the scanner could not complete')
  console.log(`Secret scan passed for ${paths.length} tracked files; allowlist and redaction controls passed.`)
} finally {
  rmSync(temporary, { recursive: true, force: true })
}

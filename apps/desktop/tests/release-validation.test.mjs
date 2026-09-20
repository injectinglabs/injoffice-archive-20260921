import { test } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { collect, preflight, REQUIRED_ASSETS, SIGNING_ENV, readDesktopVersion } from '../scripts/release-validation.mjs'

const version = readDesktopVersion()
const onTag = { GITHUB_REF_TYPE: 'tag', GITHUB_REF_NAME: `desktop-v${version}` }
const macSecrets = Object.fromEntries(Object.keys(SIGNING_ENV.mac).map(name => [name, 'set']))

test('preflight passes only on a matching desktop-v tag with every platform secret present', () => {
  assert.deepEqual(preflight('mac', { ...onTag, ...macSecrets }), [])
  assert.deepEqual(preflight('windows', { ...onTag, CSC_LINK: 'x', CSC_KEY_PASSWORD: 'y' }), [])
  assert.deepEqual(preflight('linux', onTag), [])
  const branch = preflight('linux', { GITHUB_REF_TYPE: 'branch', GITHUB_REF_NAME: 'main' })
  assert.equal(branch.length, 1)
  assert.match(branch[0], /desktop-vX\.Y\.Z tag/)
  assert.match(preflight('linux', { GITHUB_REF_TYPE: 'tag', GITHUB_REF_NAME: 'v9.9.9' })[0], /desktop-vX\.Y\.Z tag/)
  assert.match(preflight('linux', { GITHUB_REF_TYPE: 'tag', GITHUB_REF_NAME: 'desktop-v99.0.0' })[0], new RegExp(`does not match apps/desktop/package.json version ${version.replace(/\./g, '\\.')}`))
  assert.match(preflight('other', onTag)[0], /unknown platform/)
})

test('preflight names each missing secret by its environment variable and repository secret without printing values', () => {
  const problems = preflight('mac', { ...onTag, CSC_LINK: 'secret-value', APPLE_TEAM_ID: '   ' })
  assert.equal(problems.length, 4)
  for (const [variable, secret] of Object.entries(SIGNING_ENV.mac)) {
    if (variable === 'CSC_LINK') continue
    assert.ok(problems.some(problem => problem.includes(`${variable} is empty`) && problem.includes(secret)), variable)
  }
  assert.ok(problems.every(problem => !problem.includes('secret-value')))
  assert.match(preflight('windows', onTag)[0], /WINDOWS_CSC_LINK/)
})

function assets(dir, { version: feedVersion = version, drop = [] } = {}) {
  const names = [`InjOffice-${version}-mac-arm64.dmg`, `InjOffice-${version}-mac-x64.dmg`, `InjOffice-${version}-mac-arm64.zip`, `InjOffice-${version}-mac-x64.zip`,
    `InjOffice-${version}-win-x64.exe`, `InjOffice-${version}-linux-x64.AppImage`, `InjOffice-${version}-linux-x64.deb`, `InjOffice-${version}-linux-x64.rpm`]
  const platform = { mac: names.slice(0, 4), windows: names.slice(4, 5), linux: names.slice(5) }
  for (const [name, files] of Object.entries(platform)) {
    const sub = path.join(dir, `desktop-release-${name}`)
    fs.mkdirSync(sub, { recursive: true })
    for (const file of files) if (!drop.includes(file)) fs.writeFileSync(path.join(sub, file), file)
    const feed = { mac: 'latest-mac.yml', windows: 'latest.yml', linux: 'latest-linux.yml' }[name]
    fs.writeFileSync(path.join(sub, feed), `version: ${feedVersion}\nfiles:\n${files.filter(file => !file.endsWith('.dmg') && !file.endsWith('.deb') && !file.endsWith('.rpm')).map(file => `  - url: ${file}\n    sha512: x\n`).join('')}`)
  }
  return dir
}

test('collect flattens per-platform downloads, requires the complete asset set and writes SHA256SUMS', () => {
  const dir = assets(fs.mkdtempSync(path.join(os.tmpdir(), 'injoffice-collect-')))
  const { problems, files } = collect(dir)
  assert.deepEqual(problems, [])
  assert.equal(files.length, 11)
  assert.equal(fs.readdirSync(dir).filter(name => fs.statSync(path.join(dir, name)).isDirectory()).length, 0)
  const sums = fs.readFileSync(path.join(dir, 'SHA256SUMS'), 'utf8').trim().split('\n')
  assert.equal(sums.length, 11)
  assert.match(sums[0], /^[0-9a-f]{64}  /)
  assert.equal(REQUIRED_ASSETS.length, 11)
})

test('collect refuses a missing installer, a feed pointing at an absent file and a version mismatch', () => {
  const missing = collect(assets(fs.mkdtempSync(path.join(os.tmpdir(), 'injoffice-collect-')), { drop: [`InjOffice-${version}-win-x64.exe`] })).problems
  assert.ok(missing.some(problem => /no file matches .*win-x64\\\.exe/.test(problem)), missing.join('\n'))
  assert.ok(missing.some(problem => /latest\.yml references InjOffice-.*-win-x64\.exe, which is not among the assets/.test(problem)), missing.join('\n'))
  const mismatch = collect(assets(fs.mkdtempSync(path.join(os.tmpdir(), 'injoffice-collect-')), { version: '0.0.1' })).problems
  assert.equal(mismatch.length, 3)
  for (const problem of mismatch) assert.match(problem, /version 0\.0\.1 !=/)
  assert.deepEqual(collect(path.join(os.tmpdir(), 'injoffice-absent')).problems.length, 1)
})

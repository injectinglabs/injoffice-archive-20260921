const { test } = require('node:test')
const assert = require('node:assert/strict')
const { DesktopUpdateProvider } = require('../electron/desktop-update-provider.cjs')

const feed = '<feed><entry><title>Npm release</title><link href="https://github.com/injectinglabs/injoffice/releases/tag/v99.0.0"/><content>Wrong release notes</content></entry></feed>'
const release = (version, extra = {}) => ({ tag_name: `desktop-v${version}`, draft: false, prerelease: false, name: 'Desktop release', body: 'Correct desktop notes', ...extra })
function setup(pages, version = '0.3.0') {
  const calls = []
  const updater = { allowPrerelease: false, currentVersion: { raw: '0.1.0' }, fullChangelog: false }
  const executor = { async request(options, token) {
    calls.push({ options, token })
    if (options.path.endsWith('.atom')) return feed
    if (options.hostname === 'api.github.com') {
      const page = Number(new URL(`https://${options.hostname}${options.path}`).searchParams.get('page'))
      return JSON.stringify(pages[page - 1] || [])
    }
    return `version: ${version}\nfiles:\n  - url: InjOffice.zip\n    sha512: test\n    size: 123\n`
  } }
  const provider = new DesktopUpdateProvider({ owner: 'untrusted', repo: 'wrong', token: 'ignored' }, updater, { executor, platform: 'darwin' })
  return { provider, updater, calls }
}

test('ignores npm, malformed tags, drafts and prereleases; chooses greatest semantic version', async () => {
  const { provider } = setup([[release('0.3.0'), release('0.2.0'), release('99.0.0', { draft: true }), release('98.0.0', { prerelease: true }), release('0.4.0-beta.1'), release('01.0.0'), { tag_name: 'v100.0.0' }]])
  assert.equal(await provider.getLatestTagName({}), 'desktop-v0.3.0')
})

test('paginates with cancellation token and cannot configure a different repository', async () => {
  const { provider, calls } = setup([Array.from({ length: 100 }, () => ({ tag_name: 'v1.0.0' })), [release('0.3.0')]])
  const token = { test: true }
  assert.equal(await provider.getLatestTagName(token), 'desktop-v0.3.0')
  assert.equal(calls.length, 2)
  for (const { options, token: actual } of calls) {
    assert.equal(options.hostname, 'api.github.com')
    assert.match(options.path, /^\/repos\/injectinglabs\/injoffice\/releases\?per_page=100&page=/)
    assert.equal(actual, token)
    assert.equal(options.headers.Authorization, undefined)
  }
})

test('limits discovery to ten pages', async () => {
  const pages = Array.from({ length: 11 }, () => Array.from({ length: 100 }, () => release('0.3.0')))
  const { provider, calls } = setup(pages)
  await provider.getLatestTagName({})
  assert.equal(calls.length, 10)
})

test('reports absence of stable desktop releases', async () => {
  const { provider } = setup([[{ tag_name: 'v1.0.0' }]])
  await assert.rejects(provider.getLatestTagName({}), { code: 'ERR_UPDATER_NO_PUBLISHED_VERSIONS' })
})

test('downloads metadata from selected desktop tag and uses its notes even when missing in feed', async () => {
  const { provider, calls } = setup([[release('0.3.0')]])
  const info = await provider.getLatestVersion()
  assert.equal(info.tag, 'desktop-v0.3.0')
  assert.equal(info.releaseNotes, 'Correct desktop notes')
  assert.equal(info.releaseName, 'Desktop release')
  assert.match(calls.at(-1).options.path, /\/desktop-v0\.3\.0\/latest-mac.yml$/)
  const [file] = provider.resolveFiles(info)
  assert.equal(file.url.href, 'https://github.com/injectinglabs/injoffice/releases/download/desktop-v0.3.0/InjOffice.zip')
})

test('rejects metadata belonging to another version', async () => {
  const { provider } = setup([[release('0.3.0')]], '99.0.0')
  await assert.rejects(provider.getLatestVersion(), { code: 'ERR_UPDATER_INVALID_UPDATE_INFO' })
})

test('refuses prerelease lookup and custom channels instead of selecting npm releases', async () => {
  const { provider, updater, calls } = setup([])
  updater.allowPrerelease = true
  await assert.rejects(provider.getLatestVersion(), { code: 'ERR_UPDATER_UNSUPPORTED_CHANNEL' })
  updater.allowPrerelease = false
  updater.channel = 'beta'
  await assert.rejects(provider.getLatestVersion(), { code: 'ERR_UPDATER_UNSUPPORTED_CHANNEL' })
  assert.equal(calls.length, 0)
})

test('bounds selected release text', async () => {
  const { provider } = setup([[release('0.3.0', { body: 'a'.repeat(30000), name: 'b'.repeat(300) })]])
  const info = await provider.getLatestVersion()
  assert.equal(info.releaseNotes.length, 20000)
  assert.equal(info.releaseName.length, 200)
})

test('invalid API responses and network errors fail closed', async () => {
  const { provider } = setup([])
  provider.httpRequest = async () => '{bad'
  await assert.rejects(provider.getLatestTagName({}), { code: 'ERR_UPDATER_INVALID_RELEASE_FEED' })
  provider.httpRequest = async () => '{"message":"rate limited"}'
  await assert.rejects(provider.getLatestTagName({}), { code: 'ERR_UPDATER_INVALID_RELEASE_FEED' })
  provider.httpRequest = async () => { throw new Error('cancelled') }
  await assert.rejects(provider.getLatestTagName({}), /cancelled/)
})

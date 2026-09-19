const { GitHubProvider } = require('electron-updater/out/providers/GitHubProvider')
const semver = require('semver')

const TAG = /^desktop-v(\d+\.\d+\.\d+)$/
const MAX_PAGES = 10
const PAGE_SIZE = 100

function providerError(message, code) {
  return Object.assign(new Error(message), { code })
}

// GitHub's /latest endpoint also returns npm releases in this monorepo.
// Keep electron-updater's download/signature implementation, but select only
// stable desktop tags. This provider is intentionally public and stable-only.
class DesktopUpdateProvider extends GitHubProvider {
  constructor(options, updater, runtimeOptions) {
    super({ provider: 'github', owner: 'injectinglabs', repo: 'injoffice', channel: 'latest' }, updater, runtimeOptions)
    this.desktopRelease = null
  }

  async getLatestTagName(cancellationToken) {
    this.desktopRelease = null
    for (let page = 1; page <= MAX_PAGES; page += 1) {
      const url = new URL(`https://api.github.com/repos/injectinglabs/injoffice/releases?per_page=${PAGE_SIZE}&page=${page}`)
      const raw = await this.httpRequest(url, { Accept: 'application/vnd.github+json', 'X-GitHub-Api-Version': '2022-11-28' }, cancellationToken)
      let releases
      try { releases = JSON.parse(raw) } catch {
        throw providerError('GitHub returned invalid desktop release data.', 'ERR_UPDATER_INVALID_RELEASE_FEED')
      }
      if (!Array.isArray(releases)) {
        throw providerError('GitHub did not return a release list.', 'ERR_UPDATER_INVALID_RELEASE_FEED')
      }
      for (const release of releases) {
        if (!release || release.draft || release.prerelease) continue
        const match = typeof release.tag_name === 'string' && TAG.exec(release.tag_name)
        if (!match || !semver.valid(match[1])) continue
        if (!this.desktopRelease || semver.gt(match[1], this.desktopRelease.version)) {
          this.desktopRelease = {
            tag: release.tag_name,
            version: match[1],
            name: typeof release.name === 'string' ? release.name.slice(0, 200) : release.tag_name,
            notes: typeof release.body === 'string' ? release.body.slice(0, 20000) : '',
          }
        }
      }
      if (releases.length < PAGE_SIZE) break
    }
    if (!this.desktopRelease) {
      throw providerError('No stable InjOffice desktop release was found in the latest 1,000 GitHub releases.', 'ERR_UPDATER_NO_PUBLISHED_VERSIONS')
    }
    return this.desktopRelease.tag
  }

  async getLatestVersion() {
    if (this.updater.allowPrerelease || (this.updater.channel && this.updater.channel !== 'latest')) {
      throw providerError('InjOffice currently supports only stable desktop updates.', 'ERR_UPDATER_UNSUPPORTED_CHANNEL')
    }
    this.desktopRelease = null
    const info = await super.getLatestVersion()
    const release = this.desktopRelease
    if (!release || info.tag !== release.tag || info.version !== release.version) {
      throw providerError('Desktop update metadata does not match its release tag.', 'ERR_UPDATER_INVALID_UPDATE_INFO')
    }
    // The Atom feed can omit an older desktop release amid frequent npm releases.
    // Always use the selected release's notes, never another feed entry's notes.
    return { ...info, releaseName: release.name, releaseNotes: release.notes }
  }
}

module.exports = { DesktopUpdateProvider, DesktopGitHubProvider: DesktopUpdateProvider }

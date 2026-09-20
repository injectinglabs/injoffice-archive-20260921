#!/usr/bin/env node
// Gates for the signed desktop release workflow (.github/workflows/desktop-release.yml).
//
//   preflight <mac|windows|linux>   the run is on a desktop-vX.Y.Z tag matching the desktop
//                                   version and every signing secret for the platform is set
//   collect <dir>                   the downloaded per-platform artifacts form one complete
//                                   release: installers for every shipped platform, update feeds
//                                   that reference only present files, then SHA256SUMS
//
// Secret values are never printed. The workflow itself is out of scope here; this only
// refuses clearly when its inputs are incomplete.
import { createHash } from 'node:crypto'
import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

export const TAG = /^desktop-v(\d+\.\d+\.\d+)$/
// Environment variables desktop-release.yml maps from repository secrets, per platform.
export const SIGNING_ENV = {
  mac: { CSC_LINK: 'MAC_CSC_LINK', CSC_KEY_PASSWORD: 'MAC_CSC_KEY_PASSWORD', APPLE_ID: 'APPLE_ID', APPLE_APP_SPECIFIC_PASSWORD: 'APPLE_APP_SPECIFIC_PASSWORD', APPLE_TEAM_ID: 'APPLE_TEAM_ID' },
  windows: { CSC_LINK: 'WINDOWS_CSC_LINK', CSC_KEY_PASSWORD: 'WINDOWS_CSC_KEY_PASSWORD' },
  linux: {},
}
// artifactName in electron-builder.release.cjs: ${productName}-${version}-${os}-${arch}.${ext}
export const REQUIRED_ASSETS = [
  /-mac-arm64\.dmg$/, /-mac-x64\.dmg$/, /-mac-arm64\.zip$/, /-mac-x64\.zip$/, /^latest-mac\.yml$/,
  /-win-x64\.exe$/, /^latest\.yml$/,
  /-linux-x64\.AppImage$/, /-linux-x64\.deb$/, /-linux-x64\.rpm$/, /^latest-linux\.yml$/,
]

const here = path.dirname(fileURLToPath(import.meta.url))
export const readDesktopVersion = () => JSON.parse(fs.readFileSync(path.join(here, '../package.json'), 'utf8')).version

export function preflight(platform, env = process.env, { desktopVersion = readDesktopVersion() } = {}) {
  const problems = []
  const secrets = SIGNING_ENV[platform]
  if (!secrets) return [`unknown platform "${platform}"; expected ${Object.keys(SIGNING_ENV).join(', ')}`]
  const tag = env.GITHUB_REF_TYPE === 'tag' ? TAG.exec(env.GITHUB_REF_NAME ?? '') : null
  if (!tag) {
    problems.push(`signed desktop releases build only from a desktop-vX.Y.Z tag; this run is on ${env.GITHUB_REF_TYPE ?? 'an unknown ref type'} "${env.GITHUB_REF_NAME ?? ''}". Tag the commit or dispatch the workflow against the tag.`)
  } else if (tag[1] !== desktopVersion) {
    problems.push(`tag ${env.GITHUB_REF_NAME} does not match apps/desktop/package.json version ${desktopVersion}; electron-updater compares the two and would refuse the update.`)
  }
  for (const [variable, secret] of Object.entries(secrets)) {
    if (!env[variable] || env[variable].trim() === '') problems.push(`${variable} is empty: the ${secret} secret is not available to this run (desktop-release environment). ${platform} signing cannot proceed.`)
  }
  return problems
}

const yamlScalars = (text, key) => [...text.matchAll(new RegExp(`^\\s*-?\\s*${key}:\\s*(.+?)\\s*$`, 'gm'))].map(match => match[1].replace(/^['"]|['"]$/g, ''))

export function collect(directory, { desktopVersion = readDesktopVersion() } = {}) {
  const problems = []
  if (!fs.existsSync(directory)) return { problems: [`${directory} does not exist`], files: [] }
  // download-artifact places each platform in its own subdirectory; the draft step uploads top-level files.
  for (const entry of fs.readdirSync(directory, { withFileTypes: true })) {
    if (!entry.isDirectory()) continue
    for (const name of fs.readdirSync(path.join(directory, entry.name))) {
      const target = path.join(directory, name)
      if (fs.existsSync(target)) problems.push(`${name} was produced by more than one platform job`)
      else fs.renameSync(path.join(directory, entry.name, name), target)
    }
    fs.rmSync(path.join(directory, entry.name), { recursive: true, force: true })
  }
  const files = fs.readdirSync(directory).filter(name => name !== 'SHA256SUMS' && fs.statSync(path.join(directory, name)).isFile()).sort()
  for (const pattern of REQUIRED_ASSETS) if (!files.some(name => pattern.test(name))) problems.push(`no file matches ${pattern}`)
  for (const feed of files.filter(name => /^latest.*\.yml$/.test(name))) {
    const text = fs.readFileSync(path.join(directory, feed), 'utf8')
    const version = yamlScalars(text, 'version')[0]
    if (version !== desktopVersion) problems.push(`${feed} version ${version} != apps/desktop/package.json ${desktopVersion}`)
    const urls = yamlScalars(text, 'url')
    if (urls.length === 0) problems.push(`${feed} lists no files`)
    for (const url of urls) if (!files.includes(url)) problems.push(`${feed} references ${url}, which is not among the assets`)
  }
  if (problems.length === 0) {
    const sums = files.map(name => `${createHash('sha256').update(fs.readFileSync(path.join(directory, name))).digest('hex')}  ${name}\n`).join('')
    fs.writeFileSync(path.join(directory, 'SHA256SUMS'), sums)
  }
  return { problems, files }
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const [command, argument] = process.argv.slice(2)
  let problems
  if (command === 'preflight' && argument) problems = preflight(argument)
  else if (command === 'collect' && argument) problems = collect(path.resolve(argument)).problems
  else {
    console.error('usage: release-validation.mjs preflight <mac|windows|linux> | collect <dir>')
    process.exit(2)
  }
  if (problems.length > 0) {
    console.error(`Release ${command} refused:\n  ${problems.join('\n  ')}`)
    process.exit(1)
  }
  console.log(`Release ${command} ok`)
}

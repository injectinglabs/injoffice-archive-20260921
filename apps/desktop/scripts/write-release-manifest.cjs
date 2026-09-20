#!/usr/bin/env node
// Record what a desktop build produced: installer names, sizes, SHA-256 digests
// and the source revision, as release/manifest.json and release/SHA256SUMS.
//   node write-release-manifest.cjs [releaseDir]
const crypto = require('node:crypto');
const { execFileSync } = require('node:child_process');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

const INSTALLER = /\.(?:dmg|zip|exe|AppImage|deb|rpm)$/;
const SCHEMA = 'injoffice.desktop.release-manifest/1';

function sourceRevision(env = process.env) {
  try { return execFileSync('git', ['rev-parse', 'HEAD'], { encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'] }).trim(); } catch { /* not a checkout */ }
  if (env.GITHUB_SHA) return env.GITHUB_SHA;
  throw new Error('source revision unknown: not a git checkout and GITHUB_SHA is unset');
}

function writeReleaseManifest(releaseDir, { env = process.env, revision = sourceRevision(env) } = {}) {
  if (!fs.existsSync(releaseDir)) throw new Error(`${releaseDir} does not exist`);
  const names = fs.readdirSync(releaseDir).filter(name => INSTALLER.test(name) && fs.statSync(path.join(releaseDir, name)).isFile()).sort();
  if (names.length === 0) throw new Error(`no installers (${INSTALLER}) in ${releaseDir}: ${fs.readdirSync(releaseDir).join(', ') || '(empty)'}`);
  const artifacts = names.map(name => {
    const bytes = fs.readFileSync(path.join(releaseDir, name));
    return { name, bytes: bytes.length, sha256: crypto.createHash('sha256').update(bytes).digest('hex') };
  });
  const manifest = {
    schema: SCHEMA,
    product: 'InjOffice desktop (unsigned developer preview unless stated by the release workflow)',
    revision,
    ref: env.GITHUB_REF_NAME ?? null,
    workflowRun: env.GITHUB_RUN_ID ? `${env.GITHUB_SERVER_URL ?? 'https://github.com'}/${env.GITHUB_REPOSITORY}/actions/runs/${env.GITHUB_RUN_ID}` : null,
    builder: { platform: process.platform, arch: os.arch(), node: process.version },
    artifacts,
  };
  fs.writeFileSync(path.join(releaseDir, 'manifest.json'), `${JSON.stringify(manifest, null, 2)}\n`);
  fs.writeFileSync(path.join(releaseDir, 'SHA256SUMS'), artifacts.map(item => `${item.sha256}  ${item.name}\n`).join(''));
  return manifest;
}

if (require.main === module) {
  const releaseDir = path.resolve(process.argv[2] ?? path.join(__dirname, '../release'));
  try {
    const manifest = writeReleaseManifest(releaseDir);
    console.log(`revision ${manifest.revision}`);
    for (const item of manifest.artifacts) console.log(`${item.sha256}  ${String(item.bytes).padStart(11)}  ${item.name}`);
  } catch (error) {
    console.error(error.message);
    process.exit(1);
  }
}

module.exports = { INSTALLER, SCHEMA, writeReleaseManifest, sourceRevision };

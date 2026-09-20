// Unit coverage for the packaged-app helpers with a synthetic release directory and
// asar, so PR CI exercises them without electron-builder. The real packages are
// checked by test:packaged and check-packaged-updater.cjs in the desktop workflows.
const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { findPackagedApps, readAsar } = require('../scripts/packaged-app.cjs');
const { checkPackagedUpdater } = require('../scripts/check-packaged-updater.cjs');
const { writeReleaseManifest, SCHEMA } = require('../scripts/write-release-manifest.cjs');

const desktop = require('../package.json');

function writeAsar(file, files) {
  const header = { files: {} };
  const chunks = [];
  let offset = 0;
  for (const [entryPath, content] of Object.entries(files)) {
    const bytes = Buffer.from(content);
    let node = header;
    const parts = entryPath.split('/');
    for (const part of parts.slice(0, -1)) node = (node.files[part] ??= { files: {} });
    node.files[parts.at(-1)] = { size: bytes.length, offset: String(offset) };
    chunks.push(bytes);
    offset += bytes.length;
  }
  const json = Buffer.from(JSON.stringify(header));
  const padding = (4 - (json.length % 4)) % 4;
  const head = Buffer.alloc(16);
  head.writeUInt32LE(4, 0);
  head.writeUInt32LE(8 + json.length + padding, 4);
  head.writeUInt32LE(4 + json.length + padding, 8);
  head.writeUInt32LE(json.length, 12);
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, Buffer.concat([head, json, Buffer.alloc(padding), ...chunks]));
}

function fakeRelease({ release = false, feedConfig = release, electronUpdater = desktop.dependencies['electron-updater'] } = {}) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'injoffice-release-'));
  const metadata = { name: '@injoffice/desktop', version: desktop.version, main: 'electron/main.cjs', dependencies: desktop.dependencies, ...(release ? { injofficeRelease: true } : {}) };
  const files = {
    'package.json': JSON.stringify(metadata),
    'electron/main.cjs': '', 'electron/updates.cjs': '', 'electron/desktop-update-provider.cjs': '',
    'node_modules/electron-updater/package.json': JSON.stringify({ name: 'electron-updater', version: electronUpdater }),
    'node_modules/semver/package.json': JSON.stringify({ name: 'semver', version: desktop.dependencies.semver }),
    'renderer/index.html': '<script src="./assets/index-abc123.js"></script>',
  };
  writeAsar(path.join(dir, 'mac-arm64/InjOffice.app/Contents/Resources/app.asar'), files);
  writeAsar(path.join(dir, 'linux-unpacked/resources/app.asar'), files);
  fs.mkdirSync(path.join(dir, 'not-an-app'));
  if (feedConfig) {
    for (const resources of ['mac-arm64/InjOffice.app/Contents/Resources', 'linux-unpacked/resources']) fs.writeFileSync(path.join(dir, resources, 'app-update.yml'), 'provider: github\nowner: injectinglabs\nrepo: injoffice\n');
  }
  if (release) {
    fs.writeFileSync(path.join(dir, 'InjOffice-0.1.0-mac-arm64.zip'), 'zip');
    fs.writeFileSync(path.join(dir, 'InjOffice-0.1.0-linux-x64.AppImage'), 'appimage');
    fs.writeFileSync(path.join(dir, 'latest-mac.yml'), `version: ${desktop.version}\nfiles:\n  - url: InjOffice-0.1.0-mac-arm64.zip\n`);
    fs.writeFileSync(path.join(dir, 'latest-linux.yml'), `version: ${desktop.version}\nfiles:\n  - url: InjOffice-0.1.0-linux-x64.AppImage\n`);
  } else {
    fs.writeFileSync(path.join(dir, 'InjOffice-0.1.0-arm64.dmg'), 'dmg bytes');
    fs.writeFileSync(path.join(dir, 'InjOffice-0.1.0-arm64-mac.zip'), 'zip bytes');
    fs.writeFileSync(path.join(dir, 'builder-debug.yml'), 'ignored');
  }
  return dir;
}

test('findPackagedApps recognises the per-OS unpacked layouts and reads app.asar in place', () => {
  const dir = fakeRelease();
  const apps = findPackagedApps(dir);
  assert.deepEqual(apps.map(app => [app.platform, app.arch, app.name]), [['linux', 'x64', 'linux-unpacked'], ['mac', 'arm64', 'mac-arm64/InjOffice.app']]);
  const asar = readAsar(apps[1].asar);
  assert.equal(asar.has('electron/main.cjs'), true);
  assert.equal(JSON.parse(asar.read('package.json').toString()).main, 'electron/main.cjs');
  assert.equal(asar.size('renderer/index.html'), Buffer.byteLength('<script src="./assets/index-abc123.js"></script>'));
  assert.throws(() => asar.read('missing'), /not in/);
  asar.close();
  assert.deepEqual(findPackagedApps(path.join(dir, 'absent')), []);
});

test('unsigned previews bundle the updater modules but carry no update channel', () => {
  const ok = checkPackagedUpdater(fakeRelease());
  assert.deepEqual(ok.problems, []);
  assert.equal(ok.notes.length, 2);
  assert.match(ok.notes[0], /updater deliberately disabled/);
  const leaking = checkPackagedUpdater(fakeRelease({ feedConfig: true }));
  assert.ok(leaking.problems.every(problem => /must not ship .*app-update\.yml/.test(problem)), leaking.problems.join('\n'));
  assert.equal(leaking.problems.length, 2);
  const flagged = checkPackagedUpdater(fakeRelease({ release: true }));
  assert.ok(flagged.problems.some(problem => /must not carry injofficeRelease/.test(problem)));
  assert.ok(flagged.problems.some(problem => /produced update feeds latest-linux\.yml, latest-mac\.yml/.test(problem)));
  const stale = checkPackagedUpdater(fakeRelease({ electronUpdater: '0.0.1' }));
  assert.ok(stale.problems.every(problem => /packaged electron-updater@0\.0\.1/.test(problem)));
  assert.equal(checkPackagedUpdater(path.join(os.tmpdir(), 'injoffice-none')).problems[0].startsWith('no unpacked application'), true);
});

test('release builds need the flag, app-update.yml and feeds that reference present files', () => {
  assert.deepEqual(checkPackagedUpdater(fakeRelease({ release: true }), { release: true }).problems, []);
  const preview = checkPackagedUpdater(fakeRelease(), { release: true }).problems;
  assert.ok(preview.some(problem => /lacks injofficeRelease/.test(problem)));
  assert.ok(preview.some(problem => /app-update\.yml is missing/.test(problem)));
  assert.ok(preview.some(problem => /latest-mac\.yml was not produced/.test(problem)));
  const dir = fakeRelease({ release: true });
  fs.rmSync(path.join(dir, 'InjOffice-0.1.0-mac-arm64.zip'));
  assert.ok(checkPackagedUpdater(dir, { release: true }).problems.some(problem => /references missing InjOffice-0\.1\.0-mac-arm64\.zip/.test(problem)));
});

test('the release manifest records every installer with size, sha256 and the source revision', () => {
  const dir = fakeRelease();
  const manifest = writeReleaseManifest(dir, { revision: 'abc123', env: { GITHUB_REF_NAME: 'ci/x', GITHUB_RUN_ID: '7', GITHUB_REPOSITORY: 'injectinglabs/injoffice' } });
  assert.equal(manifest.schema, SCHEMA);
  assert.equal(manifest.revision, 'abc123');
  assert.equal(manifest.workflowRun, 'https://github.com/injectinglabs/injoffice/actions/runs/7');
  assert.deepEqual(manifest.artifacts.map(item => item.name), ['InjOffice-0.1.0-arm64-mac.zip', 'InjOffice-0.1.0-arm64.dmg']);
  assert.equal(manifest.artifacts[1].sha256, crypto.createHash('sha256').update('dmg bytes').digest('hex'));
  assert.equal(manifest.artifacts[1].bytes, 9);
  assert.deepEqual(JSON.parse(fs.readFileSync(path.join(dir, 'manifest.json'), 'utf8')), manifest);
  assert.equal(fs.readFileSync(path.join(dir, 'SHA256SUMS'), 'utf8'), manifest.artifacts.map(item => `${item.sha256}  ${item.name}\n`).join(''));
  assert.throws(() => writeReleaseManifest(fs.mkdtempSync(path.join(os.tmpdir(), 'injoffice-empty-')), { revision: 'x' }), /no installers/);
});

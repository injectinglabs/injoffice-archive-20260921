const test = require('node:test');
const assert = require('node:assert/strict');
const { EventEmitter } = require('node:events');
const fs = require('node:fs/promises');
const path = require('node:path');
const os = require('node:os');
const { createUpdateService, updateAvailability } = require('../electron/updates.cjs');

async function fixture(t, options = {}) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'injoffice-updates-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const settingsPath = path.join(directory, 'updates.json');
  if (options.settings) await fs.writeFile(settingsPath, options.settings);
  const updater = new EventEmitter();
  updater.nativeUpdater = new EventEmitter();
  let loads = 0, checks = 0, downloads = 0, installs = 0, cancellations = 0;
  updater.checkForUpdates = async () => { checks++; updater.emit('update-available', { version: '0.2.0', releaseNotes: '<b>Changes</b>' }); };
  updater.downloadUpdate = async () => { downloads++; updater.emit('download-progress', { percent: 27 }); updater.emit('update-downloaded', { version: '0.2.0' }); };
  updater.quitAndInstall = () => { installs++; };
  const tasks = new Map(); let next = 0;
  const timers = { setTimeout(fn, ms) { const id = ++next; tasks.set(id, { fn, ms }); return id; }, setInterval(fn, ms) { const id = ++next; tasks.set(id, { fn, ms }); return id; }, clearTimeout(id) { tasks.delete(id); }, clearInterval(id) { tasks.delete(id); } };
  const service = await createUpdateService({ appVersion: '0.1.0', settingsPath, timers, loadUpdater: () => { loads++; return updater; }, prepareInstall: fn => fn(), cancelInstall: () => cancellations++, ...options });
  t.after(() => service.dispose());
  return { service, updater, tasks, settingsPath, count: () => ({ loads, checks, downloads, installs, cancellations }) };
}

test('preview and package-manager installations never load updater or schedule network', async t => {
  assert.match(updateAvailability({ packaged: true, release: false, platform: 'darwin' }), /local preview/);
  assert.match(updateAvailability({ packaged: true, release: true, platform: 'linux', appImage: false }), /DEB\/RPM/);
  assert.equal(updateAvailability({ packaged: true, release: true, platform: 'linux', appImage: true }), null);
  const f = await fixture(t, { unavailable: 'Local preview' });
  f.service.start(); await f.service.check(); await f.service.download(); await f.service.install();
  assert.equal(f.service.getState().status, 'disabled'); assert.equal(f.tasks.size, 0); assert.equal(f.count().loads, 0);
});

test('automatic checks are delayed, periodic, persisted, and never automatically download/install', async t => {
  const f = await fixture(t); f.service.start();
  assert.deepEqual([...f.tasks.values()].map(t => t.ms), [15000, 21600000]);
  await f.service.check();
  assert.deepEqual(f.count(), { loads: 1, checks: 1, downloads: 0, installs: 0, cancellations: 0 });
  assert.equal(f.updater.autoDownload, false); assert.equal(f.updater.autoInstallOnAppQuit, false);
  assert.equal(f.updater.allowPrerelease, false); assert.equal(f.updater.allowDowngrade, false);
  assert.equal(f.updater.disableWebInstaller, true);
  await f.service.setAutomaticUpdates(false);
  assert.equal(f.tasks.size, 0); assert.deepEqual(JSON.parse(await fs.readFile(f.settingsPath, 'utf8')), { version: 1, autoCheck: false });
  await f.service.check(); assert.equal(f.count().checks, 2, 'manual checking still works');
  await assert.rejects(f.service.setAutomaticUpdates('false'), /Invalid/);
});

test('corrupt preferences fail closed, and preference writes remain ordered', async t => {
  const f = await fixture(t, { settings: '{broken' }); f.service.start();
  assert.equal(f.service.getState().autoCheck, false); assert.equal(f.tasks.size, 0);
  await Promise.all([f.service.setAutomaticUpdates(true), f.service.setAutomaticUpdates(false)]);
  assert.equal(f.service.getState().autoCheck, false); assert.equal(f.tasks.size, 0);
  assert.equal(JSON.parse(await fs.readFile(f.settingsPath, 'utf8')).autoCheck, false);
});

test('check and download calls serialize and a downloaded version cannot be replaced by a timer check', async t => {
  const f = await fixture(t); let finish;
  f.updater.checkForUpdates = () => new Promise(resolve => { finish = () => { f.updater.emit('update-available', { version: '0.2.0' }); resolve(); }; });
  const first = f.service.check();
  assert.equal((await f.service.check()).status, 'checking');
  assert.equal((await f.service.download()).status, 'checking');
  finish(); await first;
  await f.service.download(); assert.equal(f.service.getState().status, 'downloaded');
  await f.service.check(); assert.equal(f.service.getState().status, 'downloaded');
  await f.service.install(); assert.equal(f.service.getState().status, 'installing');
  await f.service.install(); assert.equal(f.count().installs, 1);
});

test('network and verification errors remain recoverable and redact server data', async t => {
  const f = await fixture(t);
  f.updater.checkForUpdates = async () => { throw new Error('secret URL/path/response'); };
  await f.service.check(); assert.equal(f.service.getState().status, 'error');
  assert.doesNotMatch(f.service.getState().message, /secret/);
  f.updater.checkForUpdates = async () => f.updater.emit('update-available', { version: '0.2.0' });
  await f.service.check();
  f.updater.downloadUpdate = async () => { f.updater.emit('error', new Error('signature failed')); throw Error('signature failed'); };
  await f.service.download(); assert.equal(f.service.getState().status, 'error');
  await f.service.install(); assert.equal(f.count().installs, 0);
});

test('workspace refusal never starts installer or cancels a separate normal close attempt', async t => {
  const f = await fixture(t, { prepareInstall: () => { throw Object.assign(Error('Save all files first.'), { code: 'UPDATE_WORKSPACE_UNSAFE' }); } });
  await f.service.check(); await f.service.download(); await f.service.install();
  assert.equal(f.service.getState().status, 'downloaded');
  assert.equal(f.service.getState().message, 'Save all files first.');
  assert.equal(f.count().installs, 0); assert.equal(f.count().cancellations, 0);
});

for (const synchronous of [true, false]) test(`${synchronous ? 'synchronous' : 'asynchronous'} Mac install error removes late quit callback before editing resumes`, async t => {
  const f = await fixture(t); let lateQuits = 0, permanentCalls = 0;
  f.updater.nativeUpdater.on('update-downloaded', () => permanentCalls++);
  f.updater.quitAndInstall = () => {
    f.updater.nativeUpdater.on('update-downloaded', () => lateQuits++);
    if (synchronous) f.updater.emit('error', Error('native install failed'));
  };
  await f.service.check(); await f.service.download(); await f.service.install();
  if (!synchronous) f.updater.emit('error', Error('native install failed'));
  assert.equal(f.service.getState().status, 'downloaded'); assert.ok(f.count().cancellations >= 1);
  f.updater.nativeUpdater.emit('update-downloaded');
  assert.equal(lateQuits, 0); assert.equal(permanentCalls, 1);
});

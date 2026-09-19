const fs = require('node:fs/promises');
const path = require('node:path');
const { atomicWrite } = require('./file-store.cjs');

// Local previews never contact the release service. Only the signed-release
// configuration embeds this marker; DEB/RPM installs remain package-managed.
function updateAvailability({ packaged, release, platform, appImage }) {
  if (!packaged || !release) return 'Updates are available in official release builds. This is a local preview.';
  if (platform === 'linux' && !appImage) return 'Install newer DEB/RPM packages from InjOffice releases, or use your configured package repository. In-app updates are available for AppImage installations.';
  if (!['darwin', 'win32', 'linux'].includes(platform)) return 'In-app updates are unavailable on this platform.';
  return null;
}

async function createUpdateService({ appVersion, settingsPath, unavailable, loadUpdater, notify = () => {}, prepareInstall, cancelInstall = () => {}, timers = { setTimeout, clearTimeout, setInterval, clearInterval } }) {
  let autoCheck = true, preferenceMessage;
  try {
    const info = await fs.stat(settingsPath);
    if (info.size > 1024) throw new Error('Invalid settings');
    const settings = JSON.parse(await fs.readFile(settingsPath, 'utf8'));
    if (settings.version !== 1 || typeof settings.autoCheck !== 'boolean') throw new Error('Invalid settings');
    autoCheck = settings.autoCheck;
  } catch (error) {
    if (error.code !== 'ENOENT') { autoCheck = false; preferenceMessage = 'Update preferences could not be read. Automatic checks are off; choose your preference again to save it.'; }
  }
  let state = { status: unavailable ? 'disabled' : 'idle', appVersion, autoCheck, ...(unavailable || preferenceMessage ? { message: unavailable || preferenceMessage } : {}) };
  let updater, timer, interval, operation, disposed = false, preferenceQueue = Promise.resolve();
  let nativeInstallListeners = [];
  const snapshot = () => ({ ...state });
  const emit = change => { state = { ...state, ...change }; if (!disposed) notify(snapshot()); return snapshot(); };
  const clearSchedule = () => { if (timer) timers.clearTimeout(timer); if (interval) timers.clearInterval(interval); timer = interval = undefined; };
  function restoreInstall() {
    // MacUpdater 6.8 installs a native callback during quitAndInstall. Remove
    // that callback after an error so a late native event cannot quit an editor
    // that has resumed work. Keep the updater's permanent listeners intact.
    for (const listener of nativeInstallListeners) updater?.nativeUpdater?.removeListener('update-downloaded', listener);
    nativeInstallListeners = [];
    cancelInstall();
  }
  function failure() {
    const installing = state.status === 'installing';
    if (installing) restoreInstall();
    return emit({ status: installing ? 'downloaded' : 'error', percent: undefined,
      message: installing ? 'The update could not be installed. Your workspace remains open. Try again later.' : 'Could not complete the update request. Check your connection and try again. Your current version is still available.' });
  }
  const details = info => ({ version: String(info.version).slice(0, 100), releaseNotes: typeof info.releaseNotes === 'string' ? info.releaseNotes.slice(0, 16000) : '' });
  function getUpdater() {
    if (updater) return updater;
    updater = loadUpdater();
    updater.autoDownload = false;
    updater.autoInstallOnAppQuit = false;
    updater.allowPrerelease = false;
    updater.allowDowngrade = false;
    updater.disableWebInstaller = true;
    updater.fullChangelog = false;
    // No request URLs, local paths, credentials, or server error bodies in UI.
    updater.logger = { info() {}, warn() {}, error() {}, debug() {} };
    updater.on('error', failure);
    updater.on('update-available', info => emit({ status: 'available', ...details(info), message: undefined, percent: undefined }));
    updater.on('update-not-available', () => emit({ status: 'not-available', version: undefined, releaseNotes: undefined, message: undefined }));
    updater.on('download-progress', progress => {
      if (state.status === 'downloading') emit({ percent: Math.max(0, Math.min(100, Number(progress.percent) || 0)) });
    });
    updater.on('update-downloaded', info => emit({ status: 'downloaded', ...details(info), percent: 100, message: undefined }));
    return updater;
  }
  async function check() {
    if (unavailable || disposed || operation || ['downloaded', 'installing'].includes(state.status)) return snapshot();
    operation = 'check';
    emit({ status: 'checking', message: undefined, version: undefined, releaseNotes: undefined, percent: undefined });
    try { await getUpdater().checkForUpdates(); }
    catch { if (state.status !== 'error') failure(); }
    finally { operation = undefined; }
    return snapshot();
  }
  async function download() {
    if (unavailable || disposed || operation || state.status !== 'available') return snapshot();
    operation = 'download';
    emit({ status: 'downloading', percent: 0, message: undefined });
    try { await getUpdater().downloadUpdate(); }
    catch { if (state.status !== 'error') failure(); }
    finally { operation = undefined; }
    return snapshot();
  }
  async function install() {
    if (unavailable || disposed || operation || state.status !== 'downloaded') return snapshot();
    operation = 'install';
    let started = false;
    try {
      await prepareInstall(() => {
        started = true;
        emit({ status: 'installing', message: 'Installing the update…' });
        const before = new Set(updater.nativeUpdater?.listeners('update-downloaded') ?? []);
        try { updater.quitAndInstall(false, true); }
        finally {
          nativeInstallListeners = (updater.nativeUpdater?.listeners('update-downloaded') ?? []).filter(listener => !before.has(listener));
          if (state.status !== 'installing') restoreInstall();
        }
      });
    } catch (error) {
      if (started) restoreInstall();
      emit({ status: 'downloaded', message: error?.code === 'UPDATE_WORKSPACE_UNSAFE' ? error.message : 'The update could not be installed. Your workspace remains open. Try again later.' });
    } finally { operation = undefined; }
    return snapshot();
  }
  function schedule() {
    clearSchedule();
    if (unavailable || disposed || !state.autoCheck) return;
    timer = timers.setTimeout(() => { timer = undefined; void check(); }, 15000);
    interval = timers.setInterval(() => void check(), 6 * 60 * 60 * 1000);
    timer?.unref?.(); interval?.unref?.();
  }
  function setAutomaticUpdates(enabled) {
    if (typeof enabled !== 'boolean') return Promise.reject(new Error('Invalid update preference.'));
    preferenceQueue = preferenceQueue.catch(() => {}).then(async () => {
      if (disposed || unavailable) return snapshot();
      await fs.mkdir(path.dirname(settingsPath), { recursive: true });
      await atomicWrite(settingsPath, Buffer.from(JSON.stringify({ version: 1, autoCheck: enabled }) + '\n'));
      emit({ autoCheck: enabled, message: undefined });
      schedule();
      return snapshot();
    });
    return preferenceQueue;
  }
  return { getState: snapshot, check, download, install, setAutomaticUpdates, start: schedule,
    dispose() { disposed = true; clearSchedule(); } };
}

function loadReleaseUpdater(platform) {
  const { MacUpdater, NsisUpdater, AppImageUpdater } = require('electron-updater');
  const { DesktopGitHubProvider } = require('./desktop-update-provider.cjs');
  const Updater = platform === 'darwin' ? MacUpdater : platform === 'win32' ? NsisUpdater : AppImageUpdater;
  return new Updater({ provider: 'custom', updateProvider: DesktopGitHubProvider, owner: 'injectinglabs', repo: 'injoffice' });
}

module.exports = { createUpdateService, updateAvailability, loadReleaseUpdater };

const { app, BrowserWindow, dialog, ipcMain, Menu, nativeTheme, net, protocol, session } = require('electron');
const path = require('node:path');
const fs = require('node:fs/promises');
const { pathToFileURL } = require('node:url');
const { FileStore, validateFormat, validateOpenFormat, normalizeSaveDestination, validateBytes, atomicWrite } = require('./file-store.cjs');
const {exportDocxPdf,cancelDocxPdf}=require('./docx-pdf-service.cjs');
const { replacePdfText } = require('./pdf-text-service.cjs');
const { createBlankDocument } = require('./new-document.cjs');
const { RecoveryStore } = require('./recovery-store.cjs');
const { RecentFiles, openRecentFile } = require('./recent-files.cjs');
const { resolveAsset, allowRequest, isTrustedSender } = require('./security.cjs');
const { createUpdateService, updateAvailability, loadReleaseUpdater } = require('./updates.cjs');

protocol.registerSchemesAsPrivileged([{ scheme: 'injoffice', privileges: { standard: true, secure: true, supportFetchAPI: true, corsEnabled: true, stream: true } }]);
const rendererRoot = path.resolve(__dirname, '../renderer');
const entryURL = 'injoffice://app/index.html';
const store = new FileStore();
let window;
let dirty = false;
let dialogBusy = false;
let editorBusy = false;
let recovery;
let pdfExportRequest;
let closeAttempt;
let closeSequence = 0;
let updates;
let prepareUpdateInstall;
let cancelUpdateInstall = () => {};
const checkpointErrors = new Map();
const pendingPaths = [];
function queueExternalPaths(filenames) {
  for (const filename of filenames) if (typeof filename === 'string' && path.isAbsolute(filename) && /\.(docx|docm|xlsx|pptx|pdf)$/i.test(filename) && !pendingPaths.includes(filename) && pendingPaths.length < 32) pendingPaths.push(filename);
  if (pendingPaths.length) menuAction('externalOpen');
}
const primaryInstance = app.requestSingleInstanceLock ? app.requestSingleInstanceLock() : true;
if (!primaryInstance) app.quit();
app.on('open-file', (event, filename) => { event.preventDefault(); queueExternalPaths([filename]); });
app.on('second-instance', (_event, argv) => {
  queueExternalPaths(argv);
  if (window && !window.isDestroyed()) { if (window.isMinimized()) window.restore(); window.focus(); }
});
if (primaryInstance) queueExternalPaths(process.argv.slice(app.isPackaged ? 1 : 2));
const filters = [{ name: 'Office documents', extensions: ['xlsx', 'docx', 'docm', 'pptx', 'pdf'] }];
const csp = "default-src 'self'; script-src 'self' 'wasm-unsafe-eval'; style-src 'self' 'unsafe-inline'; img-src 'self' data: blob:; font-src 'self' data:; worker-src 'self' blob:; connect-src 'self'; object-src 'none'; base-uri 'none'; frame-src 'none'";

const themePreferences = ['system', 'light', 'dark'];
// Matches --canvas in the renderer's styles.css so the window never flashes the other theme.
const canvasColor = () => (nativeTheme.shouldUseDarkColors ? '#1e1e1e' : '#e9edf2');

function trusted(event) {
  if (!isTrustedSender(event, window, entryURL)) {
    throw new Error('Untrusted document request.');
  }
}

function discardChanges(message) {
  if (!dirty) return true;
  return dialog.showMessageBoxSync(window, {
    type: 'warning', title: 'Unsaved changes', message,
    detail: 'Local recovery copies keep the last applied edits. Save first to update your original files. Supported pending text drafts are also kept in local recovery copies.',
    buttons: ['Keep editing', 'Close and keep recovery'], defaultId: 0, cancelId: 0, noLink: true,
  }) === 1;
}

async function withDialog(operation, editorOwned = false) {
  if (closeAttempt || pdfExportRequest || dialogBusy || (editorBusy && !editorOwned)) return null;
  dialogBusy = true;
  try { return await operation(); } finally { dialogBusy = false; }
}

function menuAction(action) { if (window && !window.isDestroyed()) window.webContents.send('document:menu', action); }

function installMenu() {
  const mac = process.platform === 'darwin';
  const template = [
    ...(mac ? [{ label: app.name, submenu: [{ role: 'about' }, { label: 'Check for Updates…', click: () => menuAction('updates') }, { type: 'separator' }, { role: 'hide' }, { role: 'hideOthers' }, { role: 'unhide' }, { type: 'separator' }, { role: 'quit' }] }] : []),
    { label: 'File', submenu: [
      { label: 'New…', accelerator: 'CmdOrCtrl+N', click: () => menuAction('new') },
      { label: 'Open…', accelerator: 'CmdOrCtrl+O', click: () => menuAction('open') },
      {label:'Import CSV / TSV…',click:()=>menuAction('importText')},
      { label: 'Save', accelerator: 'CmdOrCtrl+S', click: () => menuAction('save') },
      { label: 'Save As…', accelerator: 'CmdOrCtrl+Shift+S', click: () => menuAction('saveAs') },
      { type: 'separator' }, { role: mac ? 'close' : 'quit' },
    ] },
    { label: 'Edit', submenu: [{ label: 'Undo', accelerator: 'CmdOrCtrl+Z', click: () => menuAction('undo') }, { label: 'Redo', accelerator: mac ? 'Cmd+Shift+Z' : 'Ctrl+Y', click: () => menuAction('redo') }, { type: 'separator' }, { role: 'cut' }, { role: 'copy' }, { role: 'paste' }, { role: 'selectAll' }] },
    { label: 'View', submenu: [{ role: 'resetZoom' }, { role: 'zoomIn' }, { role: 'zoomOut' }, { type: 'separator' }, { role: 'togglefullscreen' }] },
    ...(!mac ? [{ label: 'Help', submenu: [{ label: 'Check for Updates…', click: () => menuAction('updates') }] }] : []),
  ];
  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

// macOS: Office keeps one title row. The system title bar is hidden and its traffic lights are
// inset into the renderer's own 40px .app-titlebar (see styles.css), which is a drag region.
// Windows and Linux keep the standard frame, where the app title bar sits below it.
const macTitleBar = process.platform === 'darwin' ? { titleBarStyle: 'hiddenInset', trafficLightPosition: { x: 13, y: 12 } } : {};

function createWindow() {
  dirty = false;
  editorBusy = false;
  window = new BrowserWindow({
    width: 1440, height: 960, minWidth: 900, minHeight: 640, title: 'InjOffice', icon: path.join(__dirname, 'icon.png'), backgroundColor: canvasColor(),
    ...macTitleBar,
    webPreferences: { preload: path.join(__dirname, 'preload.cjs'), contextIsolation: true, nodeIntegration: false, sandbox: true, webSecurity: true, spellcheck: false },
  });
  window.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
  window.webContents.on('will-navigate', (event) => event.preventDefault());
  window.webContents.on('will-attach-webview', (event) => event.preventDefault());
  const closingWindow = window;
  let closeApproved = false, drainingRecovery = false;
  function unsafeUpdate(message) { const error = new Error(message); error.code = 'UPDATE_WORKSPACE_UNSAFE'; return error; }
  cancelUpdateInstall = () => {
    closeApproved = false;
    drainingRecovery = false;
    const token = closeAttempt?.token;
    closeAttempt = undefined;
    if (!closingWindow.isDestroyed()) closingWindow.webContents.send('document:cancel-close', token);
  };
  prepareUpdateInstall = async install => {
    if (closeAttempt || drainingRecovery || pdfExportRequest || dialogBusy || editorBusy) throw unsafeUpdate('Finish the current operation before restarting to update.');
    if (dirty) throw unsafeUpdate('Save or close all unsaved files before restarting to update.');
    if (closingWindow.isDestroyed()) throw unsafeUpdate('Open the workspace before installing the update.');
    drainingRecovery = true;
    const token = ++closeSequence;
    let timer;
    try {
      const prepared = new Promise(resolve => {
        closeAttempt = { token, resolve };
        timer = setTimeout(() => resolve(false), 10000);
      });
      closingWindow.webContents.send('document:prepare-close', token);
      if (!await prepared) throw unsafeUpdate('The workspace is not ready to restart. Finish editing and try again.');
      if (dirty || editorBusy || dialogBusy || pdfExportRequest) throw unsafeUpdate('Save your files and finish the current operation before restarting to update.');
      await recovery.flush();
      if (dirty || editorBusy || dialogBusy || pdfExportRequest || checkpointErrors.size) throw unsafeUpdate('Save your files successfully before restarting. Recovery is not ready.');
      // Keep the renderer frozen and all file dialogs blocked until the updater
      // quits. Its error handler cancels this approval and restores editing.
      closeApproved = true;
      install();
    } catch (error) { cancelUpdateInstall(); throw error; }
    finally { clearTimeout(timer); }
  };
  window.on('close', (event) => {
    if (closeApproved) return;
    event.preventDefault();
    if (drainingRecovery || pdfExportRequest || dialogBusy || editorBusy || !discardChanges('Close the workspace without saving all files?')) return;
    drainingRecovery = true;
    const token = ++closeSequence;
    let timer;
    const prepared = new Promise(resolve => {
      closeAttempt = { token, resolve };
      // A missing or unresponsive renderer must never be treated as ready.
      timer = setTimeout(() => resolve(false), 10000);
    });
    closingWindow.webContents.send('document:prepare-close', token);
    void prepared.then(async ready => {
      if (!ready || editorBusy || dialogBusy) return;
      await recovery.flush();
      if (editorBusy || dialogBusy) return;
      if (checkpointErrors.size) {
        await dialog.showMessageBox(closingWindow, { type: 'error', title: 'Recovery could not be saved', message: 'Keep this workspace open and save your files before closing.', detail: [...checkpointErrors.values()].join('\n').slice(0, 2000), buttons: ['Keep editing'], noLink: true });
        return;
      }
      closeApproved = true;
      closingWindow.close();
    }).catch(error => { console.warn('Workspace close was canceled:', error.message); }).finally(() => {
      clearTimeout(timer);
      closeAttempt = undefined;
      drainingRecovery = false;
      if (!closeApproved && !closingWindow.isDestroyed()) closingWindow.webContents.send('document:cancel-close', token);
    });
  });
  window.on('closed', () => { if(pdfExportRequest)cancelDocxPdf(pdfExportRequest); store.releaseAll(); checkpointErrors.clear(); window = undefined; });
  window.loadURL(entryURL);
}

app.whenReady().then(async () => {
  if (!primaryInstance) return;
  recovery = new RecoveryStore(path.join(app.getPath('userData'), 'recovery'));
  const recentFiles = new RecentFiles(path.join(app.getPath('userData'), 'recent-files.json'));
  let release = false;
  try {
    const metadata = JSON.parse(await fs.readFile(path.join(app.getAppPath(), 'package.json'), 'utf8'));
    await fs.access(path.join(process.resourcesPath, 'app-update.yml'));
    release = metadata.injofficeRelease === true;
  } catch { /* Source builds and unsigned local previews have no update channel. */ }
  updates = await createUpdateService({
    appVersion: app.getVersion?.() ?? '0.1.0',
    settingsPath: path.join(app.getPath('userData'), 'updates.json'),
    unavailable: updateAvailability({ packaged: app.isPackaged, release, platform: process.platform, appImage: Boolean(process.env.APPIMAGE) }),
    loadUpdater: () => loadReleaseUpdater(process.platform),
    notify: state => { if (window && !window.isDestroyed()) window.webContents.send('updates:state', state); },
    prepareInstall: install => {
      if (!window || window.isDestroyed() || !prepareUpdateInstall) throw new Error('Open the workspace before installing the update.');
      return prepareUpdateInstall(install);
    },
    cancelInstall: () => cancelUpdateInstall(),
  });
  for (const [channel, action] of Object.entries({
    'updates:get': () => updates.getState(), 'updates:check': () => updates.check(),
    'updates:download': () => updates.download(), 'updates:install': () => updates.install(),
    'updates:automatic': enabled => updates.setAutomaticUpdates(enabled),
  })) ipcMain.handle(channel, (event, input) => { trusted(event); return action(input); });
  async function rememberDocument(document) {
    try { await recentFiles.record(store.get(document.id).filename); }
    catch (error) { console.warn('Recent file history could not be updated:', error.message); }
    return document;
  }
  session.defaultSession.setPermissionRequestHandler((_webContents, _permission, callback) => callback(false));
  session.defaultSession.setPermissionCheckHandler(() => false);
  session.defaultSession.on('will-download', (event) => event.preventDefault());
  session.defaultSession.webRequest.onBeforeRequest((details, callback) => {
    callback({ cancel: !allowRequest(rendererRoot, details.url) });
  });
  protocol.handle('injoffice', async (request) => {
    if (request.method !== 'GET') return new Response('Forbidden', { status: 403 });
    let filename;
    try { filename = resolveAsset(rendererRoot, request.url); }
    catch { return new Response('Bad request', { status: 400 }); }
    try {
      const result = await net.fetch(pathToFileURL(filename).href);
      const headers = new Headers(result.headers);
      headers.set('Content-Security-Policy', csp);
      headers.set('X-Content-Type-Options', 'nosniff');
      if (filename.endsWith('.wasm')) headers.set('Content-Type', 'application/wasm');
      return new Response(result.body, { status: result.status, headers });
    } catch { return new Response('Not found', { status: 404 }); }
  });
  ipcMain.on('document:close-ready', (event, input) => {
    if (isTrustedSender(event, window, entryURL) && input?.token === closeAttempt?.token) closeAttempt?.resolve(input.ready === true);
  });
  ipcMain.handle('document:text-history', (event, direction) => { trusted(event); if (direction !== 'undo' && direction !== 'redo') throw new Error('Invalid history action.'); if (!closeAttempt && !editorBusy && !dialogBusy) window.webContents[direction](); });
  ipcMain.handle('document:create', (event, format) => {
    trusted(event);
    return withDialog(async () => {
      validateFormat(format);
      const created = store.create(format, await createBlankDocument(format));
      dirty = true;
      window.setDocumentEdited(true);
      return created;
    });
  });
  ipcMain.handle('document:open', (event) => {
    trusted(event);
    return withDialog(async () => {
      const result = await dialog.showOpenDialog(window, { properties: ['openFile'], filters });
      if (result.canceled || !result.filePaths[0]) return null;
      return rememberDocument(await store.open(result.filePaths[0]));
    });
  });
  ipcMain.handle('document:pick-asset', (event, kind) => {
    trusted(event);
    if (kind !== 'pdf' && kind !== 'image') throw new Error('Choose a PDF or image.');
    return withDialog(async () => {
      const result = await dialog.showOpenDialog(window, { properties: ['openFile'], filters: [{ name: kind === 'pdf' ? 'PDF document' : 'Image', extensions: kind === 'pdf' ? ['pdf'] : ['png', 'jpg', 'jpeg'] }] });
      if (result.canceled || !result.filePaths[0]) return null;
      const filename = result.filePaths[0];
      const info = await fs.stat(filename);
      if (!info.isFile() || info.size > 32 * 1024 * 1024) throw new Error('Choose a file smaller than 32 MB.');
      const bytes = await fs.readFile(filename);
      if (bytes.length > 32 * 1024 * 1024) throw new Error('Choose a file smaller than 32 MB.');
      const png = bytes.subarray(0, 8).equals(Buffer.from([137,80,78,71,13,10,26,10]));
      const jpeg = bytes[0] === 255 && bytes[1] === 216 && bytes[2] === 255;
      if (kind === 'pdf' ? !bytes.subarray(0, 1024).includes(Buffer.from('%PDF-')) : !png && !jpeg) throw new Error('The selected file is not a supported PDF or PNG/JPEG image.');
      return { name: path.basename(filename), bytes };
    }, true);
  });
  ipcMain.handle('document:pick-delimited-import',event=>{
    trusted(event);
    return withDialog(async()=>{
      const result=await dialog.showOpenDialog(window,{properties:['openFile'],filters:[{name:'CSV or TSV text',extensions:['csv','tsv']}]});
      if(result.canceled||!result.filePaths[0])return null;
      const filename=result.filePaths[0],format=path.extname(filename).slice(1).toLowerCase();
      if(!['csv','tsv'].includes(format))throw new Error('Choose a CSV or TSV file.');
      const info=await fs.stat(filename);if(!info.isFile()||info.size>4*1024*1024)throw new Error('Text imports must be smaller than 4 MiB.');
      const bytes=await fs.readFile(filename);
      if(bytes.length>4*1024*1024||!Buffer.from(bytes.toString('utf8'),'utf8').equals(bytes))throw new Error('Choose a UTF-8 CSV or TSV file smaller than 4 MiB.');
      return {name:path.basename(filename),format,bytes,seed:await createBlankDocument('xlsx')};
    });
  });
  ipcMain.handle('document:export-docx-pdf',async(event,input)=>{
    trusted(event);
    if(closeAttempt||dialogBusy||pdfExportRequest)throw new Error('Wait for the current operation before exporting.');
    pdfExportRequest=input?.requestId;
    try{return await exportDocxPdf(input,{onProgress:stage=>{if(window&&!window.isDestroyed())window.webContents.send('document:pdf-export-progress',{requestId:input.requestId,stage})}})}
    finally{pdfExportRequest=undefined}
  });
  ipcMain.handle('document:cancel-docx-pdf',(event,requestId)=>{trusted(event);return cancelDocxPdf(requestId)});
  ipcMain.handle('document:replace-pdf-text' , (event, input) => { trusted(event); return replacePdfText(input); });
  ipcMain.handle('document:export-bytes', (event, input) => {
    trusted(event);
    if (!input || typeof input.name !== 'string' || input.name !== path.basename(input.name) || !['.pdf', '.svg', '.csv', '.tsv'].includes(path.extname(input.name).toLowerCase())) throw new Error('Export requires a PDF, SVG, CSV or TSV filename.');
    const bytes = validateBytes(input.bytes);
    const format = path.extname(input.name).slice(1).toLowerCase();
    if (format === 'pdf' && !bytes.subarray(0, 1024).includes(Buffer.from('%PDF-'))) throw new Error('Export requires PDF data.');
    if((format==='csv'||format==='tsv')&&(bytes.length>16*1024*1024||!Buffer.from(bytes.toString('utf8'),'utf8').equals(bytes)))throw new Error('Text exports require UTF-8 data no larger than 16 MiB.');
    if (format === 'svg') {
      // This validates the local adapter's export envelope, not arbitrary SVG safety.
      // Exported bytes are written only; the host never renders or executes them.
      if (bytes.length > 32 * 1024 * 1024) throw new Error('SVG exports must be smaller than 32 MB.');
      const svg = bytes.toString('utf8');
      if (!Buffer.from(svg, 'utf8').equals(bytes) || !svg.startsWith('<svg xmlns="http://www.w3.org/2000/svg" ') || !svg.trimEnd().endsWith('</svg>') || /<!|<\?|<(?:script|foreignObject)\b|\son[a-z]+\s*=/i.test(svg)) throw new Error('Export requires a self-contained SVG produced by the slide exporter.');
    }
    return withDialog(async () => {
      const result = await dialog.showSaveDialog(window, { defaultPath: input.name, filters: [{ name: format.toUpperCase(), extensions: [format] }] });
      if (result.canceled || !result.filePath) return null;
      const destination = normalizeSaveDestination(result.filePath, format);
      if (destination !== result.filePath) {
        let exists = false;
        try { await fs.stat(destination); exists = true; } catch (error) { if (error.code !== 'ENOENT') throw error; }
        if (exists) {
          const choice = await dialog.showMessageBox(window, { type: 'warning', title: `Replace existing ${format.toUpperCase()}?`, message: `“${path.basename(destination)}” already exists.`, buttons: ['Cancel', 'Replace'], defaultId: 0, cancelId: 0, noLink: true });
          if (choice.response !== 1) return null;
        }
      }
      await atomicWrite(destination, bytes);
      return { name: path.basename(destination) };
    }, true);
  });
  ipcMain.handle('document:import', (event, input) => {
    trusted(event);
    return withDialog(async () => {
      if (!input || typeof input.name !== 'string' || input.name !== path.basename(input.name)) throw new Error('Drop a supported document file.');
      const format = validateOpenFormat(path.extname(input.name).slice(1).toLowerCase());
      const created = store.create(format, input.bytes);
      store.get(created.id).name = input.name; created.name = input.name;
      return created;
    });
  });
  ipcMain.handle('document:checkpoint', async (event, input) => {
    trusted(event);
    const document = store.get(input?.id);
    try {
      const name = document.name || path.basename(document.filename);
      if (input.bytes === null) await recovery.writeDraft(input.id, name, input.revision, input.draft ?? null);
      else await recovery.write(input.id, name, input.bytes, input.draft ?? null, input.revision);
      checkpointErrors.delete(input.id);
    } catch (error) { checkpointErrors.set(input.id, error.message); throw error; }
  });
  ipcMain.handle('theme:set', (event, theme) => {
    trusted(event);
    if (!themePreferences.includes(theme)) throw new Error('Invalid theme preference.');
    nativeTheme.themeSource = theme;
    if (window && !window.isDestroyed()) window.setBackgroundColor(canvasColor());
    return nativeTheme.shouldUseDarkColors;
  });
  ipcMain.handle('document:recovery-list', event => { trusted(event); return recovery.list(); });
  ipcMain.handle('document:recover', (event, id) => {
    trusted(event);
    return withDialog(async () => {
      const entry = await recovery.read(id);
      const created = store.create(path.extname(entry.name).slice(1).toLowerCase(), entry.bytes);
      store.get(created.id).name = `Recovered ${entry.name}`;
      created.name = `Recovered ${entry.name}`;
      created.recoveryDraft = entry.draft;
      // Transfer durably before removing the old identity. Recovery always opens a copy.
      await recovery.write(created.id, created.name, created.bytes, entry.draft);
      await recovery.remove(id);
      return created;
    });
  });
  ipcMain.handle('document:discard-recovery', (event, id) => { trusted(event); if (closeAttempt) throw new Error('Workspace is closing.'); return recovery.remove(id); });
  ipcMain.handle('document:close', async (event, id) => { trusted(event); if (closeAttempt) throw new Error('Workspace is closing.'); store.get(id); await recovery.remove(id); checkpointErrors.delete(id); store.release(id); });
  ipcMain.handle('document:next-external', async event => {
    trusted(event);
    const opened = await withDialog(async () => {
      const filename = pendingPaths.shift();
      return filename ? rememberDocument(await store.open(filename)) : null;
    });
    // Busy editors/dialogs leave paths queued. Do not report that as an empty
    // queue: the renderer would stop draining and silently strand the last file.
    return opened ?? (pendingPaths.length ? { pending: true } : null);
  });
  ipcMain.handle('document:recent', (event) => { trusted(event); return recentFiles.list(); });
  ipcMain.handle('document:remove-recent', (event, id) => { trusted(event); return recentFiles.remove(id); });
  ipcMain.handle('document:open-recent', (event, id) => {
    trusted(event);
    return withDialog(async () => {
      return rememberDocument(await openRecentFile(recentFiles, store, id));
    });
  });
  ipcMain.handle('document:save', (event, input) => {
    trusted(event);
    return withDialog(async () => {
      if (!input || typeof input.saveAs !== 'boolean') throw new Error('Invalid save request.');
      const document = store.get(input.id);
      let destination;
      let overwrite = false;
      if (input.saveAs || document.untitled) {
        const suggestedPath = document.filename || document.name;
        const extension = path.extname(suggestedPath).slice(1);
        const result = await dialog.showSaveDialog(window, { defaultPath: suggestedPath, filters: [{ name: 'Document', extensions: [extension] }] });
        if (result.canceled || !result.filePath) return null;
        destination = normalizeSaveDestination(result.filePath, extension);
        // If we added the suffix, the native dialog may not have checked the
        // actual destination for overwrite. Confirm before replacing that file.
        if (destination !== result.filePath && destination !== document.filename) {
          let exists = false;
          try { await fs.stat(destination); exists = true; }
          catch (error) { if (error.code !== 'ENOENT') throw error; }
          if (exists) {
            const confirmation = await dialog.showMessageBox(window, { type: 'warning', title: 'Replace existing file?', message: `“${path.basename(destination)}” already exists.`, detail: 'Saving here will replace the existing file.', buttons: ['Cancel', 'Replace'], defaultId: 0, cancelId: 0, noLink: true });
            if (confirmation.response !== 1) return null;
          }
        }
      }
      if ((!destination || destination === document.filename) && await store.changed(input.id)) {
        const result = await dialog.showMessageBox(window, { type: 'warning', title: 'File changed', message: 'This file changed outside InjOffice.', detail: 'Overwriting will replace the version on disk. Choose Cancel and Save As to keep both versions.', buttons: ['Cancel', 'Overwrite'], defaultId: 0, cancelId: 0, noLink: true });
        if (result.response !== 1) return null;
        overwrite = true;
      }
      const saved = await rememberDocument(await store.save(input.id, input.bytes, destination, overwrite));
      await recovery.remove(input.id);
      checkpointErrors.delete(input.id);
      return saved;
    });
  });
  ipcMain.on('document:dirty', (event, value) => {
    if (!isTrustedSender(event, window, entryURL)) return;
    dirty = value === true;
    window.setDocumentEdited(dirty);
  });
  ipcMain.on('document:busy', (event, value) => {
    if (isTrustedSender(event, window, entryURL)) editorBusy = value === true;
  });
  // Packaged macOS apps keep the bundle icon shown during launch.
  if (process.platform === 'darwin' && !app.isPackaged) app.dock?.setIcon(path.join(__dirname, 'icon.png'));
  installMenu();
  createWindow();
  updates.start();
  app.on('will-quit', () => updates.dispose());
  app.on('activate', () => { if (BrowserWindow.getAllWindows().length === 0) createWindow(); });
});

app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit(); });

// Exercise the real main-process handlers with dialog/window adapters, without Electron or a GUI.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const path = require('node:path');
const os = require('node:os');
const vm = require('node:vm');
const { createRequire } = require('node:module');

async function harness(t, {pdfExporter, recoveryAdapter, installEffect} = {}) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'injoffice-creation-ipc-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const handlers = new Map();
  const listeners = new Map();
  const appEvents = new Map(), windowEvents = new Map();
  let updateOptions;
  let installCalls = 0;
  const updateCalls = [];
  const updateAdapter = {
    updateAvailability: () => null,
    loadReleaseUpdater: () => { throw Error('Tests must not load a real updater'); },
    createUpdateService: async options => {
      updateOptions = options;
      return {
        getState: () => { updateCalls.push('get'); return {status: 'downloaded'}; },
        check: () => { updateCalls.push('check'); },
        download: () => { updateCalls.push('download'); },
        setAutomaticUpdates: value => { updateCalls.push(['automatic', value]); },
        install: () => { updateCalls.push('install'); return options.prepareInstall(() => { installCalls++; installEffect?.(); }); },
        start() {}, dispose() {},
      };
    },
  };
  const nativeTheme = { themeSource: 'system', get shouldUseDarkColors() { return this.themeSource === 'dark'; } };
  const backgrounds = [];
  let closed = false;
  let prepareClose = () => true;
  let prepareToken;
  let canceledCloses = 0;
  let ready;
  let window;
  let saveResult = { canceled: true };
  let saveCalls = 0;
  let openResult={canceled:true,filePaths:[]};
  let discardResponse = 0;
  const noop = () => {};
  class Window {
    constructor() {
      window = this;
      this.webContents = { mainFrame: { url: 'injoffice://app/index.html' }, on: noop, send: (channel, token) => { if (channel === 'document:prepare-close') { prepareToken = token; Promise.resolve().then(() => prepareClose()).then(ready => { if (ready !== undefined) listeners.get('document:close-ready')({sender: this.webContents, senderFrame: this.webContents.mainFrame}, {token, ready}); }); } if (channel === 'document:cancel-close') canceledCloses++; }, setWindowOpenHandler: noop };
    }
    isDestroyed() { return false; }
    on(name, callback) { windowEvents.set(name, callback); }
    close() { const event = { prevented: false, preventDefault() { this.prevented = true; } }; windowEvents.get('close')(event); if (!event.prevented) closed = true; }
    loadURL() {}
    setDocumentEdited() {}
    setBackgroundColor(color) { backgrounds.push(color); }
  }
  const electron = {
    app: { name: 'InjOffice', isPackaged: true, getPath: () => directory, on: (name, callback) => appEvents.set(name, callback), whenReady: () => ({ then(callback) { ready = Promise.resolve().then(callback); return ready; } }) },
    BrowserWindow: Window,
    dialog: { showOpenDialog:async()=>openResult, showSaveDialog: async () => { saveCalls++; return saveResult; }, showMessageBoxSync: () => discardResponse, showMessageBox: async () => ({ response: 0 }) },
    ipcMain: { handle: (name, callback) => handlers.set(name, callback), on: (name, callback) => listeners.set(name, callback) },
    Menu: { buildFromTemplate: value => value, setApplicationMenu: noop },
    nativeTheme,
    protocol: { registerSchemesAsPrivileged: noop, handle: noop },
    session: { defaultSession: { setPermissionRequestHandler: noop, setPermissionCheckHandler: noop, on: noop, webRequest: { onBeforeRequest: noop } } },
  };
  const filename = path.resolve(__dirname, '../electron/main.cjs');
  const localRequire = createRequire(filename);
  vm.runInNewContext(await fs.readFile(filename, 'utf8'), {
    require: name => name === 'electron' ? electron : name === './updates.cjs' ? updateAdapter : name === './recovery-store.cjs' && recoveryAdapter ? recoveryAdapter : name === './docx-pdf-service.cjs' && pdfExporter ? pdfExporter : localRequire(name),
    __dirname: path.dirname(filename), process, console, Buffer, setTimeout, clearTimeout,
  }, { filename });
  await ready;
  const event = { sender: window.webContents, senderFrame: window.webContents.mainFrame };
  return {
    directory,
    installCalls: () => installCalls,
    updateCalls: () => updateCalls,
    cancelUpdate: () => updateOptions.cancelInstall(),
    untrustedUpdate: (name, value) => handlers.get(name)({}, value),
    setPrepareClose: callback => { prepareClose = callback; },
    ready: (ready, token = prepareToken) => listeners.get('document:close-ready')(event, {token, ready}),
    canceledCloses: () => canceledCloses,
    close: () => window.close(),
    closed: () => closed,
    emitApp: (name, ...args) => appEvents.get(name)(...args),
    invoke: (name, ...args) => handlers.get(name)(event, ...args),
    emit: (name, value) => listeners.get(name)(event, value),
    untrusted: () => handlers.get('document:create')({}, 'docx'),
    setOpenResult:value=>{openResult=value},
    setSaveResult: value => { saveResult = value; },
    setDiscardResponse: value => { discardResponse = value; },
    saveCalls: () => saveCalls,
    nativeTheme,
    backgrounds,
  };
}

test('first Save can be canceled, then saves the untitled file and adds history exactly once', async t => {
  const host = await harness(t);
  const document = await host.invoke('document:create', 'docx');
  assert.equal(document.untitled, true);
  assert(document.bytes.byteLength > 0);
  assert.deepEqual(await host.invoke('document:recent'), []);
  assert.deepEqual(await fs.readdir(host.directory), []);
  const input = { id: document.id, bytes: document.bytes, saveAs: false };
  assert.equal(await host.invoke('document:save', input), null);
  assert.deepEqual(await fs.readdir(host.directory), []);
  const destination = path.join(host.directory, 'My document.docx');
  host.setSaveResult({ canceled: false, filePath: destination });
  const saved = await host.invoke('document:save', input);
  assert.equal(saved.untitled, false);
  assert.equal(saved.name, 'My document.docx');
  assert.deepEqual(await fs.readFile(destination), Buffer.from(document.bytes));
  const history = await host.invoke('document:recent');
  assert.equal(history.length, 1);
  assert.equal(history[0].name, 'My document.docx');
  await host.invoke('document:save', input);
  assert.equal(host.saveCalls(), 2, 'subsequent Save must not open a dialog');
  assert.equal((await host.invoke('document:recent')).length, 1);
});

test('New respects trust and busy state while keeping existing unsaved sessions', async t => {
  const host = await harness(t);
  assert.throws(() => host.untrusted(), /Untrusted/);
  host.emit('document:busy', true);
  assert.equal(await host.invoke('document:create', 'xlsx'), null);
  host.emit('document:busy', false);
  host.emit('document:dirty', true);
  const first = await host.invoke('document:create', 'xlsx');
  assert.equal(first.untitled, true);
  host.setDiscardResponse(1);
  assert.equal((await host.invoke('document:create', 'xlsx')).untitled, true);
  assert.deepEqual(await fs.readdir(host.directory), []);
});

test('first Save enforces the format and protects an existing file reached by an added extension', async t => {
  const host = await harness(t);
  const document = await host.invoke('document:create', 'docx');
  const input = { id: document.id, bytes: document.bytes, saveAs: false };
  host.setSaveResult({ canceled: false, filePath: path.join(host.directory, 'Report.xlsx') });
  await assert.rejects(host.invoke('document:save', input), /DOCX|docx/);
  assert.deepEqual(await fs.readdir(host.directory), []);
  const existing = path.join(host.directory, 'Report.docx');
  await fs.writeFile(existing, 'keep this file');
  host.setSaveResult({ canceled: false, filePath: path.join(host.directory, 'Report') });
  assert.equal(await host.invoke('document:save', input), null);
  assert.equal(await fs.readFile(existing, 'utf8'), 'keep this file');
  assert.deepEqual(await host.invoke('document:recent'), []);
  host.setSaveResult({ canceled: false, filePath: path.join(host.directory, 'New report') });
  const saved = await host.invoke('document:save', input);
  assert.equal(saved.name, 'New report.docx');
  assert.deepEqual(await fs.readFile(path.join(host.directory, saved.name)), Buffer.from(document.bytes));
});

test('recovery checkpoint survives host restart and recovers an independent unsaved copy', async t => {
  const host = await harness(t);
  const original = await host.invoke('document:create', 'docx');
  const revision=require('node:crypto').randomUUID(),draft={version:1,format:'docx',text:'Pending draft'};
  await Promise.all([host.invoke('document:checkpoint',{id:original.id,bytes:original.bytes,revision}),host.invoke('document:checkpoint',{id:original.id,bytes:null,revision,draft})]);
  const entries = await host.invoke('document:recovery-list');
  assert.equal(entries.length, 1);
  const recovered = await host.invoke('document:recover', entries[0].id);
  assert.notEqual(recovered.id, original.id);
  assert.equal(recovered.untitled, true);
  assert.equal(recovered.name, 'Recovered Untitled.docx');
  assert.deepEqual(recovered.recoveryDraft,draft);
  assert.deepEqual(Buffer.from(recovered.bytes), Buffer.from(original.bytes));
  assert.equal((await host.invoke('document:recovery-list'))[0].id, recovered.id);
  await host.invoke('document:close', recovered.id);
  assert.deepEqual(await host.invoke('document:recovery-list'), []);
  await assert.rejects(host.invoke('document:checkpoint', { id: recovered.id, bytes: original.bytes }), /Unknown document/);
});

test('closing one session cannot invalidate another session or its saved bytes', async t => {
  const host = await harness(t);
  const first = await host.invoke('document:create', 'docx');
  const second = await host.invoke('document:create', 'xlsx');
  await host.invoke('document:close', first.id);
  host.setSaveResult({ canceled: false, filePath: path.join(host.directory, 'second.xlsx') });
  assert.equal((await host.invoke('document:save', { id: second.id, bytes: second.bytes, saveAs: false })).name, 'second.xlsx');
});

test('drop import creates an unsaved copy and PDF export honors cancellation and normalized overwrite protection', async t => {
  const host = await harness(t);
  const { createBlankDocument } = require('../electron/new-document.cjs');
  const bytes = await createBlankDocument('pdf');
  const imported = await host.invoke('document:import', { name: 'Dropped.pdf', bytes });
  assert.equal(imported.name, 'Dropped.pdf'); assert.equal(imported.untitled, true);
  assert.deepEqual(await host.invoke('document:recent'), []);
  assert.equal(await host.invoke('document:export-bytes', { name: 'Pages.pdf', bytes }), null);
  await assert.rejects(async () => host.invoke('document:export-bytes', { name: '../Pages.pdf', bytes }), /filename/);
  const destination = path.join(host.directory, 'Pages.pdf');
  await fs.writeFile(destination, 'keep original');
  host.setSaveResult({ canceled: false, filePath: destination.slice(0, -4) });
  assert.equal(await host.invoke('document:export-bytes', { name: 'Pages.pdf', bytes }), null);
  assert.equal(await fs.readFile(destination, 'utf8'), 'keep original');
  host.setSaveResult({ canceled: false, filePath: path.join(host.directory, 'New pages.pdf') });
  assert.equal((await host.invoke('document:export-bytes', { name: 'Pages.pdf', bytes })).name, 'New pages.pdf');
  assert.deepEqual(await fs.readFile(path.join(host.directory, 'New pages.pdf')), bytes);
});


test('OS open requests remain queued while an editor is busy and preserve source identity', async t => {
  const host = await harness(t);
  const { createBlankDocument } = require('../electron/new-document.cjs');
  const filename = path.join(host.directory, 'External.docx');
  const bytes = await createBlankDocument('docx'); await fs.writeFile(filename, bytes);
  let prevented = false;
  host.emit('document:busy', true);
  host.emitApp('open-file', { preventDefault() { prevented = true; } }, filename);
  host.emitApp('open-file', { preventDefault() {} }, filename);
  assert.equal(prevented, true);
  assert.equal(await host.invoke('document:next-external'), null);
  host.emit('document:busy', false);
  const opened = await host.invoke('document:next-external');
  assert.equal(opened.name, 'External.docx'); assert.equal(opened.untitled, undefined);
  assert.deepEqual(Buffer.from(opened.bytes), bytes);
  assert.equal(await host.invoke('document:next-external'), null, 'duplicate queued paths open once');
});

test('window close drains accepted recovery writes and refuses a failed checkpoint until saved', async t => {
  const host = await harness(t);
  const document = await host.invoke('document:create', 'docx');
  host.setDiscardResponse(1);
  const pending = host.invoke('document:checkpoint', { id: document.id, bytes: document.bytes });
  host.close(); assert.equal(host.closed(), false);
  await pending;
  await new Promise(resolve => setImmediate(resolve));
  assert.equal(host.closed(), true);
  assert.deepEqual((await host.invoke('document:recovery-list')).map(entry => entry.id), [document.id]);
  const failed = await harness(t);
  const other = await failed.invoke('document:create', 'docx'); failed.setDiscardResponse(1);
  await assert.rejects(failed.invoke('document:checkpoint', { id: other.id, bytes: Buffer.alloc(64 * 1024 * 1024 + 1) }), /64 MB/);
  failed.close(); await new Promise(resolve => setImmediate(resolve)); assert.equal(failed.closed(), false);
  failed.setSaveResult({ canceled: false, filePath: path.join(failed.directory, 'Saved.docx') });
  await failed.invoke('document:save', { id: other.id, bytes: other.bytes, saveAs: false });
  failed.close(); await new Promise(resolve => setImmediate(resolve)); assert.equal(failed.closed(), true);
});

test('SVG export validates local output and preserves cancellation and destination extension rules',async t=>{
  const host=await harness(t);
  const bytes=Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" width="720pt" height="540pt" viewBox="0 0 720 540"><rect width="720" height="540" fill="#ffffff"/></svg>');
  assert.equal(await host.invoke('document:export-bytes',{name:'Slide.svg',bytes}),null);
  for(const invalid of [Buffer.from('<svg></svg>'),Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" ><script/></svg>'),Buffer.concat([bytes.subarray(0,bytes.length-6),Buffer.from([255]),Buffer.from('</svg>')])]) {
    await assert.rejects(async()=>host.invoke('document:export-bytes',{name:'Slide.svg',bytes:invalid}),/SVG/);
  }
  await assert.rejects(async()=>host.invoke('document:export-bytes',{name:'../Slide.svg',bytes}),/filename/);
  const destination=path.join(host.directory,'Slide');
  host.setSaveResult({canceled:false,filePath:destination});
  assert.equal((await host.invoke('document:export-bytes',{name:'Slide.svg',bytes})).name,'Slide.svg');
  assert.deepEqual(await fs.readFile(destination+'.svg'),bytes);
  host.setSaveResult({canceled:false,filePath:path.join(host.directory,'Wrong.pdf')});
  await assert.rejects(async()=>host.invoke('document:export-bytes',{name:'Slide.svg',bytes}),/extension/);
  await assert.rejects(async()=>host.invoke('document:create','svg'),/DOCX/,'SVG is export-only');
});

test('editor-owned export dialog remains available while editor reports its export operation busy',async t=>{
  const host=await harness(t),{createBlankDocument}=require('../electron/new-document.cjs');
  host.emit('document:busy',true);
  host.setSaveResult({canceled:false,filePath:path.join(host.directory,'Export.pdf')});
  const bytes=await createBlankDocument('pdf');
  assert.equal((await host.invoke('document:export-bytes',{name:'Export.pdf',bytes})).name,'Export.pdf');
  assert.equal(host.saveCalls(),1);
  const assetPath=path.join(host.directory,'source.pdf');await fs.writeFile(assetPath,bytes);host.setOpenResult({canceled:false,filePaths:[assetPath]});
  const asset=await host.invoke('document:pick-asset','pdf');assert.equal(asset.name,'source.pdf');assert.deepEqual(asset.bytes,Buffer.from(bytes));
  assert.equal(await host.invoke('document:create','docx'),null,'unrelated new-file requests still respect editor busy state');
});

test('CSV/TSV picker returns a bounded UTF-8 source and independent XLSX seed; export stays separate',async t=>{
 const host=await harness(t),csv=Buffer.from('name,value\r\n"local, row",00123\r\n');
 const source=path.join(host.directory,'Data.csv');await fs.writeFile(source,csv);host.setOpenResult({canceled:false,filePaths:[source]});
 const picked=await host.invoke('document:pick-delimited-import');assert.equal(picked.format,'csv');assert.deepEqual(picked.bytes,csv);assert.equal(picked.seed[0],80);assert.equal(picked.seed[1],75);
 assert.deepEqual(await host.invoke('document:recent'),[],'picking does not open/register the original text file');
 host.setSaveResult({canceled:false,filePath:path.join(host.directory,'Export')});
 const exported=await host.invoke('document:export-bytes',{name:'Data.tsv',bytes:Buffer.from('name\tvalue\r\nlocal\t00123')});assert.equal(exported.name,'Export.tsv');assert.deepEqual(await fs.readFile(source),csv);
 await assert.rejects(async()=>host.invoke('document:export-bytes',{name:'Invalid.csv',bytes:Buffer.from([255])}),/UTF-8/);
 await fs.writeFile(source,Buffer.from([255]));await assert.rejects(async()=>host.invoke('document:pick-delimited-import'),/UTF-8/);
 host.setOpenResult({canceled:true,filePaths:[]});assert.equal(await host.invoke('document:pick-delimited-import'),null);
});


test('window close waits for renderer freeze and its late checkpoint, rejects stale acknowledgment, and resumes after refusal', async t => {
  const host = await harness(t);
  const document = await host.invoke('document:create', 'docx');
  host.setDiscardResponse(1);
  host.setPrepareClose(() => undefined);
  host.close();
  await new Promise(resolve => setImmediate(resolve));
  assert.equal(host.closed(), false, 'no acknowledgment cannot destroy the renderer');
  assert.equal(await host.invoke('document:create', 'docx'), null, 'new dialogs cannot race shutdown');
  host.ready(true, -1);
  await new Promise(resolve => setImmediate(resolve));
  assert.equal(host.closed(), false, 'stale acknowledgment is ignored');
  host.ready(false);
  await new Promise(resolve => setImmediate(resolve));
  assert.equal(host.canceledCloses(), 1);
  assert.equal(host.closed(), false);
  host.close();
  const late = host.invoke('document:checkpoint', {id: document.id, bytes: document.bytes, draft: {version: 1, format: 'docx', text: 'last edit'}});
  await late;
  assert.equal(host.closed(), false);
  host.ready(true);
  await new Promise(resolve => setImmediate(resolve));
  assert.equal(host.closed(), true);
  assert.equal((await host.invoke('document:recovery-list')).length, 1);
});


test('host PDF export blocks conflicting file actions and window close until canceled',async t=>{
 let rejectExport;const cancellations=[];
 const host=await harness(t,{pdfExporter:{exportDocxPdf:()=>new Promise((_resolve,reject)=>{rejectExport=reject}),cancelDocxPdf:requestId=>{cancellations.push(requestId);rejectExport(Error('PDF export canceled.'));return true}}});
 const document=await host.invoke('document:create','docx');host.setDiscardResponse(1);
 const pending=host.invoke('document:export-docx-pdf',{requestId:'export-one',bytes:document.bytes,revision:'unused-by-test-adapter'});
 assert.equal(await host.invoke('document:create','xlsx'),null);
 host.close();await new Promise(resolve=>setImmediate(resolve));assert.equal(host.closed(),false);
 assert.equal(await host.invoke('document:cancel-docx-pdf','export-one'),true);await assert.rejects(pending,/canceled/);assert.deepEqual(cancellations,['export-one']);
 assert.equal((await host.invoke('document:create','xlsx')).untitled,true,'canceling export releases the file-action lock');
});

const settle = () => new Promise(resolve => setImmediate(resolve));
const deferredUpdate = () => { let resolve; const promise = new Promise(done => { resolve = done; }); return {promise, resolve}; };

test('every update IPC rejects untrusted senders before invoking the update service', async t => {
  const host = await harness(t);
  for (const channel of ['updates:get', 'updates:check', 'updates:download', 'updates:install', 'updates:automatic']) {
    assert.throws(() => host.untrustedUpdate(channel, true), /Untrusted/);
  }
  assert.deepEqual(host.updateCalls(), []);
  assert.equal((await host.invoke('updates:get')).status, 'downloaded');
  await host.invoke('updates:check'); await host.invoke('updates:download'); await host.invoke('updates:automatic', false);
  assert.deepEqual(host.updateCalls(), ['get', 'check', 'download', ['automatic', false]]);
});

test('update restart never uses ordinary-close discard permission and refuses busy/export work', async t => {
  let rejectExport;
  const host = await harness(t, {pdfExporter: {exportDocxPdf: () => new Promise((_resolve, reject) => {rejectExport = reject;}), cancelDocxPdf: () => { rejectExport(Error('Canceled')); return true; }}});
  const document = await host.invoke('document:create', 'docx');
  host.emit('document:dirty', true); host.setDiscardResponse(1);
  await assert.rejects(host.invoke('updates:install'), /unsaved files/);
  assert.equal(host.installCalls(), 0);
  host.emit('document:dirty', false); host.emit('document:busy', true);
  await assert.rejects(host.invoke('updates:install'), /current operation/);
  host.emit('document:busy', false);
  const exportPromise = host.invoke('document:export-docx-pdf', {requestId: 'update-export', bytes: document.bytes, revision: 'test'});
  await assert.rejects(host.invoke('updates:install'), /current operation/);
  await host.invoke('document:cancel-docx-pdf', 'update-export'); await assert.rejects(exportPromise, /Canceled/);
  assert.equal(host.installCalls(), 0); assert.equal(host.closed(), false);
});

test('update freeze rejects stale or refused acknowledgment, and late dirtiness restores editing', async t => {
  const host = await harness(t); host.setPrepareClose(() => undefined);
  let pending = host.invoke('updates:install');
  host.ready(true, -1); await settle(); assert.equal(host.installCalls(), 0);
  assert.equal(await host.invoke('document:create', 'docx'), null);
  host.ready(false); await assert.rejects(pending, /not ready/);
  assert.equal(host.canceledCloses(), 1);
  assert.equal((await host.invoke('document:create', 'docx')).untitled, true);
  host.emit('document:dirty', false);
  pending = host.invoke('updates:install');
  host.emit('document:dirty', true); host.ready(true);
  await assert.rejects(pending, /Save your files/);
  assert.equal(host.installCalls(), 0); assert.equal(host.canceledCloses(), 2);
  host.close(); await settle(); assert.equal(host.closed(), false, 'canceled update cannot bypass normal dirty-close protection');
});

test('update waits for the recovery drain and rejects edits arriving during it', async t => {
  const entered = deferredUpdate(), release = deferredUpdate();
  const {RecoveryStore} = require('../electron/recovery-store.cjs');
  class DelayedRecovery extends RecoveryStore { async flush() { entered.resolve(); await release.promise; return super.flush(); } }
  const host = await harness(t, {recoveryAdapter: {RecoveryStore: DelayedRecovery}});
  const document = await host.invoke('document:create', 'docx');
  host.emit('document:dirty', false);
  const checkpoint = host.invoke('document:checkpoint', {id: document.id, bytes: document.bytes});
  const pending = host.invoke('updates:install');
  await entered.promise; assert.equal(host.installCalls(), 0);
  host.emit('document:dirty', true); release.resolve();
  await checkpoint; await assert.rejects(pending, /Recovery is not ready/);
  assert.equal(host.installCalls(), 0); assert.equal(host.canceledCloses(), 1);
});

test('failed recovery checkpoint refuses update until the document has been saved successfully', async t => {
  const host = await harness(t);
  const document = await host.invoke('document:create', 'docx');
  await assert.rejects(host.invoke('document:checkpoint', {id: document.id, bytes: Buffer.alloc(64 * 1024 * 1024 + 1)}), /64 MB/);
  host.emit('document:dirty', false);
  await assert.rejects(host.invoke('updates:install'), /Recovery is not ready/);
  assert.equal(host.installCalls(), 0);
  host.setSaveResult({canceled: false, filePath: path.join(host.directory, 'Saved.docx')});
  await host.invoke('document:save', {id: document.id, bytes: document.bytes, saveAs: false});
  await host.invoke('updates:install'); assert.equal(host.installCalls(), 1);
  host.cancelUpdate();
});

test('clean frozen update hands off exactly once and cancellation revokes close approval', async t => {
  const host = await harness(t); host.setPrepareClose(() => undefined);
  const pending = host.invoke('updates:install');
  await assert.rejects(host.invoke('updates:install'), /current operation/);
  host.ready(true); await pending;
  assert.equal(host.installCalls(), 1);
  assert.equal(await host.invoke('document:create', 'docx'), null, 'workspace remains frozen until installer takes over');
  host.cancelUpdate();
  assert.equal(host.canceledCloses(), 1);
  assert.equal((await host.invoke('document:create', 'docx')).untitled, true);
  host.emit('document:dirty', true); host.close(); await settle(); assert.equal(host.closed(), false);
  host.emit('document:dirty', false); host.setPrepareClose(() => true); host.close(); await settle();
  assert.equal(host.closed(), true, 'ordinary clean window close still works after update cancellation');
});

test('synchronous installer failure cancels the freeze and leaves normal close protection intact', async t => {
  const host = await harness(t, {installEffect: () => {throw Error('Installer unavailable');}});
  await assert.rejects(host.invoke('updates:install'), /Installer unavailable/);
  assert.equal(host.canceledCloses(), 1);
  assert.equal((await host.invoke('document:create', 'docx')).untitled, true);
  host.emit('document:dirty', true); host.close(); await settle(); assert.equal(host.closed(), false);
});

test('theme:set requires a trusted sender and a known preference before touching nativeTheme', async t => {
  const host = await harness(t);
  assert.throws(() => host.untrustedUpdate('theme:set', 'dark'), /Untrusted/);
  for (const theme of ['Dark', 'auto', '', null, undefined, {}]) assert.throws(() => host.invoke('theme:set', theme), /Invalid theme preference/);
  assert.equal(host.nativeTheme.themeSource, 'system', 'rejected requests leave the native theme untouched');
  assert.deepEqual(host.backgrounds, []);
  assert.equal(await host.invoke('theme:set', 'dark'), true);
  assert.equal(host.nativeTheme.themeSource, 'dark');
  assert.deepEqual(host.backgrounds, ['#1e1e1e'], 'the window background follows the dark canvas token');
  assert.equal(await host.invoke('theme:set', 'light'), false);
  assert.equal(await host.invoke('theme:set', 'system'), false);
  assert.equal(host.nativeTheme.themeSource, 'system');
  assert.deepEqual(host.backgrounds, ['#1e1e1e', '#e9edf2', '#e9edf2']);
});

const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('injDesktop', {
  getUpdateState: () => ipcRenderer.invoke('updates:get'),
  checkForUpdates: () => ipcRenderer.invoke('updates:check'),
  downloadUpdate: () => ipcRenderer.invoke('updates:download'),
  installUpdate: () => ipcRenderer.invoke('updates:install'),
  setAutomaticUpdates: enabled => ipcRenderer.invoke('updates:automatic', enabled),
  onUpdateState: callback => {
    const listener = (_event, state) => callback(state);
    ipcRenderer.on('updates:state', listener);
    return () => ipcRenderer.removeListener('updates:state', listener);
  },
  setTheme: theme => ipcRenderer.invoke('theme:set', theme),
  textHistory: direction => ipcRenderer.invoke('document:text-history', direction),
  nextExternal: () => ipcRenderer.invoke('document:next-external'),
  open: () => ipcRenderer.invoke('document:open'),
  create: (format) => ipcRenderer.invoke('document:create', format),
  pickDelimitedImport:()=>ipcRenderer.invoke('document:pick-delimited-import'),
  pickAsset: kind => ipcRenderer.invoke('document:pick-asset', kind),
  exportDocxPdf:input=>ipcRenderer.invoke('document:export-docx-pdf',input),
  cancelDocxPdf:requestId=>ipcRenderer.invoke('document:cancel-docx-pdf',requestId),
  onPdfExportProgress:callback=>{const listener=(_event,value)=>callback(value);ipcRenderer.on('document:pdf-export-progress',listener);return()=>ipcRenderer.removeListener('document:pdf-export-progress',listener)},
  replacePdfText: input => ipcRenderer.invoke('document:replace-pdf-text', input),
  exportBytes: input => ipcRenderer.invoke('document:export-bytes', input),
  importDocument: input => ipcRenderer.invoke('document:import', input),
  checkpoint: (id, bytes, draft = null, revision) => ipcRenderer.invoke('document:checkpoint', { id, bytes, draft, revision }),
  recovery: () => ipcRenderer.invoke('document:recovery-list'),
  recover: id => ipcRenderer.invoke('document:recover', id),
  discardRecovery: id => ipcRenderer.invoke('document:discard-recovery', id),
  close: id => ipcRenderer.invoke('document:close', id),
  recent: () => ipcRenderer.invoke('document:recent'),
  openRecent: (id) => ipcRenderer.invoke('document:open-recent', id),
  removeRecent: (id) => ipcRenderer.invoke('document:remove-recent', id),
  save: ({ id, bytes, saveAs = false }) => ipcRenderer.invoke('document:save', { id, bytes, saveAs }),
  setDirty: (dirty) => ipcRenderer.send('document:dirty', Boolean(dirty)),
  setBusy: (busy) => ipcRenderer.send('document:busy', Boolean(busy)),
  onPrepareClose: (prepare, cancel) => {
    const listener = async (_event, token) => {
      let ready = false;
      try { ready = await prepare(); } catch { /* Keep the workspace open. */ }
      ipcRenderer.send('document:close-ready', { token, ready: ready === true });
    };
    const canceled = () => cancel();
    ipcRenderer.on('document:prepare-close', listener);
    ipcRenderer.on('document:cancel-close', canceled);
    return () => { ipcRenderer.removeListener('document:prepare-close', listener); ipcRenderer.removeListener('document:cancel-close', canceled); };
  },
  onMenuAction: (callback) => {
    const listener = (_event, action) => callback(action);
    ipcRenderer.on('document:menu', listener);
    return () => ipcRenderer.removeListener('document:menu', listener);
  },
});

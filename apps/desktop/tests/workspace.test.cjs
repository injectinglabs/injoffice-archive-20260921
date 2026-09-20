const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const React = require('react');
const { create, act } = require('react-test-renderer');
global.IS_REACT_ACT_ENVIRONMENT = true;
async function loadApp({busyOnMount=false} = {}) {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({ input: path.resolve(__dirname, '../src/App.tsx'), platform: 'node', external: id => /^react(?:\/|$)/.test(id), transform: { jsx: { runtime: 'automatic' } }, plugins: [{ name: 'workspace-boundaries', resolveId(id) { if(id==='./UpdatesDialog')return '\0mock-updates'; if(id==='./spreadsheetDelimited')return '\0mock-delimited'; if (id.endsWith('.css')) return '\0css'; if (['./OfficeEditor', './PdfEditor', './PresentationEditor', './SpreadsheetEditor', './StartPage'].includes(id) || id.endsWith('.png')) return '\0mock:' + id; }, load(id) { if(id==='\0css')return 'export default ""'; if(id==='\0mock-updates')return 'export const UpdateNotice=()=>null; export default ()=>null;'; if(id==='\0mock-delimited')return 'export const importDelimitedWorkbook=(...args)=>globalThis.__importDelimited(...args)'; if (id.startsWith('\0mock:')) return id.endsWith('.png') ? 'export default "logo.png";' : id.endsWith('StartPage') ? `import React from 'react'; export default function StartPage(props) { return React.createElement('test-start', props); }` : `import React from 'react'; import { WorkspaceFileGroupsContext } from ${JSON.stringify(path.resolve(__dirname, '../src/Ribbon.tsx'))}; export function SpreadsheetEditor(props) { ${busyOnMount ? 'React.useEffect(()=>{props.onBusyChange?.(true)},[]);' : ''} const file = React.useContext(WorkspaceFileGroupsContext); return React.createElement('test-editor', props, React.createElement('test-file-tab', null, ...[...file.before, ...file.after].map(group => React.createElement('test-file-group', { key: group.id, label: group.label }, group.children)))); } export default SpreadsheetEditor;`; } }] });
  try { const { output } = await bundle.generate({ format: 'cjs', codeSplitting: false }); const result = { exports: {} }; new Function('require', 'module', 'exports', output[0].code)(require, result, result.exports); return result.exports.default ?? result.exports; } finally { await bundle.close(); }
}
test('workspace tabs retain independent bytes and editor instances; save and close target only their session', async () => {
  const App = await loadApp();
  const saves = [], checkpoints = [], closed = [], histories = [];
  let menu, next = 0;
  const preferences = new Map();
  const themes = [];
  global.window = { localStorage: { getItem: key => preferences.get(key) ?? null, setItem: (key,value) => preferences.set(key,value) }, document: { title: '' }, addEventListener() {}, removeEventListener() {}, injDesktop: {
    textHistory: async direction => histories.push(`text-${direction}`), nextExternal: async () => null, recent: async () => [], recovery: async () => [], setDirty() {}, setBusy() {}, setTheme: async theme => { themes.push(theme); return theme === 'dark'; },
    onMenuAction(callback) { menu = callback; return () => {}; },
    create: async format => ({ id: `id-${++next}`, name: `Untitled.${format}`, bytes: new Uint8Array([next]), untitled: true }),
    checkpoint: async (id, bytes, draft, revision) => checkpoints.push({ id, bytes: bytes ? [...bytes] : null, draft, revision }),
    save: async input => { saves.push({ ...input, bytes: [...input.bytes] }); return { id: input.id, name: `Saved-${input.id}.docx`, untitled: false }; },
    close: async id => closed.push(id),
  } };
  let renderer;
  await act(async () => { renderer = create(React.createElement(App)); });
  await act(async () => { renderer.root.findByType('test-start').props.onCreate('docx'); });
  const first = renderer.root.findByType('test-editor');
  await act(async () => renderer.root.findByProps({'aria-label':'Document zoom'}).props.onChange({target:{value:'125'}}));
  assert.equal(first.props.viewOptions.zoom,125);
  await act(async () => first.props.onChange(new Uint8Array([11])));
  await act(async () => { first.props.onDraftChange(true); first.props.onBusyChange(false); first.props.onRecoveryDraftChange({ version: 1, format: 'docx', text: 'pending' }); });
  assert.equal(checkpoints.at(-1).bytes,null,'draft-only updates do not resend native file bytes');
  assert.equal(typeof checkpoints.at(-1).revision,'string');
  first.props.registerHistory({undo:()=>histories.push('native-undo'),redo:()=>histories.push('native-redo')});
  await act(async () => menu('undo'));
  assert.deepEqual(histories,['text-undo']);
  await act(async () => first.props.onBusyChange(true));
  await act(async () => menu('save'));
  assert.equal(saves.length,0,'a native operation must block Save even while a draft exists');
  await act(async () => menu('undo'));
  assert.deepEqual(histories,['text-undo'],'native work must block text undo too');
  await act(async () => first.props.onBusyChange(false));
  first.props.registerCommit(async () => { first.props.onChange(new Uint8Array([33])); first.props.onDraftChange(false); first.props.onBusyChange(false); return true; });
  await act(async () => menu('new'));
  await act(async () => renderer.root.findByType('test-start').props.onCreate('xlsx'));
  let editors = renderer.root.findAllByType('test-editor');
  assert.equal(editors.length, 2);
  assert.equal(editors[0].props.viewOptions.zoom,125);
  assert.equal(editors[1].props.viewOptions.zoom,100);
  assert.equal(editors[0], first, 'first editor stays mounted with history intact');
  await act(async () => editors[1].props.onChange(new Uint8Array([22])));
  await act(async () => menu('save'));
  assert.deepEqual(saves[0].bytes, [22]); assert.equal(saves[0].id, 'id-2');
  const firstTab = renderer.root.findAllByType('button').find(button => button.props.title === 'Untitled.docx');
  await act(async () => firstTab.props.onClick());
  assert.equal(renderer.root.findByProps({'aria-label':'Document zoom'}).props.value,125);
  await act(async () => menu('save'));
  await act(async () => menu('undo'));
  assert.deepEqual(histories,['text-undo','native-undo']);
  assert.equal(saves[1].id, 'id-1'); assert.deepEqual(saves[1].bytes, [33], 'Save commits the selected tab draft before passing bytes to host');
  await act(async()=>first.props.onRecoveryDraftChange({version:1,format:'docx',text:'after save'}));
  assert.deepEqual(checkpoints.at(-1).bytes,[33],'the first draft after Save republishes native bytes because the old journal was removed');
  await act(async()=>first.props.onRecoveryDraftChange(null));
  assert(checkpoints.some(entry => entry.id === 'id-1' && entry.bytes?.[0] === 11));
  await act(async () => renderer.root.findByProps({ 'aria-label': 'Close Saved-id-1.docx' }).props.onClick());
  assert.deepEqual(closed, ['id-1']);
  assert.equal(renderer.root.findAllByType('test-editor').length, 1);
  assert.equal(renderer.root.findByType('test-editor').props.name, 'Untitled.xlsx');
  await act(async () => renderer.root.findByProps({ title: 'Search commands (Ctrl/⌘ K)' }).props.onClick());
  let search = renderer.root.findByProps({ 'aria-label': 'Search workspace commands' });
  await act(async () => search.props.onChange({target:{value:'preferences'}}));
  await act(async () => search.props.onKeyDown({key:'Enter',preventDefault(){}}));
  const [zoomSelect, themeSelect] = renderer.root.findAllByType('select');
  await act(async () => zoomSelect.props.onChange({target:{value:'150'}}));
  await act(async () => themeSelect.props.onChange({target:{value:'dark'}}));
  await act(async () => renderer.root.findAllByType('button').find(button => button.children.includes('Save preferences')).props.onClick());
  assert.equal(JSON.parse(preferences.get('injoffice.preferences.v1')).defaultZoom,150);
  assert.equal(JSON.parse(preferences.get('injoffice.preferences.v1')).theme,'dark');
  assert.deepEqual(themes,['system','dark'],'the host follows the saved appearance for native dialogs');
  await act(async () => renderer.root.findByProps({ title: 'Search commands (Ctrl/⌘ K)' }).props.onClick());
  search = renderer.root.findByProps({ 'aria-label': 'Search workspace commands' });
  await act(async () => search.props.onChange({target:{value:'blank pdf'}}));
  assert.equal(renderer.root.findAllByProps({role:'option'}).length,1);
  await act(async () => search.props.onKeyDown({key:'Enter',preventDefault(){}}));
  assert.equal(renderer.root.findAllByType('test-editor').length,2);
  assert.equal(renderer.root.findAllByType('test-editor')[1].props.name,'Untitled.pdf');
  assert.equal(renderer.root.findAllByType('test-editor')[1].props.viewOptions.zoom,150);
  assert.equal(renderer.root.findAllByType('test-editor')[0].props.viewOptions.zoom,100);
  assert.equal(renderer.root.findAllByType('dialog').length,0);
  await act(async () => renderer.unmount());
  delete global.window;
});

test('CSV import adopts a new workbook only after conversion and preserves existing unsaved drafts',async()=>{
 const App=await loadApp(),imports=[],checkpoints=[];let menu,renderer,convertFails=true,pickCanceled=false,waitForCancel=false;
 global.__importDelimited=async(seed,source,format,options)=>{assert.deepEqual([...seed],[80,75]);assert.equal(format,'csv');assert.deepEqual([...source],[65]);if(waitForCancel){options.onProgress(10,100);await new Promise((resolve,reject)=>options.signal.addEventListener('abort',()=>reject(new Error('Import aborted')),{once:true}))}if(convertFails)throw Error('Invalid CSV');return new Uint8Array([7,8,9])};
 global.window={localStorage:{getItem(){return null},setItem(){}},document:{title:''},addEventListener(){},removeEventListener(){},injDesktop:{
  recent:async()=>[],recovery:async()=>[],setDirty(){},setBusy(){},onMenuAction(callback){menu=callback;return()=>{}},
  create:async()=>({id:'original',name:'Untitled.docx',bytes:new Uint8Array([1]),untitled:true}),
  checkpoint:async(id,bytes,draft)=>checkpoints.push({id,bytes,draft}),
  pickDelimitedImport:async()=>pickCanceled?null:{name:'Budget.csv',format:'csv',seed:new Uint8Array([80,75]),bytes:new Uint8Array([65])},
  importDocument:async input=>{imports.push(input);return{id:'imported',name:input.name,bytes:input.bytes,untitled:true}},
 }};
 try{
  await act(async()=>{renderer=create(React.createElement(App))});
  await act(async()=>renderer.root.findByType('test-start').props.onCreate('docx'));
  const first=renderer.root.findByType('test-editor');await act(async()=>{first.props.onDraftChange(true);first.props.onRecoveryDraftChange({version:1,format:'docx',target:'text',text:'Unsaved words'})});
  await act(async()=>menu('new'));await act(async()=>renderer.root.findByType('test-start').props.onImportText());
  assert.equal(imports.length,0);assert.equal(renderer.root.findAllByType('test-editor').length,1);
  pickCanceled=true;await act(async()=>renderer.root.findByType('test-start').props.onImportText());assert.equal(imports.length,0);
  pickCanceled=false;convertFails=false;waitForCancel=true;await act(async()=>renderer.root.findByType('test-start').props.onImportText());
  const cancel=renderer.root.findAllByType('button').find(button=>button.children.includes('Cancel import'));assert(cancel);await act(async()=>cancel.props.onClick());assert.equal(imports.length,0);assert.equal(renderer.root.findAllByType('test-editor').length,1);
  waitForCancel=false;await act(async()=>renderer.root.findByType('test-start').props.onImportText());
  const editors=renderer.root.findAllByType('test-editor');assert.equal(editors.length,2);assert.equal(editors[0],first);assert.equal(editors[0].props.initialRecoveryDraft.text,'Unsaved words');
  assert.equal(imports[0].name,'Budget.xlsx');assert.deepEqual([...imports[0].bytes],[7,8,9]);assert.equal(editors[1].props.name,'Budget.xlsx');assert(checkpoints.some(value=>value.id==='imported'));
 }finally{if(renderer)await act(async()=>renderer.unmount());delete global.window;delete global.__importDelimited;}
});

function dropBridge(importDocument) {
  let hostBusy=false;
  global.window={localStorage:{getItem:()=>null,setItem(){}},document:{title:''},addEventListener(){},removeEventListener(){},injDesktop:{
    recent:async()=>[],recovery:async()=>[],nextExternal:async()=>null,setDirty(){},setBusy:value=>{hostBusy=value},onMenuAction(){return()=>{}},checkpoint:async()=>{},
    importDocument:input=>importDocument(input,hostBusy),
  }};
}
function droppedFile(name, read=async()=>new Uint8Array([1]).buffer) {return {name,size:1,arrayBuffer:read};}
function drop(renderer,files) {renderer.root.findAllByType('div').find(node=>node.props.onDrop).props.onDrop({defaultPrevented:false,preventDefault(){},dataTransfer:{files}});}

test('multi-file drop stages host imports before any editor can block the remaining files',async()=>{
  const App=await loadApp({busyOnMount:true});const imported=[];
  dropBridge(async(input,busy)=>{if(busy)return null;imported.push(input.name);return {id:'drop-'+imported.length,...input,untitled:true};});
  let renderer,releaseSecond;const second=new Promise(resolve=>{releaseSecond=resolve});
  try {
    await act(async()=>{renderer=create(React.createElement(App));});
    await act(async()=>drop(renderer,[droppedFile('One.docx'),droppedFile('Two.docx',async()=>{await second;return new Uint8Array([2]).buffer;})]));
    assert.deepEqual(imported,['One.docx']);
    assert.equal(renderer.root.findAllByType('test-editor').length,0,'staged copies do not start editor loading before the batch finishes');
    await act(async()=>{releaseSecond();await new Promise(resolve=>setImmediate(resolve));});
    assert.deepEqual(imported,['One.docx','Two.docx']);
    assert.equal(renderer.root.findAllByType('test-editor').length,2);
    assert.equal(renderer.root.findAllByProps({role:'alert'}).length,0);
    assert.match(renderer.root.findByProps({className:'status-message'}).children.join(''),/Working/);
  } finally {if(renderer)await act(async()=>renderer.unmount());delete global.window;}
});

test('partial multi-file drops retain successful copies and report every refused or invalid file',async()=>{
  const App=await loadApp();let sequence=0;
  dropBridge(async input=>{if(input.name==='Busy.docx')return null;if(input.name==='Broken.xlsx')throw new Error('Invalid ZIP package');return {id:'partial-'+(++sequence),...input,untitled:true};});
  let renderer;
  try {
    await act(async()=>{renderer=create(React.createElement(App));});
    await act(async()=>drop(renderer,['One.docx','Busy.docx','Broken.xlsx','Unsupported.txt','Two.docx'].map(name=>droppedFile(name))));
    assert.deepEqual(renderer.root.findAllByType('test-editor').map(editor=>editor.props.name),['One.docx','Two.docx']);
    const error=renderer.root.findByProps({role:'alert'}).findByType('span').children.join('');
    assert.match(error,/Could not import 3 files/);assert.match(error,/Busy.docx.*workspace was busy/);assert.match(error,/Broken.xlsx.*Invalid ZIP/);assert.match(error,/Unsupported.txt/);
    assert.match(renderer.root.findByProps({className:'status-message'}).children.join(''),/Imported 2 of 5 files/);
  } finally {if(renderer)await act(async()=>renderer.unmount());delete global.window;}
});

test('staged drops enforce the workspace limit and identify files that were not imported',async()=>{
  const App=await loadApp();let imports=0;
  dropBridge(async input=>({id:'limit-'+(++imports),...input,untitled:true}));
  let renderer;
  try {
    await act(async()=>{renderer=create(React.createElement(App));});
    await act(async()=>drop(renderer,Array.from({length:13},(_,index)=>droppedFile(`File ${index+1}.docx`))));
    assert.equal(imports,12);assert.equal(renderer.root.findAllByType('test-editor').length,12);
    assert.match(renderer.root.findByProps({role:'alert'}).findByType('span').children.join(''),/Could not import 1 file.*12-document limit/);
    assert.match(renderer.root.findByProps({className:'status-message'}).children.join(''),/Imported 12 of 13 files/);
  } finally {if(renderer)await act(async()=>renderer.unmount());delete global.window;}
});

test('workspace freezes during close preparation, waits for late drafts, and restores editing after cancellation',async()=>{
  const App=await loadApp();let prepare,cancel,menu,finish;const pending=[];
  global.window={localStorage:{getItem:()=>null},document:{title:''},addEventListener(){},removeEventListener(){},injDesktop:{
    nextExternal:async()=>null,recent:async()=>[],recovery:async()=>[],setDirty(){},setBusy(){},
    onPrepareClose(p,c){prepare=p;cancel=c;return()=>{}},onMenuAction(callback){menu=callback;return()=>{}},
    create:async()=>({id:'close-doc',name:'Untitled.docx',bytes:new Uint8Array([1]),untitled:true}),
    checkpoint:()=>new Promise(resolve=>{pending.push(resolve);finish=resolve}),
  }};
  let renderer;await act(async()=>{renderer=create(React.createElement(App))});
  try{
    await act(async()=>renderer.root.findByType('test-start').props.onCreate('docx'));
    const editor=renderer.root.findByType('test-editor');let ready=false,result;
    await act(async()=>{result=prepare().then(value=>{ready=value});});
    assert.equal(ready,false);
    assert.equal(renderer.root.findByProps({className:'desktop-app'}).props.inert,true);
    await act(async()=>menu('new'));
    assert.equal(renderer.root.findAllByType('test-start').length,0,'menu must not change frozen workspace');
    await act(async()=>{editor.props.onRecoveryDraftChange({text:'late draft'});pending[0]();});
    assert.equal(ready,false,'a draft accepted while draining extends the wait');
    await act(async()=>{finish();await result});
    assert.equal(ready,true);
    await act(async()=>cancel());
    assert.equal(renderer.root.findByProps({className:'desktop-app'}).props.inert,undefined);
    await act(async()=>editor.props.onBusyChange(true));
    assert.equal(await prepare(),false,'native operation refuses close');
    await act(async()=>editor.props.onBusyChange(false));
    await act(async()=>menu('new'));
    assert.equal(renderer.root.findAllByType('test-start').length,1);
  }finally{await act(async()=>renderer.unmount());delete global.window;}
});

test('unsupported open replaces the start page with OpenError instead of an empty home', async () => {
  const App = await loadApp();
  global.window = { localStorage: { getItem: () => null, setItem() {} }, document: { title: '' }, addEventListener() {}, removeEventListener() {}, injDesktop: {
    recent: async () => [], recovery: async () => [], setDirty() {}, setBusy() {}, onMenuAction() { return () => {}; },
    open: async () => { throw new Error('Choose a DOCX document, XLSX workbook, PPTX presentation, or PDF.'); },
  } };
  let renderer;
  try {
    await act(async () => { renderer = create(React.createElement(App)); });
    await act(async () => renderer.root.findByType('test-start').props.onOpen());
    assert.equal(renderer.root.findAllByType('test-start').length, 0);
    assert.match(renderer.root.findByProps({ id: 'open-error-title' }).children.join(''), /not supported/i);
    await act(async () => renderer.root.findAllByType('button').find(button => button.props.children === 'Back to start').props.onClick());
    assert.equal(renderer.root.findAllByType('test-start').length, 1);
  } finally { if (renderer) await act(async () => renderer.unmount()); delete global.window; }
});

test('the File tab backstage carries Home, New, Open, Save, Save as, Close and app entries into every editor ribbon', async () => {
  const App = await loadApp(); const saves = [], closed = []; let renderer;
  global.window = { localStorage: { getItem: () => null, setItem() {} }, document: { title: '' }, addEventListener() {}, removeEventListener() {}, injDesktop: {
    recent: async () => [], recovery: async () => [], nextExternal: async () => null, setDirty() {}, setBusy() {}, onMenuAction() { return () => {}; }, checkpoint: async () => {},
    create: async format => ({ id: `id-${format}`, name: `Untitled.${format}`, bytes: new Uint8Array([1]), untitled: true }),
    save: async input => { saves.push(input.saveAs); return { id: input.id, name: 'Report.docx', untitled: false }; },
    close: async id => closed.push(id),
  } };
  try {
    await act(async () => { renderer = create(React.createElement(App)); });
    await act(async () => renderer.root.findByType('test-start').props.onCreate('docx'));
    const groups = renderer.root.findAllByType('test-file-group');
    assert.deepEqual(groups.map(group => group.props.label), ['Start', 'New', 'Open & Save', 'Close', 'InjOffice']);
    const button = title => renderer.root.findAllByType('button').find(node => node.props.title === title);
    assert.equal(button('Save (⌘S)').props.disabled, false, 'a freshly created document is unsaved, so Save is enabled');
    await act(async () => button('Save as… (⌘⇧S)').props.onClick());
    assert.deepEqual(saves, [true]);
    assert.equal(button('Save (⌘S)').props.disabled, true, 'Save greys out once the document is saved');
    await act(async () => button('New spreadsheet (XLSX)').props.onClick());
    assert.deepEqual(renderer.root.findAllByType('test-editor').map(editor => editor.props.name), ['Untitled.docx', 'Untitled.xlsx']);
    await act(async () => button('Start page: create, open recent or recover').props.onClick());
    assert.equal(renderer.root.findAllByType('test-start').length, 1, 'File › Home shows the start page');
    await act(async () => renderer.root.findByType('test-start').props.onResume());
    await act(async () => button('Close document').props.onClick());
    assert.equal(renderer.root.findAllByProps({ role: 'alertdialog' }).length, 1, 'closing an unsaved workbook asks first');
    await act(async () => renderer.root.findAllByType('button').find(node => node.children.includes('Discard changes')).props.onClick());
    assert.deepEqual(closed, ['id-xlsx']);
    assert.deepEqual(renderer.root.findAllByType('test-editor').map(editor => editor.props.name), ['Untitled.docx']);
    // The palette still lists every backstage command.
    await act(async () => renderer.root.findByProps({ title: 'Search commands (Ctrl/⌘ K)' }).props.onClick());
    const labels = renderer.root.findAllByProps({ role: 'option' }).map(option => option.props.id);
    for (const id of ['command-home', 'command-new-docx', 'command-open', 'command-save', 'command-save-as', 'command-close', 'command-updates', 'command-preferences']) assert.ok(labels.includes(id), id);
  } finally { if (renderer) await act(async () => renderer.unmount()); delete global.window; }
});

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const React = require('react');
const { create, act } = require('react-test-renderer');
global.IS_REACT_ACT_ENVIRONMENT = true;

async function load() {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src/ContextMenu.tsx'),
    platform: 'node',
    external: id => /^react(?:\/|$)/.test(id),
    transform: { jsx: { runtime: 'automatic' } },
    plugins: [{ name: 'css', resolveId(id) { if (id.endsWith('.css')) return '\0css'; }, load(id) { if (id === '\0css') return 'export default ""'; } }],
  });
  try {
    const { output } = await bundle.generate({ format: 'cjs', codeSplitting: false });
    const mod = { exports: {} };
    new Function('require', 'module', 'exports', output[0].code)(require, mod, mod.exports);
    return mod.exports;
  } finally {
    await bundle.close();
  }
}

// Menus spell shortcuts for this platform, so the expectations come from shortcuts.ts itself.
async function loadShortcuts() {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({ input: path.resolve(__dirname, '../src/shortcuts.ts'), platform: 'node' });
  try { const { output } = await bundle.generate({ format: 'cjs' }); const mod = { exports: {} }; new Function('require', 'module', 'exports', output[0].code)(require, mod, mod.exports); return mod.exports; } finally { await bundle.close(); }
}

const anchor = { x: 40, y: 50, source: null };
const items = menu => menu.root.findAll(node => typeof node.type === 'string' && /^menuitem/.test(node.props.role ?? ''));

test('context menu is an accessible menu with shortcut hints, disabled explanations and check marks', async () => {
  assert.equal(fs.existsSync(path.resolve(__dirname, '../src/context-menu.css')), true);
  const { default: ContextMenu, SHORTCUTS } = await load();
  const { shortcutLabel } = await loadShortcuts();
  assert.equal(SHORTCUTS.copy, shortcutLabel('copy'), 'menus spell the shortcut for this platform, not "Ctrl / ⌘ C"');
  assert.match(SHORTCUTS.cut, /^(⌘X|Ctrl\+X)$/);
  const runs = [];
  const closes = [];
  let view;
  await act(async () => {
    view = create(React.createElement(ContextMenu, { anchor, label: 'Document', onClose: () => closes.push('close'), items: [
      { id: 'copy', label: 'Copy', shortcut: SHORTCUTS.copy, run: () => runs.push('copy') },
      { id: 'paste', label: 'Paste', shortcut: SHORTCUTS.paste, disabled: true, title: 'Clipboard policy', run: () => runs.push('paste') },
      { separator: true },
      { id: 'bold', label: 'Bold', checked: true, run: () => runs.push('bold') },
    ] }));
  });
  const menu = view.root.findByProps({ role: 'menu' });
  assert.equal(menu.props['aria-label'], 'Document');
  assert.deepEqual(menu.props.style, { left: 40, top: 50 });
  const buttons = items(view);
  assert.deepEqual(buttons.map(button => button.props.role), ['menuitem', 'menuitem', 'menuitemcheckbox']);
  assert.equal(view.root.findAllByProps({ role: 'separator' }).length, 1);
  assert.equal(buttons[0].findByType('kbd').children.join(''), SHORTCUTS.copy);
  assert.equal(buttons[1].props['aria-disabled'], true);
  assert.equal(buttons[1].props.title, 'Clipboard policy');
  assert.equal(buttons[2].props['aria-checked'], true);
  // Disabled entries never run; enabled entries close the menu before running.
  await act(async () => buttons[1].props.onClick());
  assert.deepEqual(runs, []); assert.deepEqual(closes, []);
  await act(async () => buttons[0].props.onClick());
  assert.deepEqual(runs, ['copy']); assert.deepEqual(closes, ['close']);
  // Escape closes; a right-click on the menu itself never opens a native menu.
  await act(async () => menu.props.onKeyDown({ key: 'Escape', preventDefault() {}, stopPropagation() {} }));
  assert.deepEqual(closes, ['close', 'close']);
  let prevented = false;
  menu.props.onContextMenu({ preventDefault() { prevented = true; } });
  assert.equal(prevented, true);
  await act(async () => view.unmount());
});

test('menu geometry stays inside the viewport and arrow navigation skips separators and disabled entries', async () => {
  const { clampMenuPosition, nextMenuIndex } = await load();
  assert.deepEqual(clampMenuPosition(100, 100, 200, 300, 1000, 800), { x: 100, y: 100 });
  assert.deepEqual(clampMenuPosition(950, 100, 200, 300, 1000, 800), { x: 750, y: 100 });
  assert.deepEqual(clampMenuPosition(100, 700, 200, 300, 1000, 800), { x: 100, y: 400 });
  assert.deepEqual(clampMenuPosition(2, 2, 200, 300, 1000, 800), { x: 8, y: 8 });
  assert.deepEqual(clampMenuPosition(500, 500, 2000, 2000, 1000, 800), { x: 8, y: 8 });
  const list = [{ id: 'a', label: 'A', run() {} }, { id: 'b', label: 'B', disabled: true, run() {} }, { separator: true }, { id: 'c', label: 'C', run() {} }];
  assert.equal(nextMenuIndex(list, -1, 1), 0);
  assert.equal(nextMenuIndex(list, 0, 1), 3);
  assert.equal(nextMenuIndex(list, 3, 1), 0);
  assert.equal(nextMenuIndex(list, 0, -1), 3);
  assert.equal(nextMenuIndex([{ separator: true }], 0, 1), 0);
  assert.equal(nextMenuIndex([], 0, 1), -1);
});

test('document menu mirrors the ribbon: formatting goes through onFormat, disabled states carry the toolbar explanation', async () => {
  const { documentContextMenu, SHORTCUTS, PASTE_UNAVAILABLE } = await load();
  const patches = [];
  const tabs = [];
  let found = 0;
  const context = { anchor, target: true, values: { bold: false, italic: true, underline: false, alignment: 'center', characterEditable: true }, disabled: false, link: true, table: false,
    onFormat: patch => patches.push(patch), onFind: () => found++, onRibbonTab: tab => tabs.push(tab) };
  const menu = documentContextMenu(context);
  const byId = id => menu.find(item => item.id === id);
  // Insert table… is hidden: this selection cannot take a table. Office hides what it cannot run.
  assert.deepEqual(menu.filter(item => !('separator' in item)).map(item => item.label), ['Cut', 'Copy', 'Paste', 'Bold', 'Italic', 'Underline', 'Align left', 'Center', 'Align right', 'Justify', 'Hyperlink…', 'Find / replace']);
  assert.equal(byId('paste').disabled, true);
  assert.equal(byId('paste').title, PASTE_UNAVAILABLE);
  // No DOM selection exists under node: cut/copy are disabled with an explanation instead of pretending.
  assert.equal(byId('cut').disabled, true);
  assert.equal(byId('copy').disabled, true);
  assert.equal(byId('bold').checked, false);
  assert.equal(byId('italic').checked, true);
  assert.equal(byId('align-center').checked, true);
  assert.equal(byId('align-left').checked, false);
  assert.equal(byId('table'), undefined, 'an unavailable command is left out, not greyed');
  assert.equal(byId('hyperlink').disabled, false);
  byId('bold').run(); byId('italic').run(); byId('underline').run(); byId('align-both').run();
  assert.deepEqual(patches, [{ bold: true }, { italic: false }, { underline: true }, { alignment: 'both' }]);
  byId('find').run();
  assert.equal(found, 1);
  byId('hyperlink').run();
  assert.deepEqual(tabs, ['Insert']);
  assert.deepEqual(menu.filter(item => item.shortcut).map(item => [item.id, item.shortcut]), [['cut', SHORTCUTS.cut], ['copy', SHORTCUTS.copy], ['paste', SHORTCUTS.paste], ['bold', SHORTCUTS.bold], ['italic', SHORTCUTS.italic], ['underline', SHORTCUTS.underline], ['find', SHORTCUTS.find]]);
  // Character formatting follows the toolbar's characterEditable flag; alignment stays available.
  const locked = documentContextMenu({ ...context, values: { ...context.values, characterEditable: false } });
  assert.equal(locked.find(item => item.id === 'bold'), undefined, 'text that cannot be formatted hides the formatting entries');
  assert.equal(locked.find(item => item.id === 'align-left').disabled, false);
  // No separator is left leading, trailing or doubled once entries drop out.
  const shape = locked.map(item => 'separator' in item ? '|' : 'x').join('');
  assert.ok(!/^\||\|\||\|$/.test(shape), shape);
  // Busy only means the workspace is working: Office greys those entries, it does not remove them.
  const busy = documentContextMenu({ ...context, disabled: true });
  assert.deepEqual(busy.filter(item => !('separator' in item)).map(item => item.label), menu.filter(item => !('separator' in item)).map(item => item.label));
  assert.equal(busy.filter(item => !('separator' in item) && !item.disabled).map(item => item.id).join(), 'find');
  // Cut/Copy/Paste always stay visible, greyed, the way Office shows them.
  for (const id of ['cut', 'copy', 'paste']) assert.ok(busy.find(item => item.id === id), id);
});

test('worksheet menu reuses clear/copy and leaves out the commands the native transaction lacks', async () => {
  const { spreadsheetContextMenu, contextMenuCellKey, copyThroughGrid, PASTE_UNAVAILABLE } = await load();
  let cleared = 0;
  const source = { focus() {}, closest: () => null };
  // Without a DOM there is no clipboard: copy reports failure instead of throwing or inventing a payload.
  assert.equal(copyThroughGrid(null), false);
  assert.equal(copyThroughGrid(source), false);
  const menu = spreadsheetContextMenu({ anchor: { x: 0, y: 0, source }, disabled: false, onClear: () => cleared++ });
  const byId = id => menu.find(item => item.id === id);
  assert.deepEqual(menu.filter(item => !('separator' in item)).map(item => item.label), ['Cut', 'Copy', 'Paste', 'Clear contents', 'Row height…', 'Column width…']);
  for (const id of ['row-insert', 'row-delete', 'column-insert', 'column-delete']) assert.equal(byId(id), undefined, `${id} is not in the native transaction, so it is not offered`);
  assert.equal(byId('paste').title, PASTE_UNAVAILABLE);
  byId('clear').run();
  assert.equal(cleared, 1);
  byId('cut').run();
  assert.equal(cleared, 2);
  const locked = spreadsheetContextMenu({ anchor: { x: 0, y: 0, source }, disabled: true, onClear: () => cleared++ });
  assert.deepEqual(locked.filter(item => !('separator' in item) && !item.disabled).map(item => item.id), ['copy']);
  const cell = title => ({ closest: () => ({ title }) });
  assert.equal(contextMenuCellKey(cell('B3')), 'B3');
  assert.equal(contextMenuCellKey(cell('AZ10 · merged A1:B2')), 'AZ10');
  assert.equal(contextMenuCellKey(cell('not-a-cell')), undefined);
  assert.equal(contextMenuCellKey({ closest: () => null }), undefined);
  assert.equal(contextMenuCellKey(null), undefined);
});

test('slide menu drives the existing slide and object commands only', async () => {
  const { presentationContextMenu, PPTX_CLIPBOARD_UNAVAILABLE } = await load();
  const calls = [];
  const context = { object: true, slide: true, disabled: false, canDeleteSlide: true, onDeleteObject: () => calls.push('delete-object'), onNewSlide: () => calls.push('new'), onDuplicateSlide: () => calls.push('duplicate'), onDeleteSlide: () => calls.push('delete-slide') };
  const menu = presentationContextMenu(context);
  const byId = id => menu.find(item => item.id === id);
  assert.deepEqual(menu.filter(item => !('separator' in item)).map(item => item.label), ['Cut', 'Copy', 'Paste', 'Delete object', 'New slide', 'Duplicate slide', 'Delete slide']);
  for (const id of ['cut', 'copy']) { assert.equal(byId(id).disabled, true); assert.equal(byId(id).title, PPTX_CLIPBOARD_UNAVAILABLE); }
  for (const id of ['delete-object', 'new-slide', 'duplicate-slide', 'delete-slide']) byId(id).run();
  assert.deepEqual(calls, ['delete-object', 'new', 'duplicate', 'delete-slide']);
  const single = presentationContextMenu({ ...context, object: false, canDeleteSlide: false });
  assert.equal(single.find(item => item.id === 'delete-object'), undefined, 'with nothing selected the object command is left out');
  assert.equal(single.find(item => item.id === 'delete-slide'), undefined, 'the last slide cannot be deleted, so the entry is not offered');
  // Cut/Copy/Paste stay, greyed, as Office shows them.
  assert.deepEqual(single.filter(item => !('separator' in item)).map(item => item.id), ['cut', 'copy', 'paste', 'new-slide', 'duplicate-slide']);
  assert.deepEqual(presentationContextMenu({ ...context, disabled: true }).filter(item => !('separator' in item) && !item.disabled), []);
});

test('editors attach the menu with a right-click handler and render it from their own state', () => {
  const read = file => fs.readFileSync(path.resolve(__dirname, '../src', file), 'utf8');
  const office = read('OfficeEditor.tsx');
  assert.match(office, /onContextMenu=\{event=>\{activateRunAt\(event\);menu\.open\(event\)\}\}/);
  assert.match(office, /documentContextMenu\(\{anchor:menu\.anchor,target:!!target,values:toolbarValues/);
  assert.match(office, /onFormat:patch=>void changeFormatting\(patch\)/);
  const sheet = read('SpreadsheetEditor.tsx');
  assert.match(sheet, /contextMenuCellKey\(event\.target\)/);
  assert.match(sheet, /spreadsheetContextMenu\(\{ anchor: menu\.anchor, disabled, onClear/);
  const slides = read('PresentationEditor.tsx');
  assert.match(slides, /onContextMenu=\{event => \{ selectObjectAt\(event\); menu\.open\(event\); \}\}/);
  assert.match(slides, /presentationContextMenu\(\{ object: /);
});

test('worksheet size menu opens Home and the current Format cells disclosure', async () => {
  const { spreadsheetContextMenu } = await load();
  const calls = []; const disclosure = { open: false };
  const field = { focus: () => calls.push('focus'), select: () => calls.push('select') };
  const editor = { querySelectorAll: () => [{ textContent: 'Home', click: () => calls.push('home') }], querySelector: selector => selector === '[aria-label="Format cells"]' ? { closest: () => disclosure } : field };
  global.window = { requestAnimationFrame: callback => callback() };
  try {
    const menu = spreadsheetContextMenu({ anchor: { source: { closest: () => editor } }, disabled: false, onClear() {} });
    menu.find(item => item.id === 'row-height').run();
    assert.equal(disclosure.open, true);
    assert.deepEqual(calls, ['home', 'focus', 'select']);
  } finally { delete global.window; }
});

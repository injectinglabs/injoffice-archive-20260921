const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const React = require('react');
const { create, act } = require('react-test-renderer');
global.IS_REACT_ACT_ENVIRONMENT = true;

async function load(file) {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src', file),
    platform: 'node',
    external: id => /^react(?:\/|$)/.test(id),
    transform: { jsx: { runtime: 'automatic' } },
    plugins: [{
      name: 'css',
      resolveId(id) { if (id.endsWith('.css')) return '\0css'; },
      load(id) { if (id === '\0css') return 'export default ""'; },
    }],
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

const h = React.createElement;
const text = node => node.children.map(child => typeof child === 'string' ? child : text(child)).join('');

test('shortcut labels come from one table and render per platform', async () => {
  const shortcuts = await load('shortcuts.ts');
  assert.equal(shortcuts.shortcutLabel('bold', 'mac'), '⌘B');
  assert.equal(shortcuts.shortcutLabel('bold', 'other'), 'Ctrl+B');
  assert.equal(shortcuts.shortcutLabel('redo', 'mac'), '⌘⇧Z');
  assert.equal(shortcuts.shortcutLabel('redo', 'other'), 'Ctrl+Y');
  assert.equal(shortcuts.shortcutLabel('apply', 'mac'), '⌘↩');
  assert.equal(shortcuts.shortcutLabel('apply', 'other'), 'Ctrl+Enter');
  assert.equal(shortcuts.shortcutLabel('cancel', 'mac'), 'Esc');
  assert.equal(shortcuts.shortcutTooltip('Find / replace', 'find', 'mac'), 'Find / replace (⌘F)');
  assert.equal(shortcuts.shortcutTooltip('Insert table', undefined, 'mac'), 'Insert table');
  assert.equal(shortcuts.shortcutKeys('redo', 'mac'), 'Meta+Shift+Z');
  assert.equal(shortcuts.shortcutKeys('undo', 'other'), 'Control+Z');
});

test('ribbon renders Office tabs with labelled groups and drops empty tabs and groups', async () => {
  assert.equal(fs.existsSync(path.resolve(__dirname, '../src/ribbon.css')), true);
  const { default: Ribbon, RibbonButton, visibleRibbonTabs } = await load('Ribbon.tsx');
  const changes = [];
  const tabs = [
    { id: 'File', label: 'File', groups: [{ id: 'export', label: 'Export', children: false }] },
    { id: 'Home', label: 'Home', groups: [
      { id: 'font', label: 'Font', children: h(RibbonButton, { icon: 'bold', label: 'Bold', shortcut: 'bold', labelHidden: true }) },
      { id: 'clipboard', label: 'Clipboard', children: [false, null] },
      { id: 'editing', label: 'Editing', children: h(RibbonButton, { icon: 'replace', label: 'Find / replace', shortcut: 'find' }) },
    ] },
    { id: 'Insert', label: 'Insert', groups: [{ id: 'tables', label: 'Tables', children: h(RibbonButton, { icon: 'table', label: 'Insert table' }) }] },
  ];
  assert.deepEqual(visibleRibbonTabs(tabs).map(tab => [tab.id, tab.groups.map(group => group.label)]), [['Home', ['Font', 'Editing']], ['Insert', ['Tables']]]);

  let view;
  await act(async () => {
    view = create(h(Ribbon, { label: 'Document tools', tabs, active: 'Home', onChange: tab => changes.push(tab), quickAccess: h(RibbonButton, { icon: 'undo', label: 'Undo', shortcut: 'undo' }) }));
  });
  const tablist = view.root.findByProps({ role: 'tablist' });
  assert.equal(tablist.props['aria-label'], 'Document tools tabs');
  const tabButtons = view.root.findAllByProps({ role: 'tab' });
  assert.deepEqual(tabButtons.map(text), ['Home', 'Insert']);
  assert.deepEqual(tabButtons.map(tab => tab.props['aria-selected']), [true, false]);
  assert.deepEqual(tabButtons.map(tab => tab.props.tabIndex), [0, -1]);
  const panels = view.root.findAllByProps({ role: 'tabpanel' });
  assert.deepEqual(panels.map(panel => panel.props.hidden), [false, true]);
  assert.equal(panels[0].props['aria-labelledby'], tabButtons[0].props.id);
  assert.equal(tabButtons[0].props['aria-controls'], panels[0].props.id);
  const groups = view.root.findAllByProps({ role: 'group' });
  assert.deepEqual(groups.map(group => group.props['aria-label']), ['Font', 'Editing', 'Tables']);
  assert.deepEqual(groups.map(group => text(group.findByProps({ className: 'ribbon-group-label' }))), ['Font', 'Editing', 'Tables']);
  assert.equal(view.root.findByProps({ role: 'toolbar' }).props['aria-label'], 'Quick access');

  // Arrow keys move between tabs (wrapping) and Home/End jump to the ends.
  const keyEvent = key => ({ key, preventDefault() {}, currentTarget: { parentElement: { children: [] } } });
  tabButtons[0].props.onKeyDown(keyEvent('ArrowRight'));
  tabButtons[0].props.onKeyDown(keyEvent('ArrowLeft'));
  tabButtons[0].props.onKeyDown(keyEvent('End'));
  tabButtons[1].props.onKeyDown(keyEvent('Home'));
  tabButtons[1].props.onKeyDown(keyEvent('Tab'));
  tabButtons[1].props.onClick();
  assert.deepEqual(changes, ['Insert', 'Insert', 'Insert', 'Home', 'Insert']);
  await act(async () => view.unmount());
});

test('ribbon buttons carry an icon, a label, and the shortcut in their tooltip', async () => {
  const { RibbonButton } = await load('Ribbon.tsx');
  const { shortcutLabel, shortcutKeys } = await load('shortcuts.ts');
  const { ribbonIconNames } = await load('RibbonIcons.tsx');
  assert.ok(ribbonIconNames.length >= 40);
  let view;
  await act(async () => {
    view = create(h('div', null,
      h(RibbonButton, { icon: 'undo', label: 'Undo', shortcut: 'undo', disabled: true }),
      h(RibbonButton, { icon: 'bold', label: 'Bold', shortcut: 'bold', labelHidden: true, 'aria-pressed': true }),
      h(RibbonButton, { icon: 'pdf', label: 'Export PDF…', title: 'Export vector PDF' }),
    ));
  });
  const buttons = view.root.findAllByType('button');
  assert.equal(buttons.length, 3);
  for (const button of buttons) assert.equal(button.findAllByType('svg').length, 1, 'every button has an icon');
  assert.equal(buttons[0].props.title, `Undo (${shortcutLabel('undo')})`);
  assert.equal(buttons[0].props['aria-keyshortcuts'], shortcutKeys('undo'));
  assert.equal(buttons[0].props.disabled, true);
  assert.equal(text(buttons[0]), 'Undo');
  assert.equal(buttons[1].props['aria-label'], 'Bold');
  assert.equal(buttons[1].props.title, `Bold (${shortcutLabel('bold')})`);
  assert.equal(buttons[1].props['aria-pressed'], true);
  assert.equal(text(buttons[1]), '', 'hidden label leaves only the icon');
  assert.equal(buttons[2].props.title, 'Export vector PDF');
  assert.equal(buttons[2].props['aria-keyshortcuts'], undefined);
  assert.equal(text(buttons[2]), 'Export PDF…');
  await act(async () => view.unmount());
});

test('the shell contributes Office backstage groups to every editor File tab through context', async () => {
  const { default: Ribbon, RibbonButton, WorkspaceFileGroupsContext, withWorkspaceFileGroups, visibleRibbonTabs } = await load('Ribbon.tsx');
  const clicks = [];
  const workspace = {
    before: [{ id: 'workspace-home', label: 'Start', children: h(RibbonButton, { icon: 'home', label: 'Home', onClick: () => clicks.push('home') }) }, { id: 'workspace-open-save', label: 'Open & Save', children: h(RibbonButton, { icon: 'open', label: 'Open', shortcut: 'open' }) }],
    after: [{ id: 'workspace-close', label: 'Close', children: h(RibbonButton, { icon: 'closeDocument', label: 'Close document' }) }],
  };
  const editorTabs = [
    { id: 'File', label: 'File', groups: [{ id: 'export', label: 'Export', children: h(RibbonButton, { icon: 'export', label: 'Export sheet' }) }] },
    { id: 'Home', label: 'Home', groups: [{ id: 'font', label: 'Font', children: h(RibbonButton, { icon: 'bold', label: 'Bold' }) }] },
  ];
  // Editor groups (Export) sit between the shell's leading and trailing groups; the editor's other tabs are untouched.
  assert.deepEqual(visibleRibbonTabs(withWorkspaceFileGroups(editorTabs, workspace)).map(tab => [tab.id, tab.groups.map(group => group.label)]), [['File', ['Start', 'Open & Save', 'Export', 'Close']], ['Home', ['Font']]]);
  // An editor whose File tab is empty (no export available) still gets a File tab, first.
  const emptyFile = [{ id: 'File', label: 'File', groups: [{ id: 'export', label: 'Export', children: false }] }, editorTabs[1]];
  assert.deepEqual(visibleRibbonTabs(withWorkspaceFileGroups(emptyFile, workspace)).map(tab => tab.groups.map(group => group.label)), [['Start', 'Open & Save', 'Close'], ['Font']]);
  // An editor without a File tab gets one prepended; without a provider nothing changes.
  assert.deepEqual(withWorkspaceFileGroups([editorTabs[1]], workspace).map(tab => tab.id), ['File', 'Home']);
  assert.equal(withWorkspaceFileGroups(editorTabs, null), editorTabs);

  let view;
  await act(async () => { view = create(h(WorkspaceFileGroupsContext, { value: workspace }, h(Ribbon, { label: 'Spreadsheet tools', tabs: editorTabs, active: 'File', onChange() {} }))); });
  assert.deepEqual(view.root.findAllByProps({ role: 'tab' }).map(text), ['File', 'Home']);
  assert.deepEqual(view.root.findAllByProps({ role: 'group' }).map(group => group.props['aria-label']), ['Start', 'Open & Save', 'Export', 'Close', 'Font']);
  const open = view.root.findAllByType('button').find(button => text(button) === 'Open');
  assert.equal(open.props.title, 'Open (⌘O)');
  await act(async () => view.root.findAllByType('button').find(button => button.props.role !== 'tab' && text(button) === 'Home').props.onClick());
  assert.deepEqual(clicks, ['home']);
  await act(async () => { view = create(h(Ribbon, { label: 'Spreadsheet tools', tabs: emptyFile, active: 'File', onChange() {} })); });
  assert.deepEqual(view.root.findAllByProps({ role: 'tab' }).map(text), ['Home'], 'without the shell provider an empty File tab still drops out');
});

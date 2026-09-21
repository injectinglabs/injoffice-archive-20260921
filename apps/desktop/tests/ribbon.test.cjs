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
  const { shortcutTooltip } = await load('shortcuts.ts');
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
  assert.equal(open.props.title, shortcutTooltip('Open', 'open'), 'the tooltip carries the platform shortcut');
  await act(async () => view.root.findAllByType('button').find(button => button.props.role !== 'tab' && text(button) === 'Home').props.onClick());
  assert.deepEqual(clicks, ['home']);
  await act(async () => { view = create(h(Ribbon, { label: 'Spreadsheet tools', tabs: emptyFile, active: 'File', onChange() {} })); });
  assert.deepEqual(view.root.findAllByProps({ role: 'tab' }).map(text), ['Home'], 'without the shell provider an empty File tab still drops out');
});

test('every ribbon glyph is a drawable path, including the PDF tool set', async () => {
  const { default: RibbonIcon, ribbonIconNames } = await load('RibbonIcons.tsx');
  for (const name of ['note', 'highlight', 'strikethrough', 'rectangle', 'ellipse', 'line', 'arrow', 'form', 'rotate', 'pageDelete', 'import', 'pagePrevious', 'pageNext', 'goToPage']) {
    assert.ok(ribbonIconNames.includes(name), `${name} is a ribbon glyph`);
  }
  let view;
  await act(async () => { view = create(h('div', null, ribbonIconNames.map(name => h(RibbonIcon, { key: name, name })))); });
  const paths = view.root.findAllByType('path');
  assert.equal(paths.length, ribbonIconNames.length);
  paths.forEach((path, index) => {
    assert.match(path.props.d, /^[Mm][\d.\s-]/, `${ribbonIconNames[index]} starts with a move command`);
    assert.equal(/[^MmLlHhVvCcSsQqTtAaZz\d.,\s-]/.test(path.props.d), false, `${ribbonIconNames[index]} uses path commands only`);
  });
  await act(async () => view.unmount());
});

test('the ribbon panel keeps Office geometry: 94px, one row centred, labels on one baseline', () => {
  const css = fs.readFileSync(path.resolve(__dirname, '../src/ribbon.css'), 'utf8');
  const rule = selector => {
    const match = css.match(new RegExp(`${selector.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\s*\\{([^}]*)\\}`));
    assert.ok(match, `ribbon.css declares ${selector}`);
    return match[1];
  };
  const panel = rule('.ribbon-panel');
  const body = rule('.ribbon-group-body');
  const label = rule('.ribbon-group-label');
  const rows = rule('.ribbon-rows > *');
  const px = (declarations, property) => Number(declarations.match(new RegExp(`(?:^|;|\\s)${property}:\\s*(\\d+)px`))?.[1]);
  // 4 + 66 + 18 + 6 = 94: Office's ribbon height, two 30px rows or one row centred in the body.
  assert.equal(px(panel, 'min-height'), 94);
  assert.equal(px(body, 'min-height'), 66);
  assert.match(label, /flex:\s*0 0 18px/, 'the label strip is a fixed band, so every group label shares a baseline');
  assert.equal(px(rows, 'min-height'), 30);
  assert.match(body, /align-items:\s*center/, 'a single row of controls is vertically centred');
  assert.match(rule('.ribbon-rows'), /justify-content:\s*center/, 'stacked rows are centred as a block');
  // Fixed control widths keep a group's geometry when the selection changes.
  for (const combo of ['.ribbon-combo-font', '.ribbon-combo-size', '.ribbon-combo-style', '.ribbon-combo-list']) {
    assert.ok(px(rule(combo), 'width') > 0, `${combo} has a fixed width`);
  }
});

test('a narrow ribbon collapses to icons, then moves the trailing groups behind an overflow chevron', async () => {
  const { default: Ribbon, RibbonButton, ribbonGroupsThatFit, RIBBON_OVERFLOW_WIDTH } = await load('Ribbon.tsx');
  // Everything fits: no overflow. Otherwise fill from the left, keeping room for the chevron.
  assert.equal(ribbonGroupsThatFit([100, 100, 100], 400), 3);
  assert.equal(ribbonGroupsThatFit([100, 100, 100], 300), 3);
  assert.equal(ribbonGroupsThatFit([100, 100, 100], 260, 44), 2);
  assert.equal(ribbonGroupsThatFit([100, 100, 100], 150, 44), 1);
  assert.equal(ribbonGroupsThatFit([300], 100, 44), 1, 'one group always stays on the ribbon');
  assert.equal(RIBBON_OVERFLOW_WIDTH > 0, true);

  const groups = ['Font', 'Paragraph', 'Styles', 'Editing'].map(label => ({ id: label.toLowerCase(), label, children: h(RibbonButton, { icon: 'bold', label }) }));
  const tabs = [{ id: 'Home', label: 'Home', groups }];
  let view;
  // Without a layout the panel renders every group with its labels (the wide case).
  await act(async () => { view = create(h(Ribbon, { label: 'Document tools', tabs, active: 'Home', onChange() {} })); });
  let panel = view.root.findByProps({ role: 'tabpanel' });
  assert.equal(panel.props.className, 'ribbon-panel');
  assert.deepEqual(panel.findAllByProps({ role: 'group' }).map(group => group.props['aria-label']), ['Font', 'Paragraph', 'Styles', 'Editing']);
  assert.equal(view.root.findAllByProps({ className: 'ribbon-overflow' }).length, 0);

  // Compact only: icons, no overflow chevron yet.
  await act(async () => { view.update(h(Ribbon, { label: 'Document tools', tabs, active: 'Home', onChange() {}, panelLayout: { compact: true, visible: -1 } })); });
  panel = view.root.findByProps({ role: 'tabpanel' });
  assert.equal(panel.props.className, 'ribbon-panel ribbon-panel-compact');
  assert.equal(view.root.findAllByProps({ className: 'ribbon-overflow' }).length, 0);

  // Too narrow: the last two groups move into the popover behind the chevron.
  await act(async () => { view.update(h(Ribbon, { label: 'Document tools', tabs, active: 'Home', onChange() {}, panelLayout: { compact: true, visible: 2 } })); });
  panel = view.root.findByProps({ role: 'tabpanel' });
  assert.deepEqual(panel.findAllByProps({ role: 'group' }).map(group => group.props['aria-label']), ['Font', 'Paragraph'], 'hidden groups leave the panel row');
  const chevron = view.root.findByProps({ className: 'ribbon-button ribbon-overflow-button' });
  assert.equal(chevron.props['aria-label'], 'More commands (2 groups)');
  assert.equal(chevron.props['aria-expanded'], false);
  assert.equal(view.root.findAllByProps({ className: 'ribbon-overflow-popover' }).length, 0);
  await act(async () => chevron.props.onClick());
  assert.equal(view.root.findByProps({ className: 'ribbon-button ribbon-overflow-button' }).props['aria-expanded'], true);
  const popover = view.root.findByProps({ className: 'ribbon-overflow-popover' });
  assert.deepEqual(popover.findAllByProps({ role: 'group' }).map(group => group.props['aria-label']), ['Styles', 'Editing'], 'the hidden groups are reachable in the popover');
  await act(async () => view.unmount());
});

test('the shared ribbon primitives: alignment toggles and a colour button, ready for the spreadsheet ribbon', async () => {
  const { AlignmentToggles, ColorButton, RIBBON_ALIGNMENTS, RIBBON_VERTICAL_ALIGNMENTS } = await load('Ribbon.tsx');
  assert.deepEqual(RIBBON_VERTICAL_ALIGNMENTS.map(option => option.value), ['top', 'middle', 'bottom']);
  assert.deepEqual(RIBBON_ALIGNMENTS.xlsx.map(option => option.value), ['left', 'center', 'right']);
  assert.deepEqual(RIBBON_ALIGNMENTS.docx.map(option => option.value), ['left', 'center', 'right', 'both', 'distribute']);

  const horizontal = [], vertical = [];
  let view;
  await act(async () => {
    view = create(h(AlignmentToggles, { horizontal: 'center', vertical: 'bottom', onHorizontal: value => horizontal.push(value), onVertical: value => vertical.push(value) }));
  });
  // Six toggles, not two dropdowns, and the alignment in effect is the pressed one.
  const buttons = view.root.findAllByType('button');
  assert.deepEqual(buttons.map(button => button.props['aria-label']), ['Top align', 'Middle align', 'Bottom align', 'Align left', 'Center', 'Align right']);
  assert.deepEqual(buttons.map(button => button.props['aria-pressed']), [false, false, true, false, true, false]);
  for (const button of buttons) assert.equal(button.findAllByType('svg').length, 1, `${button.props['aria-label']} has an icon`);
  await act(async () => buttons[0].props.onClick());
  await act(async () => buttons[3].props.onClick());
  assert.deepEqual(vertical, ['top']);
  assert.deepEqual(horizontal, ['left']);
  // Excel's "general" alignment presses nothing, exactly like Excel's own ribbon.
  await act(async () => view.update(h(AlignmentToggles, { horizontal: 'general', onHorizontal: () => {}, disabled: true })));
  const only = view.root.findAllByType('button');
  assert.deepEqual(only.map(button => button.props['aria-label']), ['Align left', 'Center', 'Align right'], 'a host without vertical alignment gets one row');
  assert.deepEqual(only.map(button => button.props['aria-pressed']), [false, false, false]);
  assert.deepEqual(only.map(button => button.props.disabled), [true, true, true]);

  // ColorButton: an icon button with the colour as an underline and a palette behind the chevron.
  const picked = [];
  await act(async () => { view.update(h(ColorButton, { label: 'Fill color', icon: 'fill', value: '#217447', colors: [['#B33C3C', 'Red'], ['#217447', 'Green']], onChange: color => picked.push(color) })); });
  const trigger = view.root.findByProps({ 'aria-label': 'Fill color' });
  assert.equal(trigger.props['aria-expanded'], false);
  assert.equal(trigger.props['aria-haspopup'], 'true');
  assert.deepEqual(view.root.findByProps({ className: 'ribbon-color-bar' }).props.style, { background: '#217447' });
  assert.equal(view.root.findAllByProps({ className: 'ribbon-chevron' }).length, 1, 'the chevron is part of the button, not a native colour input');
  await act(async () => trigger.props.onClick());
  assert.equal(view.root.findByProps({ 'aria-label': 'Green' }).props['aria-pressed'], true);
  await act(async () => view.root.findByProps({ 'aria-label': 'Red' }).props.onClick());
  assert.deepEqual(picked, ['#B33C3C']);
  assert.equal(view.root.findByProps({ 'aria-label': 'Fill color' }).props['aria-expanded'], false, 'picking closes the palette');
  await act(async () => view.unmount());
});

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const React = require('react');
const { create, act } = require('react-test-renderer');
global.IS_REACT_ACT_ENVIRONMENT = true;

function isExternal(id) {
  return /^react(?:\/|$)/.test(id) || id === '@injoffice/xlsx-wasm' || id === '@injoffice/sheets/browser' || id.includes('packages/formulas/');
}

function adaptWorkbookMutationBatchV1(workbook, batch) {
  const operations = batch.operations ?? [];
  const cells = operations.filter(operation => String(operation.kind).startsWith('cell.'));
  const structure = operations.filter(operation => ['row.insert','row.delete','column.insert','column.delete'].includes(operation.kind));
  const merges = operations.filter(operation => ['range.merge', 'range.unmerge'].includes(operation.kind));
  const styles = operations.filter(operation => operation.kind === 'style.patch');
  return { expected_revision: workbook.revision, ...(structure.length ? {structure} : {}), ...(merges.length ? { merges } : {}), ...(cells.length ? { cells } : {}), ...(styles.length ? { styles } : {}) };
}

function stubRequire(id) {
  if (id === '@injoffice/xlsx-wasm') {
    return {
      createXlsxWasmClient: () => globalThis.__xlsxClient,
      adaptWorkbookMutationBatchV1,
    };
  }
  if (id === '@injoffice/sheets/browser') {
    return {
      formatNativeSheetCellDisplayV2: () => { throw new Error('WASM display is mocked'); },
      editableDefinedName: name => !String(name).startsWith('_'),
    };
  }
  return require(id);
}

async function loadEditor() {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src/SpreadsheetEditor.tsx'),
    platform: 'node',
    external: isExternal,
    transform: { jsx: { runtime: 'automatic' } },
    plugins: [{
      name: 'spreadsheet-editor-test',
      resolveId(id) {
        if (id.endsWith('.css')) return '\0css';
        if (id.includes('spreadsheetCalculation.worker')) return '\0worker';
      },
      load(id) {
        if (id === '\0css') return 'export default ""';
        if (id === '\0worker') return 'export default ""';
      },
    }],
  });
  try {
    const { output } = await bundle.generate({ format: 'cjs', codeSplitting: false });
    const mod = { exports: {} };
    new Function('require', 'module', 'exports', output[0].code)(stubRequire, mod, mod.exports);
    return mod.exports.SpreadsheetEditor ?? mod.exports.default ?? mod.exports;
  } finally {
    await bundle.close();
  }
}

function workbook() {
  return {
    revision: 'rev-1',
    source: { package_sha256: 'rev-1' },
    styles: [],
    unsupported: [],
    defined_names: [],
    sheets: [{
      id: '1',
      name: 'Sheet1',
      state: 'visible',
      editable: true,
      merged_ranges: [],
      cells: [],
      rows: [],
      columns: [],
    }],
  };
}

function mockClient() {
  const applied = [];
  const model = workbook();
  return {
    applied,
    extract: async () => structuredClone(model),
    apply: async (_bytes, _workbook, request) => {
      applied.push(request);
      if (request.merges?.length) {
        const merge = request.merges[0];
        model.sheets[0].merged_ranges = merge.kind === 'range.merge' ? [{...merge.range,ref:'A1:B2',editable:false}] : [];
      }
      const operation = request.cells?.[0];
      if (operation?.kind === 'cell.set_value') {
        model.sheets[0].cells = [{ row: operation.cell.row, column: operation.cell.column, value: { kind: 'string', text: String(operation.value) } }];
        model.revision = 'rev-2';
        model.source.package_sha256 = 'rev-2';
      }
      return new Uint8Array([2]);
    },
    terminate() { applied.push('terminate'); },
  };
}

async function until(predicate) {
  for (let count = 0; count < 200; count++) {
    if (predicate()) return;
    await act(async () => new Promise(resolve => setTimeout(resolve, 5)));
  }
  throw new Error('Editor did not settle');
}

const text = node => node.children.map(child => typeof child === 'string' ? child : text(child)).join('');
function button(view, label) {
  return view.root.findAllByType('button').find(node => node.props.children === label || text(node) === label);
}

test('SpreadsheetEditor mounts, reports busy, and applies a cell edit', async () => {
  assert.equal(fs.existsSync(path.resolve(__dirname, '../src/spreadsheet.css')), true);
  const source = fs.readFileSync(path.resolve(__dirname, '../src/SpreadsheetEditor.tsx'), 'utf8');
  const calculateSnapshot = source.slice(source.indexOf('function calculateSnapshot'), source.indexOf('function exportSheet'));
  assert.equal(/from ['"]\.\/OfficeEditor['"]/.test(source), false);
  assert.equal(/\breadCharts\b/.test(source), false);
  assert.equal(/definedNameCaseKey\s+from\s+['"]@injoffice\/sheets\/browser['"]/.test(source), false);
  assert.equal(/\.apply\(/.test(calculateSnapshot), false);

  const client = mockClient();
  globalThis.__xlsxClient = client;
  const SpreadsheetEditor = await loadEditor();
  const busy = [];
  const changes = [];
  let view;
  try {
    await act(async () => {
      view = create(React.createElement(SpreadsheetEditor, {
        name: 'Book.xlsx',
        bytes: new Uint8Array([1]),
        onChange: value => changes.push(value),
        onBusyChange: value => busy.push(value),
      }));
    });
    await until(() => busy.at(-1) === false && view.root.findAllByProps({ className: 'sheet-grid' }).length > 0);
    assert.equal(view.root.findByProps({ 'aria-label': 'Book.xlsx spreadsheet' }) != null, true);
    assert.equal(busy.includes(true), true);
    assert.equal(busy.at(-1), false);

    const formula = view.root.findByProps({ 'aria-label': 'Cell value or formula' });
    await act(async () => formula.props.onChange({ target: { value: 'Hello' } }));
    const apply = button(view, 'Apply');
    assert.equal(apply.props.disabled, false);
    await act(async () => apply.props.onClick());
    await until(() => changes.length === 1 && busy.at(-1) === false && client.applied.length >= 1);
    const transaction = client.applied[0];
    assert.equal(transaction.operations, undefined);
    assert.equal(transaction.calculation, undefined);
    assert.equal(transaction.expected_revision, 'rev-1');
    assert.equal(transaction.cells[0].kind, 'cell.set_value');
    assert.equal(transaction.cells[0].value, 'Hello');
    assert.deepEqual([...changes[0]], [2]);
  } finally {
    if (view) await act(async () => view.unmount());
    delete globalThis.__xlsxClient;
  }
});

test('SpreadsheetEditor arranges its controls as an Excel ribbon with labelled groups and icons', async () => {
  const client = mockClient();
  globalThis.__xlsxClient = client;
  const SpreadsheetEditor = await loadEditor();
  const busy = [];
  let view;
  try {
    await act(async () => {
      view = create(React.createElement(SpreadsheetEditor, { name: 'Book.xlsx', bytes: new Uint8Array([1]), onChange: () => {}, onBusyChange: value => busy.push(value) }));
    });
    await until(() => busy.at(-1) === false && view.root.findAllByProps({ className: 'sheet-grid' }).length > 0);
    assert.deepEqual(view.root.findAllByProps({ role: 'tab' }).map(text), ['File', 'Home', 'Insert', 'Formulas', 'Data', 'View']);
    const panels = view.root.findAllByProps({ role: 'tabpanel' });
    const groups = panel => panel.findAllByProps({ role: 'group' }).map(group => group.props['aria-label']);
    assert.deepEqual(groups(panels[0]), ['Export']);
    assert.deepEqual(groups(panels[1]), ['Font', 'Alignment', 'Number', 'Cells']);
    assert.deepEqual(groups(panels[2]), ['Charts']);
    assert.deepEqual(groups(panels[3]), ['Defined Names', 'Calculation']);
    assert.deepEqual(groups(panels[4]), ['Sort & Filter']);
    assert.deepEqual(groups(panels[5]), ['Window']);
    assert.equal(view.root.findAllByProps({ role: 'toolbar' }).filter(toolbar => toolbar.props['aria-label'] === 'Quick access').length, 0, 'Undo/Redo live in the shell title bar, not in the ribbon');
    const commandButtons = view.root.findAllByProps({ role: 'group' }).flatMap(group => group.findAllByType('button'));
    assert.ok(commandButtons.length >= 18, `command buttons: ${commandButtons.length}`);
    for (const node of commandButtons) {
      assert.equal(node.findAllByType('svg').length, 1, `${node.props.title} has an icon`);
      assert.ok(node.props.title, 'every command button has a tooltip');
    }
    for (const summary of view.root.findAllByType('summary')) assert.equal(summary.findAllByType('svg').length, 1, `${text(summary)} popover has an icon`);
    const byLabel = Object.fromEntries(commandButtons.map(node => [node.props['aria-label'] ?? text(node), node]));
    assert.equal(byLabel.Charts.props.disabled, true);
    assert.match(byLabel.Charts.props.title, /not supported by the native XLSX transaction/);
    assert.match(byLabel['Sort range'].props.title, /not supported by the native XLSX transaction/);
    assert.equal(byLabel.Recalculate.props.disabled, false);
    assert.equal(view.root.findByProps({ 'aria-label': 'Cell value or formula' }) != null, true, 'name box and fx bar are untouched');
  } finally {
    if (view) await act(async () => view.unmount());
    delete globalThis.__xlsxClient;
  }
});

test('the sheet area is the only scroller and renders rows continuously instead of paging', async () => {
  const css = fs.readFileSync(path.resolve(__dirname, '../src/spreadsheet.css'), 'utf8');
  const rule = name => css.match(new RegExp(`\\${name}\\s*\\{([^}]*)\\}`))[1];
  assert.match(rule('.sheet-editor'), /height:\s*100%/, 'the editor is bounded by the workspace');
  assert.match(rule('.sheet-content'), /flex:\s*1/);
  assert.match(rule('.sheet-content'), /min-height:\s*0/, 'the sheet area is a height-bounded flex child');
  assert.match(rule('.sheet-grid-scroll'), /overflow:\s*auto/);
  assert.match(rule('.sheet-grid-scroll'), /min-height:\s*0/);
  assert.match(rule('.sheet-bottom'), /flex-shrink:\s*0/, 'sheet tabs stay pinned at the bottom');
  assert.match(rule('.sheet-status'), /flex-shrink:\s*0/, 'the status row stays pinned at the bottom');
  assert.match(css, /\.sheet-grid thead th\{position:sticky;top:0/, 'column headers pin to the top');
  assert.match(css, /\.sheet-grid tbody th\{position:sticky;left:0/, 'row headers pin to the left');
  assert.equal(/sheet-window-controls/.test(css), false, 'the row pager is gone');

  const client = mockClient();
  globalThis.__xlsxClient = client;
  const SpreadsheetEditor = await loadEditor();
  const busy = [];
  let view;
  try {
    await act(async () => {
      view = create(React.createElement(SpreadsheetEditor, { name: 'Book.xlsx', bytes: new Uint8Array([1]), onChange: () => {}, onBusyChange: value => busy.push(value) }));
    });
    await until(() => busy.at(-1) === false && view.root.findAllByProps({ className: 'sheet-grid' }).length > 0);
    const rows = view.root.findAllByType('tr').filter(node => node.props['data-sheet-row'] !== undefined);
    assert.ok(rows.length >= 200, `initial rows rendered: ${rows.length}`);
    assert.equal(rows[0].props['data-sheet-row'], 0);
    for (const label of ['Next rows', 'Previous rows', 'Next columns', 'Previous columns']) {
      assert.equal(view.root.findAllByProps({ 'aria-label': label }).length, 0, `${label} pager button is gone`);
    }
    assert.equal(view.root.findAllByType('div').some(node => typeof node.props.onScroll === 'function' && node.props.className === 'sheet-grid-scroll'), true, 'the grid grows as it is scrolled');
  } finally {
    if (view) await act(async () => view.unmount());
    delete globalThis.__xlsxClient;
  }
});

test('typing edits the selected cell, Enter commits and moves down, and the status row reads like Excel', async () => {
  const css = fs.readFileSync(path.resolve(__dirname, '../src/spreadsheet.css'), 'utf8');
  assert.match(css, /\.sheet-grid th\.is-active\{/, 'the active row and column headers highlight');

  const client = mockClient();
  globalThis.__xlsxClient = client;
  const SpreadsheetEditor = await loadEditor();
  const busy = [];
  let view;
  const status = () => view.root.findByProps({ 'aria-label': 'Spreadsheet status' }).findAllByType('span').map(text);
  const grid = () => view.root.findByProps({ className: 'sheet-grid-scroll' });
  const press = async (key, extra = {}) => {
    const node = grid();
    await act(async () => node.props.onKeyDown({ key, target: node, currentTarget: node, nativeEvent: {}, preventDefault() {}, ...extra }));
  };
  try {
    await act(async () => {
      view = create(React.createElement(SpreadsheetEditor, { name: 'Book.xlsx', bytes: new Uint8Array([1]), onChange: () => {}, onBusyChange: value => busy.push(value) }));
    });
    await until(() => busy.at(-1) === false && view.root.findAllByProps({ className: 'sheet-grid' }).length > 0);
    assert.equal(status()[0], 'Ready');
    assert.equal(status()[1], '', 'no shortcut sentence in the status row');
    assert.equal(status().some(value => value.includes('Formula caches')), false, 'the caches note is a tooltip, not a second status line');
    assert.match(view.root.findByProps({ className: 'sheet-calculation' }).props.title, /Formula caches/);

    // A printable key starts the entry with that keystroke: Excel's "Enter" mode.
    await press('R');
    await until(() => view.root.findAllByProps({ className: 'sheet-inline-input' }).length === 1);
    assert.equal(view.root.findByProps({ className: 'sheet-inline-input' }).props.value, 'R');
    assert.equal(status()[0], 'Enter');

    // Enter commits the entry and moves down; the name box follows the active cell.
    const input = view.root.findByProps({ className: 'sheet-inline-input' });
    await act(async () => input.props.onKeyDown({ key: 'Enter', nativeEvent: {}, preventDefault() {} }));
    await until(() => client.applied.length >= 1 && busy.at(-1) === false && view.root.findAllByProps({ className: 'sheet-inline-input' }).length === 0);
    assert.equal(client.applied[0].cells[0].value, 'R');
    assert.equal(view.root.findByProps({ 'aria-label': 'Cell or range address' }).props.value, 'A2');
    assert.equal(status()[0], 'Ready');

    // F2 edits in place: Excel's "Edit" mode, seeded with the stored text.
    await press('F2');
    await until(() => view.root.findAllByProps({ className: 'sheet-inline-input' }).length === 1);
    assert.equal(status()[0], 'Edit');
    await act(async () => view.root.findByProps({ className: 'sheet-inline-input' }).props.onKeyDown({ key: 'Escape', nativeEvent: {}, preventDefault() {} }));
    assert.equal(status()[0], 'Ready');

    // Enter on a selected cell moves down instead of opening an editor.
    await press('Enter');
    assert.equal(view.root.findAllByProps({ className: 'sheet-inline-input' }).length, 0);
    assert.equal(view.root.findByProps({ 'aria-label': 'Cell or range address' }).props.value, 'A3');
    const headers = view.root.findAllByType('th').filter(node => node.props.className === 'is-active').map(text);
    assert.deepEqual(headers, ['A', '3'], 'the active row and column headers are marked');
  } finally {
    if (view) await act(async () => view.unmount());
    delete globalThis.__xlsxClient;
  }
});

test('the sheet follows the dark theme and renders cell text in a sans stack', () => {
  const css = fs.readFileSync(path.resolve(__dirname, '../src/spreadsheet.css'), 'utf8');
  const editor = css.match(/\.sheet-editor\{([^}]*)\}/)[1];
  for (const token of ['--sheet-cell-surface', '--sheet-cell-text', '--sheet-gridline', '--sheet-cell-font']) {
    assert.ok(editor.includes(`${token}:`), `${token} has a light value on .sheet-editor`);
  }
  assert.match(editor, /--sheet-cell-font:[^;]*sans-serif/, 'cell text falls back to a sans face, never the serif default');
  const systemDark = css.match(/@media\(prefers-color-scheme:dark\)\{:root:not\(\[data-theme="light"\]\) \.sheet-editor\{([^}]*)\}/);
  const forcedDark = css.match(/:root\[data-theme="dark"\] \.sheet-editor\{([^}]*)\}/);
  assert.ok(systemDark && forcedDark, 'both dark blocks re-point the sheet tokens');
  assert.equal(systemDark[1], forcedDark[1], 'the dark blocks agree');
  for (const token of ['--sheet-cell-surface', '--sheet-cell-text', '--sheet-gridline']) {
    assert.ok(systemDark[1].includes(`${token}:`), `${token} is darkened`);
  }
  assert.match(css, /\.sheet-grid\{[^}]*background:var\(--sheet-cell-surface\)/);
  assert.match(css, /\.sheet-grid\{[^}]*color:var\(--sheet-cell-text\)/);
  assert.match(css, /\.sheet-grid th,\.sheet-grid td\{[^}]*border-right:1px solid var\(--sheet-gridline\)/);

  const source = fs.readFileSync(path.resolve(__dirname, '../src/SpreadsheetEditor.tsx'), 'utf8');
  assert.equal(/'#fff'/.test(source), false, 'pinned cells paint on the sheet surface token');
  assert.match(source, /var\(--sheet-cell-font\)/, 'authored cell fonts keep a sans fallback');
});

test('XLSX Home uses Excel icon controls: six alignment toggles, colour buttons, Borders and Cells menus', async () => {
  const client = mockClient();
  globalThis.__xlsxClient = client;
  const SpreadsheetEditor = await loadEditor();
  const busy = [];
  let view;
  try {
    await act(async () => {
      view = create(React.createElement(SpreadsheetEditor, { name: 'Book.xlsx', bytes: new Uint8Array([1]), onChange: () => {}, onBusyChange: value => busy.push(value) }));
    });
    await until(() => busy.at(-1) === false && view.root.findAllByProps({ className: 'sheet-grid' }).length > 0);
    const home = view.root.findAllByProps({ role: 'tabpanel' })[1];
    const group = label => home.findByProps({ role: 'group', 'aria-label': label });

    const alignment = group('Alignment');
    const toggles = alignment.findAllByType('button').map(node => node.props['aria-label']);
    assert.deepEqual(toggles, ['Top align', 'Middle align', 'Bottom align', 'Wrap text', 'Align left', 'Center', 'Align right', 'Merge cells']);
    assert.equal(alignment.findAllByType('select').length, 0, 'no alignment dropdowns');
    const byLabel = Object.fromEntries(alignment.findAllByType('button').map(node => [node.props['aria-label'], node]));
    assert.equal(byLabel['Bottom align'].props['aria-pressed'], true, 'the stored vertical alignment reads back');
    assert.equal(byLabel['Align left'].props['aria-pressed'], false);
    await act(async () => byLabel.Center.props.onClick());
    await until(() => client.applied.length >= 1);
    assert.equal(client.applied[0].styles[0].style.horizontal_alignment, 'center');

    const font = group('Font');
    const colors = font.findAllByType('input').filter(node => node.props.type === 'color');
    assert.deepEqual(colors.map(node => node.props['aria-label']), ['Text color', 'Fill color', 'Border color']);
    assert.equal(font.findAllByProps({ className: 'sheet-color-underline' }).length, 2, 'each colour button shows its colour underline');
    const borders = font.findByProps({ 'aria-label': 'Borders' });
    assert.equal(borders.findAllByType('svg').length, 1, 'Borders is an icon control');
    assert.equal(borders.props.title, 'Borders');

    const cells = group('Cells');
    const cellCommands = [...cells.findAllByType('button'), ...cells.findAllByType('summary')].map(node => node.props['aria-label']).filter(Boolean);
    assert.deepEqual(cellCommands.slice(0, 3), ['Insert cells', 'Delete cells', 'Format cells']);
    for (const node of [...cells.findAllByType('button'), ...cells.findAllByType('summary')]) assert.equal(node.findAllByType('svg').length >= 1, true, 'Cells commands are icons');
    assert.equal(cells.findAllByType('select').length, 0, 'no Rows & columns / Cell size dropdowns');
  } finally {
    if (view) await act(async () => view.unmount());
    delete globalThis.__xlsxClient;
  }
});

test('sheet-state notes live in the status row, not as captions over the cells', async () => {
  const source = fs.readFileSync(path.resolve(__dirname, '../src/SpreadsheetEditor.tsx'), 'utf8');
  assert.equal(/sheet-filter-status/.test(source), false, 'the filter caption above the grid is gone');
  assert.equal(/exceeds the display limit[^`']*<\/div>/.test(source), false, 'the frozen-pane caption is no longer a banner');

  const client = mockClient();
  const model = { ...structuredClone(await client.extract()) };
  model.sheets[0].auto_filter = { ref: 'A1:C9' };
  model.sheets[0].frozen_rows = 99;
  client.extract = async () => structuredClone(model);
  globalThis.__xlsxClient = client;
  const SpreadsheetEditor = await loadEditor();
  const busy = [];
  let view;
  try {
    await act(async () => {
      view = create(React.createElement(SpreadsheetEditor, { name: 'Book.xlsx', bytes: new Uint8Array([1]), onChange: () => {}, onBusyChange: value => busy.push(value) }));
    });
    await until(() => busy.at(-1) === false && view.root.findAllByProps({ className: 'sheet-grid' }).length > 0);
    const notes = view.root.findAllByProps({ className: 'sheet-note' });
    assert.deepEqual(notes.map(text), ['Filter Mode', 'Panes not pinned']);
    assert.match(notes[0].props.title, /Text filter on A1:C9/);
    assert.match(notes[1].props.title, /exceeds the display limit/);
    const sheetArea = view.root.findByProps({ className: 'sheet-content' });
    assert.deepEqual(sheetArea.findAllByType('p'), [], 'nothing is captioned inside the sheet area');
  } finally {
    if (view) await act(async () => view.unmount());
    delete globalThis.__xlsxClient;
  }
});

test('Merge and Unmerge submit native range mutations, including a selected merged anchor', async () => {
 const client=mockClient(); globalThis.__xlsxClient=client;
 const Editor=await loadEditor(); let view;
 try {
  await act(async()=>{view=create(React.createElement(Editor,{name:'Merge.xlsx',bytes:new Uint8Array([1]),onChange:()=>{}}));});
  await until(()=>view.root.findAllByProps({'aria-label':'Cell or range address'}).length>0);
  const location=()=>view.root.findByProps({'aria-label':'Cell or range address'});
  await act(async()=>location().props.onChange({target:{value:'A1:B2'}}));
  await act(async()=>location().parent.props.onSubmit({preventDefault(){}}));
  const button=label=>view.root.findAllByType('button').find(node=>node.props['aria-label']===label);
  assert.equal(button('Merge cells').props.disabled,false);
  await act(async()=>button('Merge cells').props.onClick());
  await until(()=>client.applied.length===1);
  assert.equal(client.applied[0].merges[0].kind,'range.merge');
  assert.equal(button('Unmerge cells').props.disabled,false);
  await act(async()=>button('Unmerge cells').props.onClick());
  await until(()=>client.applied.length===2);
  assert.deepEqual(client.applied[1].merges[0].range,{row:0,column:0,end_row:1,end_column:1});
  assert.equal(client.applied[1].merges[0].kind,'range.unmerge');
 } finally {if(view)await act(async()=>view.unmount());delete globalThis.__xlsxClient;}
});

test('Rows and columns menu confirms whole-axis mutations for the selected range', async () => {
 const client=mockClient();globalThis.__xlsxClient=client;const Editor=await loadEditor();let view;
 try {
  await act(async()=>{view=create(React.createElement(Editor,{name:'Rows.xlsx',bytes:new Uint8Array([1]),onChange:()=>{}}));});
  await until(()=>view.root.findAllByProps({'aria-label':'Cell or range address'}).length>0);
  const location=()=>view.root.findByProps({'aria-label':'Cell or range address'});
  await act(async()=>location().props.onChange({target:{value:'B2:C3'}}));
  await act(async()=>location().parent.props.onSubmit({preventDefault(){}}));
  for (const [label,kind] of [['Insert sheet rows','row.insert'],['Delete sheet columns','column.delete']]) {
   const component=view.root.findByProps({label});
   assert.equal(component.props.disabled,false);
   await act(async()=>component.props.onClick());
   await act(async()=>view.root.findByProps({'aria-label':'Row and column changes'}).props.onSubmit({preventDefault(){}}));
   await until(()=>client.applied.some(value=>value.structure?.[0]?.kind===kind));
   assert.equal(client.applied.at(-1).structure[0].index,1);
   assert.equal(client.applied.at(-1).structure[0].count,2);
  }
 } finally {if(view)await act(async()=>view.unmount());delete globalThis.__xlsxClient;}
});

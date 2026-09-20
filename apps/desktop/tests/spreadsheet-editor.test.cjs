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
  const cells = (batch.operations ?? []).filter(operation => String(operation.kind).startsWith('cell.'));
  return { expected_revision: workbook.revision, ...(cells.length ? { cells } : {}) };
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

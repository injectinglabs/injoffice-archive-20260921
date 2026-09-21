const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const React = require('react');
const { renderToStaticMarkup } = require('react-dom/server');
const { create, act } = require('react-test-renderer');
global.IS_REACT_ACT_ENVIRONMENT = true;

function isExternal(id) {
  return /^react(?:\/|$)/.test(id) || id === '@injoffice/xlsx-wasm' || id === '@injoffice/sheets/browser';
}

function stubRequire(id) {
  if (id === '@injoffice/xlsx-wasm' || id === '@injoffice/sheets/browser') {
    return { formatNativeSheetCellDisplayV2: () => { throw new Error('WASM display is mocked'); } };
  }
  return require(id);
}

async function loadCharts() {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src/SpreadsheetCharts.tsx'),
    platform: 'node',
    external: isExternal,
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
    new Function('require', 'module', 'exports', output[0].code)(stubRequire, mod, mod.exports);
    return mod.exports;
  } finally {
    await bundle.close();
  }
}

function source() {
  return {
    identity: { part: 'xl/charts/chart1.xml', drawingPart: 'xl/drawings/drawing1.xml', objectId: 1 },
    sheet_id: '1',
    fingerprint_sha256: 'sha256:' + '1'.repeat(64),
    chart_type: 'column',
    title: '<Revenue & cost>',
    range: { row: 0, column: 0, end_row: 2, end_column: 1 },
    anchor: { from_row: 1, from_column: 3, to_row: 18, to_column: 10 },
    editable: true,
    categories: ['First', 'Second'],
    series: [{ name: 'Value', values: ['12', '-3'] }],
  };
}

function button(view, text) {
  return view.root.findAllByType('button').find(value => value.children.join('') === text);
}

test('empty range disables insert; a valid range emits chart.insert', async () => {
  const { SpreadsheetCharts } = await loadCharts();
  const executed = [];
  const empty = { row: 0, column: 0, end_row: 0, end_column: 0 };
  const range = { row: 0, column: 0, end_row: 2, end_column: 2 };
  const props = {
    charts: [], sheetId: '1', disabled: false,
    onExecute: (operations, message) => { executed.push({ operations, message }); return Promise.resolve(true); },
    onClose() {},
  };
  let view;
  await act(async () => {
    view = create(React.createElement(SpreadsheetCharts, { ...props, range: empty }));
  });
  const disabled = button(view, 'Insert selected range');
  assert.equal(disabled.props.disabled, true);
  assert.match(disabled.props.title, /header row/);
  assert.deepEqual(executed, []);

  await act(async () => {
    view.update(React.createElement(SpreadsheetCharts, { ...props, range }));
  });
  const insert = button(view, 'Insert selected range');
  assert.equal(insert.props.disabled, false);
  await act(async () => insert.props.onClick());
  assert.equal(executed.length, 1);
  assert.equal(executed[0].message, 'Chart inserted into workbook');
  assert.deepEqual(executed[0].operations, [{
    kind: 'chart.insert',
    chart_type: 'column',
    title: '',
    range,
    anchor: { from_row: 0, from_column: 4, to_row: 18, to_column: 12 },
  }]);
  await act(async () => view.unmount());
});

test('SpreadsheetChartPreview paints series SVG; refused charts do not', async () => {
  const { SpreadsheetChartPreview, SpreadsheetCharts } = await loadCharts();
  for (const kind of ['column', 'bar', 'line', 'pie']) {
    const chart = source();
    chart.chart_type = kind;
    const svg = renderToStaticMarkup(React.createElement(SpreadsheetChartPreview, { chart }));
    assert.match(svg, /class="sheet-chart-preview"/);
    assert.ok(!/NaN|Infinity|undefined/.test(svg));
    assert.match(svg, /&lt;Revenue &amp; cost&gt;/);
    assert.equal((svg.match(/<rect /g) || []).length, kind === 'line' || kind === 'pie' ? 0 : 2);
    if (kind === 'line') assert.match(svg, /<path[^>]+d="M[^\"]+ L/);
    if (kind === 'pie') assert.match(svg, /<path[^>]+d="M/);
  }
  const chart = source();
  chart.editable = false;
  chart.categories = [];
  chart.series = [];
  chart.refusal = 'Chart values must be finite literal numbers';
  const html = renderToStaticMarkup(React.createElement(SpreadsheetCharts, {
    charts: [chart], sheetId: '1', range: chart.range, disabled: false,
    onExecute: async () => true, onClose() {},
  }));
  assert.match(html, /class="sheet-charts"/);
  assert.match(html, /preserved in the workbook/);
  assert.ok(!html.includes('<svg'));
});

test('title, type, selected range, and delete emit chart.update or chart.delete', async () => {
  const { SpreadsheetCharts } = await loadCharts();
  const chart = source();
  const range = { row: 3, column: 1, end_row: 6, end_column: 3 };
  const calls = [];
  const props = {
    charts: [chart], sheetId: '1', range, disabled: false,
    onExecute: async (...args) => { calls.push(args); return true; },
    onClose() {},
  };
  let view;
  try {
    await act(async () => {
      view = create(React.createElement(SpreadsheetCharts, props));
    });
    assert.equal(view.root.findAllByType('svg').length, 1);

    const title = view.root.findAllByType('input').find(value => value.props['aria-label'] === 'Chart title');
    await act(async () => title.props.onChange({ target: { value: 'New title' } }));
    await act(async () => view.root.findByProps({ className: 'sheet-chart-title' }).props.onSubmit({ preventDefault() {} }));
    assert.equal(calls[0][1], 'Chart title updated');
    assert.deepEqual(calls[0][0], [{
      kind: 'chart.update',
      identity: chart.identity,
      expected_fingerprint_sha256: chart.fingerprint_sha256,
      chart_type: 'column',
      title: 'New title',
      range: chart.range,
      anchor: chart.anchor,
    }]);

    const type = view.root.findAllByType('select').find(value => value.props['aria-label'] === 'Chart 1 type');
    await act(async () => type.props.onChange({ target: { value: 'bar' } }));
    assert.equal(calls[1][1], 'Chart type updated');
    assert.equal(calls[1][0][0].kind, 'chart.update');
    assert.equal(calls[1][0][0].chart_type, 'bar');
    assert.deepEqual(calls[1][0][0].identity, chart.identity);

    await act(async () => button(view, 'Use selected range').props.onClick());
    assert.equal(calls[2][1], 'Chart source updated');
    assert.equal(calls[2][0][0].kind, 'chart.update');
    assert.deepEqual(calls[2][0][0].range, range);
    assert.deepEqual(calls[2][0][0].identity, chart.identity);
    assert.equal(calls[2][0][0].expected_fingerprint_sha256, chart.fingerprint_sha256);
    assert.deepEqual(calls[2][0][0].anchor, chart.anchor);

    await act(async () => button(view, 'Delete chart').props.onClick());
    assert.equal(calls[3][1], 'Chart deleted');
    assert.deepEqual(calls[3][0], [{
      kind: 'chart.delete',
      identity: chart.identity,
      expected_fingerprint_sha256: chart.fingerprint_sha256,
    }]);

    const revised = { ...chart, title: 'Committed title', fingerprint_sha256: 'sha256:' + '2'.repeat(64) };
    await act(async () => view.update(React.createElement(SpreadsheetCharts, { ...props, charts: [revised] })));
    assert.equal(view.root.findAllByType('input').find(value => value.props['aria-label'] === 'Chart title').props.value, 'Committed title');

    await act(async () => view.update(React.createElement(SpreadsheetCharts, { ...props, disabled: true })));
    assert.equal(button(view, 'Insert selected range').props.disabled, true);
    assert.equal(button(view, 'Use selected range').props.disabled, true);
    assert.equal(button(view, 'Delete chart').props.disabled, true);
    assert.equal(view.root.findAllByType('select').find(value => value.props['aria-label'] === 'Chart 1 type').props.disabled, true);
  } finally {
    if (view) await act(async () => view.unmount());
  }
});

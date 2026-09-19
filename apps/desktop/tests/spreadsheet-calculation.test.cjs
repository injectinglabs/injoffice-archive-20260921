const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { pathToFileURL } = require('node:url');

const ENGINE = 'univer-oss/0.25.1/injoffice-scalar-v1';

async function loadCalc() {
  return import(pathToFileURL(path.resolve(__dirname, '../src/spreadsheetCalculation.ts')).href);
}

function workbook(cells, extra = {}) {
  return { revision: 'rev-1', unsupported: [], sheets: [{ id: '1', name: 'Sheet1', cells }], ...extra };
}

function installWorker(t, handler) {
  const instances = [];
  class StubWorker {
    constructor(url, options) {
      this.url = url;
      this.options = options;
      this.terminated = false;
      instances.push(this);
    }
    postMessage(data) {
      queueMicrotask(() => { if (!this.terminated) handler(this, data); });
    }
    terminate() { this.terminated = true; }
  }
  const previous = globalThis.Worker;
  globalThis.Worker = StubWorker;
  t.after(() => { globalThis.Worker = previous; });
  return instances;
}

test('calculationInput projects source values and excludes saved formula caches', async () => {
  const { calculationInput } = await loadCalc();
  const input = calculationInput(workbook([
    { row: 0, column: 0, ref: 'A1', value: { kind: 'number', lexical: '2' } },
    { row: 0, column: 1, ref: 'B1', formula: { type: 'normal', text: 'A1*3', cached: { kind: 'number', lexical: '6' } } },
    { row: 0, column: 2, ref: 'C1', value: { kind: 'boolean', lexical: 'true' } },
    { row: 0, column: 3, ref: 'D1', value: { kind: 'boolean', lexical: '0' } },
    { row: 0, column: 4, ref: 'E1', value: { kind: 'string', text: 'hi' } },
    { row: 0, column: 5, ref: 'F1', value: { kind: 'error', lexical: '#DIV/0!' } },
    { row: 0, column: 6, ref: 'G1', value: { kind: 'date', lexical: '2020-01-01' } },
    { row: 0, column: 7, ref: 'H1' },
  ]));
  assert.deepEqual(input, {
    revision: 'rev-1',
    sheets: [{ id: '1', name: 'Sheet1', cells: [
      { row: 0, column: 0, value: { kind: 'number', value: 2 } },
      { row: 0, column: 1, formula: '=A1*3' },
      { row: 0, column: 2, value: { kind: 'boolean', value: true } },
      { row: 0, column: 3, value: { kind: 'boolean', value: false } },
      { row: 0, column: 4, value: { kind: 'string', value: 'hi' } },
      { row: 0, column: 5, value: { kind: 'error', value: '#DIV/0!' } },
      { row: 0, column: 6, unsupported: 'G1: date storage is not qualified for local calculation.' },
      { row: 0, column: 7 },
    ] }],
  });
  assert.equal(input.sheets[0].cells[1].value, undefined);
  assert.throws(() => calculationInput(workbook([{ row: 0, column: 0, ref: 'A1', formula: { type: 'shared', text: 'A1' } }])), /shared, array, or data-table/);
  assert.throws(() => calculationInput(workbook([], { unsupported: [{ capability: 'formula-groups' }] })), /shared, array, or data-table/);
});

test('calculationCache is revision-checked and nulls unverified results', async () => {
  const { calculationCache } = await loadCalc();
  const result = {
    sourceRevision: 'rev-1',
    engine: ENGINE,
    cells: [
      { sheetId: '1', row: 0, column: 1, formula: '=A1*3', status: 'calculated', value: { kind: 'number', value: 6 } },
      { sheetId: '1', row: 0, column: 2, formula: '=B1/0', status: 'error', value: { kind: 'error', value: '#DIV/0!' } },
      { sheetId: '1', row: 0, column: 3, formula: '=C1', status: 'circular' },
      { sheetId: '1', row: 0, column: 4, formula: '=UNSUPPORTED(A1)', status: 'unsupported', value: { kind: 'number', value: 1 } },
    ],
  };
  assert.deepEqual(calculationCache(workbook([]), result), {
    engine: ENGINE,
    cells: [
      { sheet_id: '1', cell: { row: 0, column: 1 }, expected_formula: '=A1*3', value: { kind: 'number', value: 6 } },
      { sheet_id: '1', cell: { row: 0, column: 2 }, expected_formula: '=B1/0', value: { kind: 'error', value: '#DIV/0!' } },
      { sheet_id: '1', cell: { row: 0, column: 3 }, expected_formula: '=C1', value: null },
      { sheet_id: '1', cell: { row: 0, column: 4 }, expected_formula: '=UNSUPPORTED(A1)', value: null },
    ],
  });
  assert.throws(() => calculationCache(workbook([], { revision: 'rev-2' }), result), /older workbook revision/);
});

test('createSpreadsheetCalculator runs through a stub Worker and refuses overlap', async t => {
  const { createSpreadsheetCalculator } = await loadCalc();
  const constructed = createSpreadsheetCalculator();
  assert.equal(typeof constructed.calculate, 'function');
  assert.equal(typeof constructed.terminate, 'function');
  constructed.terminate();

  const input = { revision: 'rev-1', sheets: [{ id: '1', name: 'Sheet1', cells: [] }] };
  let release;
  const gate = new Promise(resolve => { release = resolve; });
  const workers = installWorker(t, (worker, data) => {
    Promise.resolve(gate).then(() => {
      if (!worker.terminated) worker.onmessage({ data: { result: { sourceRevision: data.revision, engine: ENGINE, cells: [] } } });
    });
  });
  const calculator = createSpreadsheetCalculator();
  t.after(() => calculator.terminate());
  const first = calculator.calculate(input);
  assert.equal(workers.length, 1);
  assert.match(String(workers[0].url), /spreadsheetCalculation\.worker\.ts/);
  assert.equal(workers[0].options.type, 'module');
  await assert.rejects(calculator.calculate(input), /already running/);
  release();
  assert.deepEqual(await first, { sourceRevision: 'rev-1', engine: ENGINE, cells: [] });
});

test('createSpreadsheetCalculator surfaces worker errors and cancel', async t => {
  const { createSpreadsheetCalculator } = await loadCalc();
  const input = { revision: 'rev-1', sheets: [{ id: '1', name: 'Sheet1', cells: [] }] };

  const failing = createSpreadsheetCalculator();
  t.after(() => failing.terminate());
  installWorker(t, worker => { worker.onmessage({ data: { error: 'engine refused' } }); });
  await assert.rejects(failing.calculate(input), /engine refused/);

  const cancelled = createSpreadsheetCalculator();
  installWorker(t, () => {});
  const pending = cancelled.calculate(input);
  cancelled.terminate();
  await assert.rejects(pending, /cancelled/);
});

test('calculation worker parses without the missing engine and reports unavailability', () => {
  const source = fs.readFileSync(path.resolve(__dirname, '../src/spreadsheetCalculation.worker.ts'), 'utf8');
  assert.doesNotMatch(source, /localWorkbookCalculation/);
  const messages = [];
  const self = { postMessage(data) { messages.push(data); } };
  new Function('self', source)(self);
  self.onmessage({ data: { revision: 'rev-1', sheets: [] } });
  assert.deepEqual(messages, [{ error: 'Local workbook calculation engine is not available.' }]);
});

test('createSpreadsheetCalculator timeout terminates the worker and rejects', async t => {
  const { createSpreadsheetCalculator } = await loadCalc();
  t.mock.timers.enable({ apis: ['setTimeout'] });
  const workers = installWorker(t, () => {});
  const calculator = createSpreadsheetCalculator({ timeoutMs: 25 });
  t.after(() => calculator.terminate());
  const pending = calculator.calculate({ revision: 'rev-1', sheets: [{ id: '1', name: 'Sheet1', cells: [] }] });
  assert.equal(workers[0].terminated, false);
  t.mock.timers.tick(24);
  assert.equal(workers[0].terminated, false);
  t.mock.timers.tick(1);
  await assert.rejects(pending, /exceeded 30 seconds/);
  assert.equal(workers[0].terminated, true);
});

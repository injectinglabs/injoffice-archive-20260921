// Real built Go/WASM engine behind the same client interface used by the editor.
const fs = require('node:fs');
const vm = require('node:vm');
const { nativeEngineAssets } = require('../scripts/native-assets.cjs');
exports.nativeClient = async function nativeClient() {
  const assets = nativeEngineAssets().find(item => item.package === '@injoffice/xlsx-wasm').assets;
  const file = name => assets.find(asset => asset.file === name).path;
  let reply;
  const sandbox = {
    ArrayBuffer, Uint8Array, TextEncoder, TextDecoder, WebAssembly, crypto: globalThis.crypto, performance,
    setTimeout, clearTimeout, setInterval, clearInterval, queueMicrotask, console,
    postMessage: value => { reply = value; }, close() {},
    importScripts: url => vm.runInContext(fs.readFileSync(url, 'utf8'), context),
    fetch: async url => ({ ok: true, arrayBuffer: async () => Uint8Array.from(fs.readFileSync(url)).buffer }),
  };
  sandbox.self = sandbox;
  const context = vm.createContext(sandbox);
  vm.runInContext(fs.readFileSync(file('xlsxnative.worker.js'), 'utf8'), context);
  let id = 0;
  async function send(op, input) {
    await sandbox.onmessage({ data: { protocol: 'injoffice.native-wasm-worker', version: 1, format: 'xlsx', id: String(++id), op, ...input } });
    if (!reply.ok) throw new Error(reply.error.message);
    return reply.result;
  }
  await send('init', { assets: { wasmUrl: file('xlsxnative.wasm'), goRuntimeUrl: file('wasm_exec.js') } });
  return {
    extract: async bytes => JSON.parse((await send('extract', { bytes: bytes.slice().buffer })).contractJson),
    apply: async (bytes, workbook, transaction) => new Uint8Array((await send('apply', { original: bytes.slice().buffer, payload: JSON.stringify(transaction), expectedRevision: workbook.source.package_sha256 })).bytes),
    terminate() {},
  };
};

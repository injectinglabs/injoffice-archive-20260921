const { Worker } = require('node:worker_threads');
const path = require('node:path');
const { validateBytes } = require('./file-store.cjs');
let active = false;
function replacePdfText(input, workerPath = path.join(__dirname, 'pdf-text-worker.mjs'), { timeoutMs = 30000 } = {}) {
  if (active) return Promise.reject(new Error('Another PDF text edit is running.'));
  if (!input || typeof input !== 'object') return Promise.reject(new Error('Select PDF text to replace.'));
  const bytes = validateBytes(input.bytes);
  if (bytes.length > 64 * 1024 * 1024) return Promise.reject(new Error('Existing-text editing supports PDFs up to 64 MB.'));
  if (typeof input.revision !== 'string' || !/^[a-f0-9]{64}$/.test(input.revision) || typeof input.oldText !== 'string' || typeof input.newText !== 'string' || input.oldText.length > 10000 || input.newText.length > 10000 || !Array.isArray(input.rect) || input.rect.length !== 4) return Promise.reject(new Error('Invalid PDF text selection.'));
  active = true;
  return new Promise((resolve, reject) => {
    let worker;
    try { worker = new Worker(workerPath, { workerData: { bytes, revision: input.revision, page: input.page, rect: input.rect, oldText: input.oldText, newText: input.newText }, resourceLimits: { maxOldGenerationSizeMb: 384 } }); }
    catch (error) { active = false; reject(error); return; }
    let settled = false;
    const finish = (error, result) => { if (settled) return; settled = true; clearTimeout(timeout); active = false; void worker.terminate(); error ? reject(error) : resolve(result); };
    const timeout = setTimeout(() => finish(new Error('PDF text editing timed out. Your original document is unchanged.')), timeoutMs);
    worker.on('message', result => { if (result.error) finish(new Error(result.error)); else if (!(result.bytes instanceof Uint8Array) || !result.bytes.length) finish(new Error('The PDF engine returned no document.')); else finish(null, result.bytes); });
    worker.on('error', error => finish(error));
    worker.on('exit', () => finish(new Error('The PDF editor stopped before completing the change.')));
  });
}
module.exports = { replacePdfText };

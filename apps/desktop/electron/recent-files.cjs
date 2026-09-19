const fs = require('node:fs/promises');
const path = require('node:path');
const { randomUUID } = require('node:crypto');
const { atomicWrite } = require('./file-store.cjs');

const pathKey = (filename) => process.platform === 'win32' ? filename.toLowerCase() : filename;
const validEntry = (entry) => entry && typeof entry.id === 'string' && /^[0-9a-f-]{36}$/i.test(entry.id)
  && typeof entry.path === 'string' && path.isAbsolute(entry.path) && !entry.path.includes('\0')
  && Number.isFinite(entry.updatedAt) && entry.updatedAt >= 0;

class RecentFiles {
  #filename;
  #entries = [];
  #loaded = false;
  #pending = Promise.resolve();

  constructor(filename) { this.#filename = filename; }

  async #load() {
    if (this.#loaded) return;
    let data;
    try {
      const info = await fs.stat(this.#filename);
      if (info.size <= 1024 * 1024) data = JSON.parse(await fs.readFile(this.#filename, 'utf8'));
    } catch (error) {
      if (error.code !== 'ENOENT' && !(error instanceof SyntaxError)) throw error;
    }
    const paths = new Set();
    const ids = new Set();
    this.#entries = (data?.version === 1 && Array.isArray(data.files) ? data.files : [])
      .filter(validEntry).sort((a, b) => b.updatedAt - a.updatedAt)
      .filter(entry => {
        const key = pathKey(entry.path);
        if (paths.has(key) || ids.has(entry.id)) return false;
        paths.add(key); ids.add(entry.id); return true;
      }).slice(0, 30).map(entry => ({ id: entry.id, name: path.basename(entry.path), path: entry.path, updatedAt: entry.updatedAt }));
    this.#loaded = true;
  }

  #enqueue(operation) {
    const result = this.#pending.then(async () => { await this.#load(); return operation(); });
    this.#pending = result.catch(() => {});
    return result;
  }

  async #persist(entries) {
    await fs.mkdir(path.dirname(this.#filename), { recursive: true });
    await atomicWrite(this.#filename, Buffer.from(JSON.stringify({ version: 1, files: entries }, null, 2)));
    this.#entries = entries;
  }

  list() { return this.#enqueue(() => this.#entries.map(entry => ({ ...entry }))); }

  record(filename) {
    return this.#enqueue(async () => {
      const canonical = await fs.realpath(filename);
      const key = pathKey(canonical);
      const previous = this.#entries.find(entry => pathKey(entry.path) === key);
      const entry = { id: previous?.id ?? randomUUID(), name: path.basename(canonical), path: canonical, updatedAt: Date.now() };
      await this.#persist([entry, ...this.#entries.filter(item => pathKey(item.path) !== key)].slice(0, 30));
      return { ...entry };
    });
  }

  get(id) {
    return this.#enqueue(() => {
      const entry = this.#entries.find(item => typeof id === 'string' && item.id === id);
      if (!entry) throw new Error('This recent file is no longer in your history. Use Open file to find it again.');
      return { ...entry };
    });
  }

  remove(id) {
    return this.#enqueue(async () => {
      if (typeof id !== 'string') throw new Error('Invalid recent file identifier.');
      await this.#persist(this.#entries.filter(entry => entry.id !== id));
    });
  }
}

async function openRecentFile(history, store, id) {
  const entry = await history.get(id);
  try { return await store.open(entry.path); }
  catch (error) {
    if (['ENOENT', 'ENOTDIR'].includes(error.code)) throw new Error(`“${entry.name}” could not be found. It may have moved or been deleted. Use Open file to locate it, or remove it from Recent files.`);
    if (['EACCES', 'EPERM'].includes(error.code)) throw new Error(`“${entry.name}” cannot be opened. Check its file permissions, then try again.`);
    throw error;
  }
}

module.exports = { RecentFiles, openRecentFile };

const fs = require('node:fs/promises');
const path = require('node:path');
const { createHash, randomUUID } = require('node:crypto');

const fingerprint = (bytes) => createHash('sha256').update(bytes).digest('hex');
const untitledNames = { docx: 'Untitled.docx', xlsx: 'Workbook.xlsx', pptx: 'Presentation.pptx', pdf: 'Untitled.pdf' };
const openableFormats = new Set(['docx', 'xlsx', 'pptx', 'pdf', 'docm']);

function validateFormat(format) {
  if (typeof format !== 'string' || !Object.hasOwn(untitledNames, format)) throw new Error('Choose a DOCX document, XLSX workbook, PPTX presentation, or PDF.');
  return format;
}

function validateOpenFormat(format) {
  if (typeof format !== 'string' || !openableFormats.has(format.toLowerCase())) throw new Error('Choose a DOCX document, XLSX workbook, PPTX presentation, or PDF.');
  return format.toLowerCase();
}

function normalizeSaveDestination(filename, format) {
  if (typeof filename !== 'string' || !path.isAbsolute(filename)) throw new Error('Choose a valid location for this document.');
  if (typeof format !== 'string' || !['docx', 'xlsx', 'pptx', 'pdf', 'docm', 'svg', 'csv', 'tsv'].includes(format.toLowerCase())) throw new Error('This document format cannot be saved.');
  const expected = format.toLowerCase();
  const extension = path.extname(filename).toLowerCase();
  if (!extension) return `${filename}.${expected}`;
  if (extension === '.') return `${filename.slice(0, -1)}.${expected}`;
  if (extension !== `.${expected}`) throw new Error(`Save this document with a .${expected} extension. Changing the filename does not convert it to ${extension.slice(1).toUpperCase()}.`);
  return filename;
}

function validateBytes(bytes) {
  if (!(bytes instanceof Uint8Array)) throw new TypeError('Expected document bytes.');
  if (bytes.byteLength > 512 * 1024 * 1024) throw new Error('Documents must be smaller than 512 MB.');
  return Buffer.from(bytes);
}

// The temporary file is on the same volume so rename atomically replaces the target.
async function atomicWrite(filename, bytes) {
  let mode = 0o600;
  try {
    const info = await fs.lstat(filename);
    if (info.isSymbolicLink() || !info.isFile()) throw new Error('Save target must be a regular file.');
    mode = info.mode & 0o777;
  } catch (error) {
    if (error.code !== 'ENOENT') throw error;
  }
  const temporary = path.join(path.dirname(filename), `.${path.basename(filename)}.${randomUUID()}.tmp`);
  let handle;
  try {
    handle = await fs.open(temporary, 'wx', mode);
    await handle.writeFile(bytes);
    await handle.sync();
    await handle.close();
    handle = undefined;
    await fs.rename(temporary, filename);
    // POSIX rename publication needs a directory fsync before old recovery
    // generations can be pruned. Some platforms/filesystems cannot provide it.
    let directory;
    try {
      directory = await fs.open(path.dirname(filename), 'r');
      await directory.sync();
      return true;
    } catch (error) {
      if (['EINVAL', 'ENOTSUP', 'EISDIR'].includes(error.code) || (process.platform === 'win32' && ['EPERM','EACCES'].includes(error.code))) return false;
      throw error;
    } finally { await directory?.close(); }
  } finally {
    if (handle) await handle.close();
    await fs.unlink(temporary).catch((error) => { if (error.code !== 'ENOENT') throw error; });
  }
}

class FileStore {
  #documents = new Map();

  create(format, bytes) {
    validateFormat(format);
    const output = validateBytes(bytes);
    const id = randomUUID();
    const name = untitledNames[format];
    this.#documents.set(id, { name, bytes: output, untitled: true });
    return { id, name, bytes: Buffer.from(output), untitled: true };
  }

  releaseAll() { this.#documents.clear(); }

  release(id) { this.get(id); this.#documents.delete(id); }

  async open(filename) {
    const canonicalPath = await fs.realpath(filename);
    const info = await fs.stat(canonicalPath);
    if (!info.isFile()) throw new Error('Choose a regular document file.');
    if (info.size > 512 * 1024 * 1024) throw new Error('Documents must be smaller than 512 MB.');
    const bytes = validateBytes(await fs.readFile(canonicalPath));
    const id = randomUUID();
    this.#documents.set(id, { filename: canonicalPath, hash: fingerprint(bytes) });
    return { id, name: path.basename(canonicalPath), bytes };
  }

  get(id) {
    if (typeof id !== 'string' || !this.#documents.has(id)) throw new Error('Unknown document. Open the file again.');
    return this.#documents.get(id);
  }

  async changed(id) {
    const document = this.get(id);
    if (document.untitled) return false;
    try {
      return fingerprint(await fs.readFile(document.filename)) !== document.hash;
    } catch (error) {
      if (error.code === 'ENOENT') return true;
      throw error;
    }
  }

  async save(id, bytes, destination, overwrite = false) {
    const document = this.get(id);
    const output = validateBytes(bytes);
    const filename = destination || document.filename;
    if (typeof filename !== 'string' || !path.isAbsolute(filename)) throw new Error('Choose where to save this document first.');
    if (filename === document.filename && !overwrite && await this.changed(id)) {
      const error = new Error('This file changed outside InjOffice.');
      error.code = 'EXTERNAL_CHANGE';
      throw error;
    }
    await atomicWrite(filename, output);
    this.#documents.set(id, { filename, hash: fingerprint(output), untitled: false });
    return { id, name: path.basename(filename), untitled: false };
  }
}

module.exports = { FileStore, atomicWrite, validateBytes, validateFormat, validateOpenFormat, normalizeSaveDestination };

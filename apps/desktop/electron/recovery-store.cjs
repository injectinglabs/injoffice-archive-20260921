const fs = require('node:fs/promises');
const path = require('node:path');
const { createHash, randomUUID } = require('node:crypto');
const { atomicWrite, validateBytes } = require('./file-store.cjs');
const ID = /^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/;
const SHA = /^[a-f0-9]{64}$/;
const hash = bytes => createHash('sha256').update(bytes).digest('hex');
function nameCheck(name) { if (typeof name !== 'string' || name !== path.basename(name) || !/\.(docx|xlsx|pptx|pdf)$/i.test(name)) throw new Error('Invalid recovery document name.'); }
function encodeDraft(value) {
  const encoded = JSON.stringify(value);
  if (typeof encoded !== 'string' || encoded.length > 512 * 1024) throw new Error('This pending draft is too large for recovery. Apply or save your work.');
  return { draft: JSON.parse(encoded), draftSha256: hash(Buffer.from(encoded)) };
}
class RecoveryStore {
  #directory; #queue = Promise.resolve(); #draftTickets = new Map();
  constructor(directory) { this.#directory = directory; }
  #file(id) { if (typeof id !== 'string' || !ID.test(id)) throw new Error('Unknown recovery copy.'); return path.join(this.#directory, `${id}.json`); }
  #blob(id, sha) { this.#file(id); if (typeof sha !== 'string' || !SHA.test(sha)) throw new Error('Invalid recovery copy.'); return path.join(this.#directory, `${id}-${sha}.bin`); }
  #serial(operation) { const result = this.#queue.then(operation); this.#queue = result.catch(() => {}); return result; }
  async #metadata(id) {
    const filename = this.#file(id), stat = await fs.lstat(filename);
    if (!stat.isFile() || stat.isSymbolicLink() || stat.size > 90 * 1024 * 1024) throw new Error('Invalid recovery copy.');
    const data = JSON.parse(await fs.readFile(filename, 'utf8'));
    if (![1,2].includes(data.version) || data.id !== id || !Number.isFinite(data.updatedAt) || typeof data.sha256 !== 'string' || !SHA.test(data.sha256)) throw new Error('Invalid recovery copy.');
    nameCheck(data.name);
    if (data.version === 1) {
      if (typeof data.content !== 'string') throw new Error('Invalid recovery copy.');
    } else {
      if (!Number.isInteger(data.size) || data.size < 0 || data.size > 64 * 1024 * 1024 || typeof data.sourceRevision !== 'string' || !ID.test(data.sourceRevision)) throw new Error('Invalid recovery copy.');
      const blob = await fs.lstat(this.#blob(id,data.sha256));
      if (!blob.isFile() || blob.isSymbolicLink() || blob.size !== data.size) throw new Error('Recovery copy failed its integrity check.');
    }
    const draft = data.draft ?? null;
    if ((data.version === 2 || draft !== null) && (JSON.stringify(draft).length > 512 * 1024 || hash(Buffer.from(JSON.stringify(draft))) !== data.draftSha256)) throw new Error('Recovery draft failed its integrity check.');
    return {...data,draft};
  }
  async #read(id) {
    const data = await this.#metadata(id);
    const bytes = data.version === 1 ? Buffer.from(data.content,'base64') : await fs.readFile(this.#blob(id,data.sha256));
    if (bytes.length > 64 * 1024 * 1024 || hash(bytes) !== data.sha256) throw new Error('Recovery copy failed its integrity check.');
    return {id,name:data.name,updatedAt:data.updatedAt,bytes,draft:data.draft};
  }
  async #entries(withBytes = true) {
    let names;
    try { names = await fs.readdir(this.#directory); } catch (error) { if (error.code === 'ENOENT') return []; throw error; }
    const entries = [];
    for (const name of names) if (ID.test(name.replace(/\.json$/, '')) && name.endsWith('.json')) {
      try {
        const id=name.slice(0,-5);
        if (withBytes) entries.push(await this.#read(id));
        else {const data=await this.#metadata(id);entries.push({id,size:data.version===1?Buffer.byteLength(data.content,'base64'):data.size,legacy:data.version===1});}
      } catch { /* Keep damaged entries on disk; never silently erase recovery data. */ }
    }
    return entries;
  }
  list() { return this.#serial(async () => (await this.#entries()).map(({ bytes, draft, ...entry }) => ({ ...entry, size: bytes.length })).sort((a, b) => b.updatedAt - a.updatedAt)); }
  async flush() { let pending; do { pending = this.#queue; await pending; } while (pending !== this.#queue); }
  read(id) { return this.#serial(() => this.#read(id)); }
  write(id, name, input, recoveryDraft = null, sourceRevision = randomUUID()) {
    let draft;
    try { draft=encodeDraft(recoveryDraft); } catch(error) { return Promise.reject(error); }
    const bytes=validateBytes(input),filename=this.#file(id);
    if (bytes.length>64*1024*1024) return Promise.reject(new Error('Recovery supports files up to 64 MB. Save this file to keep your changes.'));
    try {nameCheck(name);if(typeof sourceRevision!=='string'||!ID.test(sourceRevision))throw new Error('Invalid recovery revision.');}catch(error){return Promise.reject(error);}
    return this.#serial(async()=>{
      const entries=(await this.#entries(false)).filter(entry=>entry.id!==id);
      if(entries.length>=20||entries.reduce((total,entry)=>total+entry.size,bytes.length)>256*1024*1024)throw new Error('Recovery storage is full. Save or discard recovered copies before continuing.');
      await fs.mkdir(this.#directory,{recursive:true,mode:0o700});
      const sha256=hash(bytes),blob=this.#blob(id,sha256);
      // Bound retained generations too when directory publication cannot be
      // synchronized (and after a failed/crashed manifest update).
      let retained=0,alreadyStored=false;
      for(const name of await fs.readdir(this.#directory))if(/^[a-f0-9-]{36}-[a-f0-9]{64}\.bin$/.test(name)){
        const info=await fs.lstat(path.join(this.#directory,name));
        if(info.isFile()&&!info.isSymbolicLink()){retained+=info.size;if(name===path.basename(blob))alreadyStored=true;}
      }
      if(retained+entries.filter(entry=>entry.legacy).reduce((total,entry)=>total+entry.size,0)+(alreadyStored?0:bytes.length)>256*1024*1024)throw new Error('Recovery storage is full. Save or discard recovered copies before continuing.');
      // Publish immutable native bytes before replacing the manifest. A failed
      // manifest write leaves the previous complete recovery snapshot readable.
      const blobPublished=await atomicWrite(blob,bytes);
      const manifestPublished=await atomicWrite(filename,Buffer.from(JSON.stringify({version:2,id,name,updatedAt:Date.now(),sha256,size:bytes.length,sourceRevision,...draft})));
      if(blobPublished && manifestPublished) await this.#pruneBlobs(id,sha256);
    });
  }
  writeDraft(id,name,sourceRevision,recoveryDraft = null) {
    let draft;
    try {this.#file(id);nameCheck(name);if(typeof sourceRevision!=='string'||!ID.test(sourceRevision))throw new Error('Invalid recovery revision.');draft=encodeDraft(recoveryDraft);}catch(error){return Promise.reject(error);}
    const ticket=Symbol();this.#draftTickets.set(id,ticket);
    return this.#serial(async()=>{
      if(this.#draftTickets.get(id)!==ticket)return;
      try {
        const current=await this.#metadata(id);
        if(current.version!==2||current.sourceRevision!==sourceRevision)throw new Error('The recovery source changed. A full checkpoint is required.');
        await atomicWrite(this.#file(id),Buffer.from(JSON.stringify({...current,name,updatedAt:Date.now(),...draft})));
      }finally{if(this.#draftTickets.get(id)===ticket)this.#draftTickets.delete(id);}
    });
  }
  async #pruneBlobs(id,keep) {
    // Only superseded blobs belonging to this explicitly updated/discarded copy.
    const names=await fs.readdir(this.#directory).catch(error=>{if(error.code==='ENOENT')return [];throw error;});
    for(const name of names)if(name.startsWith(`${id}-`)&&name.endsWith('.bin')){
      const sha=name.slice(id.length+1,-4);
      if(!SHA.test(sha)||sha===keep)continue;
      try {const filename=this.#blob(id,sha),info=await fs.lstat(filename);if(info.isFile()&&!info.isSymbolicLink())await fs.unlink(filename);}catch{/* An orphan cannot invalidate an already durable manifest. */}
    }
  }
  remove(id) {const filename=this.#file(id);return this.#serial(async()=>{await fs.unlink(filename).catch(error=>{if(error.code!=='ENOENT')throw error;});await this.#pruneBlobs(id);});}
}
module.exports={RecoveryStore};

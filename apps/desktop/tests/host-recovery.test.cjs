const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const { randomUUID } = require('node:crypto');
const { RecoveryStore } = require('../electron/recovery-store.cjs');
async function setup(t) { const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'injoffice-recovery-')); t.after(() => fs.rm(directory, { recursive: true, force: true })); return { directory, journal: new RecoveryStore(directory) }; }
test('recovery serializes updates and remains readable after restart', async t => {
  const { directory, journal } = await setup(t), id = randomUUID();
  await Promise.all([journal.write(id, 'Letter.docx', Buffer.from('first')), journal.write(id, 'Letter.docx', Buffer.from('latest'))]);
  const restarted = new RecoveryStore(directory);
  assert.equal((await restarted.list()).length, 1);
  assert.equal((await restarted.read(id)).bytes.toString(), 'latest');
  await restarted.remove(id); assert.deepEqual(await restarted.list(), []);
});
test('recovery rejects forged identities, path names, and oversized checkpoints without losing previous copy', async t => {
  const { journal } = await setup(t), id = randomUUID();
  await journal.write(id, 'Letter.docx', Buffer.from('keep'));
  assert.throws(() => journal.write('../elsewhere', 'Letter.docx', Buffer.from('bad')), /Unknown/);
  await assert.rejects(journal.write(id, '../Letter.docx', Buffer.from('bad')), /Invalid/);
  await assert.rejects(journal.write(id, 'Letter.docx', Buffer.alloc(64 * 1024 * 1024 + 1)), /64 MB/);
  assert.equal((await journal.read(id)).bytes.toString(), 'keep');
});
test('recovery detects corrupt bytes and refuses symbolic links without deleting evidence', async t => {
  const { directory, journal } = await setup(t), id = randomUUID();
  await journal.write(id, 'Letter.docx', Buffer.from('keep'));
  const filename = path.join(directory, `${id}.json`);
  const record = JSON.parse(await fs.readFile(filename, 'utf8'));
  await fs.writeFile(path.join(directory,`${id}-${record.sha256}.bin`),Buffer.from('fake'));
  await assert.rejects(journal.read(id), /integrity/);
  assert((await fs.stat(filename)).isFile());
  const link = randomUUID(); await fs.symlink(filename, path.join(directory, `${link}.json`));
  await assert.rejects(journal.read(link), /Invalid/);
});

test('pending drafts survive restart with exact native bytes and clear atomically when applied', async t => {
  const { directory, journal } = await setup(t), id = randomUUID();
  const draft = { version: 1, format: 'docx', target: 'run-id', text: 'Not applied yet' };
  await journal.write(id, 'Letter.docx', Buffer.from('native-before'), draft);
  const restored = await new RecoveryStore(directory).read(id);
  assert.deepEqual(restored.draft, draft); assert.equal(restored.bytes.toString(), 'native-before');
  await journal.write(id, 'Letter.docx', Buffer.from('native-after'), null);
  const applied = await journal.read(id); assert.equal(applied.draft, null); assert.equal(applied.bytes.toString(), 'native-after');
  await assert.rejects(journal.write(id, 'Letter.docx', Buffer.from('bad'), { text: 'a'.repeat(512 * 1024) }), /too large/);
  assert.equal((await journal.read(id)).bytes.toString(), 'native-after');
});

test('draft-only checkpoints reuse immutable bytes and reject stale revisions', async t => {
  const {directory,journal}=await setup(t),id=randomUUID(),revision=randomUUID();
  await journal.write(id,'Letter.docx',Buffer.alloc(1024*1024,65),null,revision);
  const filename=path.join(directory,`${id}.json`),manifest=JSON.parse(await fs.readFile(filename,'utf8'));
  assert.equal(manifest.version,2);assert.equal(manifest.content,undefined);
  const blob=path.join(directory,`${id}-${manifest.sha256}.bin`),before=await fs.stat(blob);
  await Promise.all([journal.writeDraft(id,'Letter.docx',revision,{text:'older'}),journal.writeDraft(id,'Letter.docx',revision,{text:'latest'})]);
  assert.deepEqual((await journal.read(id)).draft,{text:'latest'});assert.equal((await fs.stat(blob)).mtimeMs,before.mtimeMs);
  assert((await fs.stat(filename)).size<1024,'draft manifest remains small independently of document size');
  await assert.rejects(journal.writeDraft(id,'Letter.docx',randomUUID(),{text:'stale'}),/source changed/);
  assert.deepEqual((await journal.read(id)).draft,{text:'latest'});
  const next=randomUUID();await journal.write(id,'Letter.docx',Buffer.from('next'),null,next);
  await assert.rejects(journal.writeDraft(id,'Letter.docx',revision,{text:'stale'}),/source changed/);
  assert.equal((await journal.read(id)).bytes.toString(),'next');
  const retained=(await fs.readdir(directory)).filter(name=>name.endsWith('.bin')).length;
  assert(retained>=1&&retained<=2);if(process.platform!=='win32')assert.equal(retained,1);
  await journal.remove(id);assert.deepEqual(await fs.readdir(directory),[]);
});
test('legacy recovery migrates on full checkpoint and manifest failure keeps the old snapshot',async t=>{
  const {directory,journal}=await setup(t),id=randomUUID(),filename=path.join(directory,`${id}.json`);
  const {createHash}=require('node:crypto'),sha=bytes=>createHash('sha256').update(bytes).digest('hex');
  const bytes=Buffer.from('legacy');
  await fs.writeFile(filename,JSON.stringify({version:1,id,name:'Letter.docx',updatedAt:1,content:bytes.toString('base64'),sha256:sha(bytes),draft:null}));
  assert.equal((await journal.read(id)).bytes.toString(),'legacy');
  const revision=randomUUID();await journal.write(id,'Letter.docx',Buffer.from('upgraded'),null,revision);
  assert.equal(JSON.parse(await fs.readFile(filename,'utf8')).version,2);
  const originalRename=fs.rename;
  fs.rename=async(from,to)=>{if(to===filename){const error=new Error('Simulated disk full');error.code='ENOSPC';throw error;}return originalRename(from,to);};
  try {await assert.rejects(journal.write(id,'Letter.docx',Buffer.from('not committed'),null,randomUUID()),/disk full/);}finally{fs.rename=originalRename;}
  assert.equal((await new RecoveryStore(directory).read(id)).bytes.toString(),'upgraded');
  await journal.writeDraft(id,'Letter.docx',revision,{text:'kept draft'});
  assert.deepEqual((await journal.read(id)).draft,{text:'kept draft'});
});

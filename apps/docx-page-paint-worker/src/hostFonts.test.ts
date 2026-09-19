import {afterAll,describe,expect,it} from 'vitest'
import {readFileSync,statSync,writeFileSync,mkdtempSync,rmSync} from 'node:fs'
import {resolve} from 'node:path'
import {tmpdir} from 'node:os'
import {createHash} from 'node:crypto'
import {loadHostFonts,MAX_OPERATOR_FONT_FACES} from './hostFonts.js'
import {nativeDocxPagePaintWorkerErrorV1} from './protocol.js'
import {encodeNativeDOCXFontInventoryV1,nativeDOCXCanonicalWireSHA256V1} from '../../../packages/docs/src/nativeFontInventoryV1.js'
import type {NativeDocxPagePaintPrepareInputV1} from '@injoffice/docs/native-page-paint-compiler'

const root=resolve(import.meta.dirname,'../../..'),scratch=mkdtempSync(resolve(tmpdir(),'docx-host-fonts-'))
afterAll(()=>rmSync(scratch,{recursive:true,force:true}))
const path=resolve(root,'node_modules/dejavu-fonts-ttf/ttf/DejaVuSans.ttf')
const sha256=`sha256:${createHash('sha256').update(readFileSync(path)).digest('hex')}`
const entry={family:'DejaVu Sans',weight:400,style:'normal',path,sha256}
const bytesOf=(file:string)=>statSync(file).size
function input(){
 const inventory=JSON.parse(readFileSync(resolve(root,'go/docxpatch/testdata/font-inventory-v1.json'),'utf8'))
 inventory.families=[];inventory.references[0].family='DejaVu Sans'
 delete inventory.native_text_manifest;delete inventory.native_text_manifest_sha256
 delete inventory.font_table.font_relationships_part;delete inventory.font_table.font_relationships_sha256
 inventory.inventory_sha256='';inventory.inventory_sha256=nativeDOCXCanonicalWireSHA256V1(inventory)
 return {font_assets:[],font_inventory_json:encodeNativeDOCXFontInventoryV1(inventory)} as unknown as NativeDocxPagePaintPrepareInputV1
}
let id=0
function config(faces:unknown[],substitutions?:unknown){const p=resolve(scratch,`fonts-${id++}.json`);writeFileSync(p,JSON.stringify({version:1,faces,...(substitutions?{substitutions}:{})}));return p}

/** Wraps a standalone sfnt in a real `ttcf` header with `faceCount` entries, all
 * naming the same face, so a collection fixture needs no third-party TTC file.
 * Every table offset shifts by the header length; table checksums cover table
 * content, not offsets, so the pinned preflight still validates the result. */
function collection(source:string,faceCount:number):{path:string,sha256:`sha256:${string}`}{
 const sfnt=readFileSync(source)
 const header=12+faceCount*4
 const bytes=Buffer.concat([Buffer.alloc(header),sfnt])
 bytes.write('ttcf',0,'latin1')
 bytes.writeUInt32BE(0x00020000,4)
 bytes.writeUInt32BE(faceCount,8)
 for(let face=0;face<faceCount;face++)bytes.writeUInt32BE(header,12+face*4)
 const numTables=bytes.readUInt16BE(header+4)
 for(let table=0;table<numTables;table++){
  const record=header+12+table*16
  bytes.writeUInt32BE(bytes.readUInt32BE(record+8)+header,record+8)
 }
 const path=resolve(scratch,`collection-${id++}.ttc`)
 writeFileSync(path,bytes)
 return {path,sha256:`sha256:${createHash('sha256').update(bytes).digest('hex')}`}
}
const ttc=collection(path,2)
const ttcEntry={family:'DejaVu Sans',weight:400,style:'normal',path:ttc.path,sha256:ttc.sha256,collectionIndex:1}
describe('operator-owned DOCX fonts',()=>{
 it('names the operator font face the pinned preflight refuses instead of failing the request unattributed',async()=>{
  // An entry of the operator's own manifest whose bytes are not a font this
  // engine can qualify. Before this was named, the preflight's own sentence
  // escaped under the blanket COMPILATION_REFUSED code: the whole request
  // failed on a message about an sfnt table, naming none of the up-to-64
  // configured files and reading like a fact about the document.
  const broken=Buffer.from(readFileSync(path))
  const count=broken.readUInt16BE(4)
  let record=-1
  for(let index=0;index<count;index++){const at=12+index*16;if(broken.toString('latin1',at,at+4)==='cmap')record=at}
  expect(record).toBeGreaterThan(0)
  broken.writeUInt32BE(8,record+12)
  const file=resolve(scratch,'unqualified.ttf')
  writeFileSync(file,broken)
  const p=config([{family:'DejaVu Sans',weight:400,style:'normal',path:file,sha256:`sha256:${createHash('sha256').update(broken).digest('hex')}`}])
  // The prefix is what this test is about; the sentence after it is the
  // preflight's own and is free to change.
  await expect(loadHostFonts(input(),p)).rejects.toThrow(/^Host font face DejaVu Sans \/ 400 \/ normal is not a qualified font face: sfnt /)
  const error=await loadHostFonts(input(),p).then(()=>undefined,(reason:unknown)=>reason)
  expect(nativeDocxPagePaintWorkerErrorV1(error)).toEqual({code:'HOST_FONT_UNQUALIFIED',scope_id:expect.stringMatching(/^host-font-0-/),message:expect.stringContaining('DejaVu Sans / 400 / normal')})
 })

 it('requires explicit opt-in, retains source family and binds selected bytes',async()=>{
  const value=input(),inventory=JSON.parse(value.font_inventory_json)
  inventory.references.forEach((r:any)=>{r.family='Missing Family'})
  inventory.inventory_sha256='';inventory.inventory_sha256=nativeDOCXCanonicalWireSHA256V1(inventory)
  value.font_inventory_json=encodeNativeDOCXFontInventoryV1(inventory)
  const policy={version:1,mappings:[{sourceFamily:'Missing Family',targetFamily:'DejaVu Sans',weight:400,style:'normal'}]}
  const p=config([entry],policy)
  await expect(loadHostFonts(value,p)).rejects.toThrow(/Exact configured font/)
  const fonts=await loadHostFonts(value,p,true)
  const run={version:1 as const,text:'Actual source',fontSizeMilliPoints:10000,font:{families:['Missing Family'],weight:400,style:'normal' as const,stretch:100},script:'Latn',language:'en',direction:'ltr' as const}
  const selected=await fonts.resolver.resolve({manifest:fonts.manifest,run})
  if(!('face' in selected))throw new Error('Expected operator selection')
  expect(selected.face).toMatchObject({matchedFamily:'Missing Family',family:'DejaVu Sans',resolution:'substitute',contentDigest:sha256})
  expect(await fonts.resolver.load(selected.face)).toHaveProperty('bytes')
  expect(()=>fonts.resolver.load({...selected.face,matchedFamily:'Forged'})).toThrow(/binding/)
  expect(await fonts.resolver.resolve({manifest:fonts.manifest,run:{...run,text:'漢',script:'Hani'}})).toMatchObject({status:'refused'})
  await expect(loadHostFonts(value,config([entry]),true)).rejects.toThrow(/not configured/)
 })
 it('loads approximate fallback evidence faces as sidecar requests without making them required',async()=>{
  const extra={...entry,family:'Fallback Face'}
  const fonts=await loadHostFonts(input(),config([entry,extra]),'approximate',[{family:'Fallback Face',weight:400,style:'normal'},{family:'Nowhere Face',weight:400,style:'normal'}])
  expect(fonts.manifest.faces.map(f=>f.family).sort()).toEqual(['DejaVu Sans','Fallback Face'])
  // Sidecar faces never satisfy or substitute an inventory reference.
  await expect(loadHostFonts(input(),config([extra]),false,[{family:'Fallback Face',weight:400,style:'normal'}])).rejects.toThrow(/Exact configured font unavailable/)
 })
 it('loads only exact referenced fonts and owns each returned byte buffer',async()=>{
  const fonts=await loadHostFonts(input(),config([entry,{...entry,family:'Unused',path:'/nonexistent/unreferenced.ttf'}]))
  expect(fonts.manifest.faces).toHaveLength(1)
  const resource=[...fonts.resources.values()][0]!
  expect(resource.face.sourceKind).toBe('host')
  const first=await fonts.resolver.load(resource.face)
  if(!('bytes' in first))throw new Error('font refused')
  first.bytes[0]^=1
  const second=await fonts.resolver.load(resource.face)
  if(!('bytes' in second))throw new Error('font refused')
  expect(second.bytes[0]).toBe(resource.bytes[0])
  expect(()=>fonts.resolver.load({...resource.face,contentDigest:`sha256:${'0'.repeat(64)}`})).toThrow(/binding/)
 })
 it('rejects duplicate, mismatched, missing, relative, and malformed fonts',async()=>{
  await expect(loadHostFonts(input(),config([entry,{...entry,family:'dejavu sans'}]))).rejects.toThrow(/Ambiguous/)
  await expect(loadHostFonts(input(),config([{...entry,sha256:`sha256:${'0'.repeat(64)}`}]))).rejects.toThrow(/digest/)
  await expect(loadHostFonts(input(),config([]))).rejects.toThrow(/Exact configured font unavailable/)
  await expect(loadHostFonts(input(),config([{...entry,weight:'400'}]))).rejects.toThrow(/Invalid/)
  await expect(loadHostFonts(input(),config([{...entry,path:'relative.ttf'}]))).rejects.toThrow(/Invalid/)
  await expect(loadHostFonts(input(),'relative.json')).rejects.toThrow(/absolute/)
  await expect(loadHostFonts(input(),config([{...entry,extra:true}]))).rejects.toThrow(/Invalid/)
 })
 it('admits a manifest at the face cap, names the limit past it, and keeps the budgets it protects',async()=>{
  // Pins the cap itself, not only its message: the Go substitution-preview
  // helper re-validates the same operator file against this same number.
  expect(MAX_OPERATOR_FONT_FACES).toBe(64)
  // A face is only loaded when the document references it, so the filler
  // entries stay unread: this measures the manifest bound alone.
  const filler=(count:number)=>Array.from({length:count},(_,i)=>({...entry,family:`Unused ${i}`,path:`/nonexistent/unreferenced-${i}.ttf`}))
  const full=config([entry,...filler(MAX_OPERATOR_FONT_FACES-1)])
  // The manifest file is read under a 64 KiB bound; a full manifest must fit.
  expect(bytesOf(full)).toBeLessThanOrEqual(65536)
  const fonts=await loadHostFonts(input(),full)
  expect(fonts.manifest.faces).toHaveLength(1)
  await expect(loadHostFonts(input(),config([entry,...filler(MAX_OPERATOR_FONT_FACES)]))).rejects.toThrow(`Invalid host font manifest: ${MAX_OPERATOR_FONT_FACES+1} faces exceeds the ${MAX_OPERATOR_FONT_FACES}-face limit`)
 })
 it('keeps strict missing-face refusal and substitutes only loaded host faces on the approximate path',async()=>{
  const value=input(),inventory=JSON.parse(value.font_inventory_json)
  inventory.references=[{...inventory.references[0],family:'Candara'},{...inventory.references[0],family:'Trebuchet MS',scope_ids:['paragraph:2']}]
  inventory.inventory_sha256='';inventory.inventory_sha256=nativeDOCXCanonicalWireSHA256V1(inventory)
  value.font_inventory_json=encodeNativeDOCXFontInventoryV1(inventory)
  const p=config([entry])
  await expect(loadHostFonts(value,p)).rejects.toThrow(/Exact configured font unavailable: Candara \/ 400 \/ normal, Trebuchet MS \/ 400 \/ normal/)
  await expect(loadHostFonts(value,config([]),'approximate')).rejects.toThrow(/Exact configured font unavailable/)
  const fonts=await loadHostFonts(value,p,'approximate')
  expect(fonts.substitutionPolicy).toBeUndefined()
  expect(fonts.manifest.faces.some(face=>face.family==='Calibri')).toBe(false)
  expect(fonts.approximateSubstitutions).toEqual([
   expect.objectContaining({source_family:'Candara',selected_family:'DejaVu Sans',reason:'loaded-host-manifest-face-v1'}),
   expect.objectContaining({source_family:'Trebuchet MS',selected_family:'DejaVu Sans',reason:'loaded-host-manifest-face-v1'}),
  ])
  const candara={version:1 as const,text:'Linked',fontSizeMilliPoints:10000,font:{families:['Candara'],weight:400,style:'normal' as const,stretch:100},script:'Latn',language:'en',direction:'ltr' as const}
  const selected=await fonts.resolver.resolve({manifest:fonts.manifest,run:candara})
  if(!('face' in selected))throw new Error('Expected loaded host substitute')
  expect(selected.face).toMatchObject({matchedFamily:'Candara',family:'DejaVu Sans',resolution:'exact',contentDigest:sha256})
  expect(await fonts.resolver.load(selected.face)).toHaveProperty('bytes')
  const unused=config([entry,{...entry,family:'Unused',path:'/nonexistent/unreferenced.ttf'}])
  const reused=await loadHostFonts(value,unused,'approximate')
  expect(reused.approximateSubstitutions?.every(record=>record.selected_family==='DejaVu Sans')).toBe(true)
  const italic=structuredClone(value),italicInventory=JSON.parse(italic.font_inventory_json)
  italicInventory.references=[{...italicInventory.references[0],family:'Candara',style:'italic'}]
  italicInventory.inventory_sha256='';italicInventory.inventory_sha256=nativeDOCXCanonicalWireSHA256V1(italicInventory)
  italic.font_inventory_json=encodeNativeDOCXFontInventoryV1(italicInventory)
  await expect(loadHostFonts(italic,p,'approximate')).rejects.toThrow(/Candara \/ 400 \/ italic/)
 })
 it('admits equation sidecar faces without changing the body substitute or counting them as references',async()=>{
  const value=input(),inventory=JSON.parse(value.font_inventory_json)
  inventory.references=[{...inventory.references[0],family:'Candara'}]
  inventory.inventory_sha256='';inventory.inventory_sha256=nativeDOCXCanonicalWireSHA256V1(inventory)
  value.font_inventory_json=encodeNativeDOCXFontInventoryV1(inventory)
  const mathPath=resolve(root,'node_modules/dejavu-fonts-ttf/ttf/DejaVuMathTeXGyre.ttf')
  const mathSha=`sha256:${createHash('sha256').update(readFileSync(mathPath)).digest('hex')}`
  const math={family:'Cambria Math',weight:400,style:'normal',path:mathPath,sha256:mathSha}
  // Admitting the requested math face before body substitution would mark its
  // weight/style as covered, block DejaVu Sans and hand Candara body text the
  // math face; the sidecar must not change which substitute the body gets.
  const p=config([entry,math])
  const without=await loadHostFonts(value,p,'approximate')
  const withEquation=await loadHostFonts(value,p,'approximate',[{family:'Cambria Math',weight:400,style:'normal'}])
  for(const fonts of [without,withEquation]){
   expect(fonts.approximateSubstitutions).toEqual([expect.objectContaining({source_family:'Candara',selected_family:'DejaVu Sans'})])
   const selected=await fonts.resolver.resolve({manifest:fonts.manifest,run:{version:1 as const,text:'Body',fontSizeMilliPoints:10000,font:{families:['Candara'],weight:400,style:'normal' as const,stretch:100},script:'Latn',language:'en',direction:'ltr' as const}})
   if(!('face' in selected))throw new Error('Expected loaded host substitute')
   expect(selected.face).toMatchObject({family:'DejaVu Sans',contentDigest:sha256})
  }
  expect(without.manifest.faces.some(face=>face.family==='Cambria Math')).toBe(false)
  const mathFace=withEquation.manifest.faces.find(face=>face.family==='Cambria Math')
  expect(mathFace).toMatchObject({source:{kind:'host',contentDigest:mathSha}})
  expect(withEquation.resources.has(mathFace!.faceId)).toBe(true)
  // Requests for faces the operator did not configure are not failures; over-long request lists are.
  expect((await loadHostFonts(value,p,'approximate',[{family:'STIX Two Math',weight:700,style:'italic'}])).manifest.faces).toHaveLength(1)
  await expect(loadHostFonts(value,p,'approximate',Array.from({length:33},()=>({family:'Cambria Math',weight:400 as const,style:'normal' as const})))).rejects.toThrow(/exceed their bound/)
 })
 it('declares one face of a TTC collection, pins its index and counts the file once',async()=>{
  const fonts=await loadHostFonts(input(),config([ttcEntry]))
  expect(fonts.manifest.faces).toHaveLength(1)
  const face=fonts.manifest.faces[0]!
  expect(face.source).toMatchObject({kind:'host',contentDigest:ttc.sha256,collectionIndex:1})
  const resource=fonts.resources.get(face.faceId)!
  expect(resource.face.collectionIndex).toBe(1)
  // The named face's own design metrics, read by the pinned preflight, not the collection's first face by luck.
  const standalone=await loadHostFonts(input(),config([entry]))
  expect(resource.metrics).toEqual([...standalone.resources.values()][0]!.metrics)
  const selected=await fonts.resolver.resolve({manifest:fonts.manifest,run:{version:1 as const,text:'Collection',fontSizeMilliPoints:10000,font:{families:['DejaVu Sans'],weight:400,style:'normal' as const,stretch:100},script:'Latn',language:'en',direction:'ltr' as const}})
  if(!('face' in selected))throw new Error('Expected the collection face')
  expect(selected.face.collectionIndex).toBe(1)
  expect(await fonts.resolver.load(selected.face)).toHaveProperty('bytes')
  // A face identity forged with a different index is not the face the manifest pinned.
  expect(()=>fonts.resolver.load({...selected.face,collectionIndex:0})).toThrow(/binding/)
  // Two faces of one collection share its bytes: the cumulative budget counts the file once.
  const both=await loadHostFonts(input(),config([ttcEntry,{...ttcEntry,collectionIndex:0,family:'Collection Alias'}]),'approximate',[{family:'Collection Alias',weight:400,style:'normal'}])
  expect(both.manifest.faces.map(f=>f.source.collectionIndex).sort()).toEqual([0,1])
  expect([...both.resources.values()].map(r=>r.bytes.length)).toEqual([bytesOf(ttc.path),bytesOf(ttc.path)])
 })
 it('refuses a collection index that names no face, a collection with no index, and an index on a standalone sfnt',async()=>{
  await expect(loadHostFonts(input(),config([{...ttcEntry,collectionIndex:2}]))).rejects.toThrow(/collectionIndex does not identify a face in the font collection/)
  await expect(loadHostFonts(input(),config([{family:'DejaVu Sans',weight:400,style:'normal',path:ttc.path,sha256:ttc.sha256}]))).rejects.toThrow(/collectionIndex is required for TTC\/OTC bytes/)
  await expect(loadHostFonts(input(),config([{...entry,collectionIndex:0}]))).rejects.toThrow(/collectionIndex must be omitted for standalone sfnt bytes/)
  for(const collectionIndex of [-1,1.5,65536,'0',null])await expect(loadHostFonts(input(),config([{...ttcEntry,collectionIndex}]))).rejects.toThrow(/Invalid host font (face|collection index)/)
 })
})

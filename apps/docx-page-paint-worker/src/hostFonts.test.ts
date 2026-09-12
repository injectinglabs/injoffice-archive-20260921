import {afterAll,describe,expect,it} from 'vitest'
import {readFileSync,writeFileSync,mkdtempSync,rmSync} from 'node:fs'
import {resolve} from 'node:path'
import {tmpdir} from 'node:os'
import {createHash} from 'node:crypto'
import {loadHostFonts} from './hostFonts.js'
import {encodeNativeDOCXFontInventoryV1,nativeDOCXCanonicalWireSHA256V1} from '../../../packages/docs/src/nativeFontInventoryV1.js'
import type {NativeDocxPagePaintPrepareInputV1} from '@injoffice/docs/native-page-paint-compiler'

const root=resolve(import.meta.dirname,'../../..'),scratch=mkdtempSync(resolve(tmpdir(),'docx-host-fonts-'))
afterAll(()=>rmSync(scratch,{recursive:true,force:true}))
const path=resolve(root,'node_modules/dejavu-fonts-ttf/ttf/DejaVuSans.ttf')
const sha256=`sha256:${createHash('sha256').update(readFileSync(path)).digest('hex')}`
const entry={family:'DejaVu Sans',weight:400,style:'normal',path,sha256}
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
describe('operator-owned DOCX fonts',()=>{
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
})

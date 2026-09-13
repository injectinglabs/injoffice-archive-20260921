import {expect,it} from 'vitest'
import {readFileSync} from 'node:fs'
import {extractChartWorkbookReferences} from './chartWorkbookExtract.js'
import {parseChartWorkbookRange} from './chartWorkbookRange.js'

const bytes=new Uint8Array([1,2,3]),sha='039058c6f2c0cb492c533b0a4d14ef77cc0f78abccced5287d84a1a2011cfb81'
const binding={relationshipId:'rId1',part:'ppt/embeddings/data.xlsx',sha256:sha,byteLength:3}
const reference={kind:'numRef' as const,formula:'Data!A1',range:parseChartWorkbookRange('Data!A1'),cachePresent:true}

it('checks original bytes before invoking the injected engine',async()=>{
 let calls=0
 await expect(extractChartWorkbookReferences({...binding,sha256:'a'.repeat(64)},[reference],bytes,async()=>{calls++;return '{}'})).rejects.toThrow(/bytes do not match/)
 expect(calls).toBe(0)
})
it('gives the extractor a disposable copy and rejects malformed or duplicate-key results',async()=>{
 for(const result of ['{}','{"version":2,"version":2}','null']){
  const source=bytes.slice()
  await expect(extractChartWorkbookReferences(binding,[reference],source,async passed=>{expect(passed).not.toBe(source);passed.fill(0);return result})).rejects.toThrow(/invalid native V2/)
  expect(source).toEqual(bytes)
 }
})
it('bounds output before invoking the native decoder',async()=>{
 await expect(extractChartWorkbookReferences(binding,[reference],bytes,async()=> ' '.repeat(16*1024*1024+1))).rejects.toThrow(/JSON budget/)
})
it('validates a complete V2 callback contract before resolving source values',async()=>{
 // This is an explicitly trusted callback stub, not an XLSX engine proof.
 const workbook=JSON.parse(readFileSync(new URL('../../../go/xlsxpatch/testdata/native-xlsx-v2/valid/lexical-render.json',import.meta.url),'utf8'))
 workbook.source.package_sha256=`sha256:${sha}`;workbook.revision=`rev:${sha}`
 const ref={...reference,formula:'Lexical!A1',range:parseChartWorkbookRange('Lexical!A1')}
 const result=await extractChartWorkbookReferences(binding,[ref],bytes,async()=>JSON.stringify(workbook))
 expect(result[0]!.values).toEqual([workbook.sheets[0].cells[0].value.lexical])
 const duplicate=JSON.stringify(workbook).replace('"version":2','"version":2,"version":2')
 await expect(extractChartWorkbookReferences(binding,[ref],bytes,async()=>duplicate)).rejects.toThrow(/invalid native V2/)
 const mutableBinding={...binding},mutableRef={...ref,range:{...ref.range}}
 const snapshotted=await extractChartWorkbookReferences(mutableBinding,[mutableRef],bytes,async()=>{
  mutableBinding.relationshipId='changed';mutableBinding.part='changed.xlsx';mutableRef.formula='Lexical!Z9';mutableRef.range.startColumn=25
  return JSON.stringify(workbook)
 })
 expect(snapshotted[0]).toMatchObject({workbookPart:binding.part,workbookRelationshipId:'rId1',formula:'Lexical!A1'})
 workbook.source.package_sha256='sha256:'+'a'.repeat(64);workbook.revision='rev:'+'a'.repeat(64)
 await expect(extractChartWorkbookReferences(binding,[ref],bytes,async()=>JSON.stringify(workbook))).rejects.toThrow(/identity/)
})

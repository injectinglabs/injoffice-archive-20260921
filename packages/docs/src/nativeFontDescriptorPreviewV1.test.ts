import {readFileSync} from 'node:fs'
import {describe,it,expect} from 'vitest'
import {qualifyNativeDocxFontDescriptorPreviewV1} from './nativeFontDescriptorPreviewV1.js'
import {encodeNativeDOCXFontInventoryV1,nativeDOCXCanonicalWireSHA256V1} from './nativeFontInventoryV1.js'
function fixture(){
 const inventory=JSON.parse(readFileSync(new URL('../../../go/docxpatch/testdata/font-inventory-v1.json',import.meta.url),'utf8'))
 inventory.families.push({family_id:'font-family:unused',name:'Unused',faces:[]})
 inventory.inventory_sha256='';inventory.inventory_sha256=nativeDOCXCanonicalWireSHA256V1(inventory)
 const source={document_id:inventory.document_id,revision:inventory.revision,source:{package_sha256:inventory.package_sha256},passthrough_parts:[{part_name:inventory.font_table.part_name,sha256:inventory.font_table.sha256,content_type:'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml'},{part_name:inventory.font_table.main_relationships_part,sha256:inventory.font_table.main_relationships_sha256}]} as any
 const diagnostic={code:'FONT_MATCHING_METADATA_PRESERVED',severity:'unsupported',scope_id:inventory.document_id,part_name:inventory.font_table.part_name,path:'/w:fonts[1]/w:font[1]/w:charset[1]',preservation:'preserve-verbatim',message:'Validated font matching metadata is preserved; native painting requires exact supplied faces, not metadata-driven substitution'}
 const resolved:any={document_id:inventory.document_id,revision:inventory.revision,source_parts:{font_table_part:inventory.font_table.part_name},fonts:[{name:'Fixture Sans',alt_name:'Fixture Alias'},{name:'Unused'}],runs:[{properties:{font_family:'Fixture Sans'}}],paragraphs:[],diagnostics:[diagnostic]}
 const eligibility:any={protocol:'injoffice.docx.font-substitution-eligibility',version:1,document_id:inventory.document_id,revision:inventory.revision,package_sha256:inventory.package_sha256,font_table:inventory.font_table,facts:[{family:'Fixture Sans',use:'latin-matching',kind:'charset',path:diagnostic.path,values:{val:'00'},diagnostic}]}
 return {source,resolved,eligibility,inventoryJSON:encodeNativeDOCXFontInventoryV1(inventory)}
}
const qualify=(v:ReturnType<typeof fixture>)=>qualifyNativeDocxFontDescriptorPreviewV1(v.eligibility,v.source,v.resolved,v.inventoryJSON)
describe('source font descriptor preview eligibility',()=>{
 it('accepts the qualified panose1 source path with its literal digit',()=>{const v=fixture(),f=structuredClone(v.eligibility.facts[0]);f.kind='panose1';f.path=f.path.replace('charset','panose1');f.values.val='020F0502020204030204';f.diagnostic.path=f.path;v.eligibility.facts.push(f);v.resolved.diagnostics.push(f.diagnostic);expect(qualify(v).facts).toHaveLength(2)})
 it('joins complete inventory part binding and declaration order without changing source',()=>{const v=fixture(),before=structuredClone(v);expect(qualify(v)).toEqual(v.eligibility);expect(v).toEqual(before)})
 it.each([
  (v:ReturnType<typeof fixture>)=>{v.eligibility.facts[0].family='Unused';v.eligibility.facts[0].use='unused'},
  (v:ReturnType<typeof fixture>)=>{v.eligibility.facts[0].path='/w:fonts[1]/w:font[2]/w:charset[1]'},
  (v:ReturnType<typeof fixture>)=>{v.resolved.fonts.reverse()},
  (v:ReturnType<typeof fixture>)=>{v.resolved.fonts[1].alt_name='Fixture Alias';v.resolved.runs[0].properties.font_family='Fixture Alias'},
  (v:ReturnType<typeof fixture>)=>{v.eligibility.facts[0].values.val='02'},
  (v:ReturnType<typeof fixture>)=>{v.eligibility.facts[0].values.extra='1'},
  (v:ReturnType<typeof fixture>)=>{v.eligibility.facts[0].diagnostic.message='forged'},
  (v:ReturnType<typeof fixture>)=>{v.eligibility.font_table.relationship_target='Another.xml'},
  (v:ReturnType<typeof fixture>)=>{v.eligibility.facts=[]},
  (v:ReturnType<typeof fixture>)=>{v.eligibility.facts.push(structuredClone(v.eligibility.facts[0]))},
  (v:ReturnType<typeof fixture>)=>{v.source.source.package_sha256='sha256:'+'b'.repeat(64)},
 ])('refuses swapped owner, source or diagnostic %#',mutate=>{const v=fixture();v.eligibility=structuredClone(v.eligibility);mutate(v);expect(()=>qualify(v)).toThrow()})
})

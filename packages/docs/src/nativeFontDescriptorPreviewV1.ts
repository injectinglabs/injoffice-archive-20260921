import {preflightWire,canonicalWireSha256} from './nativePagePaintWireV1.js'
import type {NativeDocxDocumentV1} from './nativeContract.js'
import type {NativeDocxResolvedLayoutInputV1,NativeDocxResolutionDiagnosticV1} from './nativeResolvedLayout.js'
import {decodeNativeDOCXFontInventoryV1,type NativeDOCXFontTableBindingV1} from './nativeFontInventoryV1.js'

export const DOCX_FONT_DESCRIPTOR_PREVIEW_WARNING='Font matching descriptors are retained but not used to select replacement fonts. Explicit operator mappings are not Microsoft Word font matching; metrics and layout may differ.'
export interface NativeDocxFontDescriptorFactV1 {
 family:string
 use:'latin-matching'|'unused'
 kind:'charset'|'panose1'|'family'|'pitch'|'sig'|'notTrueType'
 path:string
 values:Record<string,string>
 diagnostic:NativeDocxResolutionDiagnosticV1
}
export interface NativeDocxFontSubstitutionEligibilityV1 {
 protocol:'injoffice.docx.font-substitution-eligibility'
 version:1
 document_id:string
 revision:string
 package_sha256:string
 font_table:NativeDOCXFontTableBindingV1|null
 facts:NativeDocxFontDescriptorFactV1[]
}
const exact=(v:unknown,keys:string)=>!!v&&typeof v==='object'&&!Array.isArray(v)&&Object.keys(v).sort().join(',')===keys
const same=(a:unknown,b:unknown)=>canonicalWireSha256(a)===canonicalWireSha256(b)
const hex=(v:unknown,n:number)=>typeof v==='string'&&v.length===n&&/^[a-fA-F0-9]+$/.test(v)

/** Requalifies every preserved descriptor, its source part, actual owner and
 * consumer set. Merely supplying a matching diagnostic code never grants use. */
export function qualifyNativeDocxFontDescriptorPreviewV1(value:unknown,document:NativeDocxDocumentV1,resolved:NativeDocxResolvedLayoutInputV1,inventoryJSON:string):NativeDocxFontSubstitutionEligibilityV1{
 if(preflightWire(value,'font descriptor eligibility',50000,1000).length)throw new TypeError('Unbounded font descriptor eligibility')
 const v=structuredClone(value) as NativeDocxFontSubstitutionEligibilityV1
 if(!exact(v,'document_id,facts,font_table,package_sha256,protocol,revision,version')||v.protocol!=='injoffice.docx.font-substitution-eligibility'||v.version!==1||v.document_id!==document.document_id||v.revision!==document.revision||v.package_sha256!==document.source.package_sha256||resolved.document_id!==v.document_id||resolved.revision!==v.revision||!Array.isArray(v.facts)||v.facts.length>1000)throw new TypeError('Font descriptor source identity mismatch')
 const diagnostics=resolved.diagnostics.filter(d=>d.code==='FONT_MATCHING_METADATA_PRESERVED')
 const inventory=decodeNativeDOCXFontInventoryV1(inventoryJSON)
 if(inventory.document_id!==document.document_id||inventory.revision!==document.revision||inventory.package_sha256!==document.source.package_sha256||!same(v.font_table,inventory.font_table??null))throw new TypeError('Descriptor inventory binding mismatch')
 if(v.font_table===null){if(v.facts.length||diagnostics.length||resolved.source_parts.font_table_part)throw new TypeError('Font descriptor source part missing');return v}
 const binding=v.font_table
 const keys=Object.keys(binding).filter(k=>!['font_relationships_part','font_relationships_sha256'].includes(k)).sort().join(',')
 if(keys!=='main_relationships_part,main_relationships_sha256,part_name,relationship_id,relationship_target,relationship_type,sha256'||binding.part_name!==resolved.source_parts.font_table_part||!document.passthrough_parts.some(p=>p.part_name===binding.part_name&&p.sha256===binding.sha256&&p.content_type==='application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml')||!document.passthrough_parts.some(p=>p.part_name===binding.main_relationships_part&&p.sha256===binding.main_relationships_sha256))throw new TypeError('Font descriptor part digest mismatch')
 const used=new Set<string>()
 const owners=new Map<string,number>()
 for(const font of resolved.fonts)for(const name of [font.name,...(font.alt_name?[font.alt_name]:[])]){const key=name.toLowerCase();owners.set(key,(owners.get(key)??0)+1)}
 const add=(family:string|undefined)=>{if(!family)throw new TypeError('Uncertain source font consumer');used.add(family.toLowerCase())}
 for(const r of resolved.runs)add(r.properties.font_family)
 for(const p of resolved.paragraphs){add(p.paragraph_mark_properties?.font_family);if(p.numbering)add(p.numbering.marker_properties.font_family)}
 for(const family of used)if((owners.get(family)??0)>1)throw new TypeError('Ambiguous source font alias')
 const fonts=new Map(resolved.fonts.map(f=>[f.name,f])),seen=new Set<string>(),groups=new Map<string,NativeDocxFontDescriptorFactV1[]>()
 const expected=new Set(diagnostics.map(d=>canonicalWireSha256(d)))
 for(const f of v.facts){
  if(!exact(f,'diagnostic,family,kind,path,use,values')||typeof f.family!=='string'||!fonts.has(f.family)||!['latin-matching','unused'].includes(f.use)||typeof f.path!=='string'||!/^\/w:fonts\[1\]\/w:font\[[1-9][0-9]*\]\/w:(?:charset|panose1|family|pitch|sig|notTrueType)\[1\]$/.test(f.path)||seen.has(f.path))throw new TypeError('Invalid descriptor owner or duplicate source')
  // Go loadFonts preserves declaration order; the eligibility producer rejects
  // foreign/non-font siblings, so this exact ordinal cannot be renumbered.
  const ordinal=Number(f.path.match(/\/w:font\[([0-9]+)\]/)![1])
  const owner=resolved.fonts[ordinal-1],inventoryOwner=inventory.families.find(x=>x.name===f.family)
  if(!Number.isSafeInteger(ordinal)||owner?.name!==f.family||!inventoryOwner||owner.alt_name!==inventoryOwner.alt_name)throw new TypeError('Descriptor ordinal does not join source font declaration')
  seen.add(f.path)
  const font=fonts.get(f.family)!,active=used.has(f.family.toLowerCase())||!!font.alt_name&&used.has(font.alt_name.toLowerCase())
  if(f.use!==(active?'latin-matching':'unused')||!f.path.endsWith(`/w:${f.kind}[1]`))throw new TypeError('Descriptor use differs from actual source consumers')
  if(!f.diagnostic||f.diagnostic.part_name!==binding.part_name||f.diagnostic.path!==f.path||f.diagnostic.scope_id!==document.document_id||!expected.delete(canonicalWireSha256(f.diagnostic)))throw new TypeError('Descriptor lacks exact preserved diagnostic identity')
  const values=f.values
  if(f.kind==='sig'){if(!exact(values,'csb0,csb1,usb0,usb1,usb2,usb3')||Object.values(values).some(v=>!hex(v,8)))throw new TypeError('Invalid font signature')}
  else if(f.kind==='notTrueType'){if(active||!exact(values,'')&&!exact(values,'val')||values.val!==undefined&&!['1','0','true','false','on','off'].includes(values.val))throw new TypeError('Active or malformed legacy descriptor')}
  else{
   if(!exact(values,'val'))throw new TypeError('Invalid descriptor value')
   const n=values.val
   if(f.kind==='charset'?!['00','EE','ee',...(!active?['02']:[])].includes(n):f.kind==='panose1'?!hex(n,20)||active&&!n.startsWith('02'):f.kind==='family'?!['auto','decorative','modern','roman','script','swiss'].includes(n):f.kind==='pitch'?!['default','fixed','variable'].includes(n):true)throw new TypeError('Unsupported font matching descriptor')
  }
  const ownerPath=f.path.slice(0,f.path.lastIndexOf('/')),group=groups.get(ownerPath)??[]
  if(group.some(x=>x.family!==f.family||x.kind===f.kind))throw new TypeError('Ambiguous descriptor owner')
  group.push(f);groups.set(ownerPath,group)
 }
 for(const group of groups.values())if(group[0]!.use==='latin-matching'&&!group.some(f=>f.kind==='charset'))throw new TypeError('Active descriptor charset missing')
 if(expected.size||v.facts.length!==diagnostics.length)throw new TypeError('Incomplete preserved font descriptor evidence')
 return v
}

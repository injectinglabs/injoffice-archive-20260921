import {readFileSync, statSync} from 'node:fs'
import {isAbsolute} from 'node:path'
import {createHash} from 'node:crypto'
import {createNativeDocxEmbeddedFontResolverV1, NativeDocxPreviewRefusalV1, type NativeDocxPagePaintPrepareInputV1, type NativeDocxHostFontsV1} from '@injoffice/docs/native-page-paint-compiler'
import {decodeNativeDOCXFontInventoryV1} from '@injoffice/docs/native-page-paint-compiler'
import {HARFBUZZ_SHAPER_LIMITS,inspectHarfBuzzFontMetricsV1} from '@injoffice/font-metrics/harfbuzz'
import type {FontResource, NativeFontManifest, ResolvedFontFace} from '@injoffice/font-metrics/layout'
import {decodeExplicitFontPolicyV1,selectExplicitFontV1} from '@injoffice/font-metrics/layout'
import {
  discloseNativeDocxHostFontApproximateSubstitutionReasonsV1,
  nativeDocxHostFontApproximateSubstitutionV1,
  selectLoadedHostManifestSubstituteV1,
  type NativeDocxHostFontApproximateSubstitutionV1,
} from './hostFontApproximateSubstitution.js'

const digest=(bytes:Uint8Array|string)=>`sha256:${createHash('sha256').update(bytes).digest('hex')}` as const
const fold=(s:string)=>s.replace(/[A-Z]/g,c=>c.toLowerCase())
const key=(family:string,weight:number,style:string)=>`${fold(family)}\0${weight}\0${style}`
function file(path:string,max:number):Uint8Array {
 if(!isAbsolute(path))throw new Error('Host font paths must be absolute operator configuration')
 const stat=statSync(path)
 if(!stat.isFile()||stat.size===0||stat.size>max)throw new Error('Host font file exceeds its size/type bound')
 const bytes=Uint8Array.from(readFileSync(path))
 if(bytes.length!==stat.size||bytes.length>max)throw new Error('Host font file changed while loading')
 return bytes
}
/** A TTC/OTC holds every face of a family group in one file, so a collection
 * carries a larger per-file bound than a standalone sfnt. The cumulative
 * budget is unchanged and each distinct file still counts exactly once. */
const STANDALONE_FONT_BYTES=16*1024*1024
const COLLECTION_FONT_BYTES=HARFBUZZ_SHAPER_LIMITS.maxFontBytes
/** The operator font manifest's face cap. `validateDOCXFontSubstitutionPreview`
 * (`go/injoffice-server/internal/officehttp/docx_font_substitution_preview.go`)
 * re-reads the same operator file on the substitution-preview path and encodes
 * this same number, so raising one without the other lets this worker load a
 * manifest that helper then refuses. 64 leaves every budget the cap protected
 * intact: the manifest file is still read under 64 KiB, each standalone face
 * under 16 MiB, a collection file under `COLLECTION_FONT_BYTES`, the cumulative
 * budget is still 64 MiB, and a face is still only read when the document
 * references it — the cap bounds the JSON, not the work. */
export const MAX_OPERATOR_FONT_FACES=64

export type HostFontLoadMode=boolean|'approximate'
/** Additional exact face identity a same-bytes sidecar asks the host to load (never a substitution). */
export interface HostFontReference {family:string;weight:number;style:'normal'|'italic'}
/** One operator-configured face: a standalone sfnt file, or one face of a TTC/OTC collection named by collectionIndex. */
interface HostFontConfiguredFace {family:string;weight:number;style:'normal'|'italic';sha256:`sha256:${string}`;path:string;collectionIndex?:number}
export type NativeDocxLoadedHostFontsV1=NativeDocxHostFontsV1 & {resources:Map<string,FontResource>;approximateSubstitutions?:NativeDocxHostFontApproximateSubstitutionV1[]}

function admitConfiguredFace(index:number,f:HostFontConfiguredFace,faces:NativeFontManifest['faces'][number][],resources:Map<string,FontResource>,occupied:Set<string>,total:number,loaded:Map<string,Uint8Array>):number{
 // A collection file admitted for a second face is already budgeted and
 // digest-checked; re-reading it would double-count the same bytes.
 let bytes=loaded.get(f.sha256)
 if(!bytes){
  bytes=file(f.path,f.collectionIndex===undefined?STANDALONE_FONT_BYTES:COLLECTION_FONT_BYTES);total+=bytes.length
  if(total>64*1024*1024||digest(bytes)!==f.sha256)throw new Error('Host font digest or cumulative byte budget failed')
  loaded.set(f.sha256,bytes)
 }
 const collection=f.collectionIndex===undefined?{}:{collectionIndex:f.collectionIndex}
 const faceId=`host-font-${index}-${f.sha256.slice(7,23)}`
 const face:ResolvedFontFace={faceId,family:f.family,weight:f.weight,style:f.style,stretch:100,sourceKind:'host',resourceId:faceId,contentDigest:f.sha256,...collection,resolution:'exact',matchedFamily:f.family}
 // The pinned HarfBuzz preflight is the only authority on the ttcf wrapper: it
 // refuses an index outside the collection and a collectionIndex on a
 // standalone sfnt, and reports that face's own design metrics.
 // A face it refuses is an unusable entry in the operator's own font manifest.
 // Letting the raw preflight sentence escape made the whole request fail under
 // the blanket COMPILATION_REFUSED code, with a message that named an sfnt table
 // and no face: the operator could not tell which of up to 64 configured files
 // to replace, and a reader could mistake it for a fact about the document.
 // Name the face and give the failure its own branchable code instead.
 let metrics:FontResource['metrics']
 try{metrics=inspectHarfBuzzFontMetricsV1({bytes,contentDigest:f.sha256,...collection})}
 catch(error){throw new NativeDocxPreviewRefusalV1('HOST_FONT_UNQUALIFIED',faceId,`Host font face ${f.family} / ${f.weight} / ${f.style}${f.collectionIndex===undefined?'':` / collection index ${f.collectionIndex}`} is not a qualified font face: ${error instanceof Error?error.message:String(error)}`)}
 resources.set(faceId,{face,bytes,metrics})
 faces.push({faceId,family:f.family,weight:f.weight,style:f.style,stretch:100,source:{kind:'host',resourceId:faceId,contentDigest:f.sha256,...collection}})
 occupied.add(key(f.family,f.weight,f.style))
 return total
}

/** Only called with a server-operator CLI path, never a document/request path. */
export async function loadHostFonts(input:NativeDocxPagePaintPrepareInputV1,path:string,allowSubstitution:HostFontLoadMode=false,extraReferences:readonly HostFontReference[]=[]):Promise<NativeDocxLoadedHostFontsV1> {
 const inventory=decodeNativeDOCXFontInventoryV1(input.font_inventory_json)
 const config=JSON.parse(Buffer.from(file(path,65536)).toString('utf8'))
 if(!config||Object.keys(config).filter(k=>k!=='substitutions').sort().join(',')!=='faces,version'||config.version!==1||!Array.isArray(config.faces))throw new Error('Invalid host font manifest')
 // The operator only ever sees the 422 body, so name the limit and the count.
 if(config.faces.length>MAX_OPERATOR_FONT_FACES)throw new Error(`Invalid host font manifest: ${config.faces.length} faces exceeds the ${MAX_OPERATOR_FONT_FACES}-face limit`)
 const configuredPolicy=config.substitutions===undefined?undefined:decodeExplicitFontPolicyV1(config.substitutions)
 const explicit=allowSubstitution===true,approximate=allowSubstitution==='approximate'
 const policy=explicit?configuredPolicy:undefined
 if(explicit&&!policy)throw new Error('Operator has not configured an explicit font substitution policy')
 const resources=new Map<string,FontResource>()
 const faces:NativeFontManifest['faces'][number][]=[...structuredClone(inventory.native_text_manifest?.faces??[])]
 const manifest:NativeFontManifest={version:1,manifestId:'docx-host-fonts',revision:inventory.revision,faces,fallbackChains:[]}
 const occupied=new Set(faces.flatMap(f=>[f.family,...(f.aliases??[])].map(family=>key(family,f.weight,f.style))))
 if(inventory.native_text_manifest){
  const embedded=createNativeDocxEmbeddedFontResolverV1(inventory,input.font_assets)
  for(const face of faces){
   if(!face.source.contentDigest)throw new Error('Embedded font digest is missing')
   const resolved:ResolvedFontFace={faceId:face.faceId,family:face.family,weight:face.weight,style:face.style,stretch:face.stretch,sourceKind:face.source.kind,resourceId:face.source.resourceId,contentDigest:face.source.contentDigest,collectionIndex:face.source.collectionIndex,resolution:'exact',matchedFamily:face.family}
   const resource=await embedded.load(resolved)
   if(!('bytes' in resource))throw new Error('Embedded font resource refused')
   resources.set(face.faceId,resource)
  }
 }
 const needed=new Set(inventory.references.map(r=>key(r.family,r.weight,r.style)))
 if(extraReferences.length>32)throw new Error('Host font sidecar requests exceed their bound')
 for(const mapping of policy?.mappings??[])if(needed.has(key(mapping.sourceFamily,mapping.weight,mapping.style)))needed.add(key(mapping.targetFamily,mapping.weight,mapping.style))
 const seen=new Set<string>();let total=[...resources.values()].reduce((n,r)=>n+r.bytes.length,0)
 if(total>64*1024*1024)throw new Error('Host font cumulative byte budget failed')
 const loaded=new Map<string,Uint8Array>()
 const configured:HostFontConfiguredFace[]=[]
 for(const [index,f] of config.faces.entries()){
  const shape=Object.keys(f??{}).sort().join(',')
  if(!f||(shape!=='family,path,sha256,style,weight'&&shape!=='collectionIndex,family,path,sha256,style,weight')||typeof f.family!=='string'||!f.family||f.family.length>128||typeof f.weight!=='number'||![400,700].includes(f.weight)||!['normal','italic'].includes(f.style)||typeof f.path!=='string'||!isAbsolute(f.path)||typeof f.sha256!=='string'||!/^sha256:[a-f0-9]{64}$/.test(f.sha256))throw new Error('Invalid host font face')
  if(f.collectionIndex!==undefined&&(!Number.isSafeInteger(f.collectionIndex)||f.collectionIndex<0||f.collectionIndex>65535))throw new Error('Invalid host font collection index')
  const identity=key(f.family,f.weight,f.style)
  if(seen.has(identity))throw new Error('Ambiguous host font family/style');seen.add(identity)
  configured[index]={family:f.family,weight:f.weight,style:f.style,sha256:f.sha256,path:f.path,...(f.collectionIndex===undefined?{}:{collectionIndex:f.collectionIndex})}
  if(occupied.has(identity)||!needed.has(identity))continue
  total=admitConfiguredFace(index,configured[index]!,faces,resources,occupied,total,loaded)
 }
 const missing=()=>inventory.references.filter(r=>!occupied.has(key(r.family,r.weight,r.style))&&!(policy&&faces.length&&selectExplicitFontV1(manifest,{version:1,text:'',fontSizeMilliPoints:1000,font:{families:[r.family],weight:r.weight,style:r.style,stretch:100},script:'Zyyy',language:'und',direction:'ltr'},policy)))
 const approximateSubstitutions:NativeDocxHostFontApproximateSubstitutionV1[]=[]
 if(approximate&&missing().length){
  const loadedHost=(weight:number,style:string)=>faces.some(face=>face.source.kind==='host'&&resources.has(face.faceId)&&face.weight===weight&&face.style===style&&face.stretch===100)
  for(const [index,f] of configured.entries()){
   if(!f||occupied.has(key(f.family,f.weight,f.style)))continue
   if(!missing().some(r=>r.weight===f.weight&&r.style===f.style)||loadedHost(f.weight,f.style))continue
   total=admitConfiguredFace(index,f,faces,resources,occupied,total,loaded)
  }
  // One authored family resolves to one target family. The loop above admits at most
  // one host face per missing weight/style, so a family needing 400 and 700 used to
  // take whichever family happened to own each slot and switched typeface mid-paragraph.
  // The first substitution for a source family fixes the target family; every later
  // weight/style of that family joins it, admitting that family's own configured face
  // when the slot is held by another family.
  const chosenFamily=new Map<string,string>()
  const loadedFamilyFace=(family:string,weight:number,style:string)=>faces.some(face=>face.source.kind==='host'&&resources.has(face.faceId)&&fold(face.family)===fold(family)&&face.weight===weight&&face.style===style&&face.stretch===100)
  for(const reference of missing()){
   const preferred=chosenFamily.get(fold(reference.family))
   if(preferred!==undefined&&!loadedFamilyFace(preferred,reference.weight,reference.style)){
    const index=configured.findIndex(f=>!!f&&fold(f.family)===fold(preferred)&&f.weight===reference.weight&&f.style===reference.style&&!occupied.has(key(f.family,f.weight,f.style)))
    // A face admitted only to keep one typeface is a nicety, not a requirement: an
    // operator entry that is oversized, mis-digested or unqualified must not turn a
    // working preview into a refusal. It falls back below and the disclosure says so.
    if(index>=0)try{total=admitConfiguredFace(index,configured[index]!,faces,resources,occupied,total,loaded)}catch{/* keep the fallback path */}
   }
   const substitute=selectLoadedHostManifestSubstituteV1(faces,new Set(resources.keys()),reference,preferred)
   if(!substitute||reference.style!=='normal'&&reference.style!=='italic')continue
   // The chosen family has no face at this weight/style anywhere in the operator
   // manifest, so this one reference cannot join it. Widening the search is disclosed
   // rather than silently splitting the typeface.
   const fallbackFrom=preferred!==undefined&&fold(preferred)!==fold(substitute.family)?preferred:undefined
   const record=nativeDocxHostFontApproximateSubstitutionV1({family:reference.family,weight:reference.weight,style:reference.style},substitute,fallbackFrom===undefined?undefined:{familyFallbackFrom:fallbackFrom})
   if(preferred===undefined)chosenFamily.set(fold(reference.family),substitute.family)
   const aliases=[...(substitute.aliases??[])]
   if(!aliases.some(alias=>fold(alias)===fold(reference.family)))aliases.push(reference.family)
   faces[faces.indexOf(substitute)]={...substitute,aliases}
   occupied.add(key(reference.family,reference.weight,reference.style))
   approximateSubstitutions.push(record)
  }
 }
 // Sidecar requests (equation faces) are admitted only now, after body
 // substitution is settled: they never count as document references, never
 // serve as body substitutes, and an unavailable one is not a failure.
 const requested=new Set(extraReferences.map(r=>key(r.family,r.weight,r.style)))
 for(const [index,f] of configured.entries()){
  if(!f||occupied.has(key(f.family,f.weight,f.style))||!requested.has(key(f.family,f.weight,f.style)))continue
  total=admitConfiguredFace(index,f,faces,resources,occupied,total,loaded)
 }
 const unavailable=missing()
 if(unavailable.length)throw new Error(`Exact configured font unavailable: ${unavailable.map(r=>`${r.family} / ${r.weight} / ${r.style}`).join(', ').slice(0,1024)}`)
 const manifestDigest=digest(JSON.stringify(manifest))
 const resolver:NativeDocxHostFontsV1['resolver']={providerId:'injoffice.docx.host-fonts',providerRevision:policy||approximateSubstitutions.length?digest(JSON.stringify([manifest,policy??null,approximateSubstitutions])):manifestDigest,
  resolve(request){
   if(digest(JSON.stringify(request.manifest))!==manifestDigest)throw new Error('Host font manifest binding changed')
   const selected=selectExplicitFontV1(manifest,request.run,policy)
   if(selected)return {status:'resolved',face:selected.face,attemptedFaceIds:[selected.face.faceId],decisions:[]}
   return {status:'refused',attemptedFaceIds:[],decisions:[{code:'font-not-found',message:'No exact configured font matches the authored family/style',recoverable:false}]}
  },
  load(face){
   const resource=resources.get(face.faceId)
   const fields=['faceId','family','postscriptName','weight','style','stretch','sourceKind','resourceId','contentDigest','collectionIndex','fallbackChainId'] as const
   if(!resource||fields.some(k=>face[k]!==resource.face[k]))throw new Error('Host font resource binding changed')
   const selected=selectExplicitFontV1(manifest,{version:1,text:'A',fontSizeMilliPoints:1000,font:{families:[face.matchedFamily],weight:face.weight,style:face.style,stretch:face.stretch},script:'Latn',language:'und',direction:'ltr'},policy)
   if(!selected||selected.face.faceId!==face.faceId||selected.face.resolution!==face.resolution)throw new Error('Host font substitution policy binding changed')
   return {...resource,face:{...face},bytes:Uint8Array.from(resource.bytes)}
  },
 }
 return {manifest,resolver,resources,...(policy?{substitutionPolicy:policy}:{}),...(approximateSubstitutions.length?{approximateSubstitutions}:{})}
}

export function discloseApproximateHostFontSubstitutions(result:unknown,fonts?:NativeDocxLoadedHostFontsV1):unknown {
 if(!fonts?.approximateSubstitutions?.length||!result||typeof result!=='object'||Array.isArray(result)||!('reasons' in result)||!Array.isArray(result.reasons)||result.reasons.some(reason=>typeof reason!=='string'))return result
 const reasons=result.reasons as string[]
 const extra=discloseNativeDocxHostFontApproximateSubstitutionReasonsV1(fonts.approximateSubstitutions).filter(reason=>!reasons.includes(reason))
 if(!extra.length||reasons.length+extra.length>264)return result
 return {...result,reasons:[...reasons,...extra]}
}

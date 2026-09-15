import {readFileSync, statSync} from 'node:fs'
import {isAbsolute} from 'node:path'
import {createHash} from 'node:crypto'
import {createNativeDocxEmbeddedFontResolverV1, type NativeDocxPagePaintPrepareInputV1, type NativeDocxHostFontsV1} from '@injoffice/docs/native-page-paint-compiler'
import {decodeNativeDOCXFontInventoryV1} from '@injoffice/docs/native-page-paint-compiler'
import {inspectHarfBuzzFontMetricsV1} from '@injoffice/font-metrics/harfbuzz'
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

export type HostFontLoadMode=boolean|'approximate'
/** Additional exact face identity a same-bytes sidecar asks the host to load (never a substitution). */
export interface HostFontReference {family:string;weight:400|700;style:'normal'|'italic'}
export type NativeDocxLoadedHostFontsV1=NativeDocxHostFontsV1 & {resources:Map<string,FontResource>;approximateSubstitutions?:NativeDocxHostFontApproximateSubstitutionV1[]}

function admitConfiguredFace(index:number,f:{family:string,weight:number,style:'normal'|'italic',sha256:`sha256:${string}`,path:string},faces:NativeFontManifest['faces'][number][],resources:Map<string,FontResource>,occupied:Set<string>,total:number):number{
 const bytes=file(f.path,16*1024*1024);total+=bytes.length
 if(total>64*1024*1024||digest(bytes)!==f.sha256)throw new Error('Host font digest or cumulative byte budget failed')
 const faceId=`host-font-${index}-${f.sha256.slice(7,23)}`
 const face:ResolvedFontFace={faceId,family:f.family,weight:f.weight,style:f.style,stretch:100,sourceKind:'host',resourceId:faceId,contentDigest:f.sha256,resolution:'exact',matchedFamily:f.family}
 resources.set(faceId,{face,bytes,metrics:inspectHarfBuzzFontMetricsV1({bytes,contentDigest:f.sha256})})
 faces.push({faceId,family:f.family,weight:f.weight,style:f.style,stretch:100,source:{kind:'host',resourceId:faceId,contentDigest:f.sha256}})
 occupied.add(key(f.family,f.weight,f.style))
 return total
}

/** Only called with a server-operator CLI path, never a document/request path. */
export async function loadHostFonts(input:NativeDocxPagePaintPrepareInputV1,path:string,allowSubstitution:HostFontLoadMode=false,extraReferences:readonly HostFontReference[]=[]):Promise<NativeDocxLoadedHostFontsV1> {
 const inventory=decodeNativeDOCXFontInventoryV1(input.font_inventory_json)
 const config=JSON.parse(Buffer.from(file(path,65536)).toString('utf8'))
 if(!config||Object.keys(config).filter(k=>k!=='substitutions').sort().join(',')!=='faces,version'||config.version!==1||!Array.isArray(config.faces)||config.faces.length>32)throw new Error('Invalid host font manifest')
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
 const configured:{family:string,weight:number,style:'normal'|'italic',sha256:`sha256:${string}`,path:string}[]=[]
 for(const [index,f] of config.faces.entries()){
  if(!f||Object.keys(f).sort().join(',')!=='family,path,sha256,style,weight'||typeof f.family!=='string'||!f.family||f.family.length>128||typeof f.weight!=='number'||![400,700].includes(f.weight)||!['normal','italic'].includes(f.style)||typeof f.path!=='string'||!isAbsolute(f.path)||typeof f.sha256!=='string'||!/^sha256:[a-f0-9]{64}$/.test(f.sha256))throw new Error('Invalid host font face')
  const identity=key(f.family,f.weight,f.style)
  if(seen.has(identity))throw new Error('Ambiguous host font family/style');seen.add(identity)
  configured[index]={family:f.family,weight:f.weight,style:f.style,sha256:f.sha256,path:f.path}
  if(occupied.has(identity)||!needed.has(identity))continue
  total=admitConfiguredFace(index,configured[index]!,faces,resources,occupied,total)
 }
 const missing=()=>inventory.references.filter(r=>!occupied.has(key(r.family,r.weight,r.style))&&!(policy&&faces.length&&selectExplicitFontV1(manifest,{version:1,text:'',fontSizeMilliPoints:1000,font:{families:[r.family],weight:r.weight,style:r.style,stretch:100},script:'Zyyy',language:'und',direction:'ltr'},policy)))
 const approximateSubstitutions:NativeDocxHostFontApproximateSubstitutionV1[]=[]
 if(approximate&&missing().length){
  const loadedHost=(weight:number,style:string)=>faces.some(face=>face.source.kind==='host'&&resources.has(face.faceId)&&face.weight===weight&&face.style===style&&face.stretch===100)
  for(const [index,f] of configured.entries()){
   if(!f||occupied.has(key(f.family,f.weight,f.style)))continue
   if(!missing().some(r=>r.weight===f.weight&&r.style===f.style)||loadedHost(f.weight,f.style))continue
   total=admitConfiguredFace(index,f,faces,resources,occupied,total)
  }
  const loadedIds=new Set(resources.keys())
  for(const reference of missing()){
   const substitute=selectLoadedHostManifestSubstituteV1(faces,loadedIds,reference)
   if(!substitute||reference.style!=='normal'&&reference.style!=='italic')continue
   const record=nativeDocxHostFontApproximateSubstitutionV1({family:reference.family,weight:reference.weight,style:reference.style},substitute)
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
  total=admitConfiguredFace(index,f,faces,resources,occupied,total)
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

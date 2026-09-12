import {readFileSync, statSync} from 'node:fs'
import {isAbsolute} from 'node:path'
import {createHash} from 'node:crypto'
import {createNativeDocxEmbeddedFontResolverV1, type NativeDocxPagePaintPrepareInputV1, type NativeDocxHostFontsV1} from '@injoffice/docs/native-page-paint-compiler'
import {decodeNativeDOCXFontInventoryV1} from '@injoffice/docs/native-page-paint-compiler'
import {inspectHarfBuzzFontMetricsV1} from '@injoffice/font-metrics/harfbuzz'
import type {FontResource, NativeFontManifest, ResolvedFontFace} from '@injoffice/font-metrics/layout'
import {decodeExplicitFontPolicyV1,selectExplicitFontV1} from '@injoffice/font-metrics/layout'

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

/** Only called with a server-operator CLI path, never a document/request path. */
export async function loadHostFonts(input:NativeDocxPagePaintPrepareInputV1,path:string,allowSubstitution=false):Promise<NativeDocxHostFontsV1 & {resources:Map<string,FontResource>}> {
 const inventory=decodeNativeDOCXFontInventoryV1(input.font_inventory_json)
 const config=JSON.parse(Buffer.from(file(path,65536)).toString('utf8'))
 if(!config||Object.keys(config).filter(k=>k!=='substitutions').sort().join(',')!=='faces,version'||config.version!==1||!Array.isArray(config.faces)||config.faces.length>32)throw new Error('Invalid host font manifest')
 const configuredPolicy=config.substitutions===undefined?undefined:decodeExplicitFontPolicyV1(config.substitutions)
 const policy=allowSubstitution?configuredPolicy:undefined
 if(allowSubstitution&&!policy)throw new Error('Operator has not configured an explicit font substitution policy')
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
 for(const mapping of policy?.mappings??[])if(needed.has(key(mapping.sourceFamily,mapping.weight,mapping.style)))needed.add(key(mapping.targetFamily,mapping.weight,mapping.style))
 const seen=new Set<string>();let total=[...resources.values()].reduce((n,r)=>n+r.bytes.length,0)
 if(total>64*1024*1024)throw new Error('Host font cumulative byte budget failed')
 for(const [index,f] of config.faces.entries()){
  if(!f||Object.keys(f).sort().join(',')!=='family,path,sha256,style,weight'||typeof f.family!=='string'||!f.family||f.family.length>128||typeof f.weight!=='number'||![400,700].includes(f.weight)||!['normal','italic'].includes(f.style)||typeof f.path!=='string'||!isAbsolute(f.path)||typeof f.sha256!=='string'||!/^sha256:[a-f0-9]{64}$/.test(f.sha256))throw new Error('Invalid host font face')
  const identity=key(f.family,f.weight,f.style)
  if(seen.has(identity))throw new Error('Ambiguous host font family/style');seen.add(identity)
  if(occupied.has(identity)||!needed.has(identity))continue
  const bytes=file(f.path,16*1024*1024);total+=bytes.length
  if(total>64*1024*1024||digest(bytes)!==f.sha256)throw new Error('Host font digest or cumulative byte budget failed')
  const faceId=`host-font-${index}-${f.sha256.slice(7,23)}`
  const face:ResolvedFontFace={faceId,family:f.family,weight:f.weight,style:f.style,stretch:100,sourceKind:'host',resourceId:faceId,contentDigest:f.sha256,resolution:'exact',matchedFamily:f.family}
  resources.set(faceId,{face,bytes,metrics:inspectHarfBuzzFontMetricsV1({bytes,contentDigest:f.sha256})})
  faces.push({faceId,family:f.family,weight:f.weight,style:f.style,stretch:100,source:{kind:'host',resourceId:faceId,contentDigest:f.sha256}})
  occupied.add(identity)
 }
 const missing=inventory.references.filter(r=>!occupied.has(key(r.family,r.weight,r.style))&&!(policy&&faces.length&&selectExplicitFontV1(manifest,{version:1,text:'',fontSizeMilliPoints:1000,font:{families:[r.family],weight:r.weight,style:r.style,stretch:100},script:'Zyyy',language:'und',direction:'ltr'},policy)))
 if(missing.length)throw new Error(`Exact configured font unavailable: ${missing.map(r=>`${r.family} / ${r.weight} / ${r.style}`).join(', ').slice(0,1024)}`)
 const manifestDigest=digest(JSON.stringify(manifest))
 const resolver:NativeDocxHostFontsV1['resolver']={providerId:'injoffice.docx.host-fonts',providerRevision:policy?digest(JSON.stringify([manifest,policy])):manifestDigest,
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
 return {manifest,resolver,resources,...(policy?{substitutionPolicy:policy}:{})}
}

import {validateFontManifest, validateTextRunInput, type NativeFontManifest, type NativeFontRequest, type NativeFontFaceManifest, type ResolvedFontFace, type TextRunInput} from './layout.js'

export const EXPLICIT_FONT_POLICY_V1 = 'explicit-whole-run-font-substitution-v1' as const
export interface ExplicitFontMappingV1 {
  readonly sourceFamily: string
  readonly weight: 400 | 700
  readonly style: 'normal' | 'italic'
  readonly targetFamily: string
}
export interface ExplicitFontPolicyV1 {
  readonly version: 1
  readonly mappings: readonly ExplicitFontMappingV1[]
}
export interface ExplicitFontSelectionV1 {
  readonly face: ResolvedFontFace
  readonly policy: typeof EXPLICIT_FONT_POLICY_V1 | null
}
const fold=(s:string)=>s.normalize('NFKC').trim().toLowerCase()
function record(value:unknown,keys:readonly string[]):Record<string,unknown>{
  if(!value||typeof value!=='object'||Array.isArray(value)||![Object.prototype,null].includes(Object.getPrototypeOf(value)))throw new TypeError('Invalid font policy record')
  const descriptors=Object.getOwnPropertyDescriptors(value)
  if(Reflect.ownKeys(value).length!==keys.length||keys.some(k=>!descriptors[k]||!('value' in descriptors[k]!)))throw new TypeError('Font policy requires exact data fields')
  return Object.fromEntries(keys.map(k=>[k,descriptors[k]!.value]))
}
/** Explicit host configuration, not inferred Office or system fallback rules. */
export function decodeExplicitFontPolicyV1(value:unknown):ExplicitFontPolicyV1 {
  const v=record(value,['version','mappings'])
  if(v.version!==1||!Array.isArray(v.mappings)||Object.getPrototypeOf(v.mappings)!==Array.prototype)throw new TypeError('Invalid bounded font substitution policy')
  const lengthDescriptor=Object.getOwnPropertyDescriptor(v.mappings,'length')
  const count=lengthDescriptor&&'value'in lengthDescriptor?lengthDescriptor.value:undefined
  if(!Number.isSafeInteger(count)||count<0||count>32||Reflect.ownKeys(v.mappings).length!==count+1)throw new TypeError('Invalid bounded font mapping array')
  const owned:unknown[]=[]
  for(let i=0;i<count;i++){
    const descriptor=Object.getOwnPropertyDescriptor(v.mappings,String(i))
    if(!descriptor||!('value'in descriptor))throw new TypeError('Font mappings must be dense data entries')
    owned.push(descriptor.value)
  }
  const seen=new Set<string>()
  const mappings:ExplicitFontMappingV1[]=[]
  for(const entry of owned){
    const m=record(entry,['sourceFamily','weight','style','targetFamily'])
    if(typeof m.sourceFamily!=='string'||typeof m.targetFamily!=='string'||[m.sourceFamily,m.targetFamily].some(n=>n.length===0||n.length>128||n.trim()!==n||!/^[\x20-\x7e]+$/.test(n))||![400,700].includes(m.weight as number)||!['normal','italic'].includes(m.style as string)||fold(m.sourceFamily)===fold(m.targetFamily))throw new TypeError('Invalid font substitution mapping')
    const identity=JSON.stringify([fold(m.sourceFamily),m.weight,m.style])
    if(seen.has(identity))throw new TypeError('Duplicate source font mapping')
    seen.add(identity)
    mappings.push(Object.freeze(m as unknown as ExplicitFontMappingV1))
  }
  return Object.freeze({version:1,mappings:Object.freeze(mappings)})
}
/** Bounded canonical policy material; hosts hash it into their provider revision. */
export function canonicalExplicitFontPolicyV1(value:unknown):string {
  return JSON.stringify(decodeExplicitFontPolicyV1(value))
}
function matches(face:NativeFontFaceManifest,font:NativeFontRequest,family:string):boolean {
  return face.weight===font.weight&&face.style===font.style&&face.stretch===font.stretch&&[face.family,...(face.aliases??[])].some(n=>fold(n)===fold(family))&&(font.postscriptName===undefined||face.postscriptName===font.postscriptName)
}
/** Exact requests always win. Substitution is whole-run and never changes source text. */
export function selectExplicitFontV1(manifest:NativeFontManifest,run:TextRunInput,policy?:ExplicitFontPolicyV1):ExplicitFontSelectionV1|null {
  const checked=policy===undefined?undefined:decodeExplicitFontPolicyV1(policy)
  if(!validateFontManifest(manifest).ok||!validateTextRunInput(run).ok)throw new TypeError('Invalid font selection input')
  const choose=(family:string,target:string,substitute:boolean):ExplicitFontSelectionV1|null=>{
    const candidates=manifest.faces.filter(f=>matches(f,run.font,target))
    if(candidates.length>1)throw new TypeError('Ambiguous explicit font selection')
    const f=candidates[0]
    if(!f)return null
    if(!f.source.contentDigest||f.source.kind==='system'||substitute&&f.source.kind!=='host')throw new TypeError('Font selection requires content-addressed supplied faces')
    return {face:{faceId:f.faceId,family:f.family,...(f.postscriptName?{postscriptName:f.postscriptName}:{}),weight:f.weight,style:f.style,stretch:f.stretch,sourceKind:f.source.kind,resourceId:f.source.resourceId,contentDigest:f.source.contentDigest,...(f.source.collectionIndex===undefined?{}:{collectionIndex:f.source.collectionIndex}),resolution:substitute?'substitute':'exact',matchedFamily:family},policy:substitute?EXPLICIT_FONT_POLICY_V1:null}
  }
  for(const family of run.font.families){const exact=choose(family,family,false);if(exact)return exact}
  if(!checked)return null
  // PostScript identity and vertical/script-specific font selection stay exact.
  if(run.font.postscriptName!==undefined||run.direction!=='ltr'||run.script!=='Latn'||run.font.stretch!==100)return null
  for(const family of run.font.families){
    if(!/^[\x20-\x7e]+$/.test(family)||family.trim()!==family)continue
    const mapping=checked.mappings.find(m=>fold(m.sourceFamily)===fold(family)&&m.weight===run.font.weight&&m.style===run.font.style)
    if(mapping){const selected=choose(family,mapping.targetFamily,true);if(selected)return selected}
  }
  return null
}

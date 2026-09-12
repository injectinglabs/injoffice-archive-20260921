import {describe,it,expect} from 'vitest'
import {decodeExplicitFontPolicyV1,selectExplicitFontV1,canonicalExplicitFontPolicyV1} from './explicitFontPolicy.js'
import type {NativeFontManifest,TextRunInput} from './layout.js'
const face=(family:string)=>({faceId:family,family,weight:400,style:'normal' as const,stretch:100,source:{kind:'host' as const,resourceId:family,contentDigest:`sha256:${'a'.repeat(64)}` as const}})
const manifest:NativeFontManifest={version:1,manifestId:'test',revision:'v1',faces:[face('Selected')],fallbackChains:[]}
const run:TextRunInput={version:1,text:'Hello',fontSizeMilliPoints:12000,font:{families:['Authored'],weight:400,style:'normal',stretch:100},script:'Latn',language:'en',direction:'ltr'}
const policy={version:1 as const,mappings:[{sourceFamily:'Authored',targetFamily:'Selected',weight:400 as const,style:'normal' as const}]}
describe('explicit whole-run font policy',()=>{
 it('requires explicit policy and retains authored identity with selected bytes',()=>{
  expect(selectExplicitFontV1(manifest,run)).toBeNull()
  const selected=selectExplicitFontV1(manifest,run,policy)!
  expect(selected.face).toMatchObject({family:'Selected',matchedFamily:'Authored',resolution:'substitute',contentDigest:manifest.faces[0]!.source.contentDigest})
  expect(run.text).toBe('Hello')
  expect(selectExplicitFontV1({...manifest,faces:[...manifest.faces,face('Authored')]},run,policy)!.face.resolution).toBe('exact')
  expect(canonicalExplicitFontPolicyV1(policy)).not.toBe(canonicalExplicitFontPolicyV1({...policy,mappings:[{...policy.mappings[0]!,targetFamily:'Another'}]}))
 })
 it('does not synthesize styles, discover fallback or substitute vertical/script-specific runs',()=>{
  expect(selectExplicitFontV1(manifest,{...run,text:'',script:'Zyyy'},policy)?.face.resolution).toBe('substitute')
  expect(selectExplicitFontV1(manifest,{...run,text:'123 .',script:'Zyyy'},policy)?.face.resolution).toBe('substitute')
  expect(selectExplicitFontV1(manifest,{...run,text:'☃',script:'Zyyy'},policy)).toBeNull()
  for(const r of [{...run,font:{...run.font,weight:700}},{...run,script:'Arab'},{...run,direction:'ttb' as const},{...run,font:{...run.font,postscriptName:'AuthoredPS'}}])expect(selectExplicitFontV1(manifest,r,policy)).toBeNull()
  expect(selectExplicitFontV1({...manifest,faces:[face('Other')]},run,policy)).toBeNull()
  expect(()=>selectExplicitFontV1({...manifest,faces:[face('Selected'),{...face('Selected'),faceId:'duplicate'}]},run,policy)).toThrow()
 })
 it('rejects malformed, duplicate, recursive and accessor policy inputs without getters',()=>{
  let invoked=false
  const getter=Object.defineProperty({},'version',{get(){invoked=true;return 1}})
  for(const input of [getter,{...policy,unknown:true},{...policy,mappings:Array(33).fill(policy.mappings[0])},{...policy,mappings:[policy.mappings[0],policy.mappings[0]]},{...policy,mappings:[{...policy.mappings[0],targetFamily:'Authored'}]},{...policy,mappings:[{...policy.mappings[0],weight:500}]},{...policy,mappings:[{...policy.mappings[0],sourceFamily:'Authored\u0000'}]}])expect(()=>decodeExplicitFontPolicyV1(input)).toThrow()
  expect(invoked).toBe(false)
  const extra=Object.assign([policy.mappings[0]],{extra:true}),symbol=[policy.mappings[0]],sparse=new Array(1)
  Object.defineProperty(symbol,Symbol('extra'),{value:1})
  for(const mappings of [extra,symbol,sparse,Object.setPrototypeOf([policy.mappings[0]],{})])expect(()=>decodeExplicitFontPolicyV1({...policy,mappings})).toThrow()
  const proxy=new Proxy([policy.mappings[0]],{get(target,key,receiver){if(key==='length'){invoked=true;throw Error('length getter')};return Reflect.get(target,key,receiver)}})
  expect(decodeExplicitFontPolicyV1({...policy,mappings:proxy}).mappings).toHaveLength(1)
  expect(invoked).toBe(false)
 })
})

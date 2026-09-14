import type {NativeLiteralRadar,NativeRadarData} from './chartRadarTypes.js'
import {validNativeLiteralArea} from './chartAreaValidation.js'
import {nativeChartDecimal} from './chartDecimalValidation.js'
const integer=(v:unknown,min:number,max:number)=>typeof v==='number'&&Number.isSafeInteger(v)&&!Object.is(v,-0)&&v>=min&&v<=max
const rgb=(v:unknown)=>typeof v==='string'&&v.length===7&&/^#[0-9A-F]{6}$/.test(v)
function record(v:unknown,required:string,optional=''):v is Record<string,unknown>{
 if(!v||typeof v!=='object'||Array.isArray(v))return false
 const keys=new Set([...required.split(' '),...optional.split(' ').filter(Boolean)])
 return required.split(' ').every(k=>Object.hasOwn(v,k))&&Object.keys(v).every(k=>keys.has(k))
}
function dense(v:unknown,min:number,max:number):v is unknown[]{
 if(!Array.isArray(v)||v.length<min||v.length>max)return false
 for(let i=0;i<v.length;i++)if(!Object.hasOwn(v,i))return false
 return true
}
/** Closed first radial profile. Cartesian validation is reused only for exact
 * category/value/axis source records, never for radial geometry or label layout. */
export function validNativeLiteralRadar(v:unknown):v is NativeLiteralRadar{
 if(!record(v,'profile dataOrigin style categories series categoryAxis valueAxis')||v.profile!=='literal-radar-v1'||v.dataOrigin!=='literal')return false
 const {profile,dataOrigin,...data}=v;return validNativeRadarData(data)
}
export function validNativeRadarData(v:unknown):v is NativeRadarData{
 if(!record(v,'style categories series categoryAxis valueAxis')||!['standard','filled'].includes(v.style as string)||!dense(v.categories,3,256)||!dense(v.series,1,16))return false
 const projected=[],orders=new Set<number>()
 for(const s of v.series){
  if(!record(s,'index order values color widthEmu','title fill')||!integer(s.order,0,v.series.length-1)||orders.has(s.order as number)||!integer(s.widthEmu,1,20116800)||!rgb(s.color)||!dense(s.values,v.categories.length,v.categories.length))return false
  if(v.style==='filled'?!rgb(s.fill):Object.hasOwn(s,'fill'))return false
  orders.add(s.order as number)
  projected.push({index:s.index,order:projected.length,values:s.values,color:s.color,...(s.title!==undefined?{title:s.title}:{})})
 }
 const x=v.categoryAxis,y=v.valueAxis
 if(!record(x,'id crossAxisId orientation position deleted','color widthEmu')||!record(y,'id crossAxisId orientation position deleted min max crossesAt','color widthEmu'))return false
 if(!validNativeLiteralArea({profile:'literal-area-v1',dataOrigin:'literal',grouping:'standard',categories:v.categories,series:projected,xAxis:x,yAxis:y}))return false
 // Outside-scale radar polygons have separate reference/clipping obligations.
 // Refuse them rather than clamp data or invent opposite-spoke behavior.
 const low=nativeChartDecimal(y.min as string)!,high=nativeChartDecimal(y.max as string)!
 const compare=(a:typeof low,b:typeof low)=>{const e=Math.min(a.exponent,b.exponent),aa=a.coefficient*10n**BigInt(a.exponent-e),bb=b.coefficient*10n**BigInt(b.exponent-e);return aa<bb?-1:aa>bb?1:0}
 return projected.every(s=>s.values.every(raw=>{const value=nativeChartDecimal(raw as string)!;return compare(value,low)>=0&&compare(value,high)<=0}))
}

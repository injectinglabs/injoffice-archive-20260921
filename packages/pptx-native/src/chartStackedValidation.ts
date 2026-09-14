import type {NativeLiteralStackedBar,NativeLiteralStackedLine} from './chartStackedTypes.js'
import {validNativeLiteralArea} from './chartAreaValidation.js'

const integer=(v:unknown,min:number,max:number)=>typeof v==='number'&&Number.isInteger(v)&&!Object.is(v,-0)&&v>=min&&v<=max
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
function validate(value:unknown,bar:boolean):boolean{
 if(!record(value,bar?'profile dataOrigin grouping overlap barDirection gapWidth categories series categoryAxis valueAxis':'profile dataOrigin grouping categories series xAxis yAxis')||value.profile!==(bar?'literal-stacked-bar-v1':'literal-stacked-line-v1')||value.dataOrigin!=='literal'||!['stacked','percentStacked'].includes(value.grouping as string)||!dense(value.categories,1,256)||!dense(value.series,1,16))return false
 if(bar&&(!['bar','column'].includes(value.barDirection as string)||value.overlap!==100||!integer(value.gapWidth,0,500)))return false
 const projected=[],orders=new Set<number>()
 for(const raw of value.series){
  if(!record(raw,bar?'index order values colors':'index order values color widthEmu','title')||!dense(raw.values,value.categories.length,value.categories.length))return false
  if(!integer(raw.order,0,value.series.length-1)||orders.has(raw.order as number))return false
  orders.add(raw.order as number)
  let color=raw.color
  if(bar){if(!dense(raw.colors,value.categories.length,value.categories.length)||raw.colors.some(c=>!rgb(c)))return false;color=raw.colors[0]}
  else if(!integer(raw.widthEmu,1,20116800))return false
  projected.push({index:raw.index,order:projected.length,values:raw.values,color,...(raw.title!==undefined?{title:raw.title}:{})})
 }
 const x=bar?value.categoryAxis:value.xAxis,y=bar?value.valueAxis:value.yAxis
 if(!record(x,'id crossAxisId orientation position deleted','color widthEmu min max crossesAt labels')||!record(y,'id crossAxisId orientation position deleted','color widthEmu min max crossesAt labels'))return false
 const xAxis=x as Record<string,unknown>,yAxis=y as Record<string,unknown>
 const horizontal=bar&&value.barDirection==='bar'
 if(xAxis.position!==(horizontal?'l':'b')||yAxis.position!==(horizontal?'b':'l'))return false
 // Standard-area admission shares dense signed decimals, closed axis records,
 // exact linear-scale guards and full supplied-font label semantics. This local
 // projection changes no source object and does not reuse area stacking rules.
 // Numeric c:order metadata is privately normalized only for that validator;
 // actual stacked geometry/paint follows the original XML array sequence.
 return validNativeLiteralArea({profile:'literal-area-v1',dataOrigin:'literal',grouping:'standard',categories:value.categories,series:projected,xAxis:{...xAxis,position:'b'},yAxis:{...yAxis,position:'l'}})
}
export function validNativeLiteralStackedBar(value:unknown):value is NativeLiteralStackedBar{return validate(value,true)}
export function validNativeLiteralStackedLine(value:unknown):value is NativeLiteralStackedLine{return validate(value,false)}

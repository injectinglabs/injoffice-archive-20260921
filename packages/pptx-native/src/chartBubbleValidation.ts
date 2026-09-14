import type {NativeLiteralBubble} from './chartBubbleTypes.js'
import type {NativeLiteralBarAxis} from './types.js'
import {nativeChartDecimal as decimal} from './chartDecimalValidation.js'
import {validNativeChartAxisLabels} from './chartAxisLabelsValidation.js'
function record(value:unknown,required:string,optional=''):value is Record<string,unknown>{
 if(!value||typeof value!=='object'||Array.isArray(value))return false
 const keys=new Set([...required.split(' '),...optional.split(' ')].filter(Boolean))
 return required.split(' ').every(k=>Object.hasOwn(value,k))&&Object.keys(value).every(k=>keys.has(k))
}
function integer(value:unknown,low:number,high:number):value is number{return Number.isSafeInteger(value)&&!Object.is(value,-0)&&(value as number)>=low&&(value as number)<=high}
function rgb(value:unknown,hash=true):value is string{return typeof value==='string'&&(hash?/^#[0-9A-F]{6}$/:/^[0-9A-F]{6}$/).exec(value)?.[0]===value}
function axis(value:unknown):value is NativeLiteralBarAxis {
 if(!record(value,'id crossAxisId orientation position deleted min max crossesAt','color widthEmu labels')||!integer(value.id,0,4294967295)||!integer(value.crossAxisId,0,4294967295)||!['minMax','maxMin'].includes(value.orientation as string)||!['b','l'].includes(value.position as string)||typeof value.deleted!=='boolean')return false
 if(value.deleted?(value.color!==undefined||value.widthEmu!==undefined||value.labels!==undefined):(!rgb(value.color)||!integer(value.widthEmu,1,20116800)))return false
 if(typeof value.min!=='string'||typeof value.max!=='string'||typeof value.crossesAt!=='string')return false
 const low=decimal(value.min),high=decimal(value.max),cross=decimal(value.crossesAt)
 if(!low||!high||!cross||cross.coefficient!==0n||low.coefficient>0n||high.coefficient<0n)return false
 const exponent=Math.min(low.exponent,high.exponent)
 if(low.coefficient*10n**BigInt(low.exponent-exponent)>=high.coefficient*10n**BigInt(high.exponent-exponent))return false
 if(value.labels!==undefined){
  const labels=value.labels
  if(!record(labels,'profile position majorTickMark style majorUnit numberFormat')||labels.profile!=='explicit-axis-labels-v1'||!['low','high'].includes(labels.position as string)||!['none','out'].includes(labels.majorTickMark as string)||typeof labels.majorUnit!=='string'||typeof labels.numberFormat!=='string'||/^0(?:\.0{1,6})?$/.exec(labels.numberFormat)?.[0]!==labels.numberFormat)return false
  const style=labels.style
  if(!record(style,'fontFamily fontSize color bold italic language')||typeof style.fontFamily!=='string'||style.fontFamily.length===0||!integer(style.fontSize,1,400000)||!rgb(style.color,false)||typeof style.bold!=='boolean'||typeof style.italic!=='boolean'||typeof style.language!=='string')return false
 }
 return true
}
export function validNativeLiteralBubble(value:unknown):value is NativeLiteralBubble {
 if(!record(value,'profile dataOrigin bubbleScale sizeRepresents series xAxis yAxis')||value.profile!=='literal-bubble-v1'||value.dataOrigin!=='literal'||!integer(value.bubbleScale,0,300)||!['area','w'].includes(value.sizeRepresents as string)||!Array.isArray(value.series)||value.series.length<1||value.series.length>16)return false
 const x=value.xAxis,y=value.yAxis
 if(!axis(x)||!axis(y)||x.position!=='b'||y.position!=='l'||x.id===y.id||x.crossAxisId!==y.id||y.crossAxisId!==x.id||!validNativeChartAxisLabels(x,y,true)||!validNativeChartAxisLabels(y,x,true))return false
 const seen=new Set<number>()
 for(let order=0;order<value.series.length;order++){
  const s=value.series[order]
  if(!record(s,'index order xValues values sizes colors','title')||!integer(s.index,0,4294967295)||seen.has(s.index)||!integer(s.order,0,15)||s.order!==order||(s.title!==undefined&&(typeof s.title!=='string'||s.title.length>1024))||!Array.isArray(s.values)||s.values.length<1||s.values.length>256||!Array.isArray(s.xValues)||!Array.isArray(s.sizes)||!Array.isArray(s.colors)||s.xValues.length!==s.values.length||s.sizes.length!==s.values.length||s.colors.length!==s.values.length)return false
  seen.add(s.index)
  for(let i=0;i<s.values.length;i++){
   for(const raw of [s.xValues[i],s.values[i]])if(typeof raw!=='string'||!decimal(raw))return false
   if(typeof s.sizes[i]!=='string'||!rgb(s.colors[i]))return false
   const size=decimal(s.sizes[i]);if(!size||size.coefficient<0n)return false
  }
 }
 return true
}

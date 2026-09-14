import type {NativeLiteralArea} from './chartAreaTypes.js'
import type {NativeLiteralBarAxis} from './types.js'
import {validNativeChartAxisLabels} from './chartAxisLabelsValidation.js'
import {nativeChartDecimal as decimal} from './chartDecimalValidation.js'

function record(value:unknown,required:string,optional=''):value is Record<string,unknown>{
 if(!value||typeof value!=='object'||Array.isArray(value))return false
 const keys=new Set([...required.split(' '),...optional.split(' ')].filter(Boolean))
 return required.split(' ').every(k=>Object.hasOwn(value,k))&&Object.keys(value).every(k=>keys.has(k))
}
function integer(value:unknown,low:number,high:number):value is number{return Number.isSafeInteger(value)&&!Object.is(value,-0)&&(value as number)>=low&&(value as number)<=high}
function rgb(value:unknown,prefix=true):value is string{return typeof value==='string'&&(prefix?/^#[0-9A-F]{6}$/:/^[0-9A-F]{6}$/).exec(value)?.[0]===value}
function axis(value:unknown,numeric:boolean):value is NativeLiteralBarAxis{
 if(!record(value,'id crossAxisId orientation position deleted','color widthEmu min max crossesAt labels'))return false
 if(!integer(value.id,0,4294967295)||!integer(value.crossAxisId,0,4294967295)||!['minMax','maxMin'].includes(value.orientation as string)||!['b','l'].includes(value.position as string)||typeof value.deleted!=='boolean')return false
 if(value.deleted?(value.color!==undefined||value.widthEmu!==undefined||value.labels!==undefined):(!rgb(value.color)||!integer(value.widthEmu,1,20116800)))return false
 if(numeric){
  if(typeof value.min!=='string'||typeof value.max!=='string'||typeof value.crossesAt!=='string')return false
  const low=decimal(value.min),high=decimal(value.max),cross=decimal(value.crossesAt)
  if(!low||!high||!cross||cross.coefficient!==0n||low.coefficient>0n||high.coefficient<0n)return false
  const exponent=Math.min(low.exponent,high.exponent)
  if(low.coefficient*10n**BigInt(low.exponent-exponent)>=high.coefficient*10n**BigInt(high.exponent-exponent))return false
 }else if(value.min!==undefined||value.max!==undefined||value.crossesAt!==undefined)return false
 if(value.labels!==undefined){
  const labels=value.labels
  if(!record(labels,'profile position majorTickMark style','majorUnit numberFormat')||labels.profile!=='explicit-axis-labels-v1'||!['low','high'].includes(labels.position as string)||!['none','out'].includes(labels.majorTickMark as string))return false
  const style=labels.style
  if(!record(style,'fontFamily fontSize color bold italic language')||typeof style.fontFamily!=='string'||style.fontFamily.length<1||!integer(style.fontSize,1,400000)||!rgb(style.color,false)||typeof style.bold!=='boolean'||typeof style.italic!=='boolean'||typeof style.language!=='string')return false
  if(numeric?(typeof labels.majorUnit!=='string'||typeof labels.numberFormat!=='string'||/^0(?:\.0{1,6})?$/.exec(labels.numberFormat)?.[0]!==labels.numberFormat):(labels.majorUnit!==undefined||labels.numberFormat!==undefined))return false
 }
 return true
}

/** Complete structural and semantic validation until generated schema attachment
 * is available. This never upgrades referenced/cache data to a literal profile. */
export function validNativeLiteralArea(value:unknown):value is NativeLiteralArea{
 if(!record(value,'profile dataOrigin grouping categories series xAxis yAxis','sourceBaseline')||value.profile!=='literal-area-v1'||value.dataOrigin!=='literal'||!['standard','stacked','percentStacked'].includes(value.grouping as string))return false
 if(!Array.isArray(value.categories)||value.categories.length<1||value.categories.length>256||!Array.isArray(value.series)||value.series.length<1||value.series.length>16)return false
 let units=0
 for(const category of value.categories){if(typeof category!=='string')return false;units+=category.length;if(units>32768)return false}
 const x=value.xAxis,y=value.yAxis
 if(value.sourceBaseline!==undefined&&(!record(value.sourceBaseline,'crossing value')||value.sourceBaseline.crossing!=='min'||!y||typeof y!=='object'||value.sourceBaseline.value!==(y as Record<string,unknown>).min))return false
 if(!axis(x,false)||!axis(y,true)||x.position!=='b'||y.position!=='l'||x.id===y.id||x.crossAxisId!==y.id||y.crossAxisId!==x.id||!validNativeChartAxisLabels(x,y,false)||!validNativeChartAxisLabels(y,x,true))return false
 const indices=new Set<number>(),orders=new Set<number>()
 for(let order=0;order<value.series.length;order++){
  const series=value.series[order]
  if(!record(series,'index order values color','title')||!integer(series.index,0,4294967295)||indices.has(series.index)||!integer(series.order,0,15)||series.order>=value.series.length||orders.has(series.order)||!rgb(series.color)||(series.title!==undefined&&(typeof series.title!=='string'||series.title.length>1024))||!Array.isArray(series.values)||series.values.length!==value.categories.length)return false
  indices.add(series.index);orders.add(series.order)
  for(const raw of series.values){if(typeof raw!=='string')return false;const number=decimal(raw);if(!number||value.grouping!=='standard'&&number.coefficient<0n&&value.sourceBaseline===undefined)return false}
 }
 return true
}

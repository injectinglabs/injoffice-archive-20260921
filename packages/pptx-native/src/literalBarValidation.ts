import {validNativeChartAxisLabels} from './chartAxisLabelsValidation'
import type {NativeLiteralBar} from './types'

import {nativeChartDecimal as decimal} from './chartDecimalValidation'
export function validNativeLiteralBar(bar:NativeLiteralBar):boolean{
 const ca=bar.categoryAxis,va=bar.valueAxis
 if(!validNativeChartAxisLabels(ca,va,false)||!validNativeChartAxisLabels(va,ca,true))return false
 if(bar.categories.reduce((sum,c)=>sum+c.length,0)>32768||ca.id===va.id||ca.crossAxisId!==va.id||va.crossAxisId!==ca.id||ca.min!==undefined||ca.max!==undefined||ca.crossesAt!==undefined||va.min===undefined||va.max===undefined||va.crossesAt===undefined)return false
 for(const axis of [ca,va])if(axis.deleted?(axis.color!==undefined||axis.widthEmu!==undefined):(axis.color===undefined||axis.widthEmu===undefined))return false
 if(bar.barDirection==='column'?(ca.position!=='b'||va.position!=='l'):(ca.position!=='l'||va.position!=='b'))return false
 const minimum=decimal(va.min),maximum=decimal(va.max),cross=decimal(va.crossesAt)
 if(!minimum||!maximum||!cross||cross.coefficient!==0n||minimum.coefficient>0n||maximum.coefficient<0n)return false
 const exponent=Math.min(minimum.exponent,maximum.exponent)
 if(minimum.coefficient*10n**BigInt(minimum.exponent-exponent)>=maximum.coefficient*10n**BigInt(maximum.exponent-exponent))return false
 const indices=new Set<number>(),orders=new Set<number>()
 for(const [,series] of bar.series.entries()){
  if(!series||!Number.isSafeInteger(series.index)||Object.is(series.index,-0)||series.index<0||series.index>4294967295||indices.has(series.index)||!Number.isSafeInteger(series.order)||Object.is(series.order,-0)||series.order<0||series.order>=bar.series.length||orders.has(series.order)||series.values.length!==bar.categories.length||series.colors.length!==bar.categories.length||series.values.some(v=>!decimal(v)))return false
  indices.add(series.index);orders.add(series.order)
 }
 return true
}

import type {NativeLiteralConnected} from './types'
import {nativeChartDecimal as decimal} from './chartDecimalValidation'
/** Semantic checks beyond generated structural/conditional schema validation. */
export function validNativeLiteralConnected(c:NativeLiteralConnected):boolean{
 const scatter=c.profile==='literal-scatter-v1',x=c.xAxis,y=c.yAxis
 if(c.categories.reduce((sum,v)=>sum+v.length,0)>32768||x.id===y.id||x.crossAxisId!==y.id||y.crossAxisId!==x.id||x.position!=='b'||y.position!=='l')return false
 if(scatter?c.categories.length!==0:(c.categories.length===0||x.min!==undefined||x.max!==undefined||x.crossesAt!==undefined))return false
 for(const [i,a]of [x,y].entries()){
  if(a.deleted?(a.color!==undefined||a.widthEmu!==undefined):(a.color===undefined||a.widthEmu===undefined))return false
  if(i===0&&!scatter)continue
  if(a.min===undefined||a.max===undefined||a.crossesAt===undefined)return false
  const low=decimal(a.min),high=decimal(a.max),cross=decimal(a.crossesAt)
  if(!low||!high||!cross||cross.coefficient!==0n||low.coefficient>0n||high.coefficient<0n)return false
  const exponent=Math.min(low.exponent,high.exponent)
  if(low.coefficient*10n**BigInt(low.exponent-exponent)>=high.coefficient*10n**BigInt(high.exponent-exponent))return false
 }
 const seen=new Set<number>()
 for(const [order,s]of c.series.entries()){
  if(seen.has(s.index)||s.order!==order||s.values.some(v=>!decimal(v)))return false
  seen.add(s.index)
  if(scatter?(!s.xValues||s.xValues.length!==s.values.length||s.xValues.some(v=>!decimal(v))):(s.xValues!==undefined||s.values.length!==c.categories.length))return false
 }
 return true
}

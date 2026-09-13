import type {NativeLiteralBarAxis} from './types'
import {nativeChartDecimal as decimal} from './chartDecimalValidation'

export function validNativeChartAxisLabels(axis:NativeLiteralBarAxis,perpendicular:NativeLiteralBarAxis,value:boolean):boolean {
 const labels=axis.labels
 if(!labels)return true
 if(axis.deleted)return false
 const s=labels.style
 if(new TextEncoder().encode(s.fontFamily).length>128||s.fontFamily.startsWith('+')||/[\p{Cc}]/u.test(s.fontFamily)||!/^\p{ASCII}+$/u.test(s.language)||/^[A-Za-z]{2,8}(?:-[A-Za-z0-9]{1,8})*$/.exec(s.language)?.[0]!==s.language)return false
 if(value){
  if(labels.majorUnit===undefined||labels.numberFormat===undefined||axis.min===undefined||axis.max===undefined)return false
  const low=decimal(axis.min),high=decimal(axis.max),step=decimal(labels.majorUnit)
  if(!low||!high||!step||step.coefficient<=0n)return false
  const exp=Math.min(low.exponent,high.exponent,step.exponent)
  const a=low.coefficient*10n**BigInt(low.exponent-exp),b=high.coefficient*10n**BigInt(high.exponent-exp),u=step.coefficient*10n**BigInt(step.exponent-exp)
  if(a>=b||a%u!==0n||(b-a)/u+1n>256n)return false
 }else if(labels.majorUnit!==undefined||labels.numberFormat!==undefined)return false
 if(labels.majorTickMark==='out'&&perpendicular.min!==undefined){
  if(perpendicular.max===undefined)return false
  const low=decimal(perpendicular.min),high=decimal(perpendicular.max)
  if(!low||!high||low.coefficient!==0n&&high.coefficient!==0n)return false
 }
 return true
}

import type {NativeLiteralBar} from './types'

// Mirrors the bounded source parser. Decimal values are strings in the native
// integer-only JSON contract; compare exact coefficients instead of Number.
function decimal(raw:string):{coefficient:bigint;exponent:number}|undefined{
 const match=/^([+-]?)(?:([0-9]+)(?:\.([0-9]*))?|\.([0-9]+))(?:[eE]([+-]?[0-9]+))?$/.exec(raw)
 if(!match||match[0]!==raw||raw.length>128)return
 const fraction=match[4]??match[3]??'',digits=(match[2]??'')+fraction,exponent=Number(match[5]??0)
 if(digits.length>32||!Number.isInteger(exponent)||exponent < -100||exponent>100)return
 return {coefficient:BigInt(digits)*(match[1]==='-'?-1n:1n),exponent:exponent-fraction.length}
}
export function validNativeLiteralBar(bar:NativeLiteralBar):boolean{
 const ca=bar.categoryAxis,va=bar.valueAxis
 if(bar.categories.reduce((sum,c)=>sum+c.length,0)>32768||ca.id===va.id||ca.crossAxisId!==va.id||va.crossAxisId!==ca.id||ca.min!==undefined||ca.max!==undefined||ca.crossesAt!==undefined||va.min===undefined||va.max===undefined||va.crossesAt===undefined)return false
 for(const axis of [ca,va])if(axis.deleted?(axis.color!==undefined||axis.widthEmu!==undefined):(axis.color===undefined||axis.widthEmu===undefined))return false
 if(bar.barDirection==='column'?(ca.position!=='b'||va.position!=='l'):(ca.position!=='l'||va.position!=='b'))return false
 const minimum=decimal(va.min),maximum=decimal(va.max),cross=decimal(va.crossesAt)
 if(!minimum||!maximum||!cross||cross.coefficient!==0n||minimum.coefficient>0n||maximum.coefficient<0n)return false
 const exponent=Math.min(minimum.exponent,maximum.exponent)
 if(minimum.coefficient*10n**BigInt(minimum.exponent-exponent)>=maximum.coefficient*10n**BigInt(maximum.exponent-exponent))return false
 const indices=new Set<number>()
 for(const [order,series] of bar.series.entries()){
  if(indices.has(series.index)||series.order!==order||series.values.length!==bar.categories.length||series.colors.length!==bar.categories.length||series.values.some(v=>!decimal(v)))return false
  indices.add(series.index)
 }
 return true
}

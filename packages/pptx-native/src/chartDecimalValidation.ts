// Mirrors the bounded source parser. Decimal values are strings in the native
// integer-only JSON contract; compare exact coefficients instead of Number.
export function nativeChartDecimal(raw:string):{coefficient:bigint;exponent:number}|undefined{
 const match=/^([+-]?)(?:([0-9]+)(?:\.([0-9]*))?|\.([0-9]+))(?:[eE]([+-]?[0-9]+))?$/.exec(raw)
 if(!match||match[0]!==raw||raw.length>128)return
 const fraction=match[4]??match[3]??'',digits=(match[2]??'')+fraction,exponent=Number(match[5]??0)
 if(digits.length>32||!Number.isInteger(exponent)||exponent < -100||exponent>100)return
 return {coefficient:BigInt(digits)*(match[1]==='-'?-1n:1n),exponent:exponent-fraction.length}
}

/** Private exact source-decimal arithmetic. Public native JSON remains integer-only. */
export interface ChartDecimal {readonly lexeme:string;readonly coefficient:bigint;readonly exponent:number}
const decimal=/^([+-]?)(?:([0-9]+)(?:\.([0-9]*))?|\.([0-9]+))(?:[eE]([+-]?[0-9]+))?$/
export function parseChartDecimal(lexeme:string):ChartDecimal {
 if(typeof lexeme!=='string'||lexeme.length<1||lexeme.length>128)throw new RangeError('chart numeric lexeme budget exceeded')
 const match=decimal.exec(lexeme)
 if(!match||match[0]!==lexeme)throw new RangeError('invalid chart decimal lexeme')
 const fraction=match[4]??match[3]??'',digits=(match[2]??'')+fraction
 if(digits.length>32)throw new RangeError('chart decimal digit budget exceeded')
 const exponent=match[5]===undefined?0:Number(match[5])
 if(!Number.isInteger(exponent)||exponent < -100||exponent>100)throw new RangeError('chart decimal exponent budget exceeded')
 const coefficient=BigInt(digits)*(match[1]==='-'?-1n:1n)
 return {lexeme,coefficient,exponent:exponent-fraction.length}
}
function scaled(value:ChartDecimal,exponent:number):bigint {return value.coefficient*10n**BigInt(value.exponent-exponent)}
export function compareChartDecimals(a:ChartDecimal,b:ChartDecimal):number {
 const exponent=Math.min(a.exponent,b.exponent),x=scaled(a,exponent),y=scaled(b,exponent)
 return x<y?-1:x>y?1:0
}
/** Map an in-range decimal to local EMU without binary-float cancellation. */
export function chartDecimalCoordinate(value:ChartDecimal,min:ChartDecimal,max:ChartDecimal,extent:number,reverse=false):number {
 if(!Number.isSafeInteger(extent)||extent<1||extent>281474976710655)throw new RangeError('invalid chart coordinate extent')
 const exponent=Math.min(value.exponent,min.exponent,max.exponent)
 const low=scaled(min,exponent),high=scaled(max,exponent),v=scaled(value,exponent),span=high-low
 if(span<=0n||v<low||v>high)throw new RangeError('chart value outside explicit scale')
 const numerator=(reverse?high-v:v-low)*BigInt(extent)
 return Number((2n*numerator+span)/(2n*span))
}

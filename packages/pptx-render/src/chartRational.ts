import {parseChartDecimal} from './chartDecimal.js'

/** Private exact arithmetic for clipping; no rational objects cross native JSON. */
export interface ChartRational {readonly numerator:bigint;readonly denominator:bigint}
const maxBits=4096
function magnitude(value:bigint):bigint{return value<0n?-value:value}
function bounded(value:bigint):void{if(magnitude(value).toString(2).length>maxBits)throw new RangeError('chart rational bit budget exceeded')}
function gcd(a:bigint,b:bigint):bigint{a=magnitude(a);while(b!==0n){const next=a%b;a=b;b=next}return a}
export function chartRational(numerator:bigint,denominator=1n):ChartRational{
 if(denominator===0n)throw new RangeError('zero chart rational denominator')
 bounded(numerator);bounded(denominator)
 if(denominator<0n){numerator=-numerator;denominator=-denominator}
 const divisor=gcd(numerator,denominator)
 return {numerator:numerator/divisor,denominator:denominator/divisor}
}
export function chartRationalDecimal(lexeme:string):ChartRational{
 const value=parseChartDecimal(lexeme)
 return value.exponent<0?chartRational(value.coefficient,10n**BigInt(-value.exponent)):chartRational(value.coefficient*10n**BigInt(value.exponent))
}
export function chartRationalAdd(a:ChartRational,b:ChartRational):ChartRational{return chartRational(a.numerator*b.denominator+b.numerator*a.denominator,a.denominator*b.denominator)}
export function chartRationalSubtract(a:ChartRational,b:ChartRational):ChartRational{return chartRational(a.numerator*b.denominator-b.numerator*a.denominator,a.denominator*b.denominator)}
export function chartRationalMultiply(a:ChartRational,b:ChartRational):ChartRational{return chartRational(a.numerator*b.numerator,a.denominator*b.denominator)}
export function chartRationalDivide(a:ChartRational,b:ChartRational):ChartRational{return chartRational(a.numerator*b.denominator,a.denominator*b.numerator)}
export function chartRationalCompare(a:ChartRational,b:ChartRational):number{
 const left=a.numerator*b.denominator,right=b.numerator*a.denominator
 return left<right?-1:left>right?1:0
}
export function chartRationalCoordinate(value:ChartRational,extent:number,reverse=false):number{
 if(!Number.isSafeInteger(extent)||extent<1||extent>281474976710655||value.denominator<=0n||value.numerator<0n||value.numerator>value.denominator)throw new RangeError('invalid unit chart coordinate')
 const numerator=(reverse?value.denominator-value.numerator:value.numerator)*BigInt(extent)
 return Number((2n*numerator+value.denominator)/(2n*value.denominator))
}

import type {ChartRational} from './chartRational.js'

/** Exact fixed-decimal display; ties round away from zero. No adaptive General
 * formatting, source-linked format lookup, locale or arbitrary Excel syntax. */
export function formatChartFixedDecimal(value:ChartRational,format:string):string{
 const match=/^0(?:\.0{1,6})?$/.exec(format)
 if(!match||match[0]!==format)throw new RangeError('unsupported chart number format')
 const places=format==='0'?0:format.length-2,scale=10n**BigInt(places),negative=value.numerator<0n,absolute=negative?-value.numerator:value.numerator
 if(value.denominator<=0n)throw new RangeError('invalid chart number denominator')
 const rounded=(2n*absolute*scale+value.denominator)/(2n*value.denominator)
 let text=rounded.toString().padStart(places+1,'0')
 if(places)text=text.slice(0,-places)+'.'+text.slice(-places)
 if(negative&&rounded!==0n)text='-'+text
 if(text.length>256)throw new RangeError('formatted chart label budget exceeded')
 return text
}

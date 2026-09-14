import {chartRational,type ChartRational} from './chartRational.js'

/** Exact floor sqrt of an integer. The caller's existing rational bit budget
 * bounds every intermediate; this is an integer operation, not a second numeric
 * representation or a floating-point approximation. */
function floorRoot(value:bigint):bigint {
 if(value<2n)return value
 let high=1n<<BigInt(Math.ceil(value.toString(2).length/2))
 for(;;){const next=(high+value/high)/2n;if(next>=high)return high;high=next}
}
/** Round sqrt(value) to an integer EMU, ties upward. The exact midpoint test
 * proves <=1/2 EMU error even for extreme decimal ratios and exact half ties.
 * No source numeric value is converted to binary floating point. */
export function chartRoundedSquareRoot(value:ChartRational):number {
 const bounded=chartRational(value.numerator,value.denominator)
 if(bounded.numerator<0n)throw new RangeError('negative chart squared radius')
 const low=floorRoot(bounded.numerator/bounded.denominator)
 const midpoint=2n*low+1n
 const rounded=4n*bounded.numerator>=midpoint*midpoint*bounded.denominator?low+1n:low
 if(rounded>BigInt(Number.MAX_SAFE_INTEGER))throw new RangeError('chart radius exceeds integer bounds')
 return Number(rounded)
}

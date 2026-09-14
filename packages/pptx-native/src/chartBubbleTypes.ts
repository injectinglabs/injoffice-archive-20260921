import type {NativeLiteralBarAxis} from './types.js'
export interface NativeLiteralBubbleSeries {
 readonly index:number
 readonly order:number
 readonly title?:string
 readonly xValues:readonly string[]
 readonly values:readonly string[]
 readonly sizes:readonly string[]
 readonly colors:readonly string[]
}
/** Dedicated literal preparation record. Referenced values must use their own
 * admitted workbook profile, never a rebranded literal record. */
export interface NativeLiteralBubble {
 readonly profile:'literal-bubble-v1'
 readonly dataOrigin:'literal'
 readonly bubbleScale:number
 readonly sizeRepresents:'area'|'w'
 readonly series:readonly NativeLiteralBubbleSeries[]
 readonly xAxis:NativeLiteralBarAxis
 readonly yAxis:NativeLiteralBarAxis
}

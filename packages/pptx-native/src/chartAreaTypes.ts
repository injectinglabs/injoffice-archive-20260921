import type {NativeLiteralBarAxis} from './types.js'

export interface NativeLiteralAreaSeries {
 readonly index:number
 readonly order:number
 readonly title?:string
 readonly values:readonly string[]
 readonly color:string
}
/** Read-only source literal area profile. Values retain their source decimal
 * spellings; workbook caches are not admitted as literal values. */
export interface NativeLiteralArea {
 /** Original category crossing/minimum; absent legacy records retain zero closure. */
 readonly sourceBaseline?:{readonly crossing:'min';readonly value:string}
 readonly profile:'literal-area-v1'
 readonly dataOrigin:'literal'
 readonly grouping:'standard'|'stacked'|'percentStacked'
 readonly categories:readonly string[]
 readonly series:readonly NativeLiteralAreaSeries[]
 readonly xAxis:NativeLiteralBarAxis
 readonly yAxis:NativeLiteralBarAxis
}

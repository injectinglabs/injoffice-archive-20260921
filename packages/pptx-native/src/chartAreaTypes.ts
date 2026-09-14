import type {NativeLiteralBarAxis} from './types.js'

export interface NativeLiteralAreaSeries {
 readonly index:number
 readonly order:number
 readonly title?:string
 readonly values:readonly string[]
 readonly color:string
}
/** Dedicated preparation record; public chart attachment requires the shared
 * schema/runtime handoff. No workbook cache is a literal source value. */
export interface NativeLiteralArea {
 readonly profile:'literal-area-v1'
 readonly dataOrigin:'literal'
 readonly grouping:'standard'|'stacked'|'percentStacked'
 readonly categories:readonly string[]
 readonly series:readonly NativeLiteralAreaSeries[]
 readonly xAxis:NativeLiteralBarAxis
 readonly yAxis:NativeLiteralBarAxis
}

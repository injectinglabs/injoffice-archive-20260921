import type {NativeLiteralBarAxis} from './types.js'

export interface NativeLiteralRadarSeries {
 readonly index:number
 /** Original source order metadata; array sequence is preserved separately. */
 readonly order:number
 readonly title?:string
 readonly values:readonly string[]
 readonly color:string
 readonly widthEmu:number
 readonly fill?:string
}
/** Literal-only source record. Values retain source spellings; workbook radar
 * has no public attachment or resolved-data authority in this slice. */
export interface NativeLiteralRadar {
 readonly profile:'literal-radar-v1'
 readonly dataOrigin:'literal'
 readonly style:'standard'|'filled'
 readonly categories:readonly string[]
 readonly series:readonly NativeLiteralRadarSeries[]
 readonly categoryAxis:NativeLiteralBarAxis
 readonly valueAxis:NativeLiteralBarAxis
}

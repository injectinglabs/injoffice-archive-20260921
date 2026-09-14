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
/** Dedicated source record, not exported or attached until radial semantics and
 * connected authority are qualified. Literal values retain source spellings. */
export interface NativeLiteralRadar {
 readonly profile:'literal-radar-v1'
 readonly dataOrigin:'literal'
 readonly style:'standard'|'filled'
 readonly categories:readonly string[]
 readonly series:readonly NativeLiteralRadarSeries[]
 readonly categoryAxis:NativeLiteralBarAxis
 readonly valueAxis:NativeLiteralBarAxis
}

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
/** Profile-free radial data. Its caller owns literal or workbook authority. */
export interface NativeRadarData {
 readonly style:'standard'|'filled'
 readonly categories:readonly string[]
 readonly series:readonly NativeLiteralRadarSeries[]
 readonly categoryAxis:NativeLiteralBarAxis
 readonly valueAxis:NativeLiteralBarAxis
}

export interface NativeLiteralRadar extends NativeRadarData { readonly profile:'literal-radar-v1'; readonly dataOrigin:'literal' }

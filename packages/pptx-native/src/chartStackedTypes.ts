import type {NativeLiteralBar,NativeLiteralConnected} from './types.js'

/** Literal source only; percentages preserve signs over absolute category totals. */
export interface NativeLiteralStackedBar extends Omit<NativeLiteralBar,'profile'|'grouping'|'overlap'> {
 readonly profile:'literal-stacked-bar-v1'
 readonly grouping:'stacked'|'percentStacked'
 readonly overlap:100
}
/** Straight algebraic cumulative lines, with explicit source axes and paint. */
export interface NativeLiteralStackedLine extends Omit<NativeLiteralConnected,'profile'> {
 readonly profile:'literal-stacked-line-v1'
 readonly grouping:'stacked'|'percentStacked'
}

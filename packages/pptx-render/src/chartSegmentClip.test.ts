import {expect,it} from 'vitest'
import {chartRational as r,chartRationalDecimal as d} from './chartRational.js'
import {clipChartUnitSegment as clip,type ChartRationalPoint} from './chartSegmentClip.js'
const p=(x:string,y:string):ChartRationalPoint=>({x:d(x),y:d(y)})
it('finds actual intersections instead of clamping each endpoint',()=>{
 const segment=clip(p('-1','0'),p('1','1'))!
 expect(segment.start).toEqual({x:r(0n),y:r(1n,2n)});expect(segment.end).toEqual(p('1','1'));expect(segment.startParameter).toEqual(r(1n,2n))
 const reversed=clip(p('1','1'),p('-1','0'))!
 expect(reversed.start).toEqual(segment.end);expect(reversed.end).toEqual(segment.start)
})
it('clips every slab and preserves corner tangency',()=>{
 expect(clip(p('-.5','.5'),p('1.5','.5'))!.start).toEqual(p('0','.5'))
 expect(clip(p('-.5','.5'),p('1.5','.5'))!.end).toEqual(p('1','.5'))
 expect(clip(p('.5','-1'),p('.5','2'))!.start).toEqual(p('.5','0'))
 expect(clip(p('.5','-1'),p('.5','2'))!.end).toEqual(p('.5','1'))
 const corner=clip(p('-1','1'),p('1','-1'))!;expect(corner.start).toEqual(p('0','0'));expect(corner.end).toEqual(p('0','0'))
})
it('rejects outside parallel segments without bridging gaps and handles zero length',()=>{
 expect(clip(p('-1','-1'),p('2','-1'))).toBeUndefined();expect(clip(p('2','0'),p('2','1'))).toBeUndefined()
 expect(clip(p('.25','.75'),p('.25','.75'))!.start).toEqual(p('.25','.75'));expect(clip(p('2','2'),p('2','2'))).toBeUndefined()
})
it('retains exact edge intersections at maximum decimal dynamic range',()=>{
 const segment=clip(p('-1e100','0'),p('1e100','1'))!
 expect(segment.start).toEqual({x:r(0n),y:r(1n,2n)})
 expect(segment.end.x).toEqual(r(1n));expect(segment.end.y).toEqual(r(10n**100n+1n,2n*10n**100n))
})

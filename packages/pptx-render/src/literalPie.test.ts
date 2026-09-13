import {expect,it} from 'vitest'
import {createNativeLiteralPiePaths} from './literalPie.js'
const pie={profile:'literal-pie-v1' as const,firstSliceAngle:0,values:[1,3],colors:['#FF0000','#00FF00']}
it('uses literal proportions, point order, centered circular fitting and bounded integer paths',()=>{
 const result=createNativeLiteralPiePaths(pie,200,100)
 expect(result[0]!.path[1]).toEqual({kind:'lineTo',x:100,y:0})
 expect(result[0]!.path.at(-2)).toEqual({kind:'lineTo',x:150,y:50})
 expect(result[1]!.path[1]).toEqual(result[0]!.path.at(-2))
 expect(result[1]!.path.at(-2)).toEqual(result[0]!.path[1])
 expect(result.map(s=>s.color)).toEqual(pie.colors)
 for(const slice of result){expect(slice.path.length).toBeLessThanOrEqual(183);for(const p of slice.path)if(p.kind==='lineTo')expect(Number.isSafeInteger(p.x)&&Number.isSafeInteger(p.y)).toBe(true)}
 expect(createNativeLiteralPiePaths({...pie,firstSliceAngle:90},200,100)[0]!.path[1]).toEqual({kind:'lineTo',x:150,y:50})
})
it('refuses invalid geometry and values at the public boundary',()=>{
 for(const bad of [{...pie,values:[0,3]},{...pie,values:[Infinity,3]},{...pie,values:[1.2,3]},{...pie,colors:['#FF0000']},{...pie,firstSliceAngle:361},{...pie,values:Array(65).fill(1),colors:Array(65).fill('#FF0000')}])expect(()=>createNativeLiteralPiePaths(bad,100,100)).toThrow()
 expect(()=>createNativeLiteralPiePaths(pie,Infinity,100)).toThrow()
 expect(()=>createNativeLiteralPiePaths(pie,100,0)).toThrow()
 expect(createNativeLiteralPiePaths({...pie,values:[1],colors:['#FF0000']},100,100)[0]!.path).toHaveLength(183)
})

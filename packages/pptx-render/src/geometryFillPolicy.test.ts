import {expect,it} from 'vitest'
import {geometryPathFill} from './geometryFillPolicy.js'
it('uses explicit linear-light tone strengths, preserving original color and no-fill',()=>{
 expect(geometryPathFill('123456','norm')).toBe('123456')
 expect(geometryPathFill('123456','none')).toBeUndefined()
 expect(geometryPathFill(undefined,'darken')).toBeUndefined()
 // IEC sRGB encoding of linear luminances .6, .8, .4, and .2 respectively.
 expect(geometryPathFill('FFFFFF','darken')).toBe('CBCBCB')
 expect(geometryPathFill('FFFFFF','darkenLess')).toBe('E7E7E7')
 expect(geometryPathFill('000000','lighten')).toBe('AAAAAA')
 expect(geometryPathFill('000000','lightenLess')).toBe('7C7C7C')
 expect(geometryPathFill('000000','darken')).toBe('000000')
 expect(geometryPathFill('FFFFFF','lighten')).toBe('FFFFFF')
})
it('bounds saturated channels and refuses non-native colors',()=>{
 expect(geometryPathFill('FF0000','lighten')).toBe('FFAAAA')
 expect(geometryPathFill('00FF00','darken')).toBe('00CB00')
 for(const color of ['#FFFFFF','red','GG0000','abcdef'])expect(()=>geometryPathFill(color,'darken')).toThrow()
})

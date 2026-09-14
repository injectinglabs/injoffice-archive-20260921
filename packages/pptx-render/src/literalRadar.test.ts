import {expect,it} from 'vitest'
import {createNativeLiteralRadarPaths} from './literalRadar.js'
it('refuses a reference record instead of promoting it to literal authority',()=>{
 expect(()=>createNativeLiteralRadarPaths({profile:'workbook-radar-v1',dataOrigin:'workbook'} as any,100,100)).toThrow('authority')
})

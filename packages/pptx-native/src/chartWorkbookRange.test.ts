import {expect,it} from 'vitest'
import {parseChartWorkbookRange as parse,chartWorkbookCellAddress as address} from './chartWorkbookRange.js'
it('preserves quoted worksheet identity while resolving exact one-dimensional coordinates',()=>{
 expect(parse("'O''Brien ! Q'!$a1:$c1")).toEqual({sheet:"O'Brien ! Q",startRow:0,startColumn:0,endRow:0,endColumn:2,count:3})
 expect(parse('Sheet1!$B$2:$B$4')).toEqual({sheet:'Sheet1',startRow:1,startColumn:1,endRow:3,endColumn:1,count:3})
 expect(parse("'Δεδομένα'!XFD1048576")).toMatchObject({startRow:1048575,startColumn:16383,count:1})
 expect(parse('1!A1:A256').count).toBe(256)
 expect(address(1048575,16383)).toBe('XFD1048576')
})
it('refuses expressions, external/named/matrix references and unbounded coordinates',()=>{
 for(const formula of ['A1:A2','Sheet!A0','Sheet!A01','Sheet!XFE1','Sheet!A1048577','Sheet!A1:A257','Sheet!A2:A1','Sheet!B1:A1','Sheet!A1:B2','[1]Sheet!A1',"'[Book.xlsx]Sheet'!A1",'SUM(Sheet!A1:A2)','Sheet!Name','Sheet!A:A','Sheet!1:2','Sheet1:Sheet2!A1','Sheet!A1,Sheet!B1','Sheet!A1\n',"'Bad/Name'!A1","'\n'!A1","'''Bad'!A1",'Sheet!$A$1#','Sheet!R1C1','=Sheet!A1',"'\ud800'!A1"])expect(()=>parse(formula)).toThrow()
})

import {expect,it} from 'vitest'
import {parseChartDecimal as parse,compareChartDecimals as compare,chartDecimalCoordinate as coordinate} from './chartDecimal.js'
it('preserves bounded source spelling and rejects unsupported numeric grammar',()=>{
 for(const text of ['0','-0','+0','1.25','.5','1.','-12.34e-5','0001.000','1E+100','1e-100','12345678901234567890123456789012'])expect(parse(text).lexeme).toBe(text)
 for(const text of ['', ' 1','1 ','1\n','NaN','INF','Infinity','0x10','1_000','1,5','.','1e','1e101','1e-101','1'.repeat(33),'1e99999999999999','0'.repeat(129)])expect(()=>parse(text)).toThrow()
})
it('compares decimal values exactly without altering their source representation',()=>{
 for(const [a,b,order] of [['-0','0',0],['1.25e2','125',0],['.001','1e-3',0],['1000000000000.000000000000000001','1000000000000.000000000000000002',-1],['-1e100','-1e99',-1],['1e-100','0',1]] as const){const left=parse(a),before=left.coefficient;expect(compare(left,parse(b))).toBe(order);expect(left.coefficient).toBe(before)}
})
it('maps signed values and precision-adversarial scales before final EMU rounding',()=>{
 expect(coordinate(parse('0'),parse('-2.5'),parse('7.5'),1000)).toBe(250)
 expect(coordinate(parse('0'),parse('-2.5'),parse('7.5'),1000,true)).toBe(750)
 const min=parse('1000000000000.000000000000000001'),max=parse('1000000000000.000000000000000003'),mid=parse('1000000000000.000000000000000002')
 expect(coordinate(mid,min,max,1001)).toBe(501)
 expect(coordinate(mid,min,max,1001,true)).toBe(501)
 expect(coordinate(min,min,max,1001)).toBe(0)
 expect(coordinate(max,min,max,1001)).toBe(1001)
 expect(coordinate(parse('1e-100'),parse('0'),parse('2e-100'),1000)).toBe(500)
 for(const [v,lo,hi,size] of [['0','1','2',10],['1','2','1',10],['1','1','1',10],['1','0','2',0],['1','0','2',Infinity]] as const)expect(()=>coordinate(parse(v),parse(lo),parse(hi),size)).toThrow()
})

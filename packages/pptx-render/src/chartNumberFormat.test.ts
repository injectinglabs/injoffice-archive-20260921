import {expect,it} from 'vitest'
import {chartRationalDecimal as decimal} from './chartRational.js'
import {formatChartFixedDecimal as format} from './chartNumberFormat.js'
it('formats exact decimals, padding, half ties and negative zero',()=>{
 for(const [value,code,want]of [['-1.25','0.0','-1.3'],['1.25','0.0','1.3'],['-.004','0.00','0.00'],['1e-6','0.000000','0.000001'],['99999999999999999999999999999999','0','99999999999999999999999999999999'],['1.999','0.00','2.00']])expect(format(decimal(value!),code!)).toBe(want)
})
it('refuses unqualified formatting instead of falling back to General',()=>{
 for(const code of ['General','0.','#,##0','0%','0.0000000','0;0','0.0\n'])expect(()=>format(decimal('1'),code)).toThrow()
})

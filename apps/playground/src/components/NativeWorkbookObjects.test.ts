import {createElement} from 'react'
import {renderToStaticMarkup} from 'react-dom/server'
import {expect,it} from 'vitest'
import {NativeCachedChartPreview} from './NativeWorkbookObjects'
import type {NativeChartPreviewV1} from '@injoffice/sheets/browser'

const chart=():NativeChartPreviewV1=>({part:'xl/charts/chart1.xml',type:'col',series:[{name:'',values:[9,11],labels:[]},{name:'Saved revenue',values:[10,12],labels:['January','February']}],warnings:['Saved caches may be stale. No recalculation.']})
it('identifies cached series and point positions without inventing authored names',()=>{
 const html=renderToStaticMarkup(createElement(NativeCachedChartPreview,{chart:chart(),index:0}))
 expect(html).toContain('Preview series legend')
 expect(html).toContain('Unnamed series 1')
 expect(html).toContain('Point 1')
 expect(html).toContain('Saved revenue · January: 10')
 expect(html).toContain('<th>Saved category</th>')
 expect(html).toContain('Not available')
 expect(html).toContain('Saved caches may be stale')
})
it('escapes cached text and bounds point labels and table rows',()=>{
 const c=chart();c.series=[{name:'<script>bad</script>',values:Array(13).fill(1),labels:Array(13).fill('<b>category</b>')}]
 const html=renderToStaticMarkup(createElement(NativeCachedChartPreview,{chart:c,index:0}))
 expect(html).not.toContain('<script>')
 expect(html).toContain('&lt;b&gt;category&lt;/b&gt;')
 expect(html).not.toContain('>Point 1</text>')
 c.series[0]!.values=Array(101).fill(1);c.series[0]!.labels=Array(101).fill('category')
 const bounded=renderToStaticMarkup(createElement(NativeCachedChartPreview,{chart:c,index:0}))
 expect(bounded.match(/<tr>/g)).toHaveLength(101)
})

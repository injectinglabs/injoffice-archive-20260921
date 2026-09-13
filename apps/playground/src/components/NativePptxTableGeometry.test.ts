import {createElement} from 'react'
import {renderToStaticMarkup} from 'react-dom/server'
import {expect,it} from 'vitest'
import {NativePptxTableGeometry} from './NativePptxTableGeometry'
import type {NativePptxTableGeometryPreview} from '@injoffice/pptx-native'

it('shows plain source text at planned positions with declared clipping and full-text access',()=>{
 const preview:NativePptxTableGeometryPreview={policy:'host-sans-12pt-clipped-v1',packageSHA256:'a'.repeat(64),sourceRevision:'rev-source',cssPixelsPerInch:96,fontSize:16,lineHeight:20,inset:2,omissions:[],slides:[{slideId:'s1',slideIndex:0,sourceBounds:{x:100,y:200,width:952500,height:476250},width:100,height:50,tables:[{objectId:'cNvPr-1',tableIndex:10,rect:{x:20,y:10,width:80,height:40},cells:[{row:0,column:0,rect:{x:0,y:0,width:80,height:40},paragraphs:['<script>bad()</script>','No Border']}]}]}]}
 const html=renderToStaticMarkup(createElement(NativePptxTableGeometry,{preview,readingId:'read'}))
 expect(html).toContain('Approximate table arrangement')
 expect(html).toContain('not a full slide')
 expect(html).toContain('Text may be clipped')
 expect(html).toContain('data-geometry-cell="11:1:1"')
 expect(html).toContain('left:28px;top:18px;width:80px;height:40px')
 expect(html).toContain('font-size:16px;line-height:20px;padding:2px')
 expect(html).toContain('&lt;script&gt;bad()&lt;/script&gt;')
 expect(html).not.toContain('<script>')
 expect(html).toContain('href="#read-10"')
 expect(html).toContain('Full text: table 11')
 expect(html).not.toContain('contenteditable')
})

it('applies only qualified centered solid borders and keeps unsupported paint visibly omitted',()=>{
 const paint={policy:'source-no-style-solid-border-v1' as const,style_id:'{01234567-89AB-CDEF-0123-456789ABCDEF}',sources:[],fill:'none' as const,border:{color:'112233',width_emu:12700}}
 const preview:NativePptxTableGeometryPreview={policy:'host-sans-12pt-clipped-v1',paintPolicy:'source-no-style-solid-border-v1',paintOmissions:[{slideId:'s1',tableIndex:1,reason:'Source dash omitted'}],packageSHA256:'a'.repeat(64),sourceRevision:'rev-source',cssPixelsPerInch:96,fontSize:16,lineHeight:20,inset:2,omissions:[],slides:[{slideId:'s1',slideIndex:0,sourceBounds:{x:0,y:0,width:952500,height:476250},width:100,height:50,tables:[{objectId:'cNvPr-1',tableIndex:0,rect:{x:0,y:0,width:100,height:50},paint,cells:[{row:0,column:0,rect:{x:0,y:0,width:100,height:50},paragraphs:['Solid']}]}]}]}
 const html=renderToStaticMarkup(createElement(NativePptxTableGeometry,{preview,readingId:'read'}))
 expect(html).toContain('data-source-paint="solid-border"')
 expect(html).toContain('stroke="#112233"')
 expect(html).toContain('stroke-width="1.3333333333333333"')
 expect(html).toContain('stroke-linejoin="round"')
 expect(html).toContain('pptx-table-qualified-paint')
 expect(html).toContain('Source dash omitted')
 expect(html).toContain('1 tables have qualified source paint')
 expect(html).toContain('not PowerPoint rendering')
})

it('replays a preset with flat caps and explicit phase/corner limitations',()=>{
 const preview:NativePptxTableGeometryPreview={policy:'host-sans-12pt-clipped-v1',paintPolicy:'source-no-style-preset-border-v1',paintOmissions:[],packageSHA256:'a'.repeat(64),sourceRevision:'rev-source',cssPixelsPerInch:96,fontSize:16,lineHeight:20,inset:2,omissions:[],slides:[{slideId:'s1',slideIndex:0,sourceBounds:{x:0,y:0,width:952500,height:476250},width:100,height:50,tables:[{objectId:'cNvPr-1',tableIndex:0,rect:{x:0,y:0,width:100,height:50},paint:{policy:'source-no-style-preset-border-v1',style_id:'{01234567-89AB-CDEF-0123-456789ABCDEF}',sources:[],fill:'none',border:{color:'112233',width_emu:19050,preset:'dashDot'}},dashArray:[8,6,2,6],cells:[{row:0,column:0,rect:{x:0,y:0,width:100,height:50},paragraphs:['Dash Dot']}]}]}]}
 const html=renderToStaticMarkup(createElement(NativePptxTableGeometry,{preview,readingId:'read'}))
 expect(html).toContain('data-source-paint="preset-border"')
 expect(html).toContain('stroke-dasharray="8 6 2 6"')
 expect(html).toContain('stroke-linecap="butt"')
 expect(html).toContain('stroke-dashoffset="0"')
 expect(html).toContain('upper-left corner with zero phase')
 expect(html).toContain('PowerPoint edge phase and corner placement are unverified')
})

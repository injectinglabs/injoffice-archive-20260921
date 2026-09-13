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

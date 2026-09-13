import {createElement} from 'react'
import {renderToStaticMarkup} from 'react-dom/server'
import {describe,it,expect} from 'vitest'
import {NativeDocxTextboxGeometry,NativeDocxTextboxShapeView} from './NativeDocxTextboxGeometry'
import type {NativeDocxTextboxShapePaintV1} from '@injoffice/docs/native-docx'
describe('rectangle geometry UI',()=>{
 it('starts with browser inspection, no helper submission or fabricated preview',()=>{
  const html=renderToStaticMarkup(createElement(NativeDocxTextboxGeometry,{bytes:new Uint8Array([1]),packageDigest:'sha256:'+'a'.repeat(64)}))
  expect(html).toContain('Inspect rectangle geometry in browser');expect(html).not.toContain('<svg');expect(html).not.toContain('Render authored rectangle')
 })
 it('renders source geometry plus outlined glyphs with outward stroke padding',()=>{
  const paint={status:'supported',width_millipoints:216000,height_millipoints:72000,line_width_millipoints:1000,fill_rgb:'FFF2CC',line_rgb:'204060',text_rgb:'102030',paths:['M8000 8000 L9000 9000 Z']} as NativeDocxTextboxShapePaintV1
  const html=renderToStaticMarkup(createElement(NativeDocxTextboxShapeView,{paint}))
  expect(html).toContain('viewBox="-500 -500 217000 73000"');expect(html).toContain('fill="#FFF2CC"');expect(html).toContain('d="M8000 8000 L9000 9000 Z"');expect(html).not.toContain('<text');expect(html).toContain('Page placement is not produced')
 })
 it('discloses authored multiline centering without page placement claims',()=>{
  const paint={status:'supported',width_millipoints:216000,height_millipoints:72000,line_width_millipoints:0,fill_rgb:'none',line_rgb:'none',text_rgb:'102030',paths:[],line_layout:{lines:[{},{}]}} as NativeDocxTextboxShapePaintV1
  const html=renderToStaticMarkup(createElement(NativeDocxTextboxShapeView,{paint}));expect(html).toContain('2 authored lines. Each authored line is centered within its specified line height.');expect(html).toContain('Page placement is not produced');expect(html).not.toContain('<text')
 })

})

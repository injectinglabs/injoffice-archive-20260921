import {createElement} from 'react'
import {renderToStaticMarkup} from 'react-dom/server'
import {describe,it,expect} from 'vitest'
import {NativeDocxPages} from './NativeDocxPages'
import {NativeDocxTextboxOnPage,decodeTextboxPageResponse} from './NativeDocxTextboxPage'
import type {NativeDocxTextboxPagePreviewV1} from '@injoffice/docs/native-docx'

describe('page-placed textbox UI',()=>{
 it('requires explicit upload and discloses the embedded-font restriction',()=>{
  const html=renderToStaticMarkup(createElement(NativeDocxPages,{bytes:new Uint8Array([1]),packageDigest:'sha256:'+'a'.repeat(64),apiBase:'/helper'}))
  expect(html).toContain('Upload to helper and preview page-placed textboxes');expect(html).toContain('Nothing is uploaded until');expect(html).toContain('exact embedded regular font');expect(html).not.toContain('<svg')
 })
 it('mounts outlined text and source rectangle only on the selected physical page',()=>{
  const textbox={page_id:'page:2',x_millipoints:72000,y_millipoints:144000,paint:{width_millipoints:216000,height_millipoints:72000,line_width_millipoints:1000,fill_rgb:'FFF2CC',line_rgb:'204060',text_rgb:'102030',paths:['M8000 8000 L9000 9000 Z']}} as NativeDocxTextboxPagePreviewV1['textbox']
  expect(renderToStaticMarkup(createElement(NativeDocxTextboxOnPage,{textbox,pageID:'page:1'}))).toBe('')
  const html=renderToStaticMarkup(createElement(NativeDocxTextboxOnPage,{textbox,pageID:'page:2'}))
  expect(html).toContain('translate(72000 144000)');expect(html).toContain('width="216000"');expect(html).toContain('fill="#FFF2CC"');expect(html).toContain('stroke="#204060"');expect(html).toContain('d="M8000 8000 L9000 9000 Z"');expect(html).not.toContain('<text')
 })
 it('refuses incomplete helper data before a page can mount',()=>{
  for(const value of [{},{document:{},evidence:{},font_inventory_json:'{}',preview:{}},{document:{},evidence:{},font_inventory_json:'{}',preview:{},font_url:'https://untrusted.invalid/font'}])expect(()=>decodeTextboxPageResponse(value,'sha256:'+'a'.repeat(64))).toThrow()
 })
})

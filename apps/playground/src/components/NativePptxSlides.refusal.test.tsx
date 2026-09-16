import {describe,it,expect} from 'vitest'
import {renderToStaticMarkup} from 'react-dom/server'
import {NativePptxVector} from './NativePptxSlides'
import {decodePptxPreview} from '../../../pptx-page-paint-worker/src/contract'

// A refused shape must read as an empty region, not as painted content. A solid
// fill claims coverage we do not have: on bullet-indent.pptx two refusal boxes
// covered 70.8% of the slide while the preview painted nothing at all.
describe('refused shape chrome',()=>{
 const preview=decodePptxPreview({version:1,package_sha256:'a'.repeat(64),slide_index:0,slide_count:1,width:914400,height:914400,background:'FFFFFF',policy:'max-run-natural-v1',diagnostics:[],font_digests:[],resources:[],nodes:[{kind:'placeholder',rect:{x:0,y:0,cx:914400,cy:914400},label:'Unsupported shape'}]})
 it('outlines the region without filling it',()=>{
  const markup=renderToStaticMarkup(<NativePptxVector preview={preview}/>)
  const rect=markup.match(/<rect[^>]*>/)
  expect(rect).not.toBeNull()
  expect(rect![0]).toContain('fill="none"')
  expect(rect![0]).not.toContain('#eee')
  expect(rect![0]).toContain('stroke')
  expect(markup).toContain('Unsupported shape')
 })
 it('keeps the refusal visible as a dashed outline at the shape bounds',()=>{
  const markup=renderToStaticMarkup(<NativePptxVector preview={preview}/>)
  expect(markup).toContain('stroke-dasharray')
  // Full-bleed refusal still reports its own extent, so the region is locatable.
  expect(markup).toMatch(/width="72"[^>]*height="72"|height="72"[^>]*width="72"/)
 })
})

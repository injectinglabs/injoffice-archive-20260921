import {describe,it,expect} from 'vitest'
import {renderToStaticMarkup} from 'react-dom/server'
import {NativePptxVector} from './NativePptxSlides'
import {decodePptxPreview} from '../../../pptx-page-paint-worker/src/contract'

const base={version:1,package_sha256:'a'.repeat(64),slide_index:0,slide_count:1,width:12192000,height:6858000,background:'FFFFFF',policy:'max-run-natural-v1',diagnostics:[],font_digests:[],resources:[],nodes:[]}
const gradient={angle:5400000,stops:[{pos:0,color:'112233'},{pos:100000,color:'AABBCC'}]}

describe('native slide gradient background',()=>{
 it('paints the gradient inside the captured vector, not as a CSS background',()=>{
  const preview=decodePptxPreview({...base,background_gradient:gradient})
  const markup=renderToStaticMarkup(<NativePptxVector preview={preview}/>)
  expect(markup).toContain('<linearGradient')
  expect(markup).toContain('stop-color="#112233"')
  expect(markup).toContain('stop-color="#AABBCC"')
  // The paint server is referenced by a local fragment only; the isolated
  // capture rejects any other url() target.
  const id=/<linearGradient id="([^"]+)"/.exec(markup)?.[1]
  expect(id).toBeTruthy()
  expect(markup).toContain(`fill="url(#${id})"`)
  expect(markup).toContain('data-native-slide-background="gradient"')
 })

 it('gives two previews on one page distinct paint-server IDs',()=>{
  const preview=decodePptxPreview({...base,background_gradient:gradient})
  const markup=renderToStaticMarkup(<><NativePptxVector preview={preview}/><NativePptxVector preview={preview}/></>)
  const ids=[...markup.matchAll(/<linearGradient id="([^"]+)"/g)].map(m=>m[1])
  expect(ids).toHaveLength(2)
  expect(new Set(ids).size).toBe(2)
 })

 it('paints no gradient when the slide states none',()=>{
  const markup=renderToStaticMarkup(<NativePptxVector preview={decodePptxPreview(base)}/>)
  expect(markup).not.toContain('<linearGradient')
 })

 it('runs the axis through the slide centre along the declared angle',()=>{
  // 5400000 = 90 degrees clockwise from +x, which in SVG user space (y down)
  // is top to bottom: the axis is vertical and spans the full height.
  const markup=renderToStaticMarkup(<NativePptxVector preview={decodePptxPreview({...base,background_gradient:gradient})}/>)
  const attrs=/<linearGradient[^>]*x1="([^"]+)" y1="([^"]+)" x2="([^"]+)" y2="([^"]+)"/.exec(markup)
  expect(attrs).toBeTruthy()
  const [x1,y1,x2,y2]=attrs!.slice(1).map(Number)
  const width=base.width/12700,height=base.height/12700
  expect(x1).toBeCloseTo(width/2,6)
  expect(x2).toBeCloseTo(width/2,6)
  expect(y1).toBeCloseTo(0,6)
  expect(y2).toBeCloseTo(height,6)
 })

 it('rejects a gradient outside the transported bounds',()=>{
  for(const bad of [
   {angle:-1,stops:gradient.stops},
   {angle:21600000,stops:gradient.stops},
   {angle:0,stops:[{pos:0,color:'112233'}]},
   {angle:0,stops:[{pos:100000,color:'112233'},{pos:0,color:'AABBCC'}]},
   {angle:0,stops:[{pos:0,color:'112233'},{pos:0,color:'AABBCC'}]},
   {angle:0,stops:[{pos:0,color:'nope'},{pos:100000,color:'AABBCC'}]},
   {angle:0,stops:[{pos:0,color:'112233'},{pos:100001,color:'AABBCC'}]},
   {angle:0,stops:gradient.stops,extra:1},
   {stops:gradient.stops},
   {angle:0},
  ])expect(()=>decodePptxPreview({...base,background_gradient:bad})).toThrow()
 })
})

import {describe,it,expect} from 'vitest'
import {renderToStaticMarkup} from 'react-dom/server'
import {NativePptxVector} from './NativePptxSlides'
import {decodePptxPreview} from '../../../pptx-page-paint-worker/src/contract'

describe('native rounded picture clip SVG',()=>{
 it('uses unique user-space clip IDs and preserves enclosing transforms',()=>{
  const preview=decodePptxPreview({version:1,package_sha256:'a'.repeat(64),slide_index:0,slide_count:1,width:1000,height:1000,background:'FFFFFF',policy:'max-run-natural-v1',diagnostics:[],font_digests:[],resources:[],nodes:[{kind:'group',transform:[2,0,0,3,100,200],clip:{x:0,y:0,cx:300,cy:200,radius:33.334},children:[{kind:'rect',rect:{x:0,y:0,cx:300,cy:200},radius:0,fill:'00FF00'}]}]})
  const markup=renderToStaticMarkup(<><NativePptxVector preview={preview}/><NativePptxVector preview={preview}/></>)
  const ids=[...markup.matchAll(/<clipPath id="([^"]+)"/g)].map(m=>m[1])
  expect(ids).toHaveLength(2);expect(new Set(ids).size).toBe(2)
  for(const id of ids)expect(markup).toContain(`url(#${id})`)
  expect(markup).toContain('clipPathUnits="userSpaceOnUse"');expect(markup).toContain(`rx="${33.334/12700}" ry="${33.334/12700}"`);expect(markup).toContain(`matrix(2 0 0 3 ${100/12700} ${200/12700})`)
 })
 it('clips to a validated source-evaluated outline path instead of a rectangle',()=>{
  const preview=decodePptxPreview({version:1,package_sha256:'a'.repeat(64),slide_index:0,slide_count:1,width:1000,height:1000,background:'FFFFFF',policy:'max-run-natural-v1',diagnostics:[],font_digests:[],resources:[],nodes:[{kind:'group',transform:[1,0,0,1,0,0],clip:{x:0,y:0,cx:25400,cy:12700,d:'M0 6350 A12700 6350 0 0 1 12700 0 A12700 6350 0 0 1 25400 6350 Z'},children:[{kind:'rect',rect:{x:0,y:0,cx:25400,cy:12700},radius:0,fill:'00FF00'}]}]})
  const markup=renderToStaticMarkup(<NativePptxVector preview={preview}/>)
  const clip=markup.match(/<clipPath id="([^"]+)" clipPathUnits="userSpaceOnUse"><path d="([^"]+)"><\/path><\/clipPath>/)
  expect(clip).not.toBeNull()
  expect(clip![2]).toBe('M 0 0.5 A 1 0.5 0 0 1 1 0 A 1 0.5 0 0 1 2 0.5 Z')
  // The CLIP must be the outline path, not a rectangle. The rect the fixture
  // paints inside the clipped group is the content being clipped and has to
  // stay; asserting the whole markup carried no rect forbade the fixture's own
  // subject.
  expect(markup).toContain(`url(#${clip![1]})`);expect(clip![0]).not.toContain('<rect')
 })
})

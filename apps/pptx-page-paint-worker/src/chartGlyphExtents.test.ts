import {readFileSync} from 'node:fs'
import {createRequire} from 'node:module'
import {expect,it} from 'vitest'
import {createHarfBuzzOutlineProviderV1} from '@injoffice/font-metrics/harfbuzz'
import type {FontResource} from '@injoffice/font-metrics/layout'
import {chartGlyphExtents} from './chartGlyphExtents.js'
const bytes=new Uint8Array(readFileSync(createRequire(import.meta.url).resolve('dejavu-fonts-ttf/ttf/DejaVuSans.ttf'))),digest='sha256:7da195a74c55bef988d0d48f9508bd5d849425c1770dba5d7bfc6ce9ed848954'
it('retains normalized outline grid and exact identity from the paint provider',()=>{
 const provider=createHarfBuzzOutlineProviderV1({bytes,contentDigest:digest}),extents=chartGlyphExtents(new Map([['face',{face:{contentDigest:digest}} as unknown as FontResource]]),new Map([['face',provider]]))
 // DejaVu Sans A has design bounds16,0..1384,1493 in2048UPM; provider grid is256x.
 expect(extents({faceId:'face',contentDigest:digest,glyphId:36})).toEqual({faceId:'face',contentDigest:digest,glyphId:36,unitsPerEm:524288,bounds:{xMin:4096,yMin:0,xMax:354304,yMax:382208}})
 expect(extents({faceId:'face',contentDigest:digest,glyphId:3}).bounds).toBeNull()
 expect(()=>extents({faceId:'face',contentDigest:'wrong',glyphId:36})).toThrow(/identity/)
})

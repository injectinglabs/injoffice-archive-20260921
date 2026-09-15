import type { NativeElement } from '@injoffice/pptx-native'
import { evaluatedGeometryPaths } from './evaluatedGeometry.js'
import { RenderCompileError, type RenderClip, type RenderPathCommand, type RenderRect } from './types.js'

/** Source policy string the Go extractor attaches to a catalog-evaluated picture outline. */
export const PICTURE_GEOMETRY_CLIP_POLICY = 'pptx.picture-geometry-preview'

/** A picture's evaluated preset outline as its clip. Only fillable subpaths
 * bound the visible image; stroke-only decoration paths never widen the clip.
 * The contract already forbids editable or clip-conflicting geometry; this
 * boundary re-checks both so a stale deck cannot promote a preview to paint. */
export function pictureGeometryClip(element:Extract<NativeElement,{kind:'picture'}>,bounds:RenderRect,check:(value:number,path:string)=>void,checkWorld:(bounds:RenderRect)=>void):RenderClip {
 const path=`$.elements.${element.id}.geometry`
 if(!element.geometry)throw new RenderCompileError('picture.geometryMissing',path,'picture clip requires evaluated geometry')
 if(element.clip!==undefined)throw new RenderCompileError('picture.clipConflict',path,'picture clip and evaluated geometry are mutually exclusive')
 if(element.compatibility.status==='editable')throw new RenderCompileError('picture.geometryAuthority',path,'evaluated picture geometry must remain read-only')
 const commands:RenderPathCommand[]=[]
 for(const part of evaluatedGeometryPaths(element.geometry,check,path,checkWorld))if(part.fillMode!=='none')commands.push(...part.path)
 if(!commands.length)throw new RenderCompileError('picture.emptyClip',path,'evaluated picture geometry has no fillable outline')
 return {kind:'path',rect:bounds,path:commands}
}

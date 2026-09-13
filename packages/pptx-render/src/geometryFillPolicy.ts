/** ECMA-376 defines relative path tones but does not specify numeric strength.
 * This explicit preview policy blends 20%/40% black or white in linear sRGB.
 * It is deterministic, not a claim of PowerPoint's unspecified tone behavior. */
export const DRAWINGML_PATH_FILL_POLICY = 'linear-srgb-path-tone-20-40-v1'
export type DrawingMLPathFillMode = 'norm'|'none'|'darken'|'darkenLess'|'lighten'|'lightenLess'

export function geometryPathFill(color:string|undefined,mode:DrawingMLPathFillMode):string|undefined {
 if(!['norm','none','darken','darkenLess','lighten','lightenLess'].includes(mode))throw new TypeError('Unknown geometry path fill mode')
 if(color===undefined||mode==='none')return undefined
 if(!/^[0-9A-F]{6}$/.test(color))throw new TypeError('Geometry path fill must be a native RGB color')
 if(mode==='norm')return color
 const amount=mode==='darkenLess'||mode==='lightenLess'?0.2:0.4
 const lighten=mode==='lighten'||mode==='lightenLess'
 const linear=(n:number)=>n<=0.04045?n/12.92:((n+0.055)/1.055)**2.4
 const encoded=(n:number)=>n<=0.0031308?12.92*n:1.055*n**(1/2.4)-0.055
 return [0,2,4].map(i=>{
  const original=linear(Number.parseInt(color.slice(i,i+2),16)/255)
  const toned=lighten?original+(1-original)*amount:original*(1-amount)
  return Math.round(encoded(toned)*255).toString(16).padStart(2,'0').toUpperCase()
 }).join('')
}

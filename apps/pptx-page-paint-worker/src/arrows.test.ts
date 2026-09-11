import {describe,it,expect} from 'vitest'
import {previewArrow,previewArrowShaftInset} from './arrows.js'
describe('explicit arrow-v1 preview geometry',()=>{
 it('insets shafts under filled endpoints instead of painting flat stubs over their tips',()=>{expect(previewArrowShaftInset({type:'triangle',len:'lg'},10)).toBe(50);expect(previewArrowShaftInset({type:'stealth',len:'lg'},10)).toBe(37.5);expect(previewArrowShaftInset({type:'diamond',len:'lg'},10)).toBe(25);expect(previewArrowShaftInset({type:'arrow'},10)).toBe(0)})
 it('preserves distinct named shapes and width/length policy without mutating descriptors',()=>{
  const rendered=[]
  for(const type of ['triangle','stealth','diamond','oval','arrow'] as const){const end={type,w:'lg',len:'sm'} as const,before=JSON.stringify(end),node=previewArrow(end,{x:100,y:50},{x:1,y:0},10,'112233')!;expect(JSON.stringify(end)).toBe(before);expect(node.kind).toBe('group');if(node.kind==='group'){expect(node.transform).toEqual([1,0,-0,1,100,50]);expect(node.sourceRole).toBe('connectorArrow');rendered.push(JSON.stringify(node.children));if(type==='oval')expect(node.children[0]).toEqual({kind:'ellipse',rect:{x:-20,y:-25,cx:20,cy:50},fill:'112233'})}}
  expect(new Set(rendered).size).toBe(5)
 })
 it('orients the head against the path and refuses degenerate lines/strokes',()=>{expect(previewArrow({type:'triangle'},{x:0,y:0},{x:-1,y:0},10,'112233')).toMatchObject({transform:[-1,0,-0,-1,0,0]});expect(previewArrow({type:'none'},{x:0,y:0},{x:0,y:0},0,'112233')).toBeUndefined();expect(()=>previewArrow({type:'triangle'},{x:0,y:0},{x:0,y:0},10,'112233')).toThrow();expect(()=>previewArrow({type:'triangle'},{x:0,y:0},{x:1,y:0},0,'112233')).toThrow()})
})

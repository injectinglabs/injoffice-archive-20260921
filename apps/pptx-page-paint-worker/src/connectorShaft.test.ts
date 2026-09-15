import {describe,expect,it} from 'vitest'
import type {RenderPathCommand} from '@injoffice/pptx-render'
import {previewConnectorShaft} from './connectorShaft.js'

const straight:RenderPathCommand[]=[{kind:'moveTo',x:0,y:0},{kind:'lineTo',x:1000,y:0}]
const elbow:RenderPathCommand[]=[{kind:'moveTo',x:0,y:0},{kind:'lineTo',x:1000,y:0},{kind:'lineTo',x:1000,y:500}]
const curve:RenderPathCommand[]=[{kind:'moveTo',x:0,y:0},{kind:'cubicBezierTo',x1:500,y1:0,x2:500,y2:500,x:1000,y:500}]

describe('previewConnectorShaft',()=>{
 it('keeps the exact straight arrow-v1 shaft inset and tangents',()=>{
  const shaft=previewConnectorShaft(straight,undefined,{type:'triangle',len:'lg'},10)
  expect(shaft).toEqual({d:'M0 0 L950 0',head:{tip:{x:0,y:0},direction:{x:-1000,y:0}},tail:{tip:{x:1000,y:0},direction:{x:1000,y:0}},straight:true})
  expect(previewConnectorShaft(straight,{type:'triangle',len:'lg'},undefined,10).d).toBe('M50 0 L1000 0')
  expect(previewConnectorShaft(straight,undefined,{type:'arrow'},10).d).toBe('M0 0 L1000 0')
 })
 it('insets only the terminal segments of an elbow and reports their tangents',()=>{
  const shaft=previewConnectorShaft(elbow,{type:'triangle',len:'sm'},{type:'triangle',len:'lg'},10)
  expect(shaft.d).toBe('M20 0 L1000 0 L1000 450')
  expect(shaft.head).toEqual({tip:{x:0,y:0},direction:{x:-1000,y:0}})
  expect(shaft.tail).toEqual({tip:{x:1000,y:500},direction:{x:0,y:500}})
  expect(shaft.straight).toBe(false)
 })
 it('trims a curved terminal segment by chord distance with a bounded split and keeps the curve tangents',()=>{
  const shaft=previewConnectorShaft(curve,undefined,{type:'triangle',len:'lg'},10)
  expect(shaft.head).toEqual({tip:{x:0,y:0},direction:{x:-500,y:0}})
  expect(shaft.tail).toEqual({tip:{x:1000,y:500},direction:{x:500,y:0}})
  const match=/^M0 0 C(\S+) (\S+) (\S+) (\S+) (\S+) (\S+)$/.exec(shaft.d)
  expect(match).not.toBeNull()
  const [,,c1y,,,endX,endY]=match!.map(Number)
  expect(c1y).toBe(0)
  expect(Math.hypot(1000-endX!,500-endY!)).toBeCloseTo(50,6)
  expect(endX!).toBeLessThan(1000)
 })
 it('refuses closed, arc, non-finite, or too-short shafts instead of drawing an approximation',()=>{
  expect(()=>previewConnectorShaft([{kind:'moveTo',x:0,y:0},{kind:'lineTo',x:1000,y:0},{kind:'close'}],undefined,undefined,10)).toThrow('open source connector path')
  expect(()=>previewConnectorShaft([{kind:'moveTo',x:0,y:0},{kind:'arcTo',x:5,y:5,rx:5,ry:5,largeArc:false,clockwise:true}],undefined,undefined,10)).toThrow('open source connector path')
  expect(()=>previewConnectorShaft([{kind:'moveTo',x:0,y:0}],undefined,undefined,10)).toThrow('open source connector path')
  expect(()=>previewConnectorShaft([{kind:'moveTo',x:0,y:0},{kind:'lineTo',x:Number.NaN,y:0}],undefined,undefined,10)).toThrow('open source connector path')
  expect(()=>previewConnectorShaft([{kind:'moveTo',x:0,y:0},{kind:'lineTo',x:60,y:0}],{type:'triangle',len:'lg'},{type:'triangle',len:'lg'},10)).toThrow('too short')
  expect(()=>previewConnectorShaft([{kind:'moveTo',x:0,y:0},{kind:'cubicBezierTo',x1:10,y1:0,x2:20,y2:0,x:30,y:0}],undefined,{type:'triangle',len:'lg'},10)).toThrow('too short')
 })
})

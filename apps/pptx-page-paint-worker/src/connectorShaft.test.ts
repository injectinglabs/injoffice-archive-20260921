import {describe,expect,it} from 'vitest'
import type {RenderPathCommand} from '@injoffice/pptx-render'
import {previewConnectorShaft} from './connectorShaft.js'

const straight:RenderPathCommand[]=[{kind:'moveTo',x:0,y:0},{kind:'lineTo',x:1000,y:0}]
const elbow:RenderPathCommand[]=[{kind:'moveTo',x:0,y:0},{kind:'lineTo',x:1000,y:0},{kind:'lineTo',x:1000,y:500}]
const curve:RenderPathCommand[]=[{kind:'moveTo',x:0,y:0},{kind:'cubicBezierTo',x1:500,y1:0,x2:500,y2:500,x:1000,y:500}]

describe('previewConnectorShaft',()=>{
 it('keeps the exact straight arrow-v1 shaft inset and tangents',()=>{
  const shaft=previewConnectorShaft(straight,undefined,{type:'triangle',len:'lg'},10)
  expect(shaft).toEqual({d:'M0 0 L950 0',head:{tip:{x:0,y:0},direction:{x:-1000,y:0}},tail:{tip:{x:1000,y:0},direction:{x:1000,y:0}},straight:true,skippedInsets:[]})
  expect(previewConnectorShaft(straight,{type:'triangle',len:'lg'},undefined,10).d).toBe('M50 0 L1000 0')
  expect(previewConnectorShaft(straight,undefined,{type:'arrow'},10).d).toBe('M0 0 L1000 0')
 })
 it('insets only the terminal segments of an elbow and reports their tangents',()=>{
  const shaft=previewConnectorShaft(elbow,{type:'triangle',len:'sm'},{type:'triangle',len:'lg'},10)
  expect(shaft.d).toBe('M20 0 L1000 0 L1000 450')
  expect(shaft.head).toEqual({tip:{x:0,y:0},direction:{x:-1000,y:0}})
  expect(shaft.tail).toEqual({tip:{x:1000,y:500},direction:{x:0,y:500}})
  expect(shaft.straight).toBe(false)
  expect(shaft.skippedInsets).toEqual([])
 })
 it('leaves a terminal segment shorter than its inset untrimmed and reports it instead of throwing',()=>{
  // bentConnector2 with cy below 3x the stroke width (tail triangle med = 30).
  const short=previewConnectorShaft([{kind:'moveTo',x:0,y:0},{kind:'lineTo',x:2000,y:0},{kind:'lineTo',x:2000,y:10}],undefined,{type:'triangle'},10)
  expect(short.d).toBe('M0 0 L2000 0 L2000 10');expect(short.tail).toEqual({tip:{x:2000,y:10},direction:{x:0,y:10}});expect(short.skippedInsets).toEqual(['tail'])
  // Straight line shorter than head+tail insets keeps the head trim and skips the tail.
  const both=previewConnectorShaft([{kind:'moveTo',x:0,y:0},{kind:'lineTo',x:60,y:0}],{type:'triangle',len:'lg'},{type:'triangle',len:'lg'},10)
  expect(both.d).toBe('M50 0 L60 0');expect(both.skippedInsets).toEqual(['tail'])
  const shortCurve=previewConnectorShaft([{kind:'moveTo',x:0,y:0},{kind:'cubicBezierTo',x1:10,y1:0,x2:20,y2:0,x:30,y:0}],undefined,{type:'triangle',len:'lg'},10)
  expect(shortCurve.d).toBe('M0 0 C10 0 20 0 30 0');expect(shortCurve.skippedInsets).toEqual(['tail'])
 })
 it('takes arrow tangents from the first non-degenerate segment and yields a zero vector only for a fully degenerate path',()=>{
  // bentConnector3 with adj1=0: M0 0 L0 0 L0 500 L1000 500.
  const elbowAtZero=previewConnectorShaft([{kind:'moveTo',x:0,y:0},{kind:'lineTo',x:0,y:0},{kind:'lineTo',x:0,y:500},{kind:'lineTo',x:1000,y:500}],{type:'triangle'},undefined,10)
  expect(elbowAtZero.head).toEqual({tip:{x:0,y:0},direction:{x:0,y:-500}});expect(elbowAtZero.skippedInsets).toEqual(['head']);expect(elbowAtZero.d).toBe('M0 0 L0 0 L0 500 L1000 500')
  const degenerate=previewConnectorShaft([{kind:'moveTo',x:5,y:5},{kind:'lineTo',x:5,y:5}],{type:'triangle'},{type:'triangle'},10)
  expect(degenerate.head.direction).toEqual({x:0,y:0});expect(degenerate.tail.direction).toEqual({x:0,y:0});expect(degenerate.skippedInsets).toEqual(['head','tail'])
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
 it('refuses closed, arc, or non-finite shafts instead of drawing an approximation',()=>{
  expect(()=>previewConnectorShaft([{kind:'moveTo',x:0,y:0},{kind:'lineTo',x:1000,y:0},{kind:'close'}],undefined,undefined,10)).toThrow('open source connector path')
  expect(()=>previewConnectorShaft([{kind:'moveTo',x:0,y:0},{kind:'arcTo',x:5,y:5,rx:5,ry:5,largeArc:false,clockwise:true}],undefined,undefined,10)).toThrow('open source connector path')
  expect(()=>previewConnectorShaft([{kind:'moveTo',x:0,y:0}],undefined,undefined,10)).toThrow('open source connector path')
  expect(()=>previewConnectorShaft([{kind:'moveTo',x:0,y:0},{kind:'lineTo',x:Number.NaN,y:0}],undefined,undefined,10)).toThrow('open source connector path')
 })
})

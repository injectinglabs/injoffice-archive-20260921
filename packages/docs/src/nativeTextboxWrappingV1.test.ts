import {describe,it,expect} from 'vitest'
import {planTextboxWrap,textboxWrapPolicyForText,type NativeTextboxWrapClusterV1} from './nativeTextboxWrappingV1.js'
function clusters(text:string):NativeTextboxWrapClusterV1[]{return [...text].map((_,i)=>({start_utf16:i,end_utf16:i+1,advance_millipoints:10,unsafe_to_break:false,glyph_start:i,glyph_end:i+1}))}
describe('bounded textbox cluster wrapping planner',()=>{
 it('retains separator advances and selects the last fitting safe word boundary',()=>{
  const text='aa bb cc',c=clusters(text),p=planTextboxWrap(text,c,60)
  expect(p.map(l=>text.slice(l.start_utf16,l.end_utf16))).toEqual(['aa bb ','cc']);expect(p.map(l=>l.advance_millipoints)).toEqual([60,20]);expect(planTextboxWrap(text,c,80)).toHaveLength(1)
 })
 it('refuses space-edge overflow and unsafe boundaries on either adjacent cluster',()=>{
  const text='aa bb';expect(()=>planTextboxWrap(text,clusters(text),20)).toThrow('wrap-word-or-space-overflow')
  for(const index of [2,3]){const c=clusters(text);c[index]!.unsafe_to_break=true;expect(()=>planTextboxWrap(text,c,30)).toThrow('wrap-word-or-space-overflow')}
  expect(planTextboxWrap(text,clusters(text),30).map(l=>l.end_utf16)).toEqual([3,5])
 })
 it('preserves multi-character ligature clusters without emergency splitting',()=>{
  const text='office office',c=clusters(text);c.splice(1,3,{start_utf16:1,end_utf16:4,advance_millipoints:20,unsafe_to_break:false,glyph_start:1,glyph_end:4})
  const p=planTextboxWrap(text,c,60);expect(p.map(l=>l.end_utf16)).toEqual([7,13]);expect(()=>planTextboxWrap(text,c,25)).toThrow('wrap-word-or-space-overflow')
 })
 it('rejects coverage gaps, malformed cluster records and output-line budget exhaustion',()=>{
  for(const mutate of [(c:NativeTextboxWrapClusterV1[])=>{c[0]!.start_utf16=1},(c:NativeTextboxWrapClusterV1[])=>{c[1]!.glyph_start=0},(c:NativeTextboxWrapClusterV1[])=>{c.pop()},(c:NativeTextboxWrapClusterV1[])=>{c[0]!.advance_millipoints=Infinity}]){const c=clusters('aa bb');mutate(c);expect(()=>planTextboxWrap('aa bb',c,50)).toThrow()}
  const text=Array(17).fill('a').join(' ');expect(()=>planTextboxWrap(text,clusters(text),20)).toThrow('wrap-line-budget')
 })
})

describe('finite punctuated source profile',()=>{
 const policy='ascii-punctuation-space-greedy-v1' as const
 it('matches the Go grammar and keeps all admitted punctuation in source coverage',()=>{
  for(const text of ['Hello, world.',"It's 3.14 miles.",'Read well-known facts/figures.','A (short phrase) ends.','Why? Yes! Next: value; done.','(One) word','U.S.A. example']){
   expect(textboxWrapPolicyForText(text)).toBe(policy);const lines=planTextboxWrap(text,clusters(text),140,policy)
   expect(lines.map(l=>text.slice(l.start_utf16,l.end_utf16)).join('')).toBe(text)
   for(const l of lines.slice(0,-1))expect(text[l.end_utf16-1]).toBe(' ')
  }
 })
 it('refuses opening/closing-space ambiguity, quotes and unsupported punctuation patterns',()=>{
  for(const text of ['( word)','(word )','word )','(word','word)','((word))','"quoted text"',"'quoted text'","word''s",'Hello!!','Wait...','word , next','a  b','a\tb','a\nb','a@b','a & b']){expect(textboxWrapPolicyForText(text)).toBeUndefined();expect(()=>planTextboxWrap(text,clusters(text),1000,policy)).toThrow()}
  expect(()=>planTextboxWrap('Alpha beta',clusters('Alpha beta'),100,policy)).toThrow();expect(()=>planTextboxWrap('Hello.',clusters('Hello.'),100)).toThrow()
 })
 it('never breaks inside apostrophes, decimals, hyphens or slashes or ignores unsafe flags',()=>{
  for(const text of ["can't stop",'3.14 value','well-known facts','facts/figures next']){expect(()=>planTextboxWrap(text,clusters(text),20,policy)).toThrow('wrap-word-or-space-overflow')}
  const text='Hello, world.',c=clusters(text);c[6]!.unsafe_to_break=true;expect(()=>planTextboxWrap(text,c,70,policy)).toThrow('wrap-word-or-space-overflow')
 })
})

from fractions import Fraction as F
from decimal import Decimal as D, localcontext
from xml.etree import ElementTree as E
from pathlib import Path
import json, math, hashlib, sys
from decimal import ROUND_HALF_UP
root=Path(__file__).resolve().parent.parent
raw=(root/'go/pptxpatch/presetdata/preset-shapes.xml').read_bytes();catalog=E.fromstring(raw);ns={'a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
def dec(x):return D(x.numerator)/D(x.denominator) if isinstance(x,F) else x
def guide_formula(words,g):
 op,*tokens=words.split();v=[g[t] if t in g else F(t) for t in tokens]
 if any(isinstance(x,D) for x in v):v=[dec(x) for x in v]
 if op=='val':return v[0]
 if op=='*/':return v[0]*v[1]/v[2]
 if op=='+-':return v[0]+v[1]-v[2]
 if op=='pin':return min(max(v[1],v[0]),v[2])
 if op=='min':return min(v)
 if op in ['sin','cos']:
  assert v[1]==2700000
  return dec(v[0])*D(2).sqrt()/2
 if op=='mod':return sum(dec(x)**2 for x in v).sqrt()
 raise ValueError(words)
def point(rx,ry,q):return [(rx,F(0)),(F(0),ry),(-rx,F(0)),(F(0),-ry)][q%4]
def serial(x):
 if isinstance(x,F):return str(x)
 if isinstance(x,D):return str(x)
 raise TypeError(type(x))
results=[]
with localcontext() as ctx:
 ctx.prec=90
 for name in ['can','leftBrace','rightBrace','leftBracket','rightBracket','cornerTabs','plaqueTabs','squareTabs']:
  shape=catalog.find(name);g={'w':F(4000000),'h':F(400000),'r':F(4000000),'b':F(400000),'l':F(0),'t':F(0),'ss':F(400000),'wd2':F(2000000),'hc':F(2000000),'cd2':F(10800000),'cd4':F(5400000),'3cd4':F(16200000)}
  for listname in ['avLst','gdLst']:
   for n in shape.findall('a:'+listname+'/a:gd',ns):g[n.get('name')]=guide_formula(n.get('fmla'),g)
  rect=shape.find('a:rect',ns);edges=[g[rect.get(k)] for k in ['l','t','r','b']];rectvalid=edges[2]>edges[0] and edges[3]>edges[1]
  record={'name':name,'frame':[4000000,400000],'textRectEdges90digits':[str(dec(x)) for x in edges],'positiveTextRect':rectvalid,'paths':[]}
  if not rectvalid:
   record['classification']='Pinned text rectangle inverted, not an evaluator arithmetic error';results.append(record);continue
  for path in shape.findall('a:pathLst/a:path',ns):
   assert path.get('w') is None and path.get('h') is None
   pen=None;origin=None;vertices=[];arcs=[];construction=0;commands=[]
   for cmd in path:
    kind=cmd.tag.split('}')[-1]
    if kind in ['moveTo','lnTo']:
     p=cmd.find('a:pt',ns);pen=(g[p.get('x')],g[p.get('y')]);assert all(isinstance(x,F) and x.denominator==1 for x in pen)
     if kind=='moveTo':origin=pen
     vertices.append(pen);commands.append({'kind':kind,'point':pen})
    elif kind=='close':pen=origin;commands.append({'kind':'close'})
    elif kind=='arcTo':
     vals=[g[t] if t in g else F(t) for t in [cmd.get('wR'),cmd.get('hR'),cmd.get('stAng'),cmd.get('swAng')]];rx,ry,start,sweep=vals
     assert all(x.denominator==1 for x in vals) and rx>0 and ry>0 and start%5400000==0 and sweep%5400000==0
     q=int(start/5400000);count=abs(int(sweep/5400000));direction=1 if sweep>0 else -1
     offset=point(rx,ry,q);center=(pen[0]-offset[0],pen[1]-offset[1]);maxr=max(rx,ry);aspect=maxr/min(rx,ry)
     construction+=1024*2.220446049250313e-16*float(max(maxr,abs(pen[0]),abs(pen[1])))*float(aspect)**3
     endpoint=pen
     for step in range(1,count+1):
      p=point(rx,ry,q+direction*step);endpoint=(center[0]+p[0],center[1]+p[1]);dx,dy=endpoint[0]-pen[0],endpoint[1]-pen[1]
      assert dx*dx*ry*ry+dy*dy*rx*rx<=4*rx*rx*ry*ry
      assert all(x.denominator==1 and abs(x)<=2**53-1 for x in [*center,*endpoint,rx,ry])
      vertices.append(endpoint);commands.append({'kind':'arcTo','radius':(rx,ry),'end':endpoint,'clockwise':direction>0});pen=endpoint
     arcs.append({'rx':rx,'ry':ry,'startQuarter':q,'sweepQuarters':direction*count,'center':center,'end':endpoint,'constructionAllowanceCumulativeEMU':construction})
    else:raise ValueError(kind)
   record['paths'].append({'exactQuarterSegments':sum(abs(a['sweepQuarters']) for a in arcs),'commands':commands,'arcs':arcs,'exactExtrema':[min(p[0] for p in vertices),min(p[1] for p in vertices),max(p[0] for p in vertices),max(p[1] for p in vertices)],'genericConstructionAllowanceEMU':construction,'allPaintCoordinatesExactlyRepresentable':True})
  record['classification']='Valid exact integer cardinal paths; generic noncardinal construction allowance causes false refusal';results.append(record)
cases=[]
for r in results[:5]:
 paths=[]
 for path in r['paths']:
  commands=[]
  for c in path['commands']:
   q={'kind':'lineTo' if c['kind']=='lnTo' else c['kind']}
   if 'point' in c:q.update(zip(['x','y'],map(int,c['point'])))
   if 'end' in c:q.update(zip(['x','y'],map(int,c['end'])))
   if 'radius' in c:
    q.update(zip(['rx','ry'],map(int,c['radius'])))
    q.update(largeArc=False,clockwise=c['clockwise'])
   commands.append(q)
  paths.append(commands)
 edges=[int(D(x).quantize(D(1),rounding=ROUND_HALF_UP)) for x in r['textRectEdges90digits']]
 cases.append(dict(name=r['name'],paths=paths,textRect=dict(x=edges[0],y=edges[1],cx=edges[2]-edges[0],cy=edges[3]-edges[1])))
result=json.dumps({'provenance':'Independent Fraction cardinal endpoints and 90-digit Decimal text rectangle oracle; pinned catalog SHA256 '+hashlib.sha256(raw).hexdigest(),'cases':cases},indent=2)+'\n'
target=root/'go/pptxpatch/testdata/exact-cardinal-arcs-oracle.json'
if '--check' in sys.argv:
 assert target.read_text()==result, 'Independent oracle differs; review pinned catalog before regenerating'
else:target.write_text(result)
print('Five exact cardinal path/text rectangle oracles verified; three inverted text rectangles retained')

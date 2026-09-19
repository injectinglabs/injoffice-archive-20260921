import type { NativeElement } from '@injoffice/pptx-native';

export default function ShapeArt({ element, width, height, fill, stroke, strokeWidth }: { element: NativeElement; width: number; height: number; fill: string; stroke: string; strokeWidth: number }) {
  const nativeStroke = element.kind === 'shape' || element.kind === 'connector' ? element.stroke : undefined;
  const props = { fill, stroke, strokeWidth, strokeLinecap: (nativeStroke?.cap === 'flat' ? 'butt' : nativeStroke?.cap === 'square' ? 'square' : 'round') as 'butt' | 'square' | 'round', strokeLinejoin: nativeStroke?.join, strokeMiterlimit: nativeStroke?.miterLimit === undefined ? undefined : nativeStroke.miterLimit / 100_000 };
  if (element.kind === 'connector') return <line x1="0" y1={element.flipH ? height : 0} x2={width} y2={element.flipH ? 0 : height} stroke={stroke} strokeWidth={strokeWidth} />;
  if (element.kind !== 'shape') return null;
  if (element.preset === 'ellipse') return <ellipse cx={width / 2} cy={height / 2} rx={width / 2} ry={height / 2} {...props} />;
  if (element.preset === 'triangle') return <polygon points={`${width / 2},0 ${width},${height} 0,${height}`} {...props} />;
  if (element.preset === 'diamond') return <polygon points={`${width / 2},0 ${width},${height / 2} ${width / 2},${height} 0,${height / 2}`} {...props} />;
  return <rect x="0" y="0" width={width} height={height} rx={element.preset === 'roundRect' ? Math.min(width, height) * .12 : 0} {...props} />;
}

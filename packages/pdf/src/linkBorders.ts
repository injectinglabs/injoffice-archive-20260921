/** Canvas counterpart of PDF.js's HTML-only link borders. No actions or URLs
 * are executed. An existing PDF appearance stream remains authoritative. */
export function paintLinkBorders(context: CanvasRenderingContext2D, annotations: readonly any[], transform: readonly number[], pixelRatio: number): void {
  for (const annotation of annotations) {
    if (annotation.subtype !== 'Link' || annotation.hasAppearance || (annotation.annotationFlags & (1 | 2 | 32)) !== 0) continue;
    // Non-rectangular and fixed-orientation annotations need their own geometry.
    if (annotation.quadPoints || annotation.noRotate || annotation.noZoom || (annotation.annotationFlags & (8 | 16)) !== 0) continue;
    const { rect, borderStyle: border, borderColor: color } = annotation;
    if (!Array.isArray(rect) || rect.length !== 4 || !rect.every(Number.isFinite) || !border || !Number.isFinite(border.width) || border.width <= 0) continue;
    if (!color || color.length !== 3 || !Array.from(color as ArrayLike<number>).every((v) => Number.isFinite(v) && v >= 0 && v <= 255)) continue;
    // PDF border styles: solid=1, dashed=2, underline=5.
    if (![1, 2, 5].includes(border.style)) continue;
    const x = Math.min(rect[0], rect[2]), y = Math.min(rect[1], rect[3]);
    const width = Math.abs(rect[2] - rect[0]), height = Math.abs(rect[3] - rect[1]);
    if (!width || !height) continue;
    context.save();
    try {
      context.setTransform(pixelRatio, 0, 0, pixelRatio, 0, 0);
      context.transform(...transform as [number, number, number, number, number, number]);
      context.strokeStyle = `rgb(${color[0]}, ${color[1]}, ${color[2]})`;
      context.lineWidth = border.width;
      context.lineCap = 'butt';
      context.lineJoin = 'miter';
      const dash = border.dashArray;
      context.setLineDash(border.style === 2 && Array.isArray(dash) && dash.length > 0 && dash.every((v: number) => Number.isFinite(v) && v >= 0) && dash.some((v: number) => v > 0) ? dash : []);
      context.lineDashOffset = 0;
      context.beginPath();
      // Stroke the authored PDF rectangle. Insetting a thick border would
      // cover the linked text and shrink the source geometry.
      if (border.style === 5) {
        context.moveTo(x, y);
        context.lineTo(x + width, y);
      } else {
        const left = x, bottom = y, right = x + width, top = y + height;
        const rx = Math.max(0, Math.min(Number.isFinite(border.horizontalCornerRadius) ? border.horizontalCornerRadius : 0, (right - left) / 2));
        const ry = Math.max(0, Math.min(Number.isFinite(border.verticalCornerRadius) ? border.verticalCornerRadius : 0, (top - bottom) / 2));
        if (rx > 0 && ry > 0) {
          context.moveTo(left + rx, bottom);
          context.lineTo(right - rx, bottom);
          context.ellipse(right - rx, bottom + ry, rx, ry, 0, -Math.PI / 2, 0);
          context.lineTo(right, top - ry);
          context.ellipse(right - rx, top - ry, rx, ry, 0, 0, Math.PI / 2);
          context.lineTo(left + rx, top);
          context.ellipse(left + rx, top - ry, rx, ry, 0, Math.PI / 2, Math.PI);
          context.lineTo(left, bottom + ry);
          context.ellipse(left + rx, bottom + ry, rx, ry, 0, Math.PI, 3 * Math.PI / 2);
          context.closePath();
        } else {
          context.rect(left, bottom, right - left, top - bottom);
        }
      }
      context.stroke();
    } finally {
      context.restore();
    }
  }
}

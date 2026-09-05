#!/usr/bin/env python3
"""Generate testdata/python_pptx_sample.pptx — a REAL .pptx built by an
independent producer (python-pptx), used to test pptxpatch.ParsePPTX against
something that isn't our own writer's output. Content is synthetic test
data; the file itself derives from python-pptx's built-in default template
(MIT license, Steve Canny).

Usage: python3 scripts/gen_fixture.py   (run from go/pptxpatch/)
"""
from pptx import Presentation
from pptx.util import Inches, Pt

prs = Presentation()  # python-pptx's built-in default template

# Slide 1: title + content layout, with a bulleted body and a real autoshape.
slide = prs.slides.add_slide(prs.slide_layouts[1])
slide.shapes.title.text = "InjOffice x python-pptx"

body = slide.placeholders[1]
tf = body.text_frame
tf.text = "First real bullet from python-pptx"
p2 = tf.add_paragraph()
p2.text = "Second bullet, independently authored"
p2.level = 0
p3 = tf.add_paragraph()
p3.text = "Nested bullet"
p3.level = 1

from pptx.enum.shapes import MSO_SHAPE
rect = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1), Inches(5), Inches(3), Inches(1))
rect.text_frame.text = "A real autoshape"

# Slide 2: a simple centered title-only layout.
slide2 = prs.slides.add_slide(prs.slide_layouts[0])
slide2.shapes.title.text = "Thanks for reading"
if slide2.placeholders and len(slide2.placeholders) > 1:
    slide2.placeholders[1].text = "python-pptx generated this slide"

prs.save("testdata/python_pptx_sample.pptx")
print("wrote testdata/python_pptx_sample.pptx")

"""Builds the public demo template and icon set of slotdeck.

    python scripts/build-demo-template.py

Writes templates/demo.pptx and examples/icons/*.png. Everything in them is drawn by this script:
no stock photo, font file, logo or icon from anywhere else is used, so both are free to reuse.
The result is committed; run this only to change the design. Needs python-pptx and Pillow.

The template is an ordinary deck of sample slides, the kind a designer hands over: free-standing
text boxes, pictures, grouped cards and a table, with filler text. slotdeck turns each slide into
a reusable pattern; nothing in here was shaped to suit it beyond being a normal slide deck.
"""
import datetime
import io
import math
import os
import re
import zipfile

from lxml import etree
from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')

INK = '1B2A41'
TEAL = '0E7C7B'
AMBER = 'F4A259'
TINT = 'E8F1F0'
MUTED = '5C6B73'
WHITE = 'FFFFFF'
HEAD = 'Cambria'
BODY = 'Calibri'
ORG = 'Organization name  |  Department'


def rgb(value):
    return tuple(int(value[i : i + 2], 16) for i in (0, 2, 4))


# --- pictures -------------------------------------------------------------------------------------


def png(image):
    data = io.BytesIO()
    image.save(data, 'PNG', optimize=True)
    data.seek(0)
    return data


def scene(width, height, sky, hills, sun):
    """A drawn landscape that stands where a photo will go."""
    scale = 2
    w, h = width * scale, height * scale
    image = Image.new('RGB', (w, h))
    draw = ImageDraw.Draw(image)
    top, bottom = rgb(sky[0]), rgb(sky[1])
    for y in range(h):
        t = y / (h - 1)
        draw.line([(0, y), (w, y)], fill=tuple(round(top[i] + (bottom[i] - top[i]) * t) for i in range(3)))
    radius = h * 0.11
    cx, cy = w * 0.72, h * 0.3
    draw.ellipse([cx - radius, cy - radius, cx + radius, cy + radius], fill=rgb(sun))
    for layer, color in enumerate(hills):
        base = h * (0.55 + layer * 0.13)
        amplitude = h * (0.09 - layer * 0.015)
        points = [(0, h)]
        for step in range(0, 101):
            x = w * step / 100
            y = base + amplitude * math.sin(step / 100 * math.pi * (1.6 + layer * 0.9) + layer * 1.7) + amplitude * 0.4 * math.sin(step / 100 * math.pi * 5 + layer)
            points.append((x, y))
        points.append((w, h))
        draw.polygon(points, fill=rgb(color))
    return png(image.resize((width, height), Image.LANCZOS))


def bar_chart(width, height, values):
    """A drawn bar chart that stands where a supplied chart or diagram will go."""
    image = Image.new('RGB', (width, height), rgb(WHITE))
    draw = ImageDraw.Draw(image)
    left, right, top, bottom = width * 0.08, width * 0.96, height * 0.1, height * 0.86
    for line in range(5):
        y = bottom - (bottom - top) * line / 4
        draw.line([(left, y), (right, y)], fill=rgb('D9E2E1'), width=max(2, height // 300))
    slot = (right - left) / len(values)
    for index, value in enumerate(values):
        x0 = left + slot * index + slot * 0.2
        x1 = left + slot * (index + 1) - slot * 0.2
        draw.rounded_rectangle([x0, bottom - (bottom - top) * value, x1, bottom], radius=slot * 0.08, fill=rgb(AMBER if index == len(values) - 2 else TEAL))
    draw.line([(left, bottom), (right, bottom)], fill=rgb(INK), width=max(3, height // 200))
    return png(image)


def flow_diagram(width, height, steps=4):
    """A drawn process diagram for the wide picture slide."""
    image = Image.new('RGB', (width, height), rgb(WHITE))
    draw = ImageDraw.Draw(image)
    box_w, box_h = width * 0.17, height * 0.42
    gap = (width - steps * box_w) / (steps + 1)
    y0 = (height - box_h) / 2
    for index in range(steps):
        x0 = gap + index * (box_w + gap)
        draw.rounded_rectangle([x0, y0, x0 + box_w, y0 + box_h], radius=box_h * 0.16, fill=rgb(TEAL if index % 2 == 0 else INK))
        radius = box_h * 0.16
        cx, cy = x0 + box_w / 2, y0 + box_h / 2
        draw.ellipse([cx - radius, cy - radius, cx + radius, cy + radius], fill=rgb(AMBER if index % 2 == 0 else TINT))
        if index < steps - 1:
            ax0, ax1, ay = x0 + box_w + gap * 0.2, x0 + box_w + gap * 0.8, height / 2
            draw.line([(ax0, ay), (ax1, ay)], fill=rgb(MUTED), width=max(4, height // 90))
            head = gap * 0.16
            draw.polygon([(ax1, ay), (ax1 - head, ay - head * 0.7), (ax1 - head, ay + head * 0.7)], fill=rgb(MUTED))
    return png(image)


def logo(color, mark):
    """The demo brand mark: two stacked slides."""
    size = 512
    image = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle([40, 120, 360, 350], radius=44, fill=rgb(mark) + (255,))
    draw.rounded_rectangle([140, 190, 472, 430], radius=44, fill=rgb(color) + (255,), outline=rgb(WHITE if color != WHITE else INK) + (255,), width=14)
    return png(image.resize((256, 256), Image.LANCZOS))


def icon(name, color=TEAL):
    """One pictogram of the demo icon set, drawn large and scaled down for smooth edges."""
    s = 1024
    image = Image.new('RGBA', (s, s), (0, 0, 0, 0))
    d = ImageDraw.Draw(image)
    c = rgb(color) + (255,)
    clear = (0, 0, 0, 0)
    w = 84  # stroke width

    def ring(box, width=w):
        d.ellipse(box, outline=c, width=width)

    if name == 'target':
        ring([112, 112, 912, 912])
        ring([292, 292, 732, 732])
        d.ellipse([432, 432, 592, 592], fill=c)
    elif name == 'check':
        ring([112, 112, 912, 912])
        d.line([(320, 530), (460, 670), (720, 380)], fill=c, width=w, joint='curve')
    elif name == 'bolt':
        d.polygon([(580, 80), (240, 560), (480, 560), (420, 944), (800, 440), (550, 440)], fill=c)
    elif name == 'heart':
        d.ellipse([130, 190, 530, 590], fill=c)
        d.ellipse([494, 190, 894, 590], fill=c)
        d.polygon([(160, 480), (864, 480), (512, 900)], fill=c)
    elif name == 'star':
        points = []
        for index in range(10):
            radius = 430 if index % 2 == 0 else 180
            angle = -math.pi / 2 + index * math.pi / 5
            points.append((512 + radius * math.cos(angle), 540 + radius * math.sin(angle)))
        d.polygon(points, fill=c)
    elif name == 'chart':
        for index, height in enumerate([320, 520, 420, 720]):
            x0 = 150 + index * 190
            d.rounded_rectangle([x0, 880 - height, x0 + 130, 880], radius=24, fill=c)
    elif name == 'clock':
        ring([112, 112, 912, 912])
        d.line([(512, 300), (512, 512), (670, 600)], fill=c, width=w, joint='curve')
    elif name == 'people':
        d.ellipse([170, 200, 450, 480], fill=c)
        d.rounded_rectangle([90, 540, 530, 900], radius=150, fill=c)
        d.ellipse([590, 260, 830, 500], fill=c)
        d.rounded_rectangle([560, 560, 940, 900], radius=130, fill=c)
    elif name == 'gear':
        for index in range(8):
            angle = index * math.pi / 4
            cx, cy = 512 + 360 * math.cos(angle), 512 + 360 * math.sin(angle)
            d.ellipse([cx - 90, cy - 90, cx + 90, cy + 90], fill=c)
        d.ellipse([172, 172, 852, 852], fill=c)
        d.ellipse([372, 372, 652, 652], fill=clear)
    elif name == 'leaf':
        d.pieslice([120, 120, 1320, 1320], 180, 270, fill=c)
        d.pieslice([-296, -296, 904, 904], 0, 90, fill=c)
        d.line([(190, 850), (640, 400)], fill=clear, width=40)
    elif name == 'globe':
        ring([112, 112, 912, 912])
        ring([332, 112, 692, 912], width=60)
        d.line([(130, 512), (894, 512)], fill=c, width=60)
    elif name == 'shield':
        d.polygon([(512, 90), (860, 220), (860, 520), (512, 940), (164, 520), (164, 220)], fill=c)
        d.line([(370, 500), (480, 610), (670, 400)], fill=clear, width=80, joint='curve')
    else:
        raise ValueError(name)
    return png(image.resize((256, 256), Image.LANCZOS))


ICONS = {
    'target': 'a target: goal, focus, aim',
    'check': 'a check mark in a circle: done, approved, quality',
    'bolt': 'a lightning bolt: energy, speed, power',
    'heart': 'a heart: care, health, loyalty',
    'star': 'a star: favourite, rating, excellence',
    'chart': 'a rising bar chart: growth, results, data',
    'clock': 'a clock: time, schedule, duration',
    'people': 'two people: team, community, customers',
    'gear': 'a gear: process, settings, operations',
    'leaf': 'a leaf: nature, sustainability, growth',
    'globe': 'a globe: world, international, online',
    'shield': 'a shield with a check: safety, security, trust',
}

# --- slide building --------------------------------------------------------------------------------

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]
NUMBER = next(shape for shape in BLANK.placeholders if shape.placeholder_format.type is not None and 'SLIDE_NUMBER' in str(shape.placeholder_format.type))


def box(slide, x, y, w, h, name):
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    shape.name = name
    frame = shape.text_frame
    frame.word_wrap = True
    frame.margin_left = frame.margin_right = frame.margin_top = frame.margin_bottom = 0
    return shape


def write(shape, paragraphs, size, color, font=BODY, bold=False, italic=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, bullets=False, gap=0):
    """Fills a shape with paragraphs that share one style."""
    frame = shape.text_frame
    frame.vertical_anchor = anchor
    for index, text in enumerate(paragraphs if isinstance(paragraphs, list) else [paragraphs]):
        paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        paragraph.alignment = align
        if gap:
            paragraph.space_after = Pt(gap)
        run = paragraph.add_run()
        run.text = text
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.italic = italic
        run.font.name = font
        run.font.color.rgb = RGBColor.from_string(color)
        # The complex-script font, so that Hebrew or Arabic text is not left to a fallback.
        cs = etree.SubElement(run._r.get_or_add_rPr(), qn('a:cs'))
        cs.set('typeface', BODY)
        if bullets:
            properties = paragraph._p.get_or_add_pPr()
            properties.set('marL', str(Inches(0.3)))
            properties.set('indent', str(-Inches(0.3)))
            color_el = etree.SubElement(properties, qn('a:buClr'))
            etree.SubElement(color_el, qn('a:srgbClr')).set('val', TEAL)
            etree.SubElement(properties, qn('a:buFont')).set('typeface', 'Arial')
            etree.SubElement(properties, qn('a:buChar')).set('char', '•')
    return shape


def text(slide, x, y, w, h, name, paragraphs, size, color, **style):
    return write(box(slide, x, y, w, h, name), paragraphs, size, color, **style)


def block(shapes, kind, x, y, w, h, fill, name, radius=None):
    shape = shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.name = name
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor.from_string(fill)
    shape.line.fill.background()
    shape.shadow.inherit = False
    if radius is not None and kind == MSO_SHAPE.ROUNDED_RECTANGLE:
        shape.adjustments[0] = radius
    return shape


def picture(shapes, data, x, y, w, h, name, rounded=None):
    data.seek(0)
    shape = shapes.add_picture(data, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.name = name
    if rounded is not None:
        shape.auto_shape_type = MSO_SHAPE.ROUNDED_RECTANGLE
        geometry = shape._element.spPr.find(qn('a:prstGeom'))
        guides = geometry.find(qn('a:avLst'))
        if guides is None:
            guides = etree.SubElement(geometry, qn('a:avLst'))
        guide = etree.SubElement(guides, qn('a:gd'))
        guide.set('name', 'adj')
        guide.set('fmla', f'val {round(rounded * 100000)}')
    return shape


def group_text(group, x, y, w, h, name):
    """A text box inside a group, set up like the free-standing ones."""
    shape = group.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    shape.name = name
    frame = shape.text_frame
    frame.word_wrap = True
    frame.margin_left = frame.margin_right = frame.margin_top = frame.margin_bottom = 0
    return shape


def background(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor.from_string(color)


def frame(slide, number=True):
    """What every content slide carries: the mark, the organization line and the page number."""
    picture(slide.shapes, LOGO, 0.5, 0.36, 0.44, 0.44, 'Logo')
    text(slide, 1.08, 0.44, 6.5, 0.3, 'Organization', ORG, 12, MUTED)
    if number:
        placeholder = slide.shapes.clone_placeholder(NUMBER)
        shape = slide.shapes[-1]
        shape.name = 'Page number'
        shape.left, shape.top, shape.width, shape.height = Inches(12.03), Inches(0.44), Inches(0.8), Inches(0.3)
        body = shape.text_frame
        body.margin_left = body.margin_right = body.margin_top = body.margin_bottom = 0
        paragraph = body.paragraphs[0]
        paragraph.alignment = PP_ALIGN.RIGHT
        field = etree.SubElement(paragraph._p, qn('a:fld'))
        field.set('id', '{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}')
        field.set('type', 'slidenum')
        properties = etree.SubElement(field, qn('a:rPr'))
        properties.set('lang', 'en-US')
        properties.set('sz', '1200')
        fill = etree.SubElement(etree.SubElement(properties, qn('a:solidFill')), qn('a:srgbClr'))
        fill.set('val', MUTED)
        etree.SubElement(field, qn('a:t')).text = str(len(prs.slides))


def title(slide, value, y=1.05, w=12.3, size=34, x=0.5, h=0.75):
    return text(slide, x, y, w, h, 'Title', value, size, INK, font=HEAD, bold=True)


def new_slide():
    return prs.slides.add_slide(BLANK)


LOGO = logo(TEAL, AMBER)
LOGO_LIGHT = logo(WHITE, AMBER)
PHOTO_A = scene(1400, 1500, ('F9D9A6', 'F4A259'), ('2F6F6E', '0E7C7B', '1B2A41'), 'FFF3D6')
PHOTO_B = scene(1500, 1450, ('BFE3E0', 'E8F1F0'), ('6FB1AE', '0E7C7B', '1B2A41'), 'F4A259')
PHOTO_C = scene(1500, 1250, ('31456A', '8FB8C9'), ('3C5A7A', '1B2A41', '0F1826'), 'F9D9A6')
CHART = bar_chart(1500, 980, [0.34, 0.46, 0.4, 0.58, 0.66, 0.6, 0.78, 0.92, 0.84])
DIAGRAM = flow_diagram(2400, 760)
SAMPLE_ICON = icon('star')

LOREM = 'Body text sits here in two or three sentences. It explains the point of the slide in plain words, and leaves the detail for the speaker.'
LOREM_SHORT = 'A sentence or two that says what this item means for the audience.'

# 1. Cover
slide = new_slide()
background(slide, INK)
block(slide.shapes, MSO_SHAPE.OVAL, 9.1, -1.6, 6.4, 6.4, TEAL, 'Circle large')
block(slide.shapes, MSO_SHAPE.OVAL, 10.9, 3.7, 2.9, 2.9, AMBER, 'Circle small')
picture(slide.shapes, LOGO_LIGHT, 0.8, 0.7, 0.6, 0.6, 'Logo')
text(slide, 0.8, 2.3, 8.0, 2.0, 'Title', 'Presentation title on two lines', 50, WHITE, font=HEAD, bold=True, anchor=MSO_ANCHOR.BOTTOM)
text(slide, 0.8, 4.55, 7.6, 0.8, 'Subtitle', 'A short subtitle that frames the story', 20, TINT)
text(slide, 0.8, 6.55, 7.0, 0.32, 'Organization', ORG, 13, TINT)

# 2. Text with a photo
slide = new_slide()
picture(slide.shapes, PHOTO_A, 7.63, 0, 5.703, 7.5, 'Photo')
frame(slide, number=False)
title(slide, 'Slide title for a text and photo layout', y=1.5, w=6.4, h=1.4)
text(slide, 0.5, 3.05, 6.3, 0.4, 'Kicker', 'A supporting line in one sentence', 17, TEAL, bold=True)
text(slide, 0.5, 3.7, 6.2, 2.9, 'Body', [LOREM, 'A second paragraph adds one more idea, an example or the next step.'], 18, MUTED, gap=12)

# 3. A figure with a photo
slide = new_slide()
frame(slide)
picture(slide.shapes, PHOTO_B, 0.5, 1.3, 5.6, 5.5, 'Photo', rounded=0.06)
title(slide, 'Title for a slide built on one figure', x=6.7, y=1.4, w=6.1, size=30, h=1.2)
text(slide, 6.7, 2.85, 6.1, 1.45, 'Figure', '87%', 84, TEAL, font=HEAD, bold=True)
text(slide, 6.7, 4.5, 5.9, 1.9, 'Figure text', 'What the figure measures and why it matters, in one or two sentences that a listener can repeat.', 19, MUTED)

# 4. Bullets with a photo
slide = new_slide()
frame(slide)
title(slide, 'Title for a list of points')
text(slide, 0.5, 2.25, 6.5, 4.3, 'Points', ['The first point, kept to one line or two', 'The second point follows the same form', 'A third point for the middle of the story', 'The last point is the one to remember'], 22, INK, bullets=True, gap=18)
picture(slide.shapes, PHOTO_C, 7.6, 2.15, 5.23, 4.45, 'Photo', rounded=0.06)

# 5. A row of icon cards
slide = new_slide()
frame(slide)
title(slide, 'Title for a row of cards')
for index in range(4):
    x = 0.5 + index * 3.15
    group = slide.shapes.add_group_shape()
    group.name = f'Card {index + 1}'
    block(group.shapes, MSO_SHAPE.ROUNDED_RECTANGLE, x, 2.3, 2.85, 4.3, TINT, 'Card background', radius=0.07)
    block(group.shapes, MSO_SHAPE.OVAL, x + 0.3, 2.65, 0.95, 0.95, WHITE, 'Icon circle')
    picture(group.shapes, SAMPLE_ICON, x + 0.525, 2.875, 0.5, 0.5, 'Icon')
    write(group_text(group, x + 0.3, 3.9, 2.25, 0.45, 'Card heading'), 'Card heading', 19, INK, font=HEAD, bold=True)
    write(group_text(group, x + 0.3, 4.45, 2.25, 1.9, 'Card text'), LOREM_SHORT, 15, MUTED)

# 6. A grid of figures
slide = new_slide()
frame(slide)
title(slide, 'Title for a grid of figures')
SAMPLE_FIGURES = ['42', '3.5x', '120', '98%', '15', '7 days']
for index in range(6):
    x = 0.5 + (index % 3) * 4.2
    y = 2.25 + (index // 3) * 2.25
    group = slide.shapes.add_group_shape()
    group.name = f'Figure card {index + 1}'
    block(group.shapes, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, 3.9, 1.95, TINT, 'Card background', radius=0.1)
    write(group_text(group, x + 0.35, y + 0.25, 3.2, 0.85, 'Figure'), SAMPLE_FIGURES[index], 40, TEAL, font=HEAD, bold=True)
    write(group_text(group, x + 0.35, y + 1.15, 3.2, 0.55, 'Figure label'), 'What this figure counts', 16, MUTED)

# 7. Two columns
slide = new_slide()
frame(slide)
title(slide, 'Title for two things side by side')
for index, heading in enumerate(['First column heading', 'Second column heading']):
    x = 0.5 + index * 6.4
    block(slide.shapes, MSO_SHAPE.ROUNDED_RECTANGLE, x, 2.25, 5.9, 4.35, TINT, f'Column {index + 1} background', radius=0.05)
    text(slide, x + 0.4, 2.6, 5.1, 0.5, f'Column {index + 1} heading', heading, 22, TEAL, font=HEAD, bold=True)
    text(slide, x + 0.4, 3.3, 5.1, 3.0, f'Column {index + 1} points', ['A point about this side of the comparison', 'A second point in the same form', 'A third point to close the column'], 18, INK, bullets=True, gap=12)

# 8. A table
slide = new_slide()
frame(slide)
title(slide, 'Title for a table')
rows, columns = 5, 4
table_shape = slide.shapes.add_table(rows, columns, Inches(0.5), Inches(2.25), Inches(12.33), Inches(3.4))
table_shape.name = 'Table'
table = table_shape.table
for r in range(rows):
    table.rows[r].height = Inches(0.68)
    for c in range(columns):
        cell = table.cell(r, c)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor.from_string(INK if r == 0 else (TINT if r % 2 == 0 else WHITE))
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.margin_left = cell.margin_right = Inches(0.15)
        paragraph = cell.text_frame.paragraphs[0]
        run = paragraph.add_run()
        run.text = ('Column heading' if r == 0 else ('Row label' if c == 0 else 'Cell text'))
        run.font.size = Pt(16 if r == 0 else 15)
        run.font.bold = r == 0
        run.font.name = BODY
        run.font.color.rgb = RGBColor.from_string(WHITE if r == 0 else INK)
text(slide, 0.5, 5.95, 12.3, 0.4, 'Table note', 'A note under the table: the source, the period or what the numbers leave out.', 13, MUTED, italic=True)

# 9. A chart with text
slide = new_slide()
frame(slide)
title(slide, 'Title for a chart and what it shows')
picture(slide.shapes, CHART, 0.5, 2.2, 7.0, 4.57, 'Chart')
text(slide, 8.0, 2.25, 4.8, 2.6, 'Chart text', 'What the chart shows, in two or three sentences. Name the trend and the one bar worth looking at.', 18, MUTED)
callout = block(slide.shapes, MSO_SHAPE.ROUNDED_RECTANGLE, 8.0, 5.2, 4.83, 1.4, 'FCE9D2', 'Takeaway', radius=0.12)
callout.text_frame.word_wrap = True
callout.text_frame.margin_left = callout.text_frame.margin_right = Inches(0.25)
write(callout, 'The one conclusion to take away from this chart', 17, INK, bold=True, anchor=MSO_ANCHOR.MIDDLE)

# 10. A wide picture
slide = new_slide()
frame(slide)
title(slide, 'Title for a wide diagram')
picture(slide.shapes, DIAGRAM, 0.5, 2.2, 12.33, 3.9, 'Diagram')
text(slide, 0.5, 6.3, 12.3, 0.4, 'Caption', 'A caption that says how to read the diagram, in one line.', 14, MUTED, italic=True)

# 11. A quotation
slide = new_slide()
background(slide, TINT)
frame(slide)
block(slide.shapes, MSO_SHAPE.OVAL, 0.9, 2.0, 0.55, 0.55, AMBER, 'Dot')
text(slide, 1.9, 1.9, 10.2, 3.0, 'Quote', 'A sentence worth a slide of its own: a quotation, a principle, or the one thing the audience should carry out of the room.', 32, INK, font=HEAD, italic=True)
text(slide, 1.9, 5.25, 10.0, 0.45, 'Source', 'Name Surname, role or source', 17, TEAL, bold=True)

# 12. A section opener
slide = new_slide()
background(slide, INK)
block(slide.shapes, MSO_SHAPE.OVAL, 10.2, 3.6, 5.2, 5.2, TEAL, 'Circle')
picture(slide.shapes, LOGO_LIGHT, 0.8, 0.7, 0.6, 0.6, 'Logo')
text(slide, 0.8, 2.7, 9.0, 1.2, 'Section title', 'Section title', 50, WHITE, font=HEAD, bold=True, anchor=MSO_ANCHOR.BOTTOM)
text(slide, 0.8, 4.15, 8.4, 0.9, 'Section text', 'One line on what this part of the deck covers', 20, TINT)

# 13. Closing
slide = new_slide()
background(slide, INK)
block(slide.shapes, MSO_SHAPE.OVAL, -1.8, 4.3, 5.0, 5.0, TEAL, 'Circle large')
block(slide.shapes, MSO_SHAPE.OVAL, 11.0, 0.7, 1.5, 1.5, AMBER, 'Circle small')
picture(slide.shapes, LOGO_LIGHT, 0.8, 0.7, 0.6, 0.6, 'Logo')
text(slide, 3.9, 2.6, 8.6, 1.2, 'Closing title', 'Thank you', 56, WHITE, font=HEAD, bold=True, anchor=MSO_ANCHOR.BOTTOM)
text(slide, 3.9, 4.0, 8.4, 0.9, 'Closing text', 'Contact details or the next step', 20, TINT)
text(slide, 3.9, 6.55, 7.0, 0.32, 'Organization', ORG, 13, TINT)

prs.core_properties.title = 'slotdeck demo template'
prs.core_properties.author = 'slotdeck'
prs.core_properties.comments = 'Drawn by scripts/build-demo-template.py. Free to reuse.'
prs.core_properties.last_modified_by = 'slotdeck'
prs.core_properties.created = prs.core_properties.modified = datetime.datetime(2026, 11, 13)

os.makedirs(os.path.join(ROOT, 'templates'), exist_ok=True)
buffer = io.BytesIO()
prs.save(buffer)
# The library's starting file carries a preview picture of an empty deck; leave it out.
with zipfile.ZipFile(buffer) as source, zipfile.ZipFile(os.path.join(ROOT, 'templates', 'demo.pptx'), 'w', zipfile.ZIP_DEFLATED) as target:
    for item in source.infolist():
        if item.filename == 'docProps/thumbnail.jpeg':
            continue
        data = source.read(item.filename)
        if item.filename == '_rels/.rels':
            data = re.sub(rb'<Relationship [^>]*thumbnail[^>]*/>', b'', data)
        target.writestr(item, data)
print(f'templates/demo.pptx: {len(prs.slides)} slides')

os.makedirs(os.path.join(ROOT, 'examples', 'icons'), exist_ok=True)
for name in ICONS:
    with open(os.path.join(ROOT, 'examples', 'icons', f'{name}.png'), 'wb') as file:
        file.write(icon(name).read())
with open(os.path.join(ROOT, 'examples', 'icons', 'icons.json'), 'w', encoding='utf-8') as file:
    import json

    json.dump([{'id': name, 'description': description} for name, description in ICONS.items()], file, indent=2)
print(f'examples/icons: {len(ICONS)} icons')

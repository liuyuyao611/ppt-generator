"""One scene representation drives SVG previews and editable native PPTX shapes."""
from __future__ import annotations

import base64
import html
import io
import re

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pptx.oxml.xmlchemy import OxmlElement

from layout_engine import build_scenes, wrap


def scene_svg(scene):
    """Browser SVG supports tspan; optional SVG download requires explicit opt-in."""
    escape = html.escape
    parts = [f'<svg xmlns="http://www.w3.org/2000/svg" width="{scene["width"]}" height="{scene["height"]}" viewBox="0 0 {scene["width"]} {scene["height"]}">',
             f'<rect width="100%" height="100%" fill="{escape(scene["bg"])}"/>']
    if scene.get("backdrop"):
        parts.append(f'<image width="100%" height="100%" href="{escape(scene["backdrop"], quote=True)}"/>')
    for e in scene["elements"]:
        if e["kind"] == "rect":
            stroke = f' stroke="{escape(e["stroke"])}" stroke-width="1"' if e.get("stroke") else ""
            dash = ' stroke-dasharray="4 3"' if e.get("dash") else ""
            parts.append(f'<rect id="{e["id"]}" x="{e["x"]}" y="{e["y"]}" width="{e["w"]}" height="{e["h"]}" rx="{12 if e.get("radius") else 0}" fill="{escape(e["fill"])}"{stroke}{dash}/>')
        elif e["kind"] == "chevron":
            x,y,w,h=e["x"],e["y"],e["w"],e["h"]
            points=f"{x},{y} {x+w*.55},{y} {x+w},{y+h/2} {x+w*.55},{y+h} {x},{y+h} {x+w*.45},{y+h/2}"
            parts.append(f'<polygon id="{e["id"]}" points="{points}" fill="{escape(e["fill"])}"/>')
        elif e["kind"] == "text":
            family = escape(e.get('font',scene["font"]), quote=True)
            align=e.get("align","left")
            x=e["x"]+(e["w"]/2 if align=="center" else e["w"] if align=="right" else 0)
            anchor={"left":"start","center":"middle","right":"end"}[align]
            parts.append(f'<text id="{e["id"]}" x="{x}" y="{e["y"]+e["size"]}" font-family="{family}, Noto Sans CJK SC, sans-serif" font-size="{e["size"]}" font-weight="{700 if e["bold"] else 400}" text-anchor="{anchor}" fill="{escape(e["color"])}">')
            lines=e["text"].split("\n")
            if e.get("runs"):
                for j, runs in enumerate(e["runs"]):
                    baseline = e["y"] + e["size"] + j * e["size"] * e["line_height"]
                    parts.append(f'<tspan x="{x}" y="{baseline}">')
                    for run in runs:
                        parts.append(f'<tspan font-weight="{700 if run["bold"] else 400}" fill="{escape(run["color"])}">'
                                     f'{escape(run["text"])}</tspan>')
                    parts.append("</tspan>")
            elif len(lines)==1:
                parts.append(escape(lines[0]))
            else:
                for j, line in enumerate(lines):
                    parts.append(f'<tspan x="{x}" y="{e["y"]+e["size"]+j*e["size"]*e["line_height"]}">{escape(line)}</tspan>')
            parts.append("</text>")
        else:
            b64 = base64.b64encode(e["data"]).decode("ascii")
            parts.append(f'<image id="{e["id"]}" x="{e["x"]}" y="{e["y"]}" width="{e["w"]}" height="{e["h"]}" href="data:image/jpeg;base64,{b64}"/>')
    parts.append("</svg>")
    return "".join(parts).encode("utf-8")


def export_pptx(scenes, title="文档汇报", *, template=None):
    if template is None:
        prs = Presentation()
        prs.slide_width = Inches(13.333333)
        prs.slide_height = Inches(13.333333 * scenes[0]["height"] / scenes[0]['width'])
        cover_layout = body_layout = prs.slide_layouts[6]
    else:
        prs = Presentation(io.BytesIO(template.data))
        cover_layout, body_layout = (prs.slides[i].slide_layout for i in range(2))
        # The examples identify native layouts. Removing the slide relationships keeps
        # layout/master parts and avoids carrying example content and speaker notes.
        ids = prs.slides._sldIdLst
        for reference in list(ids):
            prs.part.drop_rel(reference.rId)
            ids.remove(reference)
    prs.core_properties.title = title
    prs.core_properties.author = "SlideCraft"
    units_per_inch=scenes[0]['width']/prs.slide_width.inches
    points_per_unit=72/units_per_inch
    def inch(value):
        return Inches(value / units_per_inch)

    for scene in scenes:
        slide = prs.slides.add_slide(cover_layout if scene["layout"] == "cover" else body_layout)
        if template is None:
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = RGBColor.from_string(scene["bg"].lstrip("#"))
        else:
            for placeholder in list(slide.placeholders):
                placeholder._element.getparent().remove(placeholder._element)
        for e in scene["elements"]:
            if e["kind"] == "image":
                shape = slide.shapes.add_picture(io.BytesIO(e["data"]), inch(e["x"]), inch(e["y"]), inch(e["w"]), inch(e["h"]))
            elif e["kind"] in {"rect", "chevron"}:
                kind = MSO_SHAPE.CHEVRON if e["kind"] == "chevron" else MSO_SHAPE.ROUNDED_RECTANGLE if e.get("radius") else MSO_SHAPE.RECTANGLE
                shape = slide.shapes.add_shape(kind, inch(e["x"]), inch(e["y"]), inch(e["w"]), inch(e["h"]))
                shape.fill.solid()
                shape.fill.fore_color.rgb = RGBColor.from_string(e["fill"].lstrip("#"))
                if e.get("stroke"):
                    from pptx.enum.dml import MSO_LINE_DASH_STYLE
                    shape.line.color.rgb = RGBColor.from_string(e["stroke"].lstrip("#"))
                    shape.line.width = Pt(.75)
                    if e.get("dash"):
                        shape.line.dash_style = MSO_LINE_DASH_STYLE.DASH
                else:
                    shape.line.fill.background()
                if e.get("radius"):
                    shape.adjustments[0] = .12
            else:
                shape = slide.shapes.add_textbox(inch(e["x"]), inch(e["y"]), inch(e["w"]), inch(e["h"]))
                tf = shape.text_frame
                tf.clear()
                tf.word_wrap = False  # Explicit lines match the preview; no hidden font substitution layout.
                tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
                for j, line in enumerate(e["text"].split("\n")):
                    p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
                    p.space_before = Pt(0)
                    p.space_after = Pt(0)
                    p.line_spacing = Pt(e["size"] * e["line_height"] * points_per_unit)
                    p.alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}[e["align"]]
                    if e.get("runs"):
                        for span in e["runs"][j]:
                            run = p.add_run()
                            run.text = span["text"]
                            run.font.bold = span["bold"]
                            run.font.color.rgb = RGBColor.from_string(span["color"].lstrip("#"))
                    else:
                        p.text = line
                    p.font.name = e.get('font',scene["font"])
                    p.font.size = Pt(e["size"] * points_per_unit)
                    p.font.bold = e["bold"]
                    p.font.color.rgb = RGBColor.from_string(e["color"].lstrip("#"))
                    # Set East Asian font explicitly for Chinese text in Office.
                    props = p._p.get_or_add_pPr()
                    default = props.find("{http://schemas.openxmlformats.org/drawingml/2006/main}defRPr")
                    if default is not None:
                        ea = OxmlElement("a:ea")
                        ea.set("typeface", 'Noto Sans SC' if scene.get('frontend_theme') else scene["font"])
                        default.append(ea)
            shape.name = e["id"]
        slide.notes_slide.notes_text_frame.text = scene.get("notes", "")
    output = io.BytesIO()
    prs.save(output)
    return output.getvalue()


def safe_filename(title):
    return (re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", title).strip(" .")[:80] or "文档汇报") + ".pptx"

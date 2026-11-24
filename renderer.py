"""One scene representation drives SVG previews and editable native PPTX shapes."""
from __future__ import annotations

import base64
import copy
import html
import io
import math
import re
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt
from pptx.oxml.xmlchemy import OxmlElement

from core import UserError, normalize_style


def text_width(text, size):
    return sum(size * (1 if ord(c) > 255 else 0.56) for c in text)


def wrap(text, width, size):
    """Conservative line breaks shared by browser preview and PowerPoint."""
    lines = []
    for paragraph in str(text).split("\n"):
        line = ""
        tokens = re.findall(r"[A-Za-z0-9_]+(?:[./@:+%-][A-Za-z0-9_]+)*|[^\x00-\xff]|\s+|.", paragraph)
        for token in tokens:
            if line and text_width(line + token, size) > width:
                lines.append(line.rstrip())
                line = ""
            for char in token:
                if line and text_width(line + char, size) > width:
                    lines.append(line.rstrip())
                    line = ""
                line += char
        lines.append(line.rstrip())
    return lines or [""]


def title_metrics(title, width):
    for size in (42, 38, 34, 30):
        lines = wrap(title, width, size)
        if len(lines) <= 3:
            break
    return lines, size, max(130, 80 + len(lines) * size * 1.22)


def available_width(layout, has_image):
    if layout == "image" and has_image:
        return 588
    if layout == "columns":
        return 530
    return 1110


def paginate(outline, style, images):
    """Keep all user-edited bullets, flowing long content onto continuation pages."""
    h = 960 if style["ratio"] == "4:3" else 720
    pages = []
    for slide in outline["slides"]:
        slide = copy.deepcopy(slide)
        image_exists = bool(slide.get("image")) and slide["image"] in images
        if slide.get("layout") not in {"bullets", "columns", "image"}:
            slide["layout"] = "bullets"
        if slide["layout"] == "image" and not image_exists:
            slide["layout"] = "bullets"
        width = available_width(slide["layout"], image_exists)
        # Reserve space for "(续)" and footer before assigning chunks.
        _, _, top = title_metrics(slide["title"] + "（续）", 1120)
        bottom = h - 88
        line_height, size, gap = 38, 26, 24
        lines_per_column = max(3, int((bottom - top - gap) // line_height))
        max_columns = 2 if slide["layout"] == "columns" else 1
        chunks, current, used, column = [], [], 0, 0
        for bullet in slide["bullets"]:
            lines = wrap(bullet, width, size)
            while lines:
                room = lines_per_column - used
                if room < 2:
                    column += 1
                    used = 0
                    if column >= max_columns:
                        chunks.append(current)
                        current, column = [], 0
                    room = lines_per_column
                take = max(1, min(len(lines), room - 1))
                current.append({"text": "\n".join(lines[:take]), "column": column})
                used += take + 1
                lines = lines[take:]
                if lines:
                    column += 1
                    used = 0
                    if column >= max_columns:
                        chunks.append(current)
                        current, column = [], 0
        if current:
            chunks.append(current)
        for i, chunk in enumerate(chunks):
            page = copy.deepcopy(slide)
            page["title"] += "（续）" if i else ""
            page["items"] = chunk
            pages.append(page)
    if len(pages) > 100:
        raise UserError("排版后超过 100 页，请缩短要点或拆分汇报。")
    return pages


def build_scenes(outline, style, images=None, sources=None):
    style = normalize_style(style)
    images, sources = images or {}, sources or {}
    h = 960 if style["ratio"] == "4:3" else 720
    pages = paginate(outline, style, images)
    scenes = []

    def new_scene(bg):
        return {"width": 1280, "height": h, "bg": bg, "font": style["font"], "elements": []}

    def rect(scene, x, y, w, height, fill, radius=False):
        scene["elements"].append(dict(kind="rect", x=x, y=y, w=w, h=height, fill=fill, radius=radius))

    def text(scene, content, x, y, w, size, color, bold=False):
        lines = str(content).split("\n")
        scene["elements"].append(dict(kind="text", text=str(content), x=x, y=y, w=w,
            h=max(1, len(lines)) * size * 1.4 + 5, size=size, color=color, bold=bold))

    cover = new_scene(style["primary"])
    rect(cover, 72, 72, 72, 8, style["accent"])
    text(cover, "DOCUMENT TO PRESENTATION", 72, 104, 1100, 18, "#FFFFFF")
    subtitle = "\n".join(wrap(outline.get("subtitle", ""), 1040, 26))
    subtitle_h = len(subtitle.split("\n")) * 26 * 1.4 + 5
    for size in (64, 58, 50, 44, 38, 34, 30):
        lines = wrap(outline["title"], 1090, size)
        title_h = len(lines) * size * 1.4
        if title_h + subtitle_h + 32 <= h * .75 - 110:
            break
    text(cover, "\n".join(lines), 72, h * .25, 1136, size, "#FFFFFF", True)
    title_bottom = h * .25 + title_h
    text(cover, subtitle, 76, max(h * .61, title_bottom + 28), 1100, 26, "#FFFFFF")
    text(cover, f"{len(pages)+1:02d} SLIDES  /  {outline.get('mode', '文档汇报')}", 76, h-62, 1100, 16, "#FFFFFF")
    cover["notes"] = "封面\n" + outline.get("subtitle", "")
    cover["title"] = outline["title"]
    scenes.append(cover)

    for index, page in enumerate(pages, 2):
        scene = new_scene(style["background"])
        scene["title"] = page["title"]
        rect(scene, 0, 0, 1280, 7, style["primary"])
        text(scene, f"{index-1:02d} / {len(pages):02d}", 74, 28, 1100, 16, style["primary"], True)
        lines, size, top = title_metrics(page["title"], 1120)
        text(scene, "\n".join(lines), 72, 62, 1136, size, style["text"], True)
        top += 22
        layout = page.get("layout", "bullets")
        width = available_width(layout, bool(page.get("image")))
        ys = {0: top, 1: top}
        for item in page["items"]:
            column = item["column"]
            x = 76 + (610 if column else 0)
            if layout != "columns":
                x = 76
            y = ys[column]
            rect(scene, x, y+10, 6, 14, style["accent"])
            text(scene, item["text"], x+26, y, width, 26, style["text"])
            ys[column] += len(item["text"].split("\n")) * 38 + 24
        if layout == "image" and page.get("image") in images:
            asset = images[page["image"]]
            box_x, box_y, box_w, box_h = 744, top, 460, h-top-108
            rect(scene, box_x-12, box_y-12, box_w+24, box_h+24, "#FFFFFF", True)
            scale = min(box_w / asset["width"], box_h / asset["height"])
            w, ih = asset["width"]*scale, asset["height"]*scale
            scene["elements"].append(dict(kind="image", x=box_x+(box_w-w)/2, y=box_y+(box_h-ih)/2,
                                          w=w, h=ih, data=asset["data"]))
        rect(scene, 74, h-66, 1132, 1, "#DCE2EF")
        ids = page.get("source_ids", [])
        label = "来源：" + " · ".join(ids[:10]) + (f" 等 {len(ids)} 处" if len(ids)>10 else "")
        text(scene, label, 76, h-48, 1030, 14, style["muted"])
        text(scene, str(index), 1162, h-48, 50, 14, style["muted"])
        references = "\n".join(f"[{sid}] {sources.get(sid, sid)}" for sid in ids)
        scene["notes"] = page.get("notes", "") + "\n\n【原文定位】\n" + references
        scenes.append(scene)
    # Geometry check also protects SVG preview and PowerPoint from off-page objects.
    for scene in scenes:
        for element in scene["elements"]:
            if element["y"] + element["h"] > scene["height"] + 2:
                raise UserError("标题、副标题或要点过长，超出页面范围，请缩短后重试。")
    return scenes


def scene_svg(scene):
    escape = html.escape
    parts = [f'<svg xmlns="http://www.w3.org/2000/svg" width="{scene["width"]}" height="{scene["height"]}" viewBox="0 0 {scene["width"]} {scene["height"]}">',
             f'<rect width="100%" height="100%" fill="{escape(scene["bg"])}"/>']
    for e in scene["elements"]:
        if e["kind"] == "rect":
            parts.append(f'<rect x="{e["x"]}" y="{e["y"]}" width="{e["w"]}" height="{e["h"]}" rx="{12 if e.get("radius") else 0}" fill="{escape(e["fill"])}"/>')
        elif e["kind"] == "text":
            family = escape(scene["font"], quote=True)
            parts.append(f'<text font-family="{family}, Noto Sans CJK SC, sans-serif" font-size="{e["size"]}" font-weight="{700 if e["bold"] else 400}" fill="{escape(e["color"])}">')
            for j, line in enumerate(e["text"].split("\n")):
                parts.append(f'<tspan x="{e["x"]}" y="{e["y"]+e["size"]+j*e["size"]*1.4}">{escape(line)}</tspan>')
            parts.append("</text>")
        else:
            b64 = base64.b64encode(e["data"]).decode("ascii")
            parts.append(f'<image x="{e["x"]}" y="{e["y"]}" width="{e["w"]}" height="{e["h"]}" href="data:image/jpeg;base64,{b64}"/>')
    parts.append("</svg>")
    return "".join(parts).encode("utf-8")


def export_pptx(scenes, title="文档汇报"):
    prs = Presentation()
    prs.slide_width = Inches(13.333333)
    prs.slide_height = Inches(13.333333 * scenes[0]["height"] / 1280)
    prs.core_properties.title = title
    prs.core_properties.author = "SlideCraft"
    def inch(value):
        return Inches(value / 96)

    for scene in scenes:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor.from_string(scene["bg"].lstrip("#"))
        for e in scene["elements"]:
            if e["kind"] == "image":
                slide.shapes.add_picture(io.BytesIO(e["data"]), inch(e["x"]), inch(e["y"]), inch(e["w"]), inch(e["h"]))
            elif e["kind"] == "rect":
                kind = MSO_SHAPE.ROUNDED_RECTANGLE if e.get("radius") else MSO_SHAPE.RECTANGLE
                shape = slide.shapes.add_shape(kind, inch(e["x"]), inch(e["y"]), inch(e["w"]), inch(e["h"]))
                shape.fill.solid()
                shape.fill.fore_color.rgb = RGBColor.from_string(e["fill"].lstrip("#"))
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
                    p.line_spacing = Pt(e["size"] * 1.4 * .75)
                    p.text = line
                    p.font.name = scene["font"]
                    p.font.size = Pt(e["size"] * .75)
                    p.font.bold = e["bold"]
                    p.font.color.rgb = RGBColor.from_string(e["color"].lstrip("#"))
                    # Set East Asian font explicitly for Chinese text in Office.
                    props = p._p.get_or_add_pPr()
                    default = props.find("{http://schemas.openxmlformats.org/drawingml/2006/main}defRPr")
                    if default is not None:
                        ea = OxmlElement("a:ea")
                        ea.set("typeface", scene["font"])
                        default.append(ea)
        slide.notes_slide.notes_text_frame.text = scene.get("notes", "")
    output = io.BytesIO()
    prs.save(output)
    return output.getvalue()


def safe_filename(title):
    return (re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", title).strip(" .")[:80] or "文档汇报") + ".pptx"

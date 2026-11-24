"""Open the designated PowerPoint bottom layer and describe its preview."""
from __future__ import annotations

import base64
import io
from dataclasses import dataclass
from xml.etree import ElementTree as ET

from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

from core import MAX_FILE, UserError, check_package


PERSONAL_STYLE = {
    "name": "个人学术汇报", "primary": "#145396", "accent": "#0D7377",
    "background": "#FFFFFF", "text": "#1E293B", "muted": "#44546A",
    "font": "Microsoft YaHei", "ratio": "16:9",
}
WIDTH = 1280


@dataclass(frozen=True)
class PageTemplate:
    backdrop: str
    title_box: tuple
    body_box: tuple
    approximate: bool


@dataclass(frozen=True)
class PersonalTemplate:
    data: bytes
    height: int
    cover: PageTemplate
    body: PageTemplate


def _color(node, colors):
    values = node.xpath("./a:srgbClr | ./a:schemeClr | ./a:sysClr")
    if not values:
        return None, 1
    value = values[0]
    color = value.get("lastClr") or value.get("val", "")
    color = colors.get(color, color)
    if len(color) != 6 or any(c not in "0123456789abcdefABCDEF" for c in color):
        return None, 1
    alpha = value.xpath("./a:alpha/@val")
    return "#" + color, int(alpha[0]) / 100000 if alpha else 1


def _palette(layout):
    theme = layout.slide_master.part.part_related_by(RT.THEME).blob
    ns = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    scheme = ET.fromstring(theme).find(f".//{ns}clrScheme")
    colors = {node.tag.split("}")[-1]: node[0].get("lastClr") or node[0].get("val") for node in scheme}
    for source in layout.slide_master._element.xpath("./p:clrMap"):
        colors.update({role: colors.get(ref, ref) for role, ref in source.attrib.items()})
    return colors


def _page_template(layout, scale, height):
    colors = _palette(layout)
    art, approximate = [], False
    def box(s):
        return tuple(round(getattr(s, attr) * scale, 3) for attr in ("left", "top", "width", "height"))

    title = next((p for p in layout.placeholders if p.placeholder_format.type in
                  {PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE}), None)
    body = next((p for p in layout.placeholders if p.placeholder_format.type == PP_PLACEHOLDER.BODY), None)
    if title is None and body is not None:
        title = next((p for p in layout.placeholders if p.placeholder_format.type == PP_PLACEHOLDER.OBJECT
                      and box(p)[1] < 110 and box(p)[3] < 110), None)
    title_box = box(title) if title is not None else (64, 25, 1080, 60)
    body_box = box(body) if body is not None else (48, 120, 1184, height - 185)
    layers = [layout.shapes] if layout._element.get("showMasterSp") == "0" else [
        layout.slide_master.shapes, layout.shapes]
    for shapes in layers:
        for shape in shapes:
            if shape.is_placeholder:
                continue
            x, y, w, h = box(shape)
            kind = shape.shape_type
            if kind == MSO_SHAPE_TYPE.PICTURE and not any(
                    (shape.crop_left, shape.crop_right, shape.crop_top, shape.crop_bottom, shape.rotation)):
                with Image.open(io.BytesIO(shape.image.blob)) as source:
                    pixels = source.convert("RGBA")
                    pixels.thumbnail((max(1, round(w * 1.5)), max(1, round(h * 1.5))), Image.Resampling.LANCZOS)
                png = io.BytesIO()
                pixels.quantize(colors=160, method=Image.Quantize.FASTOCTREE).save(png, format="PNG", optimize=True)
                uri = base64.b64encode(png.getvalue()).decode("ascii")
                art.append(f'<image x="{x}" y="{y}" width="{w}" height="{h}" href="data:image/png;base64,{uri}"/>')
                continue
            props = shape._element
            if kind == MSO_SHAPE_TYPE.LINE:
                color, opacity = None, 1
                solid = props.xpath("./p:spPr/a:ln/a:solidFill")
                stops = props.xpath("./p:spPr/a:ln/a:gradFill/a:gsLst/a:gs")
                if solid:
                    color, opacity = _color(solid[0], colors)
                if stops:
                    color, opacity = _color(stops[0], colors)
                    approximate = True
                line_width = props.xpath("./p:spPr/a:ln/@w")
                art.append(f'<line x1="{x}" y1="{y}" x2="{x+w}" y2="{y+h}" stroke="{color or PERSONAL_STYLE["primary"]}" '
                           f'stroke-width="{int(line_width[0])*scale if line_width else 1}" opacity="{opacity}"/>')
                continue
            fills = props.xpath("./p:spPr/a:solidFill")
            color, opacity = _color(fills[0], colors) if fills else (None, 1)
            if kind == MSO_SHAPE_TYPE.FREEFORM and color:
                paths = props.xpath("./p:spPr/a:custGeom/a:pathLst/a:path")
                if len(paths) == 1 and all(
                        command.tag.rsplit("}", 1)[-1] in {"moveTo", "lnTo", "close"} for command in paths[0]):
                    path = paths[0]
                    pw, ph = int(path.get("w", "1")), int(path.get("h", "1"))
                    if pw > 0 and ph > 0:
                        points = " ".join(f'{x+int(p.get("x"))*w/pw:.3f},{y+int(p.get("y"))*h/ph:.3f}'
                                          for p in path.xpath("./*/a:pt"))
                        art.append(f'<polygon points="{points}" fill="{color}" opacity="{opacity}"/>')
                        continue
            if color and props.xpath('./p:spPr/a:prstGeom[@prst="rect"]') and not shape.text:
                art.append(f'<rect x="{x}" y="{y}" width="{w}" height="{h}" fill="{color}" opacity="{opacity}"/>')
                continue
            approximate = True
    backdrop = f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {WIDTH} {height}">' + "".join(art) + "</svg>"
    data_uri = "data:image/svg+xml;base64," + base64.b64encode(backdrop.encode()).decode("ascii")
    return PageTemplate(data_uri, title_box, body_box, approximate)


def load_personal_template(data):
    if len(data) > MAX_FILE:
        raise UserError("模板 PPTX 超过 20 MB。请使用保留母版的精简副本。")
    check_package(data)
    try:
        presentation = Presentation(io.BytesIO(data))
        if len(presentation.slides) < 2:
            raise UserError("请提供至少两张示例页。第一页用于取封面版式，第二页用于取正文版式。")
        scale = WIDTH / presentation.slide_width
        height = round(presentation.slide_height * scale)
        if height not in (720, 960):
            raise UserError("个人模板需使用 16:9 或 4:3 的标准页面比例。")
        return PersonalTemplate(data, height,
                                _page_template(presentation.slides[0].slide_layout, scale, height),
                                _page_template(presentation.slides[1].slide_layout, scale, height))
    except UserError:
        raise
    except (ValueError, KeyError, AttributeError, IndexError, OSError, TypeError):
        raise UserError("无法读取模板的母版或版式。请用 PowerPoint 另存为 PPTX 后重试。") from None

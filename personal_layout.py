"""Content-led layouts inside the personal PowerPoint template."""
from __future__ import annotations

import math
import re
from collections import defaultdict

from core import UserError
from content_structure import DIAGRAM_LABELS, comparable_values, diagram_for, diagram_points
from layout_engine import picture, rect, text, tint, wrap
from personal_template import WIDTH


PERSONAL_LAYOUTS = {
    "auto": "自动推荐", "main": "主图与说明", "side": "图文左右",
    "stacked": "图文上下", "comparison": "并排对照", "paired": "图文逐项",
    "process": "顺序步骤", "columns": "并列内容", "bullets": "简洁列表",
}
SPACE_OPTIONS = {"auto": "自动分配", "figure": "图更大", "text": "文字空间更多"}
GAP = 24
BODY_SIZE = 23
CAPTION_SIZE = 16
LINE = BODY_SIZE * 1.35


def page_figures(page, images):
    links = page.get("figures")
    if links is None:
        name = page.get("image")
        links = [{"asset": name, "caption": page.get("image_caption", "")}] if name in images else []
    return [{"caption": "", **link, "data": images[link["asset"]]} for link in links if link["asset"] in images]


def recommend_personal(page, figures):
    bullets = page["bullets"]
    title = page["title"]
    relation = page.get("relationship")
    count = len(figures)
    comparison = relation == "comparison" or (
        bool(re.search(r"对比|比较|差异|对照", title))
        and any(re.search(r"相较|相比|高于|低于|优于|\bvs\.?\b", line, re.I) for line in bullets))
    ordered = relation == "sequence" or (
        bool(re.search(r"步骤|流程|顺序", title))
        and any(re.match(r"首先|第一步", line) for line in bullets)
        and any(re.match(r"随后|然后|其次|最后|第二步", line) for line in bullets))
    if any(link.get("bullet") is not None for link in figures):
        return "paired", "配图已关联要点，按对应关系排放"
    if count:
        if comparison and count >= 2:
            return "comparison", "并排对齐已标明的对照图"
        if count == 1 and (not bullets or (len(bullets) == 1 and len(bullets[0]) <= 65)):
            return "main", "突出主图和简短说明"
        aspect = sum(f["data"]["width"] / f["data"]["height"] for f in figures) / count
        if aspect >= 1.55 or not bullets:
            return "stacked", "保留宽图尺度并把说明排在图组下方"
        return "side", "为图组与分条说明分别分配空间"
    if comparison and len(bullets) >= 2:
        return "comparison", "并排显示明确的比较对象"
    if ordered and len(bullets) >= 2:
        return "process", "按已确认的顺序标出步骤"
    if len(bullets) >= 3 and max(map(len, bullets)) <= 90:
        return "columns", "将较短的并列内容分列"
    return "bullets", "沿一条阅读线排列正文"


def _take_lines(value, width):
    """Keep the emphasis intervals tied to the original text across wraps."""
    pattern = re.compile(r"\*\*(.+?)\*\*|==(.+?)==")
    plain, spans, cursor = "", [], 0
    for match in pattern.finditer(value):
        plain += value[cursor:match.start()]
        token = match[1] or match[2]
        spans.append((len(plain), len(plain) + len(token), bool(match[2])))
        plain += token
        cursor = match.end()
    plain += value[cursor:]
    result, cursor = [], 0
    for line in wrap(plain, width * .97, BODY_SIZE):
        start = plain.find(line, cursor)
        if start == -1:
            result.append([(line, False, False)])
            continue
        end, cursor = start + len(line), start + len(line)
        cuts = [(max(start, left), min(end, right), red) for left, right, red in spans
                if left < end and right > start]
        pieces, offset = [], start
        for left, right, red in cuts:
            if offset < left:
                pieces.append((plain[offset:left], False, False))
            pieces.append((plain[left:right], True, red))
            offset = right
        if offset < end or not pieces:
            pieces.append((plain[offset:end], False, False))
        result.append(pieces)
    return result


def _body(scene, lines, index, x, y, w, style, numbered=False, show_marker=True):
    inset = 36 if show_marker else 5
    if show_marker:
        marker = f"{index + 1:02d}" if numbered else "•"
        text(scene, marker, x, y, 34, BODY_SIZE, style["primary"], numbered)
    value = "\n".join("".join(piece[0] for piece in line) for line in lines)
    element = text(scene, value, x + inset, y, w - inset, BODY_SIZE, style["text"], role="body")
    element.update(bullet_id=index, origin_bullet=index, original_text=value)
    if any(bold for line in lines for _, bold, _ in line):
        element["runs"] = [[{"text": token, "bold": bold, "color": "#C00000" if red else style["text"]}
                            for token, bold, red in line] for line in lines]


def _flow(bullets, box):
    x, y, w, h = box
    limit = int((h - 4) / LINE)
    if bullets and limit < 1:
        raise UserError("正文区没有足够空间。可改用其他版式或缩短标题。")
    chunks, current, used = [], [], 0
    for index, value in bullets:
        lines = _take_lines(value, w - 36)
        while lines:
            remaining = limit - used
            if remaining <= 0:
                chunks.append(current)
                current, used = [], 0
                remaining = limit
            portion, lines = lines[:remaining], lines[remaining:]
            current.append((portion, index, x, y + used * LINE, w))
            used += len(portion) + 1
    if current:
        chunks.append(current)
    return chunks


def _caption(figure):
    source = figure["data"].get("source", "")
    caption = figure.get("caption", "")
    if caption and figure["data"].get("generated") and source not in caption:
        return f"{caption} · {source}"
    return caption or source


def _grid(figures, box, comparison=False):
    x, y, w, h = box
    count = len(figures)
    if not count:
        return []
    # Evaluate how much of every original image is actually visible at each grid width.
    options = []
    widths = (min(count, 2),) if comparison else range(1, min(count, 4) + 1)
    for columns in widths:
        rows = math.ceil(count / columns)
        cw, ch = (w - (columns - 1) * GAP) / columns, (h - (rows - 1) * GAP) / rows
        captions = [wrap(_caption(f), cw, CAPTION_SIZE) if _caption(f) else [] for f in figures]
        heights = [ch - len(lines) * CAPTION_SIZE * 1.35 - (10 if lines else 0) for lines in captions]
        if cw < 100 or min(heights) < 80:
            continue
        visible_area = []
        for f, available_h in zip(figures, heights):
            a = f["data"]["width"] / f["data"]["height"]
            visible_area.append(min(cw, available_h * a) * min(available_h, cw / a))
        options.append((min(visible_area), sum(visible_area), columns, cw, ch, heights, captions))
    if not options:
        return []
    _, _, columns, cw, ch, heights, captions = max(options)
    return [(f, x + (i % columns) * (cw + GAP), y + (i // columns) * (ch + GAP),
             cw, heights[i], captions[i]) for i, f in enumerate(figures)]


def _draw_figures(scene, placements, style):
    for f, x, y, w, h, captions in placements:
        picture(scene, f["data"], x, y, w, h)
        drawn = scene["elements"][-1]
        if min(drawn["w"], drawn["h"]) < 65:
            scene["warnings"].append("有图像较窄，可放大核对图内标注。")
        if captions:
            text(scene, "\n".join(captions), x, drawn["y"] + drawn["h"] + 6, w, CAPTION_SIZE,
                 style["muted"], role="caption", align="center")


def _figure_batches(figures, box, comparison=False):
    x, y, w, h = box
    if not figures:
        return []
    batches, remaining = [], list(figures)
    while remaining:
        for count in range(len(remaining), 0, -1):
            placements = _grid(remaining[:count], (x, y, w, h), comparison)
            if placements:
                batches.append(placements)
                remaining = remaining[count:]
                break
        else:
            raise UserError("图或图注放不下。可改用整宽图区或缩短图注。")
    return batches


def _arrange(page, figures, layout, box, bullets=None):
    x, y, w, h = box
    group_pairs = bullets is None and any(f.get("bullet") is not None for f in figures)
    if bullets is None:
        bullets = list(enumerate(page["bullets"]))
    bias = page.get("space", "auto")
    if layout == "paired" or group_pairs:
        pairs = defaultdict(list)
        for f in figures:
            pairs[f.get("bullet")].append(f)
        if layout == "paired" and not any(key is not None for key in pairs):
            raise UserError("使用图文逐项时，请为图片选对应要点。")
        units = [(index, bullet, pairs[index]) for index, bullet in bullets]
        pages = []
        for start in range(0, len(units), 2):
            group = units[start:start + 2]
            chunks = []
            if len(group) > 1:
                horizontal = layout == "comparison"
                cw, ch = (w-GAP)/2 if horizontal else w, h if horizontal else (h-GAP)/2
                offsets = [(x+i*(cw+GAP), y) if horizontal else (x,y+i*(ch+GAP)) for i in range(2)]
                try:
                    chunks = [_arrange(page, links, "stacked" if horizontal else "side", (cx, cy, cw, ch),
                                       [(index, value)]) for (index, value, links), (cx, cy) in zip(group, offsets)]
                except UserError:
                    # A full-height page provides an accessible fallback for this pair.
                    chunks = []
            if chunks and all(len(chunk) == 1 for chunk in chunks):
                pages.append(([item for chunk in chunks for item in chunk[0][0]],
                              [place for chunk in chunks for place in chunk[0][1]], None))
                continue
            for index, bullet, links in group:
                # Only this bullet's own text and figures appear in its continuation.
                pages.extend(_arrange(page, links, "side" if layout in {"paired", "side"} else "stacked",
                                      box, [(index, bullet)]))
        if pairs[None]:
            pages.extend(_arrange(page, pairs[None], "stacked", box, []))
        return pages

    if not figures:
        if layout in {"columns", "comparison"} and len(bullets) >= 2:
            columns = (bullets[:math.ceil(len(bullets)/2)], bullets[math.ceil(len(bullets)/2):])
            cw = (w - GAP) / 2
            left, right = [_flow(col, (x + i * (cw + GAP), y, cw, h)) for i, col in enumerate(columns)]
            return [((left[i] if i < len(left) else []) + (right[i] if i < len(right) else []),
                     [], (x + w/2, y, h)) for i in range(max(len(left), len(right)))]
        return [(chunk, [], None) for chunk in _flow(bullets, box)]

    if not bullets:
        return [([], placement, None) for placement in _figure_batches(figures, box, layout == "comparison")]
    if layout == "side":
        avg = sum(min(2, f["data"]["width"]/f["data"]["height"]) for f in figures) / len(figures)
        demand = min(1, sum(len(value) for _, value in bullets) / 340)
        offset = {"auto": 0, "figure": .12, "text": -.12}[bias]
        picture_w = w * min(.72, max(.32, .56 + .08*(avg - 1) - .14*demand + offset))
        figures_box = (x, y, picture_w, h)
        text_box = (x + picture_w + GAP, y, w - picture_w - GAP, h)
        divider = (x + picture_w + GAP/2, y, h)
    else:
        needed = sum(len(_take_lines(value, w-36)) * LINE + LINE for _, value in bullets)
        fraction = {"auto": .30, "figure": .18, "text": .42}[bias]
        text_h = min(h * .5, max(LINE * 2, min(needed, h * fraction)))
        figures_box = (x, y, w, h - text_h - GAP)
        text_box = (x, y + h - text_h, w, text_h)
        divider = None
    gallery = _figure_batches(figures, figures_box, layout == "comparison")
    body = _flow(bullets, text_box)
    rows = [(body[0], gallery[0], divider)]
    linked = any(f.get("bullet") is not None for f in figures)
    rows += [(body[0] if linked else [], batch, divider) for batch in gallery[1:]]
    if linked:
        # On overflow, keep the associated figure visible alongside its text.
        return rows + [(chunk, gallery[0], divider) for chunk in body[1:]]
    if len(body) > 1:
        remaining = [(lines, index) for chunk in body[1:] for lines, index, _, _, _ in chunk]
        limit = int((h - 4)/LINE)
        current, used = [], 0
        for lines, index in remaining:
            if used + len(lines) > limit:
                rows.append((current, [], None))
                current, used = [], 0
            current.append((lines, index, x, y + used * LINE, w))
            used += len(lines) + 1
        if current:
            rows.append((current, [], None))
    return rows


def _diagram_rows(page, kind, box):
    x, y, w, h = box
    values = list(enumerate(diagram_points(page)))
    cols = 2 if kind == "comparison" else min(3, len(values))
    cw = (w - (cols - 1) * GAP) / cols
    rows = math.ceil(len(values) / cols)
    ch = (h - (rows - 1) * GAP) / rows
    cells = [_take_lines(value, cw - 70) for _, value in values]
    needed = max(134, max(len(lines) for lines in cells) * LINE + 76)
    if needed <= ch:
        ch = needed
        offset = min(78, max(0, (h-rows*ch-(rows-1)*GAP)/3))
        return [([(lines, j, x+(j%cols)*(cw+GAP)+12, y+offset+(j//cols)*(ch+GAP)+53, cw-24)
                  for j, lines in enumerate(cells)], cw, ch)]
    # Dense content uses wider panels and can continue without dropping text.
    items = _flow(values, (x+12, y+49, w-24, h-66))
    return [(chunk, w, None) for chunk in items]


def _draw_diagram(scene, items, cw, ch, kind, style):
    for slot, (lines, index, x, y, w) in enumerate(items):
        card_h = ch if ch is not None else len(lines)*LINE + 12
        top = y-53 if ch is not None else y-6
        if kind != "tasks" or ch is None:
            rect(scene, x-12, top, cw, card_h, tint(style["primary"], .96), stroke=tint(style["primary"], .75))
        if kind == "flow" and ch is not None:
            text(scene, f"{index+1:02d}", x+4, top+12, 55, 25, style["primary"], True)
        elif ch is not None:
            rect(scene, x+4, top+25, min(cw-32, 92), 3, style["primary"])
        if kind == "flow" and ch is not None and slot+1 < len(items) and items[slot+1][3] == y:
            rect(scene, x+cw-9, top+ch/2-8, GAP-6, 16, style["primary"], kind="chevron")
        _body(scene, lines, index, x, y, w, style, numbered=kind == "flow", show_marker=ch is None)


def _draw_bars(scene, values, box, style):
    x, y, w, h = box
    text(scene, values[0][1], x, y+12, w, 28, style["text"], True)
    maximum = max(float(item[2]) for item in values)
    start = x + min(w*.3, 272)
    span = w - (start-x) - 190
    row_gap = min(112, (h-82)/2)
    rect(scene, start-1, y+64, 1.3, row_gap+62, tint(style["primary"], .49))
    for i, (label, _, number, unit) in enumerate(values):
        top = y + 76 + row_gap*i
        text(scene, "\n".join(wrap(label, start-x-14, 23)), x, top-3, start-x-14, 23, style["text"])
        rect(scene, start, top+3, span, 20, tint(style["primary"], .94))
        value = float(number)
        if value:
            rect(scene, start, top+3, span*value/maximum, 20, style["primary"] if i == 0 else style["accent"])
        suffix = f"{number}{unit}" if unit == "%" else f"{number} {unit}"
        text(scene, suffix, start+span+16, top-4, 165, 23, style["text"], True)


def build_personal_scenes(outline, style, template, images, sources):
    scenes = []
    height = template.height

    def new(title, layout, page_template, reason=""):
        return {"width": WIDTH, "height": height, "title": title, "layout": layout, "layout_reason": reason,
                "font": style["font"], "bg": style["background"], "backdrop": page_template.backdrop,
                "elements": [], "warnings": [], "personal": True}

    if outline.get("include_cover", True):
        cover = new(outline["title"], "cover", template.cover)
        x, y, w, _ = template.cover.title_box
        x, y = max(x, 90), max(y, 126)
        w = min(w, WIDTH-x-90)
        for size in (64, 58, 52, 46, 40):
            lines = wrap(outline["title"], w * .97, size)
            if len(lines) <= 3:
                break
        needed = len(lines) * size * 1.35 + 4
        line_y = max(height * .60, y + needed + GAP)
        text(cover, "\n".join(lines), x, y + (line_y-y-needed-GAP)/2, w, size, "#000000", True,
             role="title", align="center")
        rect(cover, 82, line_y, 1116, 6, style["primary"])
        subtitle = outline.get("subtitle", "")
        if subtitle.strip():
            text(cover, "\n".join(wrap(subtitle, 1050, 32)), 115, line_y + GAP, 1050, 32, style["text"],
                 role="subtitle", align="center")
        cover["notes"] = subtitle
        scenes.append(cover)
    for index, page in enumerate(outline["slides"]):
        figures = page_figures(page, images)
        if "figures" in page and len(page["figures"]) != len(figures):
            raise UserError(f"第 {index+1} 页有图片已移除，请重新选择。")
        recommended, reason = recommend_personal(page, figures)
        choice = page.get("personal_layout", "auto")
        layout = recommended if choice == "auto" else choice
        if choice != "auto":
            reason = "使用你指定的版式"
        if not figures and layout in {"main", "side", "stacked", "paired"}:
            layout, reason = "bullets", "当前页没有配图，使用文字排版"
        if figures and layout in {"columns", "bullets", "process"}:
            layout, reason = "stacked", "原页有配图，保留图文并使用上下排布"
        if layout == "paired" and not any(f.get("bullet") is not None for f in figures):
            layout, reason = "stacked", "图文对应关系可在编辑大纲时设置"
        title = page["title"]
        tx, ty, tw, _ = template.body.title_box
        x, y, w, h = template.body.body_box
        x, y = min(max(x, 38), 150), max(y, 126)
        w, bottom = min(w, WIDTH - x - 38), min(y+h, height-48)
        tx, ty, tw = max(64, tx), max(20, ty), min(tw, 1060, WIDTH-max(64, tx)-75)
        title_lines = wrap(title, tw * .97, 37.333)
        long_title = bool(title) and len(title_lines) > 1
        if long_title:
            tx, ty, tw = x, y, w
            title_lines = wrap(title, tw * .97, 37.333)
            y += len(title_lines)*37.333*1.35 + GAP
        diagram = diagram_for(page) if page.get("diagram") and not figures else None
        if diagram == "bars" and bottom-y < 200:
            diagram = "comparison"
        diagram_rows = []
        try:
            if diagram == "bars":
                rows = [([], [], None)]
                layout, reason = "comparison", "文字转图 · " + DIAGRAM_LABELS[diagram]
            elif diagram:
                diagram_rows = _diagram_rows(page, diagram, (x, y, w, bottom-y))
                rows = [(items, [], None) for items, _, _ in diagram_rows]
                layout = {"flow": "process", "tasks": "columns", "comparison": "comparison"}[diagram]
                reason = "文字转图 · " + DIAGRAM_LABELS[diagram]
            else:
                rows = _arrange(page, figures, layout, (x, y, w, bottom-y))
        except UserError:
            if diagram or layout != "side":
                raise
            layout, reason = "stacked", "图或图注需要整宽图区"
            rows = _arrange(page, figures, layout, (x, y, w, bottom-y))
        title_only = bool(title.strip()) and not page["bullets"] and not figures
        if title_only:
            rows = [([], [], None)]
        for continuation, (items, placement, divider) in enumerate(rows):
            scene = new(title, layout, template.body, reason)
            if title_only:
                text(scene, "\n".join(wrap(title, w*.97, 48)), x, y+70, w, 48, style["primary"],
                     True, role="title", align="center")
            elif title:
                text(scene, "\n".join(title_lines), tx, ty, tw, 37.333, style["primary"], True, role="title")
            if divider:
                dx, dy, dh = divider
                rect(scene, dx, dy, 1.4, dh, "#FFFFFF", stroke="#B7C2CF", dash=True)
            if diagram == "bars":
                _draw_bars(scene, comparable_values(page["bullets"]), (x, y, w, bottom-y), style)
            elif diagram:
                _, cw, ch = diagram_rows[continuation]
                _draw_diagram(scene, items, cw, ch, diagram, style)
            else:
                for lines, bullet, bx, by, bw in items:
                    _body(scene, lines, bullet, bx, by, bw, style, layout == "process")
            _draw_figures(scene, placement, style)
            ids = page.get("source_ids", [])
            if ids:
                text(scene, "来源  " + " · ".join(ids), x, height-34, w-75, 14, style["muted"])
            text(scene, str(len(scenes)+1), WIDTH-89, height-34, 45, 14, style["muted"], align="right")
            refs = "\n".join(f"[{sid}] {sources.get(sid, sid)}" for sid in ids)
            figure_sources = "\n".join(_caption(f) for f, *_ in placement).strip()
            scene["notes"] = (page.get("notes", "") + ("\n\n【原文定位】\n" + refs if refs else "")
                              + ("\n\n【配图出处】\n" + figure_sources if figure_sources else ""))
            if diagram == "bars":
                scene["notes"] += "\n\n【图表原文】\n" + "\n".join(page["bullets"])
            scene.update(source_ids=ids, origin_slide=index)
            if len(rows)>1:
                scene["warnings"].append(f"本页图文分为 {len(rows)} 页，当前为第 {continuation+1} 页。")
            scenes.append(scene)
    if len(scenes) > 100:
        raise UserError("排版后超过 100 页，请减少同页的图文或拆分汇报。")
    for i, scene in enumerate(scenes, 1):
        for j, element in enumerate(scene["elements"], 1):
            element["id"] = f"slide-{i:03d}-{element.get('role', element['kind'])}-{j:03d}"
            if (min(element["w"], element["h"]) <= 0 or min(element["x"], element["y"]) < 0
                    or element["x"]+element["w"] > WIDTH+1 or element["y"]+element["h"] > height+1):
                raise UserError(f"第 {i} 页的内容超出模板可用空间，请缩短标题或图注，或换一种版式。")
    return scenes

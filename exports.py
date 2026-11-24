"""Standalone HTML and evidence bundle; native editability audit from the local skill."""
import base64
import html
import io
import json
import re
import zipfile
from xml.etree import ElementTree as ET

from layouts import LAYOUT_LABELS
from renderer import scene_svg
from skill_audit import audit


def scene_html(scene):
    esc=html.escape
    elements=[]
    for e in scene["elements"]:
        position=f'position:absolute;left:{e["x"]}px;top:{e["y"]}px;width:{e["w"]}px;height:{e["h"]}px;box-sizing:border-box;'
        identifier=esc(e["id"])
        if e["kind"] in {"rect","chevron"}:
            border=f'border:1px {"dashed" if e.get("dash") else "solid"} {e["stroke"]};' if e.get("stroke") else ""
            clip='clip-path:polygon(0 0,55% 0,100% 50%,55% 100%,0 100%,45% 50%);' if e["kind"]=="chevron" else ""
            elements.append(f'<div id="{identifier}" aria-hidden="true" style="{position}background:{e["fill"]};border-radius:{12 if e.get("radius") else 0}px;{border}{clip}"></div>')
        elif e["kind"]=="text":
            tag="h2" if e.get("role")=="title" else "p"
            font=re.sub(r"[^a-zA-Z0-9 \-\u4e00-\u9fff]", "", scene["font"])
            css=f"margin:0;font-family:'{font}',sans-serif;font-size:{e['size']}px;font-weight:{700 if e['bold'] else 400};line-height:{e['line_height']};white-space:pre;text-align:{e['align']};color:{e['color']};"
            elements.append(f'<{tag} id="{identifier}" style="{position}{css}">{esc(e["text"])}</{tag}>')
        else:
            b64=base64.b64encode(e["data"]).decode()
            elements.append(f'<img id="{identifier}" alt="用户选择的原图" src="data:image/jpeg;base64,{b64}" style="{position}object-fit:contain">')
    return f'<section class="slide" aria-label="{esc(scene["title"],quote=True)}" style="width:1280px;height:{scene["height"]}px;background:{scene["bg"]};">'+"".join(elements)+"</section>"


def deck_html(scenes,title):
    slides="".join('<div class="slide-wrap" style="aspect-ratio:1280/'+str(s["height"])+'">'+scene_html(s)+"</div>" for s in scenes)
    return ('<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
            f'<title>{html.escape(title)}</title><style>body{{margin:0;background:#e8edf6;font-family:sans-serif}}'
            'main{max-width:1280px;margin:24px auto;padding:0 16px}.slide-wrap{position:relative;overflow:hidden;margin-bottom:24px;box-shadow:0 6px 24px #17243d14}'
            '.slide{position:absolute;transform-origin:top left;overflow:hidden}p,h2{overflow:visible}@media print{body{background:white}main{margin:0;padding:0}.slide-wrap{break-after:page;margin:0;box-shadow:none}}'
            '</style></head><body><main>'+slides+'</main><script>'
            'function fit(){document.querySelectorAll(".slide-wrap").forEach(w=>{w.firstElementChild.style.transform="scale("+w.clientWidth/1280+")"})}addEventListener("resize",fit);fit();'
            '</script></body></html>').encode('utf-8')


def audit_native(scenes,pptx):
    # This SVG is an internal semantic manifest, not a visual presentation SVG.
    # It lets the skill's actual checker compare each paragraph ID with one native textbox.
    root=ET.Element('{http://www.w3.org/2000/svg}svg')
    identifiers=[]
    for scene in scenes:
        for e in scene['elements']:
            if e['kind']=='text':
                ET.SubElement(root,'{http://www.w3.org/2000/svg}text',{'id':e['id']}).text=e['text']
                identifiers.append(e['id'])
    result=audit(ET.tostring(root),pptx,identifiers)
    result['profile']='Native PPTX paragraph mapping; internal semantic manifest only'
    result['warnings']=[w for w in result['warnings'] if not w.startswith('Raw newline')]
    result['warnings'].append('HTML/SVG/PPTX use a shared scene; no native PowerPoint visual rendering or manual editing was tested.')
    result['layout_count']=len({s['layout'] for s in scenes if s['layout']!='cover'})
    return result


def evidence_record(doc,outline):
    return {'document':doc.metadata,'blocks':doc.blocks,
        'extracted_images':[{k:i.get(k,'') for k in ('name','source','pdf_page','width','height')} for i in doc.images],
        'pages':[
        {k:s.get(k,'') for k in ('title','purpose','relationship','source_ids','image','image_caption')}
        for s in outline['slides']],
        'boundary':'DOIs are text candidates, not externally verified. Figure numbers and interpretation require author review.'}


def export_bundle(scenes,outline,style,doc,pptx,check,include_svg=False):
    output=io.BytesIO()
    with zipfile.ZipFile(output,'w',zipfile.ZIP_DEFLATED) as z:
        z.writestr('presentation.pptx',pptx)
        z.writestr('presentation.html',deck_html(scenes,outline['title']))
        for name,value in [('outline.json',outline),('style.json',style),('evidence.json',evidence_record(doc,outline)),('checks.json',check)]:
            z.writestr(name,json.dumps(value,ensure_ascii=False,indent=2))
        if include_svg:
            for i,s in enumerate(scenes,1):
                z.writestr(f'svg/slide-{i:03d}.svg',scene_svg(s))
        boundary=('文稿包包括大纲、HTML、原生可编辑PPTX、证据记录和结构检查。\n'
                  '文字和图形可编辑；原图内的标签仍属于图片。字体未嵌入。\n'
                  '检查器核对文字框数量、ID与正文，不代表内容事实核对或PowerPoint视觉检查。\n'
                  'HTML与原生PPTX来自相同排版场景，不执行任意HTML到PPT转换。\n')
        if include_svg:
            boundary+='已按网页选择附加浏览器SVG：多行采用tspan；不满足严格no-tspan契约，不保证PowerPoint转换为形状后不拆字。推荐编辑原生PPTX。\n'
        else:
            boundary+='默认不附加多行SVG；可在网页明确启用浏览器SVG导出。\n'
        z.writestr('说明.txt',boundary)
    return output.getvalue()

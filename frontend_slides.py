"""A deployable Frontend Slides adapter: recipes, native scenes and offline fonts.

Upstream fixed-stage CSS and recipe references live in assets/frontend-slides.
No Codex installation or model call is needed to style an existing outline.
"""
import base64
import copy
import hashlib
import io
import json
import re
from functools import lru_cache
from pathlib import Path

from core import UserError, clean_text
from layout_engine import build_scenes, rect, text, wrap

ASSETS=Path(__file__).parent/'assets'/'frontend-slides'
THEMES={
    'electric-studio':{
        'name':'Electric Studio · 蓝白演讲','summary':'蓝白分区、大字标题，适合课程讲解。',
        'source':'STYLE_PRESETS.md / Electric Studio','background':'#FFFFFF','primary':'#4361EE',
        'accent':'#4361EE','text':'#0A0A0A','muted':'#5C6475','cover':'#4361EE',
        'display':'Manrope','body':'Manrope',
    },
    'signal':{
        'name':'Signal · 深蓝金色','summary':'深蓝、纸白与金色细线，适合组会和文献汇报。',
        'source':'bold-template-pack/templates/signal/preview.md',
        'recipe':'bold-template-pack/templates/signal/design.md',
        'background':'#F0ECE3','primary':'#1C2644','accent':'#C8A870','text':'#1A2030',
        'muted':'#657080','cover':'#1C2644','display':'Source Serif 4','body':'DM Sans',
    },
    'research-atlas':{
        'name':'Research Atlas · 实验笔记','summary':'浅纸色、深绿和坐标网格，适合科研记录。',
        'source':'SlideCraft custom wildcard','background':'#F6F5EE','primary':'#174F46',
        'accent':'#C65A32','text':'#172B29','muted':'#5F746E','cover':'#F6F5EE',
        'display':'DM Sans','body':'Manrope',
    },
}


def theme_spec(theme_id, preview=False):
    if theme_id not in THEMES:
        raise UserError('请选择网站中提供的 Frontend Slides 风格。')
    result=copy.deepcopy(THEMES[theme_id])
    if not preview and result.get('recipe'):
        # Read exactly the selected full recipe; gallery covers use preview metadata only.
        recipe=(ASSETS/result['recipe']).read_text('utf-8')
        for key in ('background','primary','accent','text'):
            match=re.search(re.escape(result[key]),recipe,re.I)
            if not match:
                raise UserError('所选设计资料与配色不一致，请更新网站的 Frontend Slides 资源。')
            result[key]=match.group().upper()
        result['recipe_sha256']=hashlib.sha256(recipe.encode()).hexdigest()
    result['id']=theme_id
    return result


def density_outline(outline,density):
    if density not in {'reading','speaking'}:
        raise UserError('未知的内容密度。')
    result=copy.deepcopy(outline)
    pages=[]
    for index,slide in enumerate(result['slides']):
        pieces=[]
        for bullet_id,bullet in enumerate(slide['bullets']):
            if density=='speaking':
                # Exact slices preserve whitespace, punctuation, numbers and every source character.
                remaining=bullet
                while remaining:
                    end=min(88,len(remaining))
                    if end<len(remaining):
                        cuts=[remaining.rfind(c,24,end) for c in ('。','；','，',' ','. ','; ')]
                        if max(cuts)>24:
                            end=max(cuts)+1
                    pieces.append((bullet_id,remaining[:end]))
                    remaining=remaining[end:]
            else:
                pieces.append((bullet_id,bullet))
        cap=2 if density=='speaking' else 6
        for offset in range(0,len(pieces),cap):
            page=copy.deepcopy(slide)
            part=pieces[offset:offset+cap]
            page['bullets']=[p[1] for p in part]
            page['_origin_slide']=index
            page['_origin_bullets']=[p[0] for p in part]
            if offset:
                page['title']=slide['title']+'（续）'
            pages.append(page)
    result['slides']=pages
    return result


def _cover(outline,theme,count):
    theme_id=theme['id']
    s={'width':1280,'height':720,'bg':theme['cover'],'font':theme['body'],
       'title':outline['title'],'layout':'cover','layout_reason':'','notes':outline.get('subtitle',''),
       'origin_slide':-1,'elements':[]}
    if theme_id=='electric-studio':
        rect(s,0,0,1280,225,'#FFFFFF')
        rect(s,72,170,104,8,theme['primary'])
        top,color=274,'#FFFFFF'
    elif theme_id=='signal':
        for x in range(80,1280,80):
            rect(s,x,0,1,720,'#232F55')
        for y in range(80,720,80):
            rect(s,0,y,1280,1,'#232F55')
        rect(s,72,130,104,4,theme['accent'])
        top,color=202,'#F0ECE3'
    else:
        rect(s,72,72,1136,576,theme['cover'],stroke='#C2CCC0')
        for x in range(72,1210,88):
            rect(s,x,72,1,576,'#E3E7DC')
        for y in range(72,650,72):
            rect(s,72,y,1136,1,'#E3E7DC')
        rect(s,72,72,12,576,theme['primary'])
        rect(s,112,148,80,5,theme['accent'])
        top,color=220,theme['text']
    width=1056 if theme_id=='research-atlas' else 1136
    x=112 if theme_id=='research-atlas' else 72
    sub=outline.get('subtitle','')
    sub_lines=wrap(sub,width,24) if sub else []
    sub_h=len(sub_lines)*24*1.35
    for size in (68,62,56,50,44,38,32):
        lines=wrap(outline['title'],width,size)
        title_h=len(lines)*size*1.35
        if top+title_h+sub_h+32<650:
            break
    e=text(s,'\n'.join(lines),x,top,width,size,color,True,role='title')
    e.update(original_text=outline['title'],font=theme['display'])
    if sub_lines:
        e=text(s,'\n'.join(sub_lines),x,top+title_h+24,width,24,color,role='subtitle')
        e['original_text']=sub
    text(s,f'1 / {count}',x,674,1080,16,color)
    return s


def frontend_scenes(outline,theme_id,density='reading',images=None,sources=None,preview=False):
    theme=theme_spec(theme_id,preview)
    if preview:
        scenes=[_cover(outline,theme,1)]
    else:
        adapted=density_outline(outline,density)
        style={k:theme[k] for k in ('background','primary','accent','text','muted')}
        style.update(name=theme['name'],font=theme['body'],ratio='16:9')
        scenes=build_scenes(adapted,style,images,sources)
        scenes[0]=_cover(outline,theme,len(scenes))
        for s in scenes[1:]:
            s['font']=theme['body']
            for e in s['elements']:
                if e['kind']=='text':
                    e['font']=theme['display'] if e.get('role')=='title' else theme['body']
                if e['kind']=='rect' and theme_id=='signal':
                    e['radius']=False
                    if e['fill'].upper() not in {theme['primary'],theme['accent'],'#FFFFFF'}:
                        e['fill']='#E6E0D4'
                    if e['w']>120 and e['h']>60 and e['fill'].upper()!=theme['primary']:
                        e['stroke']='#D2C9B8'
                if e['kind']=='text' and e.get('role')=='title':
                    e['original_text']=s['title']
    # The upstream stage is genuinely authored at 1920x1080, not a responsive reflow.
    for i,s in enumerate(scenes,1):
        s.update(width=1920,height=1080,frontend_theme=theme_id)
        for j,e in enumerate(s['elements'],1):
            for key in ('x','y','w','h'):
                e[key]*=1.5
            if e['kind']=='text':
                e['size']*=1.5
                e.setdefault('font',theme['body'])
            e['id']=f'slide-{i:03d}-{e.get("role",e["kind"])}-{j:03d}'
            if e['x']+e['w']>1922 or e['y']+e['h']>1082:
                raise UserError('Frontend Slides 内容超出画布，请缩短标题或正文。')
    return scenes,theme


@lru_cache(maxsize=8)
def _asset_bytes(relative):
    return (ASSETS/relative).read_bytes()


def font_css(scenes,cache=None):
    """Subset the open CJK font to this deck. Cache is supplied by the current user session only."""
    from fontTools import subset
    from fontTools.ttLib import TTFont
    characters=set(''.join(e['text'] for s in scenes for e in s['elements'] if e['kind']=='text'))
    characters.update('上一页下一页编辑保存退出全屏修改下载讲稿文字超出区域请精简自动本地清除')
    characters.update(chr(n) for n in range(32,127))
    key=hashlib.sha256(''.join(sorted(characters)).encode()).hexdigest()
    cache=cache if cache is not None else {}
    if key not in cache:
        font=TTFont(io.BytesIO(_asset_bytes('fonts/NotoSansSC.ttf')))
        options=subset.Options()
        options.hinting=False
        sub=subset.Subsetter(options=options)
        sub.populate(unicodes=[ord(c) for c in characters])
        sub.subset(font)
        font.flavor='woff2'
        buf=io.BytesIO();font.save(buf);font.close()
        # Do not retain multiple documents' subsets indefinitely in a live session.
        if len(cache)>=4:
            cache.pop(next(iter(cache)))
        cache[key]=base64.b64encode(buf.getvalue()).decode('ascii')
    parts=["@font-face{font-family:'SlideCraft CJK';font-style:normal;font-weight:100 900;font-display:swap;src:url(data:font/woff2;base64,"+cache[key]+") format('woff2');}"]
    for family,filename in [('Manrope','Manrope'),('DM Sans','DMSans'),('Source Serif 4','SourceSerif4')]:
        if family not in {e.get('font',s['font']) for s in scenes for e in s['elements'] if e['kind']=='text'}:
            continue
        b64=base64.b64encode(_asset_bytes('fonts/'+filename+'.woff2')).decode('ascii')
        parts.append("@font-face{font-family:'"+family+"';font-style:normal;font-weight:100 900;font-display:swap;src:url(data:font/woff2;base64,"+b64+") format('woff2');}")
    return '\n'.join(parts)


def licenses():
    result={'frontend-slides':_asset_bytes('LICENSE').decode('utf-8')}
    for name in ('NotoSansSC','Manrope','DMSans','SourceSerif4'):
        result[name]=_asset_bytes('fonts/'+name+'-OFL.txt').decode('utf-8')
    return result


def edit_bindings(scenes):
    result={}
    for s in scenes:
        for e in s['elements']:
            role=e.get('role')
            if e['kind']=='text' and role in {'title','body','subtitle','caption'}:
                result[e['id']]={'role':role,'slide':s.get('origin_slide',-1),
                    'bullet':e.get('origin_bullet',e.get('bullet_id',-1)),
                    'text':e.get('original_text',e['text']), 'display':e['text']}
    return result


def edit_signature(scenes):
    return hashlib.sha256(json.dumps(edit_bindings(scenes),sort_keys=True,ensure_ascii=False).encode()).hexdigest()


def apply_frontend_edits(outline,scenes,payload):
    if not isinstance(payload,dict) or payload.get('schema')!='slidecraft-frontend-edits/v1':
        raise UserError('请导入动态演示中下载的修改 JSON。')
    if payload.get('deck_id')!=edit_signature(scenes):
        raise UserError('修改文件不属于当前演示。请保持原来的风格和密度，或使用同一份大纲重新生成。')
    edits=payload.get('edits')
    bindings=edit_bindings(scenes)
    if not isinstance(edits,dict) or any(k not in bindings or not isinstance(v,str) or len(v)>3000 for k,v in edits.items()):
        raise UserError('修改文件中的文字或段落编号无效。')
    result=copy.deepcopy(outline)
    chunks={}
    title_changes={}
    for identifier,b in bindings.items():
        value=clean_text(edits[identifier]) if identifier in edits else b['text']
        if not value.strip():
            raise UserError('标题和正文不能留空，请在网页中补充文字后再导入。')
        if b['role']=='body' and b['slide']>=0:
            chunks.setdefault((b['slide'],b['bullet']),[]).append(value)
        elif identifier in edits:
            if b['role']=='title':
                title_changes.setdefault(b['slide'],set()).add(value)
            elif b['role']=='subtitle' and b['slide']==-1:
                if len(value)>150:
                    raise UserError('副标题超过 150 字，请精简后再导入。')
                result['subtitle']=value
            elif b['role']=='caption' and b['slide']>=0:
                if len(value)>200:
                    raise UserError('图片出处超过 200 字，请精简后再导入。')
                result['slides'][b['slide']]['image_caption']=value
    for (slide,bullet),parts in chunks.items():
        result['slides'][slide]['bullets'][bullet]=''.join(parts)
    for slide in result['slides']:
        slide['bullets']=[line for bullet in slide['bullets'] for line in bullet.splitlines() if line.strip()]
    for slide,values in title_changes.items():
        if len(values)>1:
            raise UserError('同一章节的续页标题修改不一致，请使用相同标题后再导入。')
        if slide==-1:
            result['title']=next(iter(values))
        else:
            result['slides'][slide]['title']=next(iter(values)).removesuffix('（续）')
    return result

"""Bounded compositions shared by preview and native export."""
import copy
import math
import re

from core import UserError, normalize_style
from layouts import resolve_layout


def text_width(value, size):
    return sum(size*(1 if ord(c)>255 else .62 if c.isupper() else .56) for c in value)


def wrap(value, width, size):
    lines = []
    for paragraph in str(value).split("\n"):
        line = ""
        for token in re.findall(r"[A-Za-z0-9_]+(?:[./@:+%-][A-Za-z0-9_]+)*|[^\x00-\xff]|\s+|.", paragraph):
            if line and text_width(line+token, size)>width:
                if token in "，。！？；：、）》】」』”’" and len(line)>1:
                    lines.append(line[:-1].rstrip())
                    line=line[-1]+token
                    continue
                lines.append(line.rstrip())
                line = ""
            for char in token:
                if line and text_width(line+char, size)>width:
                    lines.append(line.rstrip())
                    line = ""
                line += char
        lines.append(line.rstrip())
    # Avoid leaving a Chinese closing mark or a one-character tail on its own line.
    for i in range(1,len(lines)):
        if (0<len(lines[i].strip())<=2 and len(lines[i-1])>8
                and (ord(lines[i][0])>255) and ord(lines[i-1][-1])>255):
            tail=lines[i-1][-4:]
            if text_width(tail+lines[i],size)<=width:
                lines[i-1]=lines[i-1][:-4]
                lines[i]=tail+lines[i]
    return lines or [""]


def tint(color, amount=.92):
    values = [int(color[i:i+2], 16) for i in (1,3,5)]
    return "#"+"".join(f"{round(v+(255-v)*amount):02X}" for v in values)


def title_metrics(title):
    for size in (40,36,32,28):
        lines = wrap(title,1044,size)
        if len(lines)<=3:
            break
    return lines,size,max(152,51+len(lines)*size*1.35+25)


def rect(scene,x,y,w,h,fill,radius=False,stroke=None,dash=False,kind="rect"):
    scene["elements"].append(dict(kind=kind,x=x,y=y,w=w,h=h,fill=fill,radius=radius,stroke=stroke,dash=dash))


def text(scene,value,x,y,w,size,color,bold=False,role="label",align="left"):
    e=dict(kind="text",text=str(value),x=x,y=y,w=w,h=len(str(value).split("\n"))*size*1.35+4,
           size=size,color=color,bold=bold,role=role,align=align,line_height=1.35)
    scene["elements"].append(e)
    return e


def fit(value,w,h,preferred=26):
    for size in range(preferred,19,-1):
        lines=wrap(value,w,size)
        if len(lines)*size*1.35+4<=h:
            return lines,size
    return None


def slots_for(layout,count,top,height):
    body_h=height-82-top
    gap=22
    if layout=="focus":
        if count==1:
            return [(110,top+38,1048,body_h-90,44)]
        return [(108,top+42,606,body_h-92,40)]+[
            (826,top+45+i*((body_h-60)/(count-1)),338,(body_h-60)/(count-1)-40,25)
            for i in range(count-1)]
    if layout=="takeaway":
        main_h=body_h if count==1 else min(205,body_h*.49)
        slots=[(112,top+32,1040,main_h-62,34)]
        if count>1:
            w=(1136-(count-2)*gap)/(count-1)
            y=top+main_h+38
            slots += [(92+i*(w+gap),y+24,w-36,height-98-y-24,25) for i in range(count-1)]
        return slots
    if layout in {"cards","comparison","process"}:
        n=max(1,min(count,6))
        cols=2 if layout=="comparison" else (n if n<=3 else 2 if n==4 else 3)
        rows=math.ceil(n/cols)
        w=(1136-(cols-1)*gap)/cols
        cell_h=(body_h-(rows-1)*gap)/rows
        return [(96+(i%cols)*(w+gap),top+(i//cols)*(cell_h+gap)+62,w-48,cell_h-88,34 if n<=2 else 28 if n==3 else 25)
                for i in range(n)]
    if layout=="literature":
        n=max(1,min(count,4))
        cell_h=(body_h-(n-1)*12)/n
        return [(158,top+i*(cell_h+12)+20,544,cell_h-34,25) for i in range(n)]
    if layout=="image":
        cell_h=(body_h-20)/max(1,count)
        return [(112,top+12+i*cell_h,514,cell_h-24,27) for i in range(count)]
    if layout=="columns":
        rows=math.ceil(count/2)
        cell_h=(body_h-20)/max(1,rows)
        return [(108+(i%2)*592,top+14+(i//2)*cell_h,484,cell_h-28,28) for i in range(count)]
    cell_h=(body_h-24)/max(1,count)
    return [(112,top+16+i*cell_h,1064,cell_h-26,30) for i in range(count)]


def paginate(outline,style,images):
    """Continue long text onto new pages, keeping each body segment in one text object."""
    height=960 if style["ratio"]=="4:3" else 720
    pages,previous=[],None
    for slide_index,slide in enumerate(outline["slides"]):
        slide=copy.deepcopy(slide)
        layout,reason=resolve_layout(slide,images,previous)
        previous=layout
        if layout in {"focus","takeaway"} and len(slide["bullets"])>4:
            layout,reason="columns","内容较多，使用双栏保证可读字号"
        _,_,top=title_metrics(slide["title"]+"（续）")
        capacity=min(len(slide["bullets"]),4 if layout=="literature" else 6)
        slots=slots_for(layout,capacity,top,height)
        chunks,current,slot_index=[],[],0
        for bullet_id,original in enumerate(slide["bullets"]):
            remaining=original
            continued=False
            while remaining:
                slot=slots[slot_index]
                x,y,w,h,preferred=slot
                fitted=fit(remaining,w,h,preferred)
                if fitted:
                    lines,size=fitted
                    piece,remaining=remaining,""
                else:
                    size=20
                    # Find the longest exact source prefix that fits. No words or spaces are discarded.
                    low,high,best=1,len(remaining),0
                    while low<=high:
                        mid=(low+high)//2
                        if len(wrap(remaining[:mid],w,size))*size*1.35+4<=h:
                            best,low=mid,mid+1
                        else:
                            high=mid-1
                    if not best:
                        raise UserError("页标题过长或内容过密，请缩短标题或选择简洁列表。")
                    piece,remaining=remaining[:best],remaining[best:]
                    lines=wrap(piece,w,size)
                current.append(dict(text=piece,lines=lines,size=size,slot=slot,bullet_id=bullet_id,continued=continued))
                continued=True
                slot_index+=1
                if slot_index==len(slots):
                    chunks.append(current)
                    current,slot_index=[],0
        if current:
            chunks.append(current)
        for i,items in enumerate(chunks):
            page=copy.deepcopy(slide)
            page.update(title=slide["title"]+("（续）" if i else ""),layout=layout,layout_reason=reason,items=items,top=top)
            page['origin_slide']=slide.get('_origin_slide',slide_index)
            pages.append(page)
    if len(pages)>100:
        raise UserError("排版后超过 100 页，请缩短要点或拆分汇报。")
    return pages


def picture(scene,asset,x,y,w,h):
    scale=min(w/asset["width"],h/asset["height"])
    iw,ih=asset["width"]*scale,asset["height"]*scale
    scene["elements"].append(dict(kind="image",x=x+(w-iw)/2,y=y+(h-ih)/2,w=iw,h=ih,data=asset["data"]))


def image_panel(scene,page,images,style,x,y,w,h):
    rect(scene,x,y,w,h,"#FFFFFF",stroke=tint(style["primary"],.72),dash=True)
    asset=images.get(page.get("image"))
    caption=page.get("image_caption","") or (asset.get("source","") if asset else "")
    lines=wrap(caption,w-36,16) if caption else []
    cap_h=len(lines)*22+20 if lines else 0
    if asset:
        picture(scene,asset,x+18,y+18,w-36,h-36-cap_h)
    else:
        rect(scene,x+w*.35,y+h*.22,w*.3,h*.32,tint(style["primary"],.95),stroke=tint(style["primary"],.65))
        text(scene,"选择一张原图",x+24,y+h*.61,w-48,23,style["muted"],align="center")
    if lines:
        text(scene,"\n".join(lines),x+18,y+h-cap_h+4,w-36,16,style["muted"],role="caption")


def build_scenes(outline,style,images=None,sources=None):
    style=normalize_style(style)
    images,sources=images or {},sources or {}
    height=960 if style["ratio"]=="4:3" else 720
    pages=paginate(outline,style,images)
    scenes=[]
    def new(bg,title,layout,reason=""):
        return dict(width=1280,height=height,bg=bg,font=style["font"],title=title,layout=layout,layout_reason=reason,elements=[])
    cover=new(style["primary"],outline["title"],"cover")
    rect(cover,72,72,72,8,style["accent"])
    text(cover,"DOCUMENT TO PRESENTATION",72,104,1100,18,"#FFFFFF")
    subtitle="\n".join(wrap(outline.get("subtitle",""),1040,26))
    sub_h=len(subtitle.split("\n"))*26*1.35+4
    for size in (64,58,50,44,38,34,30):
        lines=wrap(outline["title"],1090,size)
        title_h=len(lines)*size*1.35
        if title_h+sub_h+32<=height*.75-110:
            break
    text(cover,"\n".join(lines),72,height*.25,1136,size,"#FFFFFF",True,role="title")
    text(cover,subtitle,76,max(height*.61,height*.25+title_h+28),1100,26,"#FFFFFF")
    text(cover,f"{len(pages)+1:02d} SLIDES  /  {outline.get('mode','文档汇报')}",76,height-62,1100,16,"#FFFFFF")
    cover["notes"]="封面\n"+outline.get("subtitle","")
    scenes.append(cover)
    for index,page in enumerate(pages,2):
        layout=page["layout"]
        s=new(style["background"],page["title"],layout,page["layout_reason"])
        top,body_h=page["top"],height-82-page["top"]
        rect(s,72,43,62,62,style["primary"])
        text(s,f"{index-1:02d}",75,52,56,27,"#FFFFFF",True,align="center")
        lines,size,_=title_metrics(page["title"])
        text(s,"\n".join(lines),158,44,1044,size,style["text"],True,role="title")
        rect(s,158,top-22,1050,1,tint(style["primary"],.78))
        pale,pale_accent=tint(style["primary"],.94),tint(style["accent"],.92)
        if layout=="focus":
            rect(s,72,top,1136 if len(page["bullets"])==1 else 694,body_h,pale,radius=True)
            rect(s,72,top+30,7,56,style["accent"])
            text(s,"本页焦点",110,top+10,560,16,style["primary"],True)
        elif layout=="takeaway":
            rect(s,72,top,1136,page["items"][0]["slot"][3]+62,style["primary"],radius=True)
        elif layout=="literature":
            for item in page["items"]:
                x,y,w,h,_=item["slot"]
                rect(s,72,y-20,662,h+34,"#FFFFFF",stroke=tint(style["primary"],.65),dash=True)
                rect(s,88,y-3,48,42,style["primary"])
                label=chr(65+item["bullet_id"])
                text(s,label,91,y+2,42,18,"#FFFFFF",True,align="center")
            if page.get("image") in images:
                image_panel(s,page,images,style,762,top,446,body_h)
            else:
                rect(s,762,top,446,body_h,pale)
                text(s,"原文定位",792,top+28,386,24,style["primary"],True)
                y=top+82
                for sid in page.get("source_ids",[])[:5]:
                    lines=wrap(f"[{sid}] {sources.get(sid,sid)}",378,19)
                    if len(lines)*19*1.35+4>top+body_h-24-y:
                        break
                    text(s,"\n".join(lines),792,y,378,19,style["muted"])
                    y+=len(lines)*19*1.35+20
        elif layout=="image":
            image_panel(s,page,images,style,698,top,510,body_h)
        for j,item in enumerate(page["items"]):
            x,y,w,h,_=item["slot"]
            color,bold=style["text"],False
            if layout in {"cards","comparison","process"}:
                cx,cy,cw,ch=x-24,y-62,w+48,h+88
                rect(s,cx,cy,cw,ch,pale_accent if layout=="comparison" and j%2 else pale,radius=True)
                if layout=="process":
                    rect(s,cx,cy,58,43,style["primary"],radius=True)
                    text(s,f"{item['bullet_id']+1:02d}",cx+4,cy+4,50,23,"#FFFFFF",True,align="center")
                    if item['continued']:
                        text(s,"续",cx+70,cy+8,44,16,style['muted'])
                    if (j+1<len(page["items"]) and abs(page["items"][j+1]["slot"][1]-y)<1
                            and page["items"][j+1]['bullet_id']!=item['bullet_id']):
                        rect(s,cx+cw+4,cy+ch/2-7,14,14,style["accent"],kind="chevron")
                else:
                    label=("对照 " if layout=="comparison" else "要点 ")+f"{item['bullet_id']+1:02d}"
                    text(s,label,x,cy+18,w,17,style["primary"],True)
                    rect(s,x,cy+48,min(74,w),4,style["accent"])
            elif layout=="takeaway":
                if j==0:
                    color,bold="#FFFFFF",True
                else:
                    rect(s,x,y-16,36,4,style["accent"])
            elif layout=="focus":
                bold=j==0
                if j:
                    rect(s,x,y-18,36,4,style["accent"])
            elif layout in {"bullets","columns","image"}:
                rect(s,x-25,y+10,6,16,style["accent"])
            body=text(s,"\n".join(item["lines"]),x,y,w,item["size"],color,bold,role="body")
            body.update(original_text=item["text"],bullet_id=item["bullet_id"])
            body['origin_bullet']=page.get('_origin_bullets',list(range(len(page['bullets']))))[item['bullet_id']]
        rect(s,72,height-64,1136,1,tint(style["primary"],.8))
        ids=page.get("source_ids",[])
        label="来源："+" · ".join(ids[:10])+(f" 等 {len(ids)} 处" if len(ids)>10 else "")
        text(s,label,76,height-45,980,14,style["muted"])
        text(s,f"{index} / {len(pages)+1}",1108,height-45,100,14,style["muted"],align="right")
        refs="\n".join(f"[{sid}] {sources.get(sid,sid)}" for sid in ids)
        image_source=page.get('image_caption','') or images.get(page.get('image'),{}).get('source','')
        s["notes"]=(f"【本页目的】{page.get('purpose','')}\n【内容关系】{page.get('relationship','parallel')}\n"
                    +page.get("notes","")+"\n\n【图片出处】\n"+image_source+"\n\n【原文定位】\n"+refs)
        s["source_ids"]=ids
        s['origin_slide']=page['origin_slide']
        scenes.append(s)
    for page_index,s in enumerate(scenes,1):
        for element_index,e in enumerate(s["elements"],1):
            e["id"]=f"slide-{page_index:03d}-{e.get('role',e['kind'])}-{element_index:03d}"
            if e["x"]<0 or e["y"]<0 or e["x"]+e["w"]>1282 or e["y"]+e["h"]>height+2:
                raise UserError(f"第 {page_index} 页内容超出画布，请缩短标题或切换版式。")
    return scenes

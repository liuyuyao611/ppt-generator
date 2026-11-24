"""Streamlit integration. Style discovery and explicit edits happen inside the website."""
import hashlib
import io
import json
import os
import zipfile

import streamlit as st
import streamlit.components.v1 as components

from core import UserError, validate_outline
from exports import audit_native,export_bundle
from frontend_slides import (THEMES,frontend_scenes,font_css,apply_frontend_edits,
                             edit_signature,licenses)
from frontend_player import frontend_html
from renderer import export_pptx,safe_filename


def bundle_frontend(scenes,theme,density,outline,doc,pptx,check,html_bytes):
    style={k:theme[k] for k in ('background','primary','accent','text','muted')}
    style.update(name=theme['name'],font=theme['body'],ratio='16:9')
    base=export_bundle(scenes,outline,style,doc,pptx,check)
    result=io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(base)) as source,zipfile.ZipFile(result,'w',zipfile.ZIP_DEFLATED) as out:
        for name in source.namelist():
            if name not in {'presentation.html','说明.txt'}:
                out.writestr(name,source.read(name))
        out.writestr('presentation.html',html_bytes)
        out.writestr('frontend-slides.json',json.dumps({'theme':theme,'density':density,'canvas':[1920,1080],
            'source':'https://github.com/zarazhangrui/frontend-slides','native_pptx':'static editable counterpart'},ensure_ascii=False,indent=2))
        for name,license_text in licenses().items():
            out.writestr('licenses/'+name+'.txt',license_text)
        out.writestr('说明.txt','动态 HTML 使用固定 1920×1080 画布，内嵌字体、图片、样式和控制器。\n'
            '方向键、空格、滚轮、手机滑动可翻页，全屏按钮可演示。按 E 或左上角进入文字编辑。\n'
            '网页文字修改在本机暂存。下载修改 JSON 并导回网站后，才能重新生成对应的 PPTX。\n'
            'PPTX 为原生可编辑静态版本，不包含网页动画。未嵌入字体；设备字体替换可能改变显示。\n'
            '检查记录只核对原生文字与段落映射，不代替科学事实核对或 PowerPoint 实测。\n')
    return result.getvalue()


def render_frontend(outline,doc,assets):
    st.subheader('Frontend Slides · 动态演示')
    st.write('先比较三种封面，再选一种生成整套演示。所有预览都使用当前大纲的真实标题。')
    if '_frontend_notice' in st.session_state:
        st.success(st.session_state.pop('_frontend_notice'))
    try:
        validate_outline(outline,{b['id'] for b in doc.blocks})
        font_cache=st.session_state.setdefault('frontend_font_cache',{})
        columns=st.columns(3)
        for col,(theme_id,spec) in zip(columns,THEMES.items()):
            with col:
                cover,theme=frontend_scenes(outline,theme_id,preview=True)
                css=font_css(cover,font_cache)
                components.html(frontend_html(cover,theme,css,preview=True,embedded=True).decode(),height=210,scrolling=False)
                st.markdown('**'+spec['name']+'**')
                st.caption(spec['summary'])
        selected=st.radio('选择动态演示风格',list(THEMES),format_func=lambda k:THEMES[k]['name'],
                          index=None,horizontal=True,key='frontend_theme_choice')
        density=st.radio('动态演示内容密度',['reading','speaking'],
            format_func=lambda k:'阅读汇报 · 内容较完整' if k=='reading' else '现场讲解 · 每页更精简',
            horizontal=True,key='frontend_density')
        st.caption('动态演示采用 16:9，不随手机宽度重排正文。此处的风格和密度不改动上方原生版式设置。')
        if selected is None:
            st.info('选择上方一种风格，即可生成完整动态演示；也可以继续使用“版式预览与下载”中的原生 PPTX。')
            return
        sources={b['id']:f"{b['heading']} · {b['source']}" for b in doc.blocks}
        scenes,theme=frontend_scenes(outline,selected,density,assets,sources)
        css=font_css(scenes,font_cache)
        embedded=frontend_html(scenes,theme,css,embedded=True)
        standalone=frontend_html(scenes,theme,css)
        st.caption(f"{len(scenes)} 页 · {theme['name']} · 1920×1080 · 支持全屏、方向键和手机滑动")
        components.html(embedded.decode(),height=680,scrolling=False)
        st.download_button('下载 Frontend Slides 动态 HTML',standalone,'frontend-presentation.html','text/html',
                           type='primary',use_container_width=True)
        st.caption('HTML 内嵌当前内容所需字体和图片，可离线打开。鼠标移到演示左上角或按 E 编辑；PPTX 使用静态、可编辑版本。')
        state_key=hashlib.sha256((edit_signature(scenes)+selected+density).encode()).hexdigest()
        if st.button('生成对应的可编辑 PPTX 和文稿包',use_container_width=True):
            with st.spinner('正在创建静态 PPTX 并检查段落映射…'):
                pptx=export_pptx(scenes,outline['title'])
                check=audit_native(scenes,pptx)
                if not check['passed']:
                    raise UserError('原生编辑性检查未通过：'+'；'.join(check['errors'][:3]))
                st.session_state.frontend_output={'key':state_key,'pptx':pptx,
                    'bundle':bundle_frontend(scenes,theme,density,outline,doc,pptx,check,standalone)}
        output=st.session_state.get('frontend_output',{})
        if output.get('key')==state_key:
            st.download_button('下载动态演示对应的静态 PPTX',output['pptx'],safe_filename(outline['title']),
                'application/vnd.openxmlformats-officedocument.presentationml.presentation',use_container_width=True)
            st.download_button('下载 Frontend Slides 完整文稿包',output['bundle'],'frontend-slides-bundle.zip',
                               'application/zip',use_container_width=True)
        with st.expander('把网页编辑的文字导回大纲，再生成 PPTX'):
            st.write('在演示中按 E 编辑，再点击“下载修改 JSON”。保持相同的风格和密度，将文件上传到这里。正文中的换行将按新的要点处理。')
            upload=st.file_uploader('网页修改 JSON',type=['json'],key='frontend_import_edits')
            if st.button('应用网页修改到大纲',disabled=upload is None,use_container_width=True):
                if upload.size>200_000:
                    raise UserError('修改文件超过 200 KB，请减少单次修改量。')
                try:
                    payload=json.loads(upload.getvalue())
                except (ValueError,UnicodeError):
                    raise UserError('修改文件不是有效 JSON。') from None
                updated=apply_frontend_edits(outline,scenes,payload)
                validate_outline(updated,{b['id'] for b in doc.blocks})
                st.session_state.outline=updated
                st.session_state.revision=os.urandom(4).hex()
                for key in ('export','export_sig','export_check','frontend_output'):
                    st.session_state.pop(key,None)
                st.session_state._frontend_notice='网页修改已导入大纲。请核对原文后重新下载 PPTX。'
                st.rerun()
    except UserError as exc:
        st.error(str(exc))


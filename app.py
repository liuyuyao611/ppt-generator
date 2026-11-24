"""SlideCraft — document to editable PowerPoint, ready for Streamlit Cloud."""
from __future__ import annotations

import copy
import hashlib
import hmac
import json
import os
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

from core import (SourceDoc, DEFAULT_STYLE, MAX_IMAGES, MAX_IMAGE_BYTES, UserError, ai_outline,
                  extract_style, make_outline, normalize_style, normalized_image,
                  read_document, validate_outline)
from renderer import build_scenes, export_pptx, safe_filename, scene_svg
from layouts import LAYOUT_LABELS, RELATION_LABELS, recommend_layout
from exports import audit_native, deck_html, export_bundle
from preview import preview_html
from frontend_ui import render_frontend
from personal_layout import PERSONAL_LAYOUTS, build_personal_scenes
from personal_template import PERSONAL_STYLE
from personal_ui import edit_personal_pages, personal_template_picker, restore_personal_draft

ROOT = Path(__file__).parent
st.set_page_config(page_title="SlideCraft · 文档生成 PPT", page_icon="📑", layout="wide")


def setting(name, default=""):
    try:
        return str(st.secrets.get(name, os.environ.get(name, default)))
    except (FileNotFoundError, st.errors.StreamlitSecretNotFoundError):
        return os.environ.get(name, default)


def signature(value):
    return hashlib.sha256(json.dumps(value, ensure_ascii=False, sort_keys=True).encode()).hexdigest()


def reset_result():
    for key in list(st.session_state):
        if key in {"outline", "doc", "assets", "export", "export_sig", "export_check", "result_input_sig", "personal_images_sig", "personal_assets",
                   "frontend_output", "frontend_font_cache", "frontend_theme_choice"} or key.startswith("edit_"):
            del st.session_state[key]


def collect_assets(doc, uploads, *, stable=False):
    if len(doc.images) + len(uploads) > MAX_IMAGES:
        raise UserError("文档内图片与补充图片合计最多 24 张，请减少后重试。")
    assets = {f"文档图片 {i} · {image['name']}": image for i, image in enumerate(doc.images, 1)}
    for i, upload in enumerate(uploads, 1):
        data = upload.getvalue()
        suffix = f" · {hashlib.sha256(data).hexdigest()[:10]}" if stable else ""
        label = f"补充图片 · {upload.name}" if stable else f"补充图片 {i} · {upload.name}"
        assets[label+suffix] = normalized_image(data, upload.name)
    if sum(len(a["data"]) for a in assets.values()) > MAX_IMAGE_BYTES:
        raise UserError("图片合计过大，请压缩图片后重试。")
    return assets


def main():
    access = setting("APP_ACCESS_CODE")
    if access and not st.session_state.get("authenticated"):
        st.title("SlideCraft")
        st.write("输入分享者提供的访问码后开始制作。")
        with st.form("access"):
            code = st.text_input("访问码", type="password")
            submit = st.form_submit_button("进入")
        if submit:
            if hmac.compare_digest(code.encode(), access.encode()):
                st.session_state.authenticated = True
                st.rerun()
            st.error("访问码不正确。")
        st.stop()

    st.markdown("""
    <style>
    .block-container {max-width:1380px;padding-top:2.2rem;padding-bottom:3rem}
    h1 {letter-spacing:-.04em}
    .brand {font-weight:800;color:#2147B1;letter-spacing:.12em;font-size:.9rem}
    .intro {font-size:1.12rem;color:#62708A;line-height:1.8;max-width:750px}
    .flow {padding:1rem 0 1.5rem;color:#62708A;font-size:.94rem;letter-spacing:.04em}
    [data-testid="stSidebar"] {border-right:1px solid #E1E6F1}
    [data-testid="stMetric"] {background:white;border:1px solid #E1E6F1;border-radius:12px;padding:.8rem 1rem}
    div.stButton>button[kind="primary"] {border-radius:9px}
    </style>
    <div class="brand">SLIDECRAFT / 多版式汇报 · V3 · FRONTEND SLIDES</div>
    """, unsafe_allow_html=True)
    st.title("把一份文档，变成一场清晰的汇报。")
    st.markdown('<p class="intro">上传文档，整理讲述顺序，配上你的图片与风格。最终下载可继续编辑的 PowerPoint。</p>'
                '<div class="flow">上传材料 &nbsp; / &nbsp; 检查大纲 &nbsp; / &nbsp; 预览与下载</div>', unsafe_allow_html=True)

    with st.sidebar:
        st.subheader("汇报设置")
        title = st.text_input("汇报标题（可选）", placeholder="默认使用文档标题", max_chars=120)
        subtitle = st.text_input("副标题 / 姓名 / 课程", placeholder="例如：张同学 · 课程文献汇报", max_chars=150)
        pages = st.slider("目标总页数", 5, 24, 10)
        st.caption("优先保持章节完整，长要点可自动分为续页。")
        mode = st.radio("内容整理方式", ["规则分段 · 免费", "AI 整理 · 需要 API"])
        api_key, model, consent = "", "", False
        api_base = setting("OPENAI_BASE_URL", "https://api.openai.com/v1")
        if mode.startswith("AI"):
            providers = {"站点默认": (api_base, setting("OPENAI_MODEL")),
                         "OpenAI": ("https://api.openai.com/v1", "gpt-4.1-mini"),
                         "DeepSeek": ("https://api.deepseek.com", "deepseek-flash")}
            provider = st.selectbox("模型服务商", list(providers))
            api_base, default_model = providers[provider]
            server_key = setting("OPENAI_API_KEY") if provider == "站点默认" else ""
            configured = bool(server_key)
            api_key = st.text_input("API Key（留空使用部署者配置）" if configured else "API Key",
                                    type="password", key=f"visitor_api_key_{provider}")
            api_key = api_key or server_key
            model = st.text_input("模型名称", value=default_model, key=f"model_{provider}",
                                  placeholder="填写你的服务支持的模型")
            st.caption("模型服务：" + api_base)
            consent = st.checkbox("同意将文档文字发送至所配置的模型服务", key=f"model_consent_{provider}")
            st.caption("AI 调用可能产生服务商费用；上传的图片不发送给模型。")
        st.divider()
        st.subheader("页面风格")
        configured_template = setting("PERSONAL_TEMPLATE_PATH")
        configured_template_base64 = setting("PERSONAL_TEMPLATE_BASE64") if not configured_template else ""
        presets = ["学术蓝", "组会蓝橙", "森林绿", "深紫", "暖橙", PERSONAL_STYLE["name"]]
        has_template = bool(configured_template or configured_template_base64)
        preset = st.selectbox("配色与风格", presets, index=len(presets)-1 if has_template else 0)
        personal = preset == PERSONAL_STYLE["name"]
        template_selection = None
        if personal:
            style = PERSONAL_STYLE.copy()
            template_selection = personal_template_picker(configured_template, configured_template_base64, ROOT)
            if template_selection:
                style["ratio"] = "4:3" if template_selection[1].height == 960 else "16:9"
        else:
            styles = {
                "学术蓝": DEFAULT_STYLE,
                "组会蓝橙": DEFAULT_STYLE | {"name":"组会蓝橙", "primary":"#073B71", "accent":"#E98429", "background":"#FFFFFF", "text":"#111111", "muted":"#4B4B4B"},
                "森林绿": DEFAULT_STYLE | {"name":"森林绿", "primary":"#126254", "accent":"#C18A23"},
                "深紫": DEFAULT_STYLE | {"name":"深紫", "primary":"#55349A", "accent":"#C74E8D"},
                "暖橙": DEFAULT_STYLE | {"name":"暖橙", "primary":"#9F411E", "accent":"#147E9A"},
            }
            style = copy.deepcopy(styles[preset])
            if preset == "学术蓝" and (ROOT / "style.json").exists():
                try:
                    style = normalize_style(json.loads((ROOT / "style.json").read_text("utf-8")))
                except (ValueError, UserError):
                    st.warning("本地 style.json 无效，已采用默认样式。")
            reference = st.file_uploader("参考 PPT（可选）", type=["pptx"], key="reference")
            if reference:
                digest = hashlib.sha256(reference.getvalue()).hexdigest()
                try:
                    if st.session_state.get("reference_hash") != digest:
                        st.session_state.reference_style = extract_style(reference.getvalue())
                        st.session_state.reference_hash = digest
                    style = copy.deepcopy(st.session_state.reference_style)
                    st.caption("已读取参考配色、字体和比例；复杂母版、背景及动画不复刻。")
                except UserError as exc:
                    st.error(str(exc))
            style_file = st.file_uploader("导入 style.json（可选，优先于参考 PPT）", type=["json"])
            if style_file:
                try:
                    if style_file.size > 100_000:
                        raise UserError("样式文件过大。")
                    style = normalize_style(json.loads(style_file.getvalue()))
                except (UserError, ValueError) as exc:
                    st.error("样式文件无效：" + str(exc))
            style_digest = signature(style)[:10]
            a, b = st.columns(2)
            style["primary"] = a.color_picker("主题色", style["primary"], key=f"color_primary_{style_digest}")
            style["accent"] = b.color_picker("强调色", style["accent"], key=f"color_accent_{style_digest}")
            style["font"] = st.text_input("PPT 字体", style["font"], max_chars=80, key=f"font_{style_digest}")
            ratio_options = ["16:9", "4:3"]
            style["ratio"] = st.selectbox("页面比例", ratio_options,
                                          index=ratio_options.index(style["ratio"]), key=f"ratio_{style_digest}")
            st.download_button("保存样式 JSON", json.dumps(style, ensure_ascii=False, indent=2).encode(),
                               "style.json", "application/json", use_container_width=True)
            st.divider()
        st.caption("材料只在当前服务会话的内存中处理；不写入项目文件。AI 模式另有服务商的数据处理政策。")
        if st.button("清空本次内容", use_container_width=True):
            st.session_state.clear()
            st.rerun()

    left, right = st.columns([1.35, 1], gap="large")
    with left:
        st.subheader("准备你的材料")
        source_mode = st.radio("输入方式", ["上传文档", "粘贴文字"], horizontal=True, label_visibility="collapsed")
        uploaded, pasted = None, ""
        if source_mode == "上传文档":
            uploaded = st.file_uploader("支持 Word / PDF / Markdown / TXT",
                                       type=["docx", "pdf", "md", "txt"], key="document")
            st.caption("每份文件最多 20 MB；PDF 最多 120 页。扫描件请先 OCR。")
        else:
            pasted = st.text_area("文档正文", height=230, placeholder="把正文粘贴在这里，可使用 # 和 ## 标记章节。",
                                  max_chars=100_000, key="pasted")
        supplements = st.file_uploader("补充图片（可多选）", type=["png", "jpg", "jpeg"],
                                       accept_multiple_files=True, key="pictures")
        extract_images = st.checkbox("提取 PDF 内嵌原图，供逐页手动选图", value=True)
        st.caption("图片由你分配到具体页面。" if personal else "图片由你逐页选择，原文要点可自动续页。")
    with right:
        st.subheader("先看一次完整流程")
        st.write("没有现成材料也可以先试用。")
        st.info("1. 生成带原文定位的大纲。\n\n2. 逐页修改标题、要点和备注。\n\n3. 选择版式与图片，预览后下载 PPTX。")
        demo = st.button("载入腕关节传感器示例", use_container_width=True)
        st.caption("示例用于演示工作流程，不包含未经验证的实验数据。")

    content_bytes = uploaded.getvalue() if uploaded else pasted.encode()
    content_name = uploaded.name if uploaded else "输入文字.md"
    image_identity = [(u.name, hashlib.sha256(u.getvalue()).hexdigest()) for u in supplements]
    input_sig = signature([source_mode, content_name, hashlib.sha256(content_bytes).hexdigest(),
                           [] if personal else image_identity, title, subtitle, pages, mode, extract_images, api_base, model])
    generate = st.button("生成可编辑大纲", type="primary", use_container_width=True)
    blank = personal and st.button("从空白页开始", use_container_width=True)
    if generate or demo or blank:
        try:
            if not demo and not blank and not content_bytes.strip():
                raise UserError("请先上传文档或粘贴正文。")
            if not demo and not blank and mode.startswith("AI") and not consent:
                raise UserError("使用 AI 整理前，请勾选发送文档文字的确认选项。")
            with st.spinner("正在读取材料并整理大纲…"):
                if demo:
                    demo_text = (ROOT / "examples" / "腕关节传感器.md").read_text("utf-8")
                    doc = read_document("腕关节传感器.md", demo_text.encode())
                elif blank:
                    doc = SourceDoc(title or "学术汇报", [])
                else:
                    doc = read_document(content_name, content_bytes, extract_images=extract_images)
                assets = collect_assets(doc, [] if demo else supplements, stable=personal)
                if blank:
                    outline = {"title": doc.title, "subtitle": subtitle, "mode": "逐页编辑",
                               "slides": [{"title": "", "bullets": [], "source_ids": [], "notes": ""}]}
                elif not demo and mode.startswith("AI"):
                    outline = ai_outline(doc, pages, title, subtitle, api_key, model, api_base)
                else:
                    outline = make_outline(doc, pages, title, subtitle)
                reset_result()
                st.session_state.update(outline=outline, doc=doc, assets=assets, revision=os.urandom(4).hex(),
                                        result_input_sig=input_sig)
                if personal:
                    st.session_state.update(personal_images_sig=signature([] if demo else image_identity),
                                            personal_assets=assets)
        except UserError as exc:
            st.error(str(exc))
            return
        except Exception:
            st.error("处理未完成，请检查文件后重试。")
            return

    if "outline" not in st.session_state:
        st.divider()
        cols = st.columns(3)
        cols[0].markdown("**可继续编辑**\n\n文字与图形以原生 PPT 对象导出。")
        cols[1].markdown("**风格可复用**\n\n保存 style.json，下次直接导入。")
        cols[2].markdown("**先核对再下载**\n\n查看来源与讲稿，减少内容遗漏。")
        return

    doc, assets = st.session_state.doc, st.session_state.assets
    outline = copy.deepcopy(st.session_state.outline)
    assets_error = None
    if personal:
        if st.session_state.get("personal_images_sig") != signature(image_identity):
            try:
                new_assets = collect_assets(doc, supplements, stable=True)
                st.session_state.update(personal_images_sig=signature(image_identity), personal_assets=new_assets)
            except UserError as exc:
                assets_error = str(exc)
        generated = st.session_state.get(f"edit_{st.session_state.revision}_personal_generated_assets", {})
        assets = {**st.session_state.get("personal_assets", assets), **generated}
    if st.session_state.result_input_sig != input_sig:
        st.warning("输入材料或汇报设置已改变；下方仍是上一次的大纲，点击“生成可编辑大纲”应用新输入。")
    for message in doc.warnings + outline.get("warnings", []):
        st.info(message)
    st.divider()
    metrics = st.columns(3)
    metrics[0].metric("原文内容块", len(doc.blocks))
    rev = st.session_state.revision
    page_order = st.session_state.get(f"edit_{rev}_personal_order")
    count = len(page_order) if personal and page_order else len(outline["slides"])
    metrics[1].metric("大纲页数" if personal else "大纲页数（含封面）", count + (0 if personal else 1))
    metrics[2].metric("可选图片", len(assets))
    active = None
    if personal:
        active = st.radio("工作区", ["编辑大纲", "版式预览与下载", "原文与来源"], horizontal=True,
                          label_visibility="collapsed", key=f"edit_{rev}_personal_view")
        edit = preview = source = st.container()
    else:
        edit, preview, frontend, source = st.tabs(
            ["编辑大纲", "版式预览与下载", "动态演示 · Frontend Slides", "原文与来源"])
    with edit:
        if personal:
            if active == "编辑大纲":
                edit_personal_pages(outline, assets, rev)
            else:
                restore_personal_draft(outline, rev)
        else:
            st.caption("编辑即时用于预览。长要点自动拆为续页，下载页数可能增加。请保存下载结果，刷新或休眠后会话可能丢失。")
            def auto_layouts():
                for i in range(len(outline["slides"])):
                    st.session_state[f"edit_{rev}_{i}_layout"] = "auto"
            st.button("为全部页面重新推荐版式（保留文字）", use_container_width=True, on_click=auto_layouts)
            outline["title"] = st.text_input("封面标题", outline["title"], max_chars=120, key=f"edit_{rev}_deck_title")
            outline["subtitle"] = st.text_input("封面副标题", outline.get("subtitle", ""), max_chars=150, key=f"edit_{rev}_subtitle")
            for i, slide in enumerate(outline["slides"]):
                with st.expander(f"{i+2:02d} · {slide['title']}", expanded=(i == 0)):
                    prefix = f"edit_{rev}_{i}"
                    slide["title"] = st.text_input("页标题", slide["title"], max_chars=100, key=prefix+"_title")
                    body = st.text_area("要点（每行一条，最多 6 条）", "\n".join(slide["bullets"]),
                                        height=150, key=prefix+"_body")
                    slide["bullets"] = [s.strip() for s in body.splitlines() if s.strip()]
                    slide["purpose"] = st.text_input("本页目的", slide.get("purpose", ""), max_chars=180, key=prefix+"_purpose")
                    relations = list(RELATION_LABELS)
                    relation = slide.get("relationship", "parallel")
                    slide["relationship"] = st.selectbox("内容关系（必须有原文依据）", relations,
                        index=relations.index(relation) if relation in relations else 0,
                        format_func=RELATION_LABELS.get, key=prefix+"_relationship")
                    c1, c2 = st.columns(2)
                    layouts = list(LAYOUT_LABELS)
                    choice = slide.get("layout", "auto")
                    slide["layout"] = c1.selectbox("版式", layouts, index=layouts.index(choice) if choice in layouts else 0,
                                                  format_func=LAYOUT_LABELS.get, key=prefix+"_layout")
                    image_names = [""] + list(assets)
                    slide["image"] = c2.selectbox("配图", image_names, format_func=lambda n: n or "不使用图片", key=prefix+"_image")
                    slide["image_caption"] = st.text_input("图片出处 / 图号 / PDF 页码", slide.get("image_caption", ""),
                        max_chars=200, placeholder="只填写已核对的来源；未知图号请留空", key=prefix+"_caption")
                    recommended, reason = recommend_layout(slide, assets)
                    st.caption(f"内容推荐：{LAYOUT_LABELS[recommended]} · {reason}")
                    if slide["image"] and slide["layout"] not in {"auto", "image", "literature"}:
                        st.caption("已选图片；选择“自动推荐”“图文并排”或“文献概览”后显示。")
                    if slide["layout"] == "image" and not slide["image"]:
                        st.caption("请选择配图，未选择时按重点列表排版。")
                    slide["notes"] = st.text_area("讲稿备注（保存到 PPT 备注栏）", slide.get("notes", ""), height=100, key=prefix+"_notes")
                    st.caption("原文定位：" + " · ".join(slide["source_ids"]))
    with source:
        if not personal or active == "原文与来源":
            diagnostics = doc.metadata.get("pdf_extraction", [])
            if diagnostics:
                st.caption("PDF 文字已通过编码检查。请展开原文核对，尤其留意公式、表格和双栏阅读顺序。")
                with st.expander("查看逐页文字识别结果"):
                    st.dataframe([{"PDF 页码": d["page"], "读取方式": "备用解析" if d["engine"] == "Poppler" else "文字层",
                                   "文字数": len(doc.blocks[d["page"]-1]["text"]) if len(doc.blocks) == len(diagnostics) else "见原文"}
                                  for d in diagnostics], hide_index=True, use_container_width=True)
            for block in doc.blocks:
                with st.expander(f"[{block['id']}] {block['heading']} · {block['source']}"):
                    st.text(block["text"])
            if assets:
                st.write("可选图片")
                for name, asset in assets.items():
                    st.image(asset["data"], caption=name, width=240)
    with preview:
        if not personal or active == "版式预览与下载":
            try:
                validate_outline(outline, {b["id"] for b in doc.blocks}, personal=personal)
                style = normalize_style(style)
                source_map = {b["id"]: f"{b['heading']} · {b['source']}" for b in doc.blocks}
                template_digest = None
                template = None
                if personal:
                    if assets_error:
                        raise UserError(assets_error)
                    if template_selection is None:
                        raise UserError("请在「页面风格」选择用于导出的 PPTX 模板。")
                    template_digest, template = template_selection
                    scenes = build_personal_scenes(outline, style, template, assets, source_map)
                    if template.cover.approximate or template.body.approximate:
                        st.caption("浏览器预览近似显示部分母版效果。PPTX 使用原生母版和版式。")
                    for i, scene in enumerate(scenes, 1):
                        for warning in dict.fromkeys(scene["warnings"]):
                            st.caption(f"第 {i} 页 · {warning}")
                else:
                    scenes = build_scenes(outline, style, assets, source_map)
                diversity = len({s["layout"] for s in scenes if s["layout"] != "cover"})
                st.caption(f"实际排版 {len(scenes)} 页 · 使用 {diversity} 种内容版式。字体显示取决于打开设备是否安装该字体。")
                components.html(preview_html(scenes, rev), height=700 if style["ratio"] == "4:3" else 640,
                                scrolling=False)
                with st.expander("查看整套缩略图，对比页面节奏", expanded=personal):
                    group = st.selectbox("缩略图分组", range(max(1, (len(scenes)+11)//12)),
                                         format_func=lambda n: f"第 {n*12+1}–{min(n*12+12,len(scenes))} 页")
                    columns = st.columns(3)
                    for n, s in enumerate(scenes[group*12:group*12+12]):
                        with columns[n%3]:
                            st.image(scene_svg(s).decode("utf-8"), use_container_width=True)
                            labels = PERSONAL_LAYOUTS if personal else LAYOUT_LABELS
                            st.caption(f"{group*12+n+1:02d} · {labels.get(s['layout'], '封面')}")
                export_sig = signature([outline, style, template_digest])
                if st.button("生成 PPTX 下载文件", type="primary", use_container_width=True):
                    with st.spinner("正在创建可编辑幻灯片…"):
                        result = export_pptx(scenes, outline["title"], template=template)
                        check = audit_native(scenes, result)
                        if not check["passed"]:
                            raise UserError("编辑性检查未通过：" + "；".join(check["errors"][:3]))
                        st.session_state.export = result
                        st.session_state.export_check = check
                        st.session_state.export_sig = export_sig
                    st.success("PPTX 已生成，可从下面下载。")
                if st.session_state.get("export_sig") == export_sig:
                    st.download_button("下载 PowerPoint", st.session_state.export, safe_filename(outline["title"]),
                                       "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                       type="primary", use_container_width=True)
                    st.caption("已检查原生文字框与段落映射；尚未在 PowerPoint 中逐页核对显示效果。")
                    include_svg = st.checkbox("同时附加浏览器 SVG（多行使用 tspan；编辑请用原生 PPTX）", value=False)
                    bundle = export_bundle(scenes, outline, style, doc, st.session_state.export,
                                           st.session_state.export_check, include_svg)
                    st.download_button("下载完整文稿包（PPTX + HTML + 大纲 + 来源 + 检查记录）", bundle,
                                       "presentation-bundle.zip", "application/zip", use_container_width=True)
                elif st.session_state.get("export"):
                    st.caption("内容或样式已更新，请重新生成 PPTX 下载文件。")
                st.download_button("下载大纲与讲稿 JSON", json.dumps(outline, ensure_ascii=False, indent=2).encode(),
                                   "outline.json", "application/json", use_container_width=True)
                st.download_button("下载独立 HTML 幻灯片", deck_html(scenes, outline["title"]),
                                   "presentation.html", "text/html", use_container_width=True)
            except UserError as exc:
                st.error(str(exc))
            except Exception:
                st.error("预览或导出失败，请检查页面内容与图片后重试。")
    if not personal:
        with frontend:
            render_frontend(outline,doc,assets)


if __name__ == "__main__":
    main()

"""SlideCraft — document to editable PowerPoint, ready for Streamlit Cloud."""
from __future__ import annotations

import copy
import hashlib
import hmac
import io
import json
import os
from pathlib import Path

import streamlit as st

from core import (DEFAULT_STYLE, MAX_IMAGES, MAX_IMAGE_BYTES, UserError, ai_outline,
                  extract_style, make_outline, normalize_style, normalized_image,
                  read_document, validate_outline)
from renderer import build_scenes, export_pptx, safe_filename, scene_svg

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
        if key in {"outline", "doc", "assets", "export", "export_sig", "result_input_sig"} or key.startswith("edit_"):
            del st.session_state[key]


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
    <div class="brand">SLIDECRAFT / 文档成稿</div>
    """, unsafe_allow_html=True)
    st.title("把一份文档，变成一场清晰的汇报。")
    st.markdown('<p class="intro">上传文档，整理讲述顺序，配上你的图片与风格。最终下载可继续编辑的 PowerPoint。</p>'
                '<div class="flow">上传材料 &nbsp; / &nbsp; 检查大纲 &nbsp; / &nbsp; 预览与下载</div>', unsafe_allow_html=True)

    with st.sidebar:
        st.subheader("汇报设置")
        title = st.text_input("汇报标题（可选）", placeholder="默认使用文档标题", max_chars=120)
        subtitle = st.text_input("副标题 / 姓名 / 课程", placeholder="例如：张同学 · 课程文献汇报", max_chars=150)
        pages = st.slider("目标总页数", 5, 24, 10)
        st.caption("含封面。内容较短时减少页数，长要点会自动分到续页。")
        mode = st.radio("内容整理方式", ["规则分段 · 免费", "AI 整理 · 需要 API"])
        api_key, model, consent = "", "", False
        if mode.startswith("AI"):
            configured = bool(setting("OPENAI_API_KEY"))
            api_key = st.text_input("API Key（留空使用部署者配置）" if configured else "API Key",
                                    type="password", key="visitor_api_key")
            api_key = api_key or setting("OPENAI_API_KEY")
            model = st.text_input("模型名称", value=setting("OPENAI_MODEL"),
                                  placeholder="填写你的服务支持的模型")
            st.caption("模型服务：" + setting("OPENAI_BASE_URL", "https://api.openai.com/v1"))
            consent = st.checkbox("同意将文档文字发送至所配置的模型服务")
            st.caption("AI 调用可能产生服务商费用；上传的图片不发送给模型。")
        st.divider()
        st.subheader("页面风格")
        preset = st.selectbox("配色", ["学术蓝", "森林绿", "深紫", "暖橙"])
        styles = {
            "学术蓝": DEFAULT_STYLE,
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
                           image_identity, title, subtitle, pages, mode])
    generate = st.button("生成可编辑大纲", type="primary", use_container_width=True)
    if generate or demo:
        try:
            if not demo and not content_bytes.strip():
                raise UserError("请先上传文档或粘贴正文。")
            if not demo and mode.startswith("AI") and not consent:
                raise UserError("使用 AI 整理前，请勾选发送文档文字的确认选项。")
            with st.spinner("正在读取材料并整理大纲…"):
                if demo:
                    demo_text = (ROOT / "examples" / "腕关节传感器.md").read_text("utf-8")
                    doc = read_document("腕关节传感器.md", demo_text.encode())
                else:
                    doc = read_document(content_name, content_bytes)
                assets = {}
                extra_images = [] if demo else supplements
                if len(doc.images) + len(extra_images) > MAX_IMAGES:
                    raise UserError("文档内图片与补充图片合计最多 24 张，请减少后重试。")
                for i, image in enumerate(doc.images, 1):
                    assets[f"文档图片 {i} · {image['name']}"] = image
                for i, upload in enumerate(extra_images, 1):
                    assets[f"补充图片 {i} · {upload.name}"] = normalized_image(upload.getvalue(), upload.name)
                if sum(len(a["data"]) for a in assets.values()) > MAX_IMAGE_BYTES:
                    raise UserError("图片合计过大，请压缩图片后重试。")
                if not demo and mode.startswith("AI"):
                    outline = ai_outline(doc, pages, title, subtitle, api_key, model, setting("OPENAI_BASE_URL", "https://api.openai.com/v1"))
                else:
                    outline = make_outline(doc, pages, title, subtitle)
                reset_result()
                st.session_state.update(outline=outline, doc=doc, assets=assets, revision=os.urandom(4).hex(),
                                        result_input_sig=input_sig)
        except UserError as exc:
            st.error(str(exc))
        except Exception:
            st.error("处理未完成，请检查文件后重试。")

    if "outline" not in st.session_state:
        st.divider()
        cols = st.columns(3)
        cols[0].markdown("**可继续编辑**\n\n文字与图形以原生 PPT 对象导出。")
        cols[1].markdown("**风格可复用**\n\n保存 style.json，下次直接导入。")
        cols[2].markdown("**先核对再下载**\n\n查看来源与讲稿，减少内容遗漏。")
        return

    doc, assets = st.session_state.doc, st.session_state.assets
    outline = copy.deepcopy(st.session_state.outline)
    if st.session_state.result_input_sig != input_sig:
        st.warning("输入材料或汇报设置已改变；下方仍是上一次的大纲，点击“生成可编辑大纲”应用新输入。")
    for message in doc.warnings + outline.get("warnings", []):
        st.info(message)
    st.divider()
    metrics = st.columns(3)
    metrics[0].metric("原文内容块", len(doc.blocks))
    metrics[1].metric("大纲页数（含封面）", len(outline["slides"]) + 1)
    metrics[2].metric("可选图片", len(assets))
    edit, preview, source = st.tabs(["编辑大纲", "版式预览与下载", "原文与来源"])
    rev = st.session_state.revision
    with edit:
        st.caption("编辑即时用于预览。长要点自动拆为续页，下载页数可能增加。请保存下载结果，刷新或休眠后会话可能丢失。")
        outline["title"] = st.text_input("封面标题", outline["title"], max_chars=120, key=f"edit_{rev}_deck_title")
        outline["subtitle"] = st.text_input("封面副标题", outline.get("subtitle", ""), max_chars=150, key=f"edit_{rev}_subtitle")
        for i, slide in enumerate(outline["slides"]):
            with st.expander(f"{i+2:02d} · {slide['title']}", expanded=(i == 0)):
                prefix = f"edit_{rev}_{i}"
                slide["title"] = st.text_input("页标题", slide["title"], max_chars=100, key=prefix+"_title")
                body = st.text_area("要点（每行一条，最多 6 条）", "\n".join(slide["bullets"]),
                                    height=150, key=prefix+"_body")
                slide["bullets"] = [s.strip() for s in body.splitlines() if s.strip()]
                c1, c2 = st.columns(2)
                layouts = ["bullets", "columns", "image"]
                labels = {"bullets":"重点列表", "columns":"双栏要点", "image":"图文并排"}
                slide["layout"] = c1.selectbox("版式", layouts, format_func=labels.get, key=prefix+"_layout")
                image_names = [""] + list(assets)
                slide["image"] = c2.selectbox("配图", image_names, format_func=lambda n: n or "不使用图片", key=prefix+"_image")
                if slide["image"] and slide["layout"] != "image":
                    st.caption("已选图片；将版式改为“图文并排”后显示。")
                if slide["layout"] == "image" and not slide["image"]:
                    st.caption("请选择配图，未选择时按重点列表排版。")
                slide["notes"] = st.text_area("讲稿备注（保存到 PPT 备注栏）", slide.get("notes", ""), height=100, key=prefix+"_notes")
                st.caption("原文定位：" + " · ".join(slide["source_ids"]))
    with source:
        for block in doc.blocks:
            with st.expander(f"[{block['id']}] {block['heading']} · {block['source']}"):
                st.text(block["text"])
        if assets:
            st.write("可选图片")
            for name, asset in assets.items():
                st.image(asset["data"], caption=name, width=240)
    with preview:
        try:
            validate_outline(outline, {b["id"] for b in doc.blocks})
            style = normalize_style(style)
            source_map = {b["id"]: f"{b['heading']} · {b['source']}" for b in doc.blocks}
            scenes = build_scenes(outline, style, assets, source_map)
            st.caption(f"实际排版 {len(scenes)} 页。此处为版式预览；PowerPoint 中的字体显示取决于打开设备是否安装该字体。")
            index = st.selectbox("选择预览页面", range(len(scenes)),
                                format_func=lambda i: f"{i+1:02d} · {scenes[i]['title']}", key=f"preview_{rev}")
            if index >= len(scenes):
                index = 0
            st.image(scene_svg(scenes[index]).decode("utf-8"), use_container_width=True)
            export_sig = signature([outline, style])
            if st.button("生成 PPTX 下载文件", type="primary", use_container_width=True):
                with st.spinner("正在创建可编辑幻灯片…"):
                    st.session_state.export = export_pptx(scenes, outline["title"])
                    st.session_state.export_sig = export_sig
                st.success("PPTX 已生成，可从下面下载。")
            if st.session_state.get("export_sig") == export_sig:
                st.download_button("下载 PowerPoint", st.session_state.export, safe_filename(outline["title"]),
                                   "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                   type="primary", use_container_width=True)
            elif st.session_state.get("export"):
                st.caption("内容或样式已更新，请重新生成 PPTX 下载文件。")
            st.download_button("下载大纲与讲稿 JSON", json.dumps(outline, ensure_ascii=False, indent=2).encode(),
                               "outline.json", "application/json", use_container_width=True)
        except UserError as exc:
            st.error(str(exc))
        except Exception:
            st.error("预览或导出失败，请检查页面内容与图片后重试。")


if __name__ == "__main__":
    main()

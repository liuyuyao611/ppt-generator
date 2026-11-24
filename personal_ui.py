"""Personal template input and page-level editing in the existing Streamlit flow."""
from __future__ import annotations

import base64
import binascii
import hashlib
from pathlib import Path
from uuid import uuid4

import streamlit as st

from core import MAX_FILE, MAX_IMAGES, MAX_IMAGE_BYTES, UserError
from concept_image import generate_concept_image, illustration_description
from content_structure import DIAGRAM_LABELS, diagram_for
from layouts import RELATION_LABELS
from personal_layout import PERSONAL_LAYOUTS, SPACE_OPTIONS, page_figures, recommend_personal
from personal_template import load_personal_template


def personal_template_picker(configured_path, configured_base64, root):
    try:
        if configured_path:
            path = Path(configured_path).expanduser()
            path = path if path.is_absolute() else root / path
            if path.stat().st_size > MAX_FILE:
                raise UserError("预设模板超过 20 MB。请使用精简的 PPTX 副本。")
            data = path.read_bytes()
            label = "网站指定的模板"
        elif configured_base64:
            encoded = "".join(configured_base64.split())
            if len(encoded) > 4 * ((MAX_FILE+2)//3):
                raise UserError("预设模板超过 20 MB。请使用精简的 PPTX 副本。")
            try:
                data = base64.b64decode(encoded, validate=True)
            except (ValueError, binascii.Error):
                raise UserError("预设模板的 Base64 配置无效，请联系维护者。") from None
            label = "网站指定的模板"
        else:
            upload = st.file_uploader("用于导出的 PPTX 模板", type=["pptx"], key="personal_template_file",
                                      help="第一页对应封面版式，第二页对应正文版式。")
            if upload is None:
                st.caption("排版前请上传模板，或由网站维护者预设模板。")
                return None
            data = upload.getvalue()
            label = upload.name
        digest = hashlib.sha256(data).hexdigest()
        cached = st.session_state.get("personal_template_cache")
        if cached is None or cached[0] != digest:
            cached = (digest, load_personal_template(data))
            st.session_state.personal_template_cache = cached
        st.caption(f"导出底稿 · {label}")
        return cached
    except FileNotFoundError:
        st.error("找不到网站指定的模板文件，请联系维护者核对配置。")
    except OSError:
        st.error("读取网站指定的模板文件失败，请联系维护者。")
    except UserError as exc:
        st.error(str(exc))
    return None


def edit_personal_pages(outline, assets, rev):
    prefix = f"edit_{rev}_personal_"
    order_key = prefix + "order"
    order = st.session_state.setdefault(order_key, [str(i) for i in range(len(outline["slides"]))])
    previous = {str(i): page for i, page in enumerate(outline["slides"])}
    draft = st.session_state.get(prefix+"draft", {})
    previous.update(draft.get("pages", {}))

    def add_page():
        st.session_state[order_key] = st.session_state[order_key] + [uuid4().hex]

    def move_page(index, move):
        items = list(st.session_state[order_key])
        items[index], items[index+move] = items[index+move], items[index]
        st.session_state[order_key] = items

    def remove_page(index):
        items = list(st.session_state[order_key])
        items.pop(index)
        st.session_state[order_key] = items

    def unlock_page(key):
        st.session_state[key] = "auto"

    st.button("添加一页", on_click=add_page, disabled=len(order) >= 29, key=prefix+"add")
    outline["include_cover"] = st.checkbox("生成封面", value=draft.get("include_cover", True), key=prefix+"cover")
    title = st.text_input("汇报标题", draft.get("title", outline["title"]), max_chars=120, key=prefix+"deck_title")
    outline["title"] = title if outline["include_cover"] else title.strip() or "演示文稿"
    outline["subtitle"] = draft.get("subtitle", outline.get("subtitle", ""))
    if outline["include_cover"]:
        outline["subtitle"] = st.text_input("封面副标题", outline.get("subtitle", ""), max_chars=150,
                                            key=prefix+"subtitle")
    output = []
    generated = st.session_state.setdefault(prefix+"generated_assets", {})
    for i, key in enumerate(order):
        page = previous.get(key, {"title": "", "bullets": [], "source_ids": [], "notes": ""}).copy()
        stem = prefix + key + "_"
        page_label = st.session_state.get(stem+"title", page["title"]) or "未命名"
        with st.expander(f"{i+1:02d} · {page_label}", expanded=i == 0):
            left, right, remove = st.columns(3)
            left.button("上移", disabled=i == 0, key=stem+"up", on_click=move_page, args=(i, -1))
            right.button("下移", disabled=i == len(order)-1, key=stem+"down", on_click=move_page, args=(i, 1))
            remove.button("移除本页", disabled=len(order) == 1, key=stem+"remove", on_click=remove_page, args=(i,))
            page["title"] = st.text_input("页标题（可留空）", page["title"], max_chars=100, key=stem+"title")
            body = st.text_area("正文（每行一条）", "\n".join(page["bullets"]), height=130, key=stem+"body",
                                help="输入 **词组** 加粗，输入 ==词组== 标红并加粗。")
            page["bullets"] = [line.strip() for line in body.splitlines() if line.strip()]
            # Identical text is disambiguated by occurrence. Moving a paragraph
            # preserves its association; removing it returns the image to the page grid.
            bullet_labels = {"": "随本页图组"}
            bullet_indices = {}
            occurrences = {}
            for j, value in enumerate(page["bullets"]):
                occurrences[value] = occurrences.get(value, 0) + 1
                token = f"{hashlib.sha256(value.encode()).hexdigest()}:{occurrences[value]}"
                bullet_labels[token] = f"{j+1}. {value[:38]}"
                bullet_indices[token] = j
            previous_figures = {f["asset"]: f for f in page_figures(page, assets)}
            if st.checkbox("按本页文字生成概念插图", key=stem+"add_illustration"):
                with st.container(border=True):
                    description = st.text_area("要画的概念", illustration_description(page), max_chars=900,
                        key=stem+"art_description", help="只保留图像服务需要的主题与细节，可以先删去敏感内容。")
                    key_value = st.text_input("OpenAI 图片 API Key", type="password", key=stem+"image_api_key")
                    consent = st.checkbox("同意将上面的插图描述发送给 OpenAI 图片服务，调用可能产生费用",
                                          key=stem+"image_consent")
                    if st.button("生成并添加到本页", key=stem+"make_image",
                                 disabled=not (consent and key_value.strip() and description.strip())):
                        try:
                            if len(assets) >= MAX_IMAGES:
                                raise UserError("当前会话的图片数量已达上限，请先减少图片。")
                            with st.spinner("正在生成概念插图…"):
                                asset = generate_concept_image(description, key_value)
                            if sum(len(item["data"]) for item in assets.values()) + len(asset["data"]) > MAX_IMAGE_BYTES:
                                raise UserError("当前会话的图片总大小已达上限。")
                            name = f"AI 概念示意 {len(generated)+1}"
                            asset.update(name=name, source="AI 生成 · 概念示意", generated=True)
                            generated[name] = assets[name] = asset
                            current = st.session_state.get(stem+"images", list(previous_figures))
                            st.session_state[stem+"images"] = [*current, name]
                            st.success("概念插图已添加到本页。")
                        except UserError as exc:
                            st.error(str(exc))
            selected = st.multiselect("本页图片", list(assets), default=list(previous_figures), key=stem+"images",
                                      format_func=lambda name: assets[name]["name"],
                                      help="按选入顺序排放。可以为图片填写图注或选择对应的要点。")
            figures = []
            if len(selected) > 1:
                first = st.selectbox("第一张图", selected, format_func=lambda name: assets[name]["name"], key=stem+"first")
                selected = [first] + [name for name in selected if name != first]
            for name in selected:
                fkey = stem + hashlib.sha256(name.encode()).hexdigest()[:16]
                asset = assets[name]
                prior = previous_figures.get(name, {})
                c1, c2 = st.columns([1, 3])
                c1.image(asset["data"], caption=asset["name"], width=125)
                caption = c2.text_input("图注 / 出处", prior.get("caption", ""), max_chars=200,
                                        key=fkey+"caption", placeholder=asset.get("source", ""))
                prior_bullet = prior.get("bullet")
                options = list(bullet_labels)
                default = next((token for token, index in bullet_indices.items() if index == prior_bullet), "")
                stored = st.session_state.get(fkey+"bullet_link", default)
                association = c2.selectbox("对应正文要点", options, index=options.index(stored) if stored in options else 0,
                                           format_func=bullet_labels.__getitem__, key=fkey+"bullet")
                st.session_state[fkey+"bullet_link"] = association
                figures.append({"asset": name, "caption": caption, "bullet": bullet_indices.get(association)})
            page["figures"] = figures
            page["image"] = figures[0]["asset"] if figures else ""
            page["image_caption"] = figures[0]["caption"] if figures else ""
            with st.container(border=True):
                st.caption("版式与讲稿")
                relations = list(RELATION_LABELS)
                relation = page.get("relationship", "parallel")
                page["relationship"] = st.selectbox("内容关系", relations,
                    index=relations.index(relation) if relation in relations else 0, format_func=RELATION_LABELS.get,
                    key=stem+"relation")
                choices = list(PERSONAL_LAYOUTS)
                choice = page.get("personal_layout","auto")
                page["personal_layout"] = st.selectbox("版式", choices, index=choices.index(choice),
                    format_func=PERSONAL_LAYOUTS.get, key=stem+"layout")
                st.button("本页重新推荐", on_click=unlock_page, args=(stem+"layout",),
                          key=stem+"auto", disabled=page["personal_layout"]=="auto")
                page["space"] = st.selectbox("空间侧重", list(SPACE_OPTIONS), format_func=SPACE_OPTIONS.get,
                                             key=stem+"space")
                page["notes"] = st.text_area("讲稿备注", page.get("notes", ""), height=90, key=stem+"notes")
                diagram = diagram_for(page)
                if diagram and not figures:
                    page["diagram"] = st.checkbox("把本页变成图", value=page.get("diagram", False),
                        key=stem+"diagram", help="预览和 PPTX 使用可编辑图形，保留当前正文和出处。")
                    st.caption(f"建议图形 · {DIAGRAM_LABELS[diagram]}")
                else:
                    page["diagram"] = False
            layout, reason = recommend_personal(page, page_figures(page, assets))
            if page["personal_layout"] == "auto" and not page["diagram"]:
                st.caption(f"自动推荐 · {PERSONAL_LAYOUTS[layout]} · {reason}")
            if page["source_ids"]:
                st.caption("原文定位 · " + " · ".join(page["source_ids"]))
            output.append(page)
    outline["slides"] = output
    st.session_state[prefix+"draft"] = {"title": title, "subtitle": outline["subtitle"],
        "include_cover": outline["include_cover"], "pages": dict(zip(order, output))}


def restore_personal_draft(outline, rev):
    draft = st.session_state.get(f"edit_{rev}_personal_draft")
    if draft:
        outline["include_cover"] = draft["include_cover"]
        title = draft["title"]
        outline["title"] = title if draft["include_cover"] else title.strip() or "演示文稿"
        outline["subtitle"] = draft["subtitle"]
        outline["slides"] = list(draft["pages"].values())

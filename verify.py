"""Offline checks: python verify.py. No API key or paid requests."""
import base64
import copy
import io
import json
import os
import re
import sys
import unittest
import urllib.error
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from docx import Document
from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches
from pypdf import PdfWriter
from pypdf.generic import DecodedStreamObject, DictionaryObject, NameObject, NumberObject

from core import (DEFAULT_STYLE, SourceDoc, UserError, ai_outline, extract_style,
                  make_outline, normalized_image, read_document, validate_outline)
from renderer import build_scenes, export_pptx, scene_svg
from layouts import LAYOUT_LABELS, recommend_layout
from exports import audit_native, deck_html, export_bundle
from frontend_slides import (THEMES,frontend_scenes,density_outline,edit_bindings,
    edit_signature,apply_frontend_edits,font_css,ASSETS)
from frontend_player import frontend_html
from personal_layout import build_personal_scenes
from personal_template import PERSONAL_STYLE, load_personal_template
from content_structure import diagram_for, rank_excerpts
from concept_image import IMAGE_ENDPOINT, generate_concept_image

ROOT = Path(__file__).parent


class PipelineTests(unittest.TestCase):
    def setUp(self):
        self.doc = read_document("sample.md", (
            "# 传感器汇报\n\n## 背景\n腕关节传感器用于记录与运动相关的信号。"
            "研究中应保留单位与原始数据。\n\n## 方法\n先固定佩戴位置，再记录参考动作。"
            "校准后需要独立验证，不能把预期结果写成实测结论。").encode())

    def test_document_title_uses_first_heading(self):
        self.assertEqual(self.doc.title, "传感器汇报")

    def test_rule_source_coverage_and_notes(self):
        out = make_outline(self.doc, 8)
        validate_outline(out, {b["id"] for b in self.doc.blocks})
        used = {v for s in out["slides"] for v in s["source_ids"]}
        self.assertEqual(used, {b["id"] for b in self.doc.blocks})
        self.assertIn("独立验证", "\n".join(s["notes"] for s in out["slides"]))

    def test_docx_table_and_image(self):
        d = Document()
        d.add_heading("验证方案", level=1)
        d.add_paragraph("这是测试文档正文，包含用于来源核对的中文内容。")
        t = d.add_table(rows=1, cols=2)
        t.cell(0, 0).text, t.cell(0, 1).text = "单位", "毫米"
        pic = io.BytesIO()
        Image.new("RGB", (120, 60), "blue").save(pic, "PNG")
        pic.seek(0)
        d.add_picture(pic)
        data = io.BytesIO()
        d.save(data)
        doc = read_document("test.docx", data.getvalue())
        self.assertTrue(any("毫米" in b["text"] for b in doc.blocks))
        self.assertEqual(len(doc.images), 1)

    def test_pdf_text_and_scanned_rejection(self):
        writer = PdfWriter()
        p = writer.add_blank_page(612, 792)
        font = DictionaryObject({NameObject("/Type"): NameObject("/Font"),
            NameObject("/Subtype"): NameObject("/Type1"), NameObject("/BaseFont"): NameObject("/Helvetica")})
        p[NameObject("/Resources")] = DictionaryObject({NameObject("/Font"): DictionaryObject({NameObject("/F1"): writer._add_object(font)})})
        stream = DecodedStreamObject()
        stream.set_data(b"BT /F1 14 Tf 30 700 Td (Sensor validation uses independent reference measurements.) Tj ET")
        p[NameObject("/Contents")] = writer._add_object(stream)
        buf = io.BytesIO()
        writer.write(buf)
        result = read_document("test.pdf", buf.getvalue())
        self.assertIn("independent", result.text)
        blank = PdfWriter()
        blank.add_blank_page(612, 792)
        b = io.BytesIO()
        blank.write(b)
        with self.assertRaisesRegex(UserError, "OCR"):
            read_document("scan.pdf", b.getvalue())

    def test_encoding_and_empty_input(self):
        doc = read_document("test.txt", "腕关节运动记录。测量方法需要校准，结果需要独立验证。".encode("gb18030"))
        self.assertIn("独立验证", doc.text)
        with self.assertRaises(UserError):
            read_document("x.txt", b"")
        with self.assertRaises(UserError):
            read_document("x.docx", b"broken zip")

    def test_pdf_original_images_keep_page_provenance(self):
        writer=PdfWriter()
        page=writer.add_blank_page(612,792)
        pic=io.BytesIO()
        Image.new('RGB',(200,100),'orange').save(pic,'JPEG')
        image=DecodedStreamObject()
        image.set_data(pic.getvalue())
        image.update({NameObject('/Type'):NameObject('/XObject'),NameObject('/Subtype'):NameObject('/Image'),
            NameObject('/Width'):NumberObject(200),NameObject('/Height'):NumberObject(100),
            NameObject('/ColorSpace'):NameObject('/DeviceRGB'),NameObject('/BitsPerComponent'):NumberObject(8),
            NameObject('/Filter'):NameObject('/DCTDecode')})
        font=DictionaryObject({NameObject('/Type'):NameObject('/Font'),NameObject('/Subtype'):NameObject('/Type1'),NameObject('/BaseFont'):NameObject('/Helvetica')})
        page[NameObject('/Resources')]=DictionaryObject({
            NameObject('/Font'):DictionaryObject({NameObject('/F1'):writer._add_object(font)}),
            NameObject('/XObject'):DictionaryObject({NameObject('/Im0'):writer._add_object(image)})})
        content=DecodedStreamObject()
        content.set_data(b'q 200 0 0 100 50 400 cm /Im0 Do Q BT /F1 14 Tf 30 700 Td (Independent sensor validation is needed.) Tj ET')
        page[NameObject('/Contents')]=writer._add_object(content)
        buf=io.BytesIO()
        writer.write(buf)
        doc=read_document('figures.pdf',buf.getvalue(),extract_images=True)
        self.assertEqual(len(doc.images),1)
        self.assertEqual(doc.images[0]['pdf_page'],1)
        self.assertEqual(doc.blocks[0]['pdf_page'],1)
        self.assertEqual(doc.images[0]['width']/doc.images[0]['height'],2)
        self.assertIn('PDF第1页',doc.images[0]['source'])

    def test_editable_pptx_and_sources(self):
        out = make_outline(self.doc, 6)
        scenes = build_scenes(out, DEFAULT_STYLE, sources={b["id"]: b["source"] for b in self.doc.blocks})
        data = export_pptx(scenes, out["title"])
        ppt = Presentation(io.BytesIO(data))
        self.assertEqual(len(ppt.slides), len(scenes))
        texts = [s.text for slide in ppt.slides for s in slide.shapes if s.has_text_frame]
        self.assertTrue(any("独立验证" in t for t in texts))
        self.assertIn("原文定位", ppt.slides[1].notes_slide.notes_text_frame.text)
        profile = extract_style(data)
        self.assertEqual(profile["primary"].upper(), DEFAULT_STYLE["primary"])
        self.assertEqual(profile["ratio"], "16:9")

    def test_long_bullets_paginate_without_loss(self):
        out = make_outline(self.doc, 6)
        original = "验证" * 240
        out["slides"] = [{"title":"较长的验证内容", "bullets":[original]*6, "source_ids":["S1"], "notes":"全量", "layout":"columns", "image":""}]
        scenes = build_scenes(out, DEFAULT_STYLE)
        self.assertGreater(len(scenes), 2)
        reconstructed = "".join(e["text"].replace("\n", "") for scene in scenes[1:]
            for e in scene["elements"] if e.get("role") == "body")
        self.assertEqual(reconstructed, original * 6)
        export_pptx(scenes)

    def test_all_layouts_ratios_images_and_bounds(self):
        pic = io.BytesIO()
        Image.new("RGB", (800, 1600), "green").save(pic, "PNG")
        assets = {"image": normalized_image(pic.getvalue(), "chart.png")}
        for ratio in ("16:9", "4:3"):
            for layout in LAYOUT_LABELS:
                out = make_outline(self.doc, 6)
                for s in out["slides"]:
                    s.update(layout=layout, image="image")
                scenes = build_scenes(out, DEFAULT_STYLE | {"ratio":ratio}, assets)
                for scene in scenes:
                    for e in scene["elements"]:
                        self.assertGreaterEqual(e["x"], 0)
                        self.assertGreaterEqual(e["y"], 0)
                        self.assertLessEqual(e["x"] + e["w"], scene["width"]+2)
                        self.assertLessEqual(e["y"] + e["h"], scene["height"]+2)
                ppt = Presentation(io.BytesIO(export_pptx(scenes)))
                self.assertEqual(len(ppt.slides), len(scenes))

    def test_chapters_never_mix_to_meet_page_target(self):
        doc=read_document('chapters.md', ('# 总标题\n\n'+ '\n\n'.join(
            f'## 第{i}章\n独立章节{i}的第一项内容。第二项内容保持章节完整，不得混入其他章节。' for i in range(1,9))).encode())
        out=make_outline(doc,5)
        for slide in out['slides'][:-1]:
            headings={b['heading'] for b in doc.blocks if b['id'] in slide['source_ids']}
            self.assertEqual(len(headings),1)
            self.assertIn(slide['title'],headings)

    def test_auto_variety_and_no_invented_flow(self):
        demo=read_document('demo.md',(ROOT/'examples'/'腕关节传感器.md').read_bytes())
        scenes=build_scenes(make_outline(demo,10),DEFAULT_STYLE)
        self.assertGreaterEqual(len({s['layout'] for s in scenes[1:]}),4)
        chosen,_=recommend_layout({'title':'四项独立研究','bullets':['材料A','结构B','器件C','系统D']})
        self.assertNotIn(chosen,{'process','comparison'})
        chosen,_=recommend_layout({'title':'实验流程','bullets':['记录基线','采集参考动作','独立验证']})
        self.assertEqual(chosen,'process')

    def test_all_layouts_keep_long_original_body_in_native_text(self):
        original='校准测量 123.45 mm，Independent validation; 保留实验条件。'*7
        for layout in LAYOUT_LABELS:
            out=make_outline(self.doc,6)
            out['slides']=[{'title':'长段落验证','bullets':[original]*4,'source_ids':['S1'],
                            'notes':'原文','layout':layout,'image':''}]
            scenes=build_scenes(out,DEFAULT_STYLE)
            native=Presentation(io.BytesIO(export_pptx(scenes)))
            body_names={e['id'] for s in scenes for e in s['elements'] if e.get('role')=='body'}
            actual=''.join(shape.text for slide in native.slides for shape in slide.shapes
                           if shape.name in body_names)
            self.assertEqual(re.sub(r'\s+','',actual),re.sub(r'\s+','',original*4),layout)

    def test_skill_audit_detects_broken_paragraph_mapping(self):
        out=make_outline(self.doc,6)
        scenes=build_scenes(out,DEFAULT_STYLE)
        data=export_pptx(scenes)
        check=audit_native(scenes,data)
        self.assertTrue(check['passed'],check['errors'])
        ppt=Presentation(io.BytesIO(data))
        body=next(e for s in scenes for e in s['elements'] if e.get('role')=='body')
        shape=next(sh for sl in ppt.slides for sh in sl.shapes if sh.name==body['id'])
        shape.text='故意损坏正文'
        broken=io.BytesIO()
        ppt.save(broken)
        self.assertFalse(audit_native(scenes,broken.getvalue())['passed'])

    def test_bundle_html_evidence_and_svg_opt_in(self):
        import zipfile
        out=make_outline(self.doc,6)
        out['title']='<script>bad()</script>'
        scenes=build_scenes(out,DEFAULT_STYLE)
        pptx=export_pptx(scenes)
        check=audit_native(scenes,pptx)
        markup=deck_html(scenes,out['title']).decode()
        self.assertNotIn('<script>bad()',markup)
        self.assertEqual(markup.count('<section class="slide"'),len(scenes))
        del self.doc.metadata  # Simulate a session opened before the V2 metadata field existed.
        for include in (False,True):
            bundle=export_bundle(scenes,out,DEFAULT_STYLE,self.doc,pptx,check,include)
            with zipfile.ZipFile(io.BytesIO(bundle)) as z:
                self.assertIn('evidence.json',z.namelist())
                self.assertEqual(any(n.startswith('svg/') for n in z.namelist()),include)
                evidence=json.loads(z.read('evidence.json'))
                self.assertEqual(evidence['blocks'],self.doc.blocks)

    def test_extreme_cover_and_titles(self):
        out = make_outline(self.doc, 6, title="中"*120, subtitle="文"*150)
        out["slides"][0]["title"] = "文"*100
        scenes = build_scenes(out, DEFAULT_STYLE)
        self.assertGreater(len(scenes), 1)

    def test_svg_escapes_user_markup(self):
        out = make_outline(self.doc, 6, title='<script>alert("x")</script>')
        svg = scene_svg(build_scenes(out, DEFAULT_STYLE)[0]).decode()
        self.assertNotIn("<script>", svg)
        self.assertIn("&lt;script&gt;", svg)

    def test_validation_rejects_bad_sources_and_long_bullets(self):
        out = make_outline(self.doc, 6)
        out["slides"][0]["source_ids"] = ["S9999"]
        with self.assertRaises(UserError):
            validate_outline(out, {"S1", "S2"})
        out["slides"][0]["source_ids"] = ["S1"]
        out["slides"][0]["bullets"] = ["x"*501]
        with self.assertRaises(UserError):
            validate_outline(out, {"S1", "S2"})

    def test_frontend_themes_canvas_and_native_geometry(self):
        out=make_outline(self.doc,6,subtitle='课程汇报')
        for name in THEMES:
            for density in ('reading','speaking'):
                scenes,theme=frontend_scenes(out,name,density)
                self.assertTrue(all((s['width'],s['height'])==(1920,1080) for s in scenes))
                ppt=Presentation(io.BytesIO(export_pptx(scenes)))
                self.assertAlmostEqual(ppt.slide_width/ppt.slide_height,16/9,places=5)
                self.assertTrue(audit_native(scenes,export_pptx(scenes))['passed'])
                for slide in ppt.slides:
                    for shape in slide.shapes:
                        self.assertLessEqual(shape.left+shape.width,ppt.slide_width+100)
                        self.assertLessEqual(shape.top+shape.height,ppt.slide_height+100)

    def test_frontend_density_preserves_every_body_character(self):
        out=make_outline(self.doc,6)
        out['slides'][0]['bullets']=['原始条件 123.45 mm，Independent validation; '+'校准流程。'*35]*4
        for density in ('reading','speaking'):
            result=density_outline(out,density)
            original=''.join(b for s in out['slides'] for b in s['bullets'])
            actual=''.join(b for s in result['slides'] for b in s['bullets'])
            self.assertEqual(actual,original)
            scenes,_=frontend_scenes(out,'electric-studio',density)
            native=''.join(e.get('original_text','') for s in scenes for e in s['elements'] if e.get('role')=='body')
            self.assertEqual(native,original)

    def test_frontend_edits_roundtrip_and_wrong_deck_rejection(self):
        out=make_outline(self.doc,6)
        out['slides'][0]['bullets']=['保留条件 12.5 mm。'+'独立校准与验证。'*18]
        scenes,_=frontend_scenes(out,'signal','speaking')
        bindings=edit_bindings(scenes)
        title=next(k for k,v in bindings.items() if v['role']=='title' and v['slide']==-1)
        body=next(k for k,v in bindings.items() if v['role']=='body' and v['slide']==0)
        edited={'schema':'slidecraft-frontend-edits/v1','deck_id':edit_signature(scenes),
                'edits':{title:'修改后的实验汇报',body:'新的校准内容。'}}
        result=apply_frontend_edits(out,scenes,edited)
        self.assertEqual(result['title'],'修改后的实验汇报')
        self.assertEqual(result['slides'][0]['bullets'][0],
            '新的校准内容。'+out['slides'][0]['bullets'][0][len(bindings[body]['text']):])
        self.assertEqual(result['slides'][0]['source_ids'],out['slides'][0]['source_ids'])
        split=apply_frontend_edits(out,scenes,{**edited,'edits':{body:'新的第一条。\n新的第二条。'}})
        self.assertEqual(split['slides'][0]['bullets'][:2],['新的第一条。','新的第二条。'+out['slides'][0]['bullets'][0][len(bindings[body]['text']):]])
        unchanged=apply_frontend_edits(out,scenes,{**edited,'edits':{}})
        self.assertEqual(unchanged,out)
        with self.assertRaises(UserError):
            apply_frontend_edits(out,scenes,{**edited,'deck_id':'wrong-deck'})
        with self.assertRaises(UserError):
            apply_frontend_edits(out,scenes,{**edited,'edits':{'unknown':'new text'}})

    def test_frontend_offline_css_license_and_escaping(self):
        out=make_outline(self.doc,6,title='__CONFIG__ </script><script>window.pwned=1</script>')
        scenes,theme=frontend_scenes(out,'electric-studio')
        cache={}
        css=font_css(scenes,cache)
        result=frontend_html(scenes,theme,css).decode()
        self.assertIn((ASSETS/'viewport-base.css').read_text('utf-8'),result)
        self.assertIn('data:font/woff2;base64,',result)
        self.assertNotIn('<script>window.pwned=1</script>',result)
        self.assertIn('Copyright (c) 2025 Zara Zhang',result)
        self.assertNotIn('src="https://',result)
        self.assertEqual(font_css(scenes,cache),css)

    def test_ai_mock_only_no_network(self):
        out = make_outline(self.doc, 6)
        payload = json.dumps({"choices":[{"message":{"content":json.dumps(out)}}]}).encode()
        class Response:
            def __enter__(self): return self
            def __exit__(self, *args): pass
            def read(self, *args): return payload
        with patch("urllib.request.urlopen", return_value=Response()) as request:
            result = ai_outline(self.doc, 6, "测试", "", "fake-key", "configured-model")
            self.assertEqual(result["mode"], "AI 整理")
            self.assertEqual(request.call_count, 1)
        with patch("urllib.request.urlopen", side_effect=urllib.error.HTTPError("https://example.com", 401, "bad", {}, None)):
            with self.assertRaisesRegex(UserError, "API Key"):
                ai_outline(self.doc, 6, "", "", "fake-key", "model")

    def test_deepseek_choice_shows_url_and_model_without_sending(self):
        from streamlit.testing.v1 import AppTest
        at = AppTest.from_file(str(ROOT / "app.py")).run()
        next(r for r in at.radio if r.label == "内容整理方式").set_value("AI 整理 · 需要 API").run()
        next(s for s in at.selectbox if s.label == "模型服务商").select("DeepSeek").run()
        self.assertFalse(at.exception)
        self.assertEqual(next(t.value for t in at.text_input if t.label == "模型名称"), "deepseek-flash")
        self.assertTrue(any("https://api.deepseek.com" in c.value for c in at.caption))
        self.assertFalse(next(c.value for c in at.checkbox if c.label.startswith("同意将文档文字")))

    def test_streamlit_demo_edit_export_invalidation(self):
        from streamlit.testing.v1 import AppTest
        at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=40).run()
        self.assertFalse(at.exception, str(at.exception))
        next(b for b in at.button if b.label == "载入腕关节传感器示例").click().run()
        self.assertFalse(at.exception, str(at.exception))
        self.assertFalse(at.error, [e.value for e in at.error])
        next(b for b in at.button if b.label == "生成 PPTX 下载文件").click().run()
        self.assertFalse(at.exception, str(at.exception))
        self.assertTrue(at.session_state["export"].startswith(b"PK"))
        self.assertTrue(at.session_state['export_check']['passed'])
        old_sig = at.session_state["export_sig"]
        next(t for t in at.text_input if t.label == "封面标题").set_value("修改后的课堂汇报").run()
        self.assertFalse(at.exception, str(at.exception))
        self.assertFalse(at.error, [e.value for e in at.error])
        next(b for b in at.button if b.label == "生成 PPTX 下载文件").click().run()
        self.assertNotEqual(at.session_state["export_sig"], old_sig)
        ppt = Presentation(io.BytesIO(at.session_state["export"]))
        self.assertTrue(any("修改后的课堂汇报" in s.text for s in ppt.slides[0].shapes if s.has_text_frame))
        selectors=[s for s in at.selectbox if s.label=='版式']
        selectors[0].select('comparison').run()
        self.assertFalse(at.exception,str(at.exception))
        next(b for b in at.button if b.label=='为全部页面重新推荐版式（保留文字）').click().run()
        self.assertTrue(all(s.value=='auto' for s in at.selectbox if s.label=='版式'))
        self.assertEqual(next(t for t in at.text_input if t.label=='封面标题').value,'修改后的课堂汇报')
        next(r for r in at.radio if r.label=='选择动态演示风格').set_value('signal').run()
        self.assertFalse(at.exception,str(at.exception))
        self.assertFalse(at.error,[e.value for e in at.error])
        next(b for b in at.button if b.label=='生成对应的可编辑 PPTX 和文稿包').click().run()
        self.assertFalse(at.exception,str(at.exception))
        self.assertFalse(at.error,[e.value for e in at.error])
        self.assertTrue(at.session_state['frontend_output']['pptx'].startswith(b'PK'))


class PersonalThemeTests(unittest.TestCase):
    def setUp(self):
        base = Presentation()
        for layout in list(base.slide_layouts)[:2]:
            slide = base.slides.add_slide(layout)
            slide.shapes.add_textbox(0, 0, Inches(2), Inches(1)).text = "REMOVE SAMPLE TEXT"
        mark = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(.15), Inches(.15))
        mark.name = "Native master mark"
        slide.slide_layout.shapes._spTree.insert_element_before(copy.deepcopy(mark._element), "p:extLst")
        out = io.BytesIO()
        base.save(out)
        self.template_bytes = out.getvalue()
        self.template = load_personal_template(self.template_bytes)
        self.assets = {}
        for name, size in (("wide", (1100, 270)), ("A", (780, 440)), ("B", (700, 440))):
            image = io.BytesIO()
            Image.new("RGB", size, "#B8CDE2").save(image, "PNG")
            self.assets[name] = normalized_image(image.getvalue(), name)

    def test_excerpts_keep_results_units_limits_and_source(self):
        lines = ["目录", "说明文件制作于上周", "采用三个传感器测量", "结果显示误差下降 15%",
                 "结果显示误差下降 15%", "仅在室内条件下验证", "参考文献"]
        units = [({"id": f"S{i}"}, value) for i, value in enumerate(lines)]
        picks = rank_excerpts(units, 3)
        self.assertEqual([source for source, _ in picks], ["S2", "S3", "S5"])
        self.assertIn("15%", picks[1][1])

    def test_text_diagrams_render_as_native_shapes(self):
        pages = [
            {"title": "记录", "bullets": ["先固定佩戴位置，再记录动作，最后核查信号。"]},
            {"title": "步骤流程", "bullets": ["设备调试", "整理资料", "阅读文献"]},
            {"title": "两种方案", "bullets": ["A 使用电极", "B 使用光学标记"], "relationship": "comparison"},
        ]
        for page in pages:
            page.update(diagram=True, source_ids=[])
        self.assertEqual([diagram_for(p) for p in pages], ["flow", "tasks", "comparison"])
        scenes = build_personal_scenes({"title": "测试", "include_cover": False, "slides": pages},
                                       PERSONAL_STYLE, self.template, {}, {})
        self.assertEqual([sum(e["kind"] == "chevron" for e in s["elements"]) for s in scenes], [2, 0, 0])
        pptx = export_pptx(scenes, template=self.template)
        self.assertTrue(audit_native(scenes, pptx)["passed"])
        doc = Presentation(io.BytesIO(pptx))
        self.assertEqual(len(doc.slides), 3)
        self.assertTrue(any(s.has_text_frame and "核查信号" in s.text for s in doc.slides[0].shapes))

    def test_measured_bars_keep_units_and_fallback_on_incomparable_values(self):
        page = {"title": "同条件比较", "bullets": ["对照组：平均误差 2 mm", "实验组：平均误差 1.25 mm"],
                "source_ids": [], "relationship": "comparison", "diagram": True}
        self.assertEqual(diagram_for(page), "bars")
        self.assertEqual(diagram_for({**page, "bullets": [page["bullets"][0], "实验组：平均延迟 1.25 ms"]}), "comparison")
        scenes = build_personal_scenes({"title": "测试", "include_cover": False, "slides": [page]},
                                       PERSONAL_STYLE, self.template, {}, {})
        scene = scenes[0]
        self.assertIn("1.25 mm", [e.get("text") for e in scene["elements"]])
        self.assertIn(page["bullets"][0], scene["notes"])
        bars = [e["w"] for e in scene["elements"] if e.get("fill") in {PERSONAL_STYLE["primary"], PERSONAL_STYLE["accent"]}]
        self.assertAlmostEqual(bars[1]/bars[0], 1.25/2)

    def test_concept_images_use_one_fixed_endpoint_and_stay_marked(self):
        art = self.assets["wide"]
        response = {"data": [{"b64_json": base64.b64encode(art["data"]).decode()}]}
        with patch("urllib.request.urlopen", return_value=io.BytesIO(json.dumps(response).encode())) as request:
            asset = generate_concept_image("抽象的传感器安装示意", "test-key")
            req = request.call_args.args[0]
            self.assertEqual(req.full_url, IMAGE_ENDPOINT)
            self.assertIn("抽象的传感器安装示意", json.loads(req.data)["prompt"])
        asset.update(generated=True, source="AI 生成 · 概念示意")
        page = {"title": "测量概念", "bullets": ["示意设备位置"], "source_ids": [],
                "figures": [{"asset": "art", "caption": "设备位置"}]}
        scenes = build_personal_scenes({"title": "测试", "include_cover": False, "slides": [page]},
                                       PERSONAL_STYLE, self.template, {"art": asset}, {})
        self.assertTrue(any("AI 生成" in e.get("text", "") for e in scenes[0]["elements"]))
        pptx = Presentation(io.BytesIO(export_pptx(scenes, template=self.template)))
        self.assertEqual(sum(s.shape_type == 13 for s in pptx.slides[0].shapes), 1)

    def test_concept_image_consent_and_attach_in_streamlit(self):
        from streamlit.testing.v1 import AppTest
        with TemporaryDirectory() as folder:
            template_path = Path(folder) / "template.pptx"
            template_path.write_bytes(self.template_bytes)
            with patch.dict(os.environ, {"PERSONAL_TEMPLATE_PATH": str(template_path)}):
                at = AppTest.from_file(str(ROOT / "app.py")).run()
                next(b for b in at.button if b.label == "从空白页开始").click().run()
                next(t for t in at.text_input if t.label == "页标题（可留空）").set_value("传感器安装")
                next(t for t in at.text_area if t.label == "正文（每行一条）").set_value("展示安装概念。")
                next(c for c in at.checkbox if c.label == "按本页文字生成概念插图").check().run()
                self.assertTrue(next(b for b in at.button if b.label == "生成并添加到本页").disabled)
                next(t for t in at.text_input if t.label == "OpenAI 图片 API Key").set_value("fake-key")
                next(c for c in at.checkbox if c.label.startswith("同意将上面的插图描述")).check().run()
                response = {"data": [{"b64_json": base64.b64encode(self.assets["wide"]["data"]).decode()}]}
                with patch("urllib.request.urlopen", return_value=io.BytesIO(json.dumps(response).encode())):
                    next(b for b in at.button if b.label == "生成并添加到本页").click().run()
                self.assertFalse(at.exception, str(at.exception))
                self.assertFalse(at.error, [err.value for err in at.error])
                self.assertEqual(next(m.value for m in at.multiselect if m.label == "本页图片"), ["AI 概念示意 1"])

    def test_personal_geometry_and_native_template(self):
        pages = [
            {"title": "任意标题", "bullets": ["用 **粗体** 和 ==重点== 说明"], "source_ids": [],
             "figures": [{"asset": "wide", "caption": "说明一"}]},
            {"title": "任意标题", "bullets": ["比较分组条件"], "source_ids": [], "relationship": "comparison",
             "figures": [{"asset": "A", "caption": "图A"}, {"asset": "B", "caption": "图B"}]},
            {"title": "", "bullets": ["对应甲图", "对应乙图"], "source_ids": [],
             "figures": [{"asset": "A", "caption": "甲", "bullet": 0}, {"asset": "B", "caption": "乙", "bullet": 1}]},
            {"title": "这些是校准过程", "bullets": ["测量和记录属于两类信息", "保留原图"], "source_ids": []},
        ]
        out = {"title": "学术汇报", "include_cover": False, "slides": pages}
        validate_outline(out, set(), personal=True)
        scenes = build_personal_scenes(out, PERSONAL_STYLE, self.template, self.assets, {})
        self.assertEqual([s["layout"] for s in scenes], ["main", "comparison", "paired", "bullets"])
        compare = [e for e in scenes[1]["elements"] if e["kind"] == "image"]
        self.assertLess(compare[0]["x"] + compare[0]["w"], compare[1]["x"])
        self.assertIn(b"data:image/svg+xml;base64,", scene_svg(scenes[0]))
        self.assertIn(b"#C00000", scene_svg(scenes[0]))
        result = export_pptx(scenes, template=self.template)
        self.assertTrue(audit_native(scenes, result)["passed"])
        prs = Presentation(io.BytesIO(result))
        self.assertEqual((prs.slide_width, prs.slide_height),
                         (Presentation(io.BytesIO(self.template_bytes)).slide_width,
                          Presentation(io.BytesIO(self.template_bytes)).slide_height))
        self.assertEqual(len(prs.slides), len(pages))
        self.assertEqual(sum(s.shape_type == 13 for s in prs.slides[2].shapes), 2)
        for slide in prs.slides:
            self.assertIn("Native master mark", [s.name for s in slide.slide_layout.shapes])
            self.assertFalse(any("REMOVE SAMPLE TEXT" in s.text for s in slide.shapes if s.has_text_frame))
        runs = [r for s in prs.slides[0].shapes if s.has_text_frame for p in s.text_frame.paragraphs for r in p.runs]
        self.assertTrue(any(r.text == "重点" and r.font.bold and str(r.font.color.rgb) == "C00000" for r in runs))

    def test_personal_image_sizes_and_long_paired_page(self):
        def scenes_for(page):
            return build_personal_scenes({"title": "测试", "include_cover": False, "slides": [page]},
                                         PERSONAL_STYLE, self.template, self.assets, {})
        base = {"title": "内容", "bullets": ["图是示意，文字可换行。"], "source_ids": [],
                "personal_layout": "side", "figures": [{"asset": "A", "caption": "图 A"}]}
        figure_width = lambda scenes: next(e["w"] for e in scenes[0]["elements"] if e["kind"] == "image")
        wide = {**base, "figures": [{"asset": "wide", "caption": "宽图"}]}
        self.assertGreater(figure_width(scenes_for({**wide, "space": "figure"})),
                           figure_width(scenes_for({**wide, "space": "text"})))
        long = "请保留这段文字与对应的图片。" * 28
        pair = {**base, "personal_layout": "paired", "bullets": [long, long],
                "figures": [{"asset": "A", "caption": "第一个要点", "bullet": 0},
                            {"asset": "B", "caption": "第二个要点", "bullet": 1}]}
        scenes = scenes_for(pair)
        for index in range(2):
            rendered = "".join(e["text"].replace("\n", "") for s in scenes for e in s["elements"]
                               if e.get("role") == "body" and e["bullet_id"] == index)
            self.assertEqual(rendered, long)
        for scene in scenes:
            self.assertTrue(any(e["kind"] == "image" for e in scene["elements"]))

    def test_personal_streamlit_manual_pages_preserve_edits(self):
        from streamlit.testing.v1 import AppTest
        with TemporaryDirectory() as folder:
            template_path = Path(folder) / "template.pptx"
            template_path.write_bytes(self.template_bytes)
            with patch.dict(os.environ, {"PERSONAL_TEMPLATE_PATH": str(template_path)}):
                at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=25).run()
                next(b for b in at.button if b.label == "从空白页开始").click().run()
                self.assertFalse(at.exception, str(at.exception))
                at.session_state["personal_assets"] = self.assets
                next(c for c in at.checkbox if c.label == "生成封面").uncheck()
                next(t for t in at.text_input if t.label == "页标题（可留空）").set_value("我自拟的标题")
                next(t for t in at.text_area if t.label == "正文（每行一条）").set_value("两张图是原始示意。")
                at.run()
                next(m for m in at.multiselect if m.label == "本页图片").set_value(["A", "B"]).run()
                [t for t in at.text_input if t.label == "图注 / 出处"][1].set_value("图片乙的图注")
                next(s for s in at.selectbox if s.label == "第一张图").select("B")
                next(s for s in at.selectbox if s.label == "版式").select("comparison").run()
                next(b for b in at.button if b.label == "添加一页").click().run()
                [t for t in at.text_input if t.label == "页标题（可留空）"][1].set_value("标题独立成页").run()
                [b for b in at.button if b.label == "上移"][1].click().run()
                self.assertEqual([t.value for t in at.text_input if t.label == "页标题（可留空）"],
                                 ["标题独立成页", "我自拟的标题"])
                [b for b in at.button if b.label == "本页重新推荐"][1].click().run()
                self.assertFalse(at.exception, str(at.exception))
                self.assertFalse(at.error, [e.value for e in at.error])
                next(r for r in at.radio if r.label == "工作区").set_value("版式预览与下载").run()
                next(b for b in at.button if b.label == "生成 PPTX 下载文件").click().run()
                self.assertFalse(at.error, [e.value for e in at.error])
                self.assertEqual(next(r for r in at.radio if r.label == "工作区").value, "版式预览与下载")
                self.assertTrue(at.session_state["export_check"]["passed"])
                prs = Presentation(io.BytesIO(at.session_state["export"]))
                self.assertEqual(len(prs.slides), 2)
                self.assertTrue(any(s.has_text_frame and s.text == "图片乙的图注" for s in prs.slides[1].shapes))
                self.assertEqual(sum(s.shape_type == 13 for s in prs.slides[1].shapes), 2)


if __name__ == "__main__":
    unittest.main(verbosity=2)

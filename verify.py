"""Offline checks: python verify.py. No API key or paid requests."""
import copy
import io
import json
import re
import sys
import unittest
import urllib.error
from pathlib import Path
from unittest.mock import patch

from docx import Document
from PIL import Image
from pptx import Presentation
from pypdf import PdfWriter
from pypdf.generic import DecodedStreamObject, DictionaryObject, NameObject, NumberObject

from core import (DEFAULT_STYLE, SourceDoc, UserError, ai_outline, extract_style,
                  make_outline, normalized_image, read_document, validate_outline)
from renderer import build_scenes, export_pptx, scene_svg
from layouts import LAYOUT_LABELS, recommend_layout
from exports import audit_native, deck_html, export_bundle

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


if __name__ == "__main__":
    unittest.main(verbosity=2)
